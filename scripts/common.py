# -*- coding: utf-8 -*-
"""PowerPoint 原生动画构建框架（COM / pywin32）。画布 960x540pt。

已探明语义（probe_com.py / probe_path.py）：
- MotionEffect.Path 坐标 = 相对形状"原始中心"的位移，x 分数 × 幻灯片宽、y × 高；
  每条路径效果都从原始位置出发（绝对式），多段接续时后续路径的 M 写在上一条终点偏移上。
- 结束位置在放映中持久保留；Accelerate=Decelerate=0 时严格匀速（按弧长重采样可保证匀速）。
- Wipe 方向经 effect.EffectParameters.Direction（先设方向、后设时长）。
- 自定义路径：AddEffect(shape, msoAnimEffectPathRight=149) 后覆写 Behaviors(1).MotionEffect.Path。
"""
import os
import time

import win32com.client

app = None


def boot():
    global app
    if app is None:
        app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        from win32com.client import constants as _c  # noqa: F401  (确保 gen_py 就绪)
        app.DisplayAlerts = 1  # ppAlertsNone
        app.Visible = -1
    return app


def quit_app():
    global app
    if app is not None:
        try:
            app.Quit()
        except Exception:
            pass
        app = None


# ---------------- 常量 ----------------
MTRUE, MFALSE = -1, 0
AE_APPEAR, AE_FLY, AE_FADE, AE_WIPE, AE_ZOOM = 1, 2, 10, 22, 23
AE_FADEDZOOM, AE_GROWSHRINK, AE_SPIN = 48, 59, 61
AE_PATHRIGHT = 149                      # 任意路径类 preset，Path 会被整体覆写
TRIG_CLICK, TRIG_WITH, TRIG_AFTER = 1, 2, 3
DIR_L, DIR_R, DIR_U, DIR_D = 4, 2, 1, 3
SHAPE_RECT, SHAPE_RRECT, SHAPE_OVAL = 1, 5, 9
SEG_LINE, SEG_CURVE = 0, 1
EDIT_CORNER = 0
ANCHOR_TOP, ANCHOR_MID = 1, 3
AUTOSIZE_NONE = 0
ALIGN_L, ALIGN_C, ALIGN_R = 1, 2, 3
ORIENT_H = 1
LINE_SOLID, LINE_DOT, LINE_DASH = 1, 3, 4

# 配色（与 GIF 版一致）
BG = "#F7F9FC"
INK = "#1F3864"      # NAVY
TEAL = "#2E86AB"
TEAL_BG = "#EAF3F8"
ORANGE = "#E8833A"
ORANGE_BG = "#FDF3E7"
GREEN = "#3F9D5C"
GREEN_BG = "#EBF4EE"
RED = "#C0392B"
RED_BG = "#FBEDEB"
GRAY = "#6B7280"
LIGHT = "#E4E9F0"
WHITE = "#FFFFFF"
GRID = "#E2E8F1"
AXIS = "#B9C2D0"

FONT_CN = "微软雅黑"
FONT_EN = "Microsoft YaHei"


def rgb(h):
    return int(h[1:3], 16) + (int(h[3:5], 16) << 8) + (int(h[5:7], 16) << 16)


def fr(v):
    """路径坐标格式化（去掉尾零，防科学计数）"""
    s = f"{v:.5f}".rstrip("0").rstrip(".")
    return s if s not in ("", "-") else "0"


def lerp(a, b, u):
    return a + (b - a) * u


def resample(pts, n):
    """按弧长把折线重采样为 n 段等长（保证路径匀速）。pts=[(x,y),...]"""
    segs = []
    total = 0.0
    for i in range(len(pts) - 1):
        d = ((pts[i + 1][0] - pts[i][0]) ** 2 + (pts[i + 1][1] - pts[i][1]) ** 2) ** 0.5
        segs.append(d)
        total += d
    if total <= 1e-9:
        return [pts[0], pts[-1]]
    out = [pts[0]]
    target = total / n
    acc = 0.0
    j = 0
    need = target
    while j < len(segs) and len(out) < n:
        if segs[j] >= need - 1e-9 and segs[j] > 0:
            u = need / segs[j]
            out.append((lerp(pts[j][0], pts[j + 1][0], u), lerp(pts[j][1], pts[j + 1][1], u)))
            segs[j] = segs[j] - need
            pts[j] = out[-1]
            need = target
        else:
            need -= segs[j]
            j += 1
    out.append(pts[-1])
    return out


def sig(x0, x1, mid, k):
    """生成 sigmoid 曲线采样点（绘图用）"""
    import math
    pts = []
    for i in range(101):
        t = x0 + (x1 - x0) * i / 100.0
        pts.append((t, 1.0 / (1 + math.exp(-(t - mid) / k))))
    return pts


# ---------------- 幻灯片上下文 ----------------
class Slide:
    def __init__(self, pres, mode="live"):
        self.mode = mode
        self.pres = pres
        self.s = pres.Slides.Add(pres.Slides.Count + 1, 12)
        self.s.FollowMasterBackground = MFALSE
        self.s.Background.Fill.ForeColor.RGB = rgb(BG)
        self.MS = self.s.TimeLine.MainSequence
        self._n = 0
        self.mode = mode
        self.prev_start = None
        self.prev_end = None
        self.t = 0.0            # 便捷游标（可选使用）

    # ---------- 形状 ----------
    def _name(self, prefix):
        self._n += 1
        return f"{prefix}{self._n}"

    def _set_text(self, shp, txt, size, color, bold, align, anchor, wrap, spacing=None):
        tf = shp.TextFrame2
        tf.VerticalAnchor = anchor
        tf.AutoSize = AUTOSIZE_NONE
        tf.WordWrap = MTRUE if wrap else MFALSE
        tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
        tr = tf.TextRange
        tr.Text = txt
        f = tr.Font
        f.NameFarEast = FONT_CN
        f.Name = FONT_EN
        f.Size = size
        f.Bold = MTRUE if bold else MFALSE
        f.Fill.ForeColor.RGB = rgb(color)
        tr.ParagraphFormat.Alignment = align
        if spacing:
            try:
                tr.Font.Kerning = 0
                tr.ParagraphFormat.SpaceWithin = spacing
            except Exception:
                pass
        return shp

    def text(self, x, y, w, h, txt, size=10, color=INK, bold=False,
             align="l", anchor="m", wrap=False):
        amap = {"l": ALIGN_L, "c": ALIGN_C, "r": ALIGN_R}
        amap2 = {"t": ANCHOR_TOP, "m": ANCHOR_MID}
        shp = self.s.Shapes.AddTextbox(ORIENT_H, x, y, w, h)
        shp.Name = self._name("t")
        return self._set_text(shp, txt, size, color, bold, amap[align], amap2[anchor], wrap)

    def rect(self, x, y, w, h, fill=INK, line=None, lw=1.0):
        shp = self.s.Shapes.AddShape(SHAPE_RECT, x, y, w, h)
        shp.Name = self._name("r")
        shp.Fill.ForeColor.RGB = rgb(fill if fill else WHITE)
        if fill is None:
            shp.Fill.Visible = MFALSE
        shp.Fill.Transparency = 0.0
        if line:
            shp.Line.ForeColor.RGB = rgb(line)
            shp.Line.Weight = lw
        else:
            shp.Line.Visible = MFALSE
        shp.Shadow.Visible = MFALSE
        return shp

    def card(self, x, y, w, h, r=None, fill=WHITE, line=None, lw=1.0, shadow=True,
             text=None, size=10, color=INK, bold=False, talign="c", alpha=None):
        shp = self.s.Shapes.AddShape(SHAPE_RRECT, x, y, w, h)
        shp.Name = self._name("c")
        rad = (h / 2.0 if r == "pill" else (r if r is not None else min(w, h) * 0.12))
        try:
            shp.Adjustments.SetItem(1, min(0.5, rad / min(w, h)))
        except Exception:
            pass
        shp.Fill.ForeColor.RGB = rgb(fill if fill else WHITE)
        if fill is None:
            shp.Fill.Visible = MFALSE
        if alpha is not None:
            shp.Fill.Transparency = 1.0 - alpha
        if line:
            shp.Line.ForeColor.RGB = rgb(line)
            shp.Line.Weight = lw
        else:
            shp.Line.Visible = MFALSE
        if shadow:
            shp.Shadow.Visible = MTRUE
            shp.Shadow.Blur = 9
            shp.Shadow.OffsetX = 0
            shp.Shadow.OffsetY = 2.5
            shp.Shadow.Transparency = 0.9
            shp.Shadow.ForeColor.RGB = rgb(INK)
        else:
            shp.Shadow.Visible = MFALSE
        if text is not None:
            self._set_text(shp, text, size, color, bold,
                           {"l": ALIGN_L, "c": ALIGN_C, "r": ALIGN_R}[talign], ANCHOR_MID, True)
        return shp

    def dot(self, cx, cy, d, fill=INK, line=None, lw=1.5):
        shp = self.s.Shapes.AddShape(SHAPE_OVAL, cx - d / 2, cy - d / 2, d, d)
        shp.Name = self._name("d")
        shp.Fill.ForeColor.RGB = rgb(fill)
        if line:
            shp.Line.ForeColor.RGB = rgb(line)
            shp.Line.Weight = lw
        else:
            shp.Line.Visible = MFALSE
        shp.Shadow.Visible = MFALSE
        return shp

    def ring(self, cx, cy, d, color=ORANGE, lw=3.0):
        shp = self.s.Shapes.AddShape(SHAPE_OVAL, cx - d / 2, cy - d / 2, d, d)
        shp.Name = self._name("g")
        shp.Fill.Visible = MFALSE
        shp.Line.ForeColor.RGB = rgb(color)
        shp.Line.Weight = lw
        shp.Shadow.Visible = MFALSE
        return shp

    def line(self, x1, y1, x2, y2, color=INK, w=2.0, dash=None,
             begin_arrow=False, end_arrow=False):
        shp = self.s.Shapes.AddLine(x1, y1, x2, y2)
        shp.Name = self._name("l")
        ln = shp.Line
        ln.ForeColor.RGB = rgb(color)
        ln.Weight = w
        if dash:
            ln.DashStyle = dash
        if begin_arrow:
            ln.BeginArrowheadStyle = 2
        if end_arrow:
            ln.EndArrowheadStyle = 2
        shp.Shadow.Visible = MFALSE
        return shp

    def freeform(self, pts, line_color=None, line_w=2.0, fill=None,
                 fill_alpha=None, close=False):
        """pts=[(x,y)...]pt；fill 给定时闭合填充（面积图）。"""
        b = self.s.Shapes.BuildFreeform(EDIT_CORNER, pts[0][0], pts[0][1])
        use = pts[1:] + ([pts[0]] if close and pts[0] != pts[-1] else [])
        for (x, y) in use:
            b.AddNodes(SEG_LINE, EDIT_CORNER, x, y)
        shp = b.ConvertToShape()
        shp.Name = self._name("f")
        if fill:
            shp.Fill.ForeColor.RGB = rgb(fill)
            shp.Fill.Visible = MTRUE
            if fill_alpha is not None:
                shp.Fill.Transparency = 1.0 - fill_alpha
        else:
            shp.Fill.Visible = MFALSE
        if line_color:
            shp.Line.ForeColor.RGB = rgb(line_color)
            shp.Line.Weight = line_w
        else:
            shp.Line.Visible = MFALSE
        shp.Shadow.Visible = MFALSE
        return shp

    def picture(self, path, x, y, max_w, max_h):
        """插入图片，等比缩放居中放入 (x,y,max_w,max_h) 盒内（保持 1:1 比例不拉伸）。"""
        from PIL import Image
        iw, ih = Image.open(path).size
        scale = min(max_w / iw, max_h / ih)
        w, h = iw * scale, ih * scale
        shp = self.s.Shapes.AddPicture(path, MFALSE, MTRUE,
                                       x + (max_w - w) / 2, y + (max_h - h) / 2, w, h)
        shp.Name = self._name("pic")
        return shp

    def group(self, shapes, name=None):
        """按名称收组（动画作用于组）。"""
        if len(shapes) == 1:
            return shapes[0]
        for shp in shapes:
            shp.Name = "G" + shp.Name
        rng = self.s.Shapes.Range(tuple(shp.Name for shp in shapes))
        g = rng.Group()
        g.Name = name or self._name("grp")
        return g

    def chip(self, cx, cy, txt, w=None, h=26, fill=TEAL_BG, line=TEAL,
             color=INK, size=9.5, bold=False):
        """圆角 pill + 文本 → 组。cx,cy 为中心。"""
        w = w or (len(txt) * size * 1.05 + 22)
        c1 = self.card(cx - w / 2, cy - h / 2, w, h, r="pill", fill=fill, line=line, lw=1.5)
        t1 = self.text(cx - w / 2, cy - h / 2, w, h, txt, size=size, color=color,
                       bold=bold, align="c")
        return self.group([c1, t1])

    def pill(self, cx, cy, txt, bg=TEAL, fg=WHITE, size=9.5, h=24, w=None, bold=False):
        w = w or (len(txt) * size * 1.12 + 24)
        c1 = self.card(cx - w / 2, cy - h / 2, w, h, r="pill", fill=bg, shadow=False)
        t1 = self.text(cx - w / 2, cy - h / 2, w, h, txt, size=size, color=fg,
                       bold=True or bold, align="c")
        return self.group([c1, t1])

    # ---------- 动画 ----------
    def _sched(self, T, dur):
        """返回 (trigger, delay)。

        对象模型语义（已由交付文件 XML + 实机 F5 双重验证，勿回退）：
        主序列只有单个自动开始的 click 组（onBegin mainSeq），组内每个效果
        par 的 <p:cond delay> 是**相对组起点（=整条时间线起点）的绝对时刻**，
        WithPrevious / AfterPrevious 皆然。COM 写入什么值就按什么值播放。
        因此 delay 一律写绝对时刻 T；此前"实机按增量解释、live 写 delta"的
        双引擎说法是误判，曾导致整页时间线被压缩到前 2~3 秒（错位/不显示）。
        mode 参数保留仅为兼容，两种模式现在生成完全相同的时间线。
        效果须按开始时间升序添加。
        """
        if self.prev_start is None:
            return TRIG_AFTER, round(max(T, 0.0), 3)
        if T >= self.prev_start - 1e-9:
            return TRIG_WITH, round(T, 3)
        raise RuntimeError(f"效果开始时间逆序：T={T} < prev_start={self.prev_start}（需升序添加）")

    def _finish(self, T, dur):
        self.prev_start = T
        self.prev_end = T + dur

    def fx(self, shape, eff, T=None, dur=0.5, direction=None, exit_=False,
           repeat=0, accel=None, decel=None, smooth=None):
        """通用效果。T=绝对开始时刻；返回 effect。

        缓动约定：smooth=True 时先写 SmoothStart/SmoothEnd=msoTrue，再用
        Accelerate/Decelerate 覆写精确曲线（布尔在前、数值在后，数值生效）。
        位移类效果（wipe/fly/pop/path）默认 smooth=True + (0.2, 0.5)。"""
        trig, delay = self._sched(T, dur)
        e = self.MS.AddEffect(shape, eff, 0, trig)
        if direction is not None:
            try:
                e.EffectParameters.Direction = direction
            except Exception:
                pass
        if exit_:
            e.Exit = MTRUE
        tm = e.Timing
        try:
            tm.TriggerDelayTime = max(delay, 0.0)
        except Exception:
            pass
        tm.Duration = dur
        if repeat:
            tm.RepeatCount = repeat
        if smooth is None:
            smooth = eff in (AE_WIPE, AE_FLY, AE_FADEDZOOM, AE_ZOOM, AE_PATHRIGHT)
        if smooth:
            try:
                tm.SmoothStart = MTRUE
                tm.SmoothEnd = MTRUE
            except Exception:
                pass
        tm.Accelerate = 0.2 if accel is None else accel
        tm.Decelerate = 0.5 if decel is None else decel
        self._finish(T, dur)
        return e

    def wipe(self, shape, T, dur=0.8, dir_="left", exit_=False, ease=None):
        a, d = ease if ease else (0.2, 0.5)
        dmap = {"left": DIR_L, "right": DIR_R, "up": DIR_U, "down": DIR_D}
        return self.fx(shape, AE_WIPE, T=T, dur=dur, direction=dmap[dir_], exit_=exit_,
                       accel=a, decel=d)

    def fade(self, shape, T, dur=0.5, exit_=False):
        return self.fx(shape, AE_FADE, T=T, dur=dur, exit_=exit_, smooth=False,
                       accel=0.0, decel=0.0)

    def appear(self, shape, T):
        return self.fx(shape, AE_APPEAR, T=T, dur=0.001, smooth=False)

    def pop(self, shape, T, dur=0.45, exit_=False):
        """FadedZoom：柔和弹出（放大+淡入）/ 淡出缩小。"""
        return self.fx(shape, AE_FADEDZOOM, T=T, dur=dur, exit_=exit_,
                       accel=0.25, decel=0.6)

    def fly(self, shape, T, dur=0.55, from_="right"):
        d = {"right": DIR_R, "left": DIR_L, "top": DIR_U, "bottom": DIR_D}[from_]
        return self.fx(shape, AE_FLY, T=T, dur=dur, direction=d,
                       accel=0.3, decel=0.6)

    def path(self, shape, pts, T, dur=1.0, ease=(0.22, 0.62), resample_n=0):
        """运动路径。pts=绝对pt航点，首点=形状当前视觉中心（或后续路径接续点）。
        ease=(accelerate, decelerate) 0~1；默认带进带出（平滑滑行）。
        开环减速场景传 ease=(0.0, 1.0)（全程 ease-out，先快后慢）。"""
        ocx = shape.Left + shape.Width / 2.0
        ocy = shape.Top + shape.Height / 2.0
        use = resample(pts, resample_n) if resample_n and len(pts) > 2 else pts
        cmds = []
        p0 = use[0]
        dx, dy = (p0[0] - ocx) / 960.0, (p0[1] - ocy) / 540.0
        if abs(dx) < 1e-9 and abs(dy) < 1e-9:
            cmds.append("M 0 0")
        else:
            cmds.append(f"M {fr(dx)} {fr(dy)}")
        for (x, y) in use[1:]:
            cmds.append(f"L {fr((x - ocx) / 960.0)} {fr((y - ocy) / 540.0)}")
        trig, delay = self._sched(T, dur)
        eff = self.MS.AddEffect(shape, AE_PATHRIGHT, 0, trig)
        eff.Behaviors.Item(1).MotionEffect.Path = " ".join(cmds) + " E"
        tm = eff.Timing
        try:
            tm.TriggerDelayTime = max(delay, 0.0)
        except Exception:
            pass
        tm.Duration = dur
        try:
            tm.SmoothStart = MTRUE
            tm.SmoothEnd = MTRUE
        except Exception:
            pass
        tm.Accelerate = ease[0]
        tm.Decelerate = ease[1]
        self._finish(T, dur)
        return eff

    # ---------- 页眉页脚 ----------
    def header(self, title, sub, foot, note="放映（F5）自动播放"):
        self.text(36, 12, 780, 28, title, size=19, color=INK, bold=True)
        self.rect(36, 46, 52, 3, fill=TEAL)
        self.text(36, 52, 790, 18, sub, size=12, color=GRAY)
        self.text(36, 516, 830, 16, foot, size=12, color=GRAY)
        self.text(800, 12, 142, 16, note, size=12, color=GRAY, align="r")

    def effects_report(self):
        return self.MS.Count


# ---------------- Deck ----------------
class Deck:
    def __init__(self, advance_for_video=False, mode="live"):
        a = boot()
        self.app = a
        self.pres = a.Presentations.Add()
        self.pres.PageSetup.SlideWidth = 960.0
        self.pres.PageSetup.SlideHeight = 540.0
        self.slides = []
        # True: 写自动换片时间（供 CreateVideo 定片长，页播完 3s 后自动翻页）
        # False（交付默认）：不自动翻页——动画进页自动播放，播完停在终帧，单击才翻页
        self.advance_for_video = advance_for_video
        self.mode = mode

    def slide(self):
        s = Slide(self.pres, self.mode)
        self.slides.append(s)
        return s

    def save(self, path):
        # CreateVideo 的 UseTimingsAndNarrations 读取换片计时定片长；交付版关闭
        for s in self.slides:
            tr = s.s.SlideShowTransition
            if self.advance_for_video:
                tr.AdvanceOnTime = MTRUE
                tr.AdvanceTime = int(s.prev_end or 0) + 3
            else:
                tr.AdvanceOnTime = MFALSE
        if os.path.exists(path):
            os.remove(path)
        self.pres.SaveAs(path)
        return path

    def render_video(self, mp4, timeout=300):
        """CreateVideo 无头渲染（含时间线）。返回耗时。"""
        if os.path.exists(mp4):
            os.remove(mp4)
        self.pres.CreateVideo(mp4, True, 5)
        t0 = time.time()
        while True:
            st = self.pres.CreateVideoStatus
            if st == 3:      # done
                break
            if st not in (0, 1, 2):
                raise RuntimeError(f"CreateVideo 失败 status={st}")
            if time.time() - t0 > timeout:
                raise RuntimeError("CreateVideo 超时")
            time.sleep(1.5)
        return time.time() - t0

    def close(self, save=False):
        try:
            if save:
                self.pres.Save()
            self.pres.Close()
        except Exception:
            pass
