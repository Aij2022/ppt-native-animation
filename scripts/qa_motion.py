# -*- coding: utf-8 -*-
"""渲染 QA + 运动验证通用工具（零项目硬编码；Windows + Office 16.0 + pywin32）。

管线（QA 闭环的工具化）：交付 pptx → 复制为渲染副本（只改"自动换片时间"，动画
时间线不动）→ CreateVideo 30fps 出 mp4（轮询 CreateVideoStatus，3=done）→
ffmpeg（缺则 imageio）抽帧 → 均匀采样 n 帧拼蒙太奇（帧号+页内时刻标签画在画面
下方条带，不遮内容）→ 频带内质心逐帧跟踪 → 前后半程平均速率比（减速叙事断言，
参考：ease=(0,1) 纯 Decelerate 实测 ≈4.25:1，阈值 ≥1.6）。

页时长泛化口径：不给 --durations 时用 COM 读回主序列自动计算——
max(起始时刻 + Duration × RepeatCount(>1) × 2(AutoReverse) + BounceEnd 回弹余量)。
渲染收尾务必 Quit，并 taskkill 清理残留 POWERPNT.EXE（若同时人工开着其他
PowerPoint 窗口，请先保存关闭再运行）。

CLI:
  python qa_motion.py render  deck.pptx [--out-dir DIR] [--fps 30] [--hold 2]
        [--durations 1:11.1,2:9.25] [--morph 6:2] [--reuse]
  python qa_motion.py extract render.mp4 --out-dir frames [--fps 30] [--start S] [--end E]
  python qa_motion.py montage frames --out mont.png [-n 16] [--cols 4]
        [--t0 0 --t1 12] [--fps 30] [--prefix s1] [--font PATH]
  python qa_motion.py track frames [--band 200:320] [--color 192,57,43]
        [--tol 60] [--thresh 200] [--fps 30] [--csv track.csv]
  python qa_motion.py speed track.csv [--split 0.5]
"""
import argparse
import csv
import math
import os
import re
import shutil
import subprocess
import sys
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

MTRUE, MFALSE = -1, 0


# ---------------- 进程卫生 ----------------
def kill_pp():
    """taskkill 清理残留 POWERPNT.EXE（COM 崩溃遗留进程会锁文件）。"""
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)


# ---------------- COM：渲染 ----------------
def _boot():
    try:
        import common as K
    except ImportError as e:
        raise RuntimeError("需要同目录 common.py 且已安装 pywin32（pip install pywin32）") from e
    return K


def slide_duration(slide, bounce_margin=1.0):
    """COM 读回单页动画时长（秒）：各效果 max(起始 + 有效时长)。

    有效时长 = Duration × RepeatCount(>1 时) × 2(AutoReverse=-1)
    + BounceEnd 回弹余量（回弹不计入 Duration，换片须另加 ~1s）。
    """
    end = 0.0
    seq = slide.TimeLine.MainSequence
    for i in range(1, seq.Count + 1):
        tm = seq.Item(i).Timing
        try:
            start = float(tm.TriggerDelayTime or 0.0)
        except Exception:
            start = 0.0
        try:
            dur = float(tm.Duration or 0.0)
        except Exception:
            dur = 0.0
        try:
            rep = int(tm.RepeatCount)
        except Exception:
            rep = 0
        total = dur * rep if rep > 1 else dur
        try:
            if tm.AutoReverse == MTRUE:
                total *= 2
        except Exception:
            pass
        try:
            if tm.BounceEnd == MTRUE:
                total += bounce_margin
        except Exception:
            pass
        end = max(end, start + total)
    return end


def render(pptx_path, fps=30, hold=2, out_dir=None, durations=None,
           morph_extra=None, timeout=600, default_slide=5, bounce_margin=1.0):
    """交付 pptx → 渲染副本（只改换片时间）→ CreateVideo 出 mp4。

    durations: {页号: 秒}；缺省用 COM 读回自动计算。morph_extra: {页号: 额外秒}
    （Morph 换片占用，计入该页窗口）。返回 (mp4_path, durations_dict)。
    """
    K = _boot()
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir or os.path.join(os.path.dirname(pptx_path), "qa_render"))
    os.makedirs(out_dir, exist_ok=True)
    copy = os.path.join(out_dir, "_render_copy.pptx")
    mp4 = os.path.join(out_dir, "render.mp4")
    morph_extra = morph_extra or {}
    a = K.boot()
    try:
        if os.path.exists(copy):
            os.remove(copy)
        shutil.copyfile(pptx_path, copy)
        pres = a.Presentations.Open(copy, MFALSE, MFALSE, MFALSE)
        try:
            n = int(pres.Slides.Count)
            durs = {}
            for si in range(1, n + 1):
                if durations and si in durations:
                    d = float(durations[si])
                else:
                    d = slide_duration(pres.Slides.Item(si), bounce_margin)
                durs[si] = d
                tr = pres.Slides.Item(si).SlideShowTransition
                tr.AdvanceOnTime = MTRUE
                tr.AdvanceOnClick = MFALSE
                tr.AdvanceTime = int(math.ceil(d)) + int(hold) + int(morph_extra.get(si, 0))
            pres.Save()
            if os.path.exists(mp4):
                os.remove(mp4)
            if int(fps) == 30:                     # 与已验证管线逐字一致的调用
                pres.CreateVideo(mp4, True, default_slide)
            else:
                pres.CreateVideo(mp4, True, default_slide, 720, int(fps))
            t0 = time.time()
            while True:
                st = pres.CreateVideoStatus
                if st == 3:                        # done
                    break
                if st not in (0, 1, 2):
                    raise RuntimeError(f"CreateVideo 失败 status={st}")
                if time.time() - t0 > timeout:
                    raise RuntimeError("CreateVideo 超时")
                time.sleep(2)
            return mp4, durs
        finally:
            try:
                pres.Close()
            except Exception:
                pass
    finally:
        K.quit_app()
        kill_pp()


# ---------------- 时窗与抽帧 ----------------
def probe_dur(mp4):
    """ffprobe 读容器总时长；ffprobe 不可用返回 None。"""
    try:
        r = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                            "-of", "csv=p=0", mp4], capture_output=True, text=True)
        return float(r.stdout.strip())
    except (FileNotFoundError, ValueError):
        return None


def slide_windows(total, durs, hold=2, morph_extra=None):
    """按换片窗把 mp4 总时长切成每页 (t0, anim_end)。

    窗宽 = ceil(页时长)+hold+morph；total 与理论窗和的微小偏差按比例校正。
    """
    morph_extra = morph_extra or {}
    ws, t = [], 0.0
    for si in sorted(durs):
        d = float(durs[si])
        w = math.ceil(d) + hold + morph_extra.get(si, 0)
        ws.append((t, t + d))                 # 窗起点 → 动画末时刻（终帧保持段在窗尾）
        t += w
    tot = total if total else t
    scale = tot / t if t else 1.0
    return [(a * scale, b * scale) for (a, b) in ws]


def frame_at(mp4, t, out_png):
    """抽视频第 t 秒单帧。"""
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-ss", f"{float(t):.3f}", "-i", mp4,
                    "-frames:v", "1", out_png], check=True)
    return out_png


def list_frames(frames_dir):
    fs = [os.path.join(frames_dir, f) for f in os.listdir(frames_dir)
          if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp"))]
    return sorted(fs)


def extract(mp4, out_dir, fps=None, start=None, end=None):
    """抽帧：ffmpeg 优先（缺/败则 imageio 兜底）。fps 给定时按目标帧率采样。"""
    os.makedirs(out_dir, exist_ok=True)
    pattern = os.path.join(out_dir, "f_%05d.png")
    try:
        cmd = ["ffmpeg", "-y", "-v", "error"]
        if start is not None:
            cmd += ["-ss", f"{float(start):.3f}"]
        cmd += ["-i", mp4]
        if end is not None:
            dur = float(end) - (float(start) if start is not None else 0.0)
            cmd += ["-t", f"{max(dur, 0.001):.3f}"]
        if fps:
            cmd += ["-vf", f"fps={fps}"]
        cmd += [pattern]
        subprocess.run(cmd, check=True)
        if list_frames(out_dir):
            return out_dir
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass
    return _extract_imageio(mp4, out_dir, fps, start, end)


def _extract_imageio(mp4, out_dir, fps=None, start=None, end=None):
    try:
        import imageio.v2 as imageio
    except Exception:
        raise RuntimeError("ffmpeg 不可用且未安装 imageio，无法抽帧")
    src_fps = 30.0
    try:
        with imageio.get_reader(mp4) as r0:
            src_fps = float(r0.get_meta_data().get("fps", 30.0))
    except Exception:
        pass
    step = max(1, int(round(src_fps / fps))) if fps else 1
    n = 0
    with imageio.get_reader(mp4) as r:
        for i, frame in enumerate(r):
            t = i / src_fps
            if start is not None and t < float(start) - 1e-6:
                continue
            if end is not None and t > float(end) + 1e-6:
                break
            if step > 1 and i % step:
                continue
            n += 1
            imageio.imwrite(os.path.join(out_dir, f"f_{n:05d}.png"), frame)
    if n == 0:
        raise RuntimeError(f"抽帧 0 帧：{mp4}")
    return out_dir


# ---------------- 蒙太奇 ----------------
def load_font(size=16, path=None):
    """中文字体两级回退：wqy-zenhei（Linux/WSL）→ msyh.ttc / 微软雅黑（Windows）。"""
    from PIL import ImageFont
    cands = [path] if path else []
    cands += [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "msyh.ttc"),
        "msyh.ttc",
        "微软雅黑",
    ]
    for p in cands:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _uniform(idxs_len, n):
    if idxs_len <= n:
        return list(range(idxs_len))
    n = max(n, 2)
    out, seen = [], set()
    for k in range(n):
        i = round(k * (idxs_len - 1) / (n - 1))
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def montage(frames_dir, out_png, n=16, cols=4, t0=0.0, t1=None, fps=None,
            prefix="", font=None, strip_h=30, bg="#1F3864", fg="#FFFFFF"):
    """均匀采样 n 帧拼蒙太奇；帧号+页内时刻标签画在画面下方条带（不遮内容）。

    时间标注：t1 给定 → 窗口模式，时间在 [t0,t1] 均分（t0=页窗起点时标签即页内
    时刻）；否则帧率模式 t = t0 + 帧序号/fps。
    """
    from PIL import Image, ImageDraw
    frames = list_frames(frames_dir)
    if not frames:
        raise RuntimeError(f"无帧可拼：{frames_dir}")
    n = max(int(n), 1)
    idxs = _uniform(len(frames), n)
    if t1 is not None:
        times = [t0 + (t1 - t0) * k / max(len(idxs) - 1, 1) for k in range(len(idxs))]
    else:
        base = float(fps) if fps else 30.0
        times = [t0 + idx / base for idx in idxs]
    ims = [Image.open(frames[i]) for i in idxs]
    w, h = ims[0].size
    rows = (len(idxs) + cols - 1) // cols
    m = Image.new("RGB", (w * cols, (h + strip_h) * rows), bg)
    dr = ImageDraw.Draw(m)
    fnt = load_font(15, font)
    for i, im in enumerate(ims):
        r, c = divmod(i, cols)
        x0, y0 = w * c, r * (h + strip_h)
        m.paste(im, (x0, y0))
        dr.rectangle([x0, y0 + h, x0 + w, y0 + h + strip_h], fill=bg)
        stem = os.path.splitext(os.path.basename(frames[idxs[i]]))[0]
        dm = re.search(r"(\d+)", stem)
        fno = int(dm.group(1)) if dm else idxs[i] + 1
        dr.text((x0 + 8, y0 + h + (strip_h - 15) // 2),
                f"{prefix} f{fno:02d}  t={times[i] - t0:5.2f}s", fill=fg, font=fnt)
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    m.save(out_png)
    return out_png


# ---------------- 质心跟踪与速率 ----------------
def _mask_centroid(im, color, tol, thresh):
    """帧内掩码质心 (cx, cy)；无命中返回 None。color=(r,g,b) 时按逐通道差 ≤ tol。"""
    try:
        import numpy as np
    except ImportError:
        np = None
    if np is not None:
        a = np.asarray(im, dtype=np.int16)
        if color is not None:
            mask = (np.abs(a - np.array(color, dtype=np.int16)) <= tol).all(axis=2)
        else:
            lum = 0.299 * a[..., 0] + 0.587 * a[..., 1] + 0.114 * a[..., 2]
            mask = lum >= thresh
        n = int(mask.sum())
        if n == 0:
            return None
        ys, xs = np.nonzero(mask)
        return float(xs.mean()), float(ys.mean())
    from PIL import ImageChops
    if color is not None:
        d = ImageChops.difference(im, Image.new("RGB", im.size, tuple(color)))
        mk = d.convert("L").point(lambda v: 255 if v <= tol else 0)   # 近似：灰度差≤tol
    else:
        mk = im.convert("L").point(lambda v: 255 if v >= thresh else 0)
    cnt = mk.histogram()[255]
    if cnt == 0:
        return None
    px = mk.load()
    sx = sy = 0
    for yy in range(im.height):
        for xx in range(im.width):
            if px[xx, yy]:
                sx += xx
                sy += yy
    return sx / cnt, sy / cnt


def track_centroid(frames, band=None, color=None, tol=60, thresh=None,
                   fps=None, t0=0.0, min_pixels=6, out_csv=None):
    """频带内质心逐帧跟踪，输出 [(t, x, y)...]。

    frames: 帧目录或帧路径列表。band=(y0,y1) 像素行（频带，None=整帧）。
    color=(r,g,b) 色带（推荐：给目标形状填色）；二者至少给一个，色带优先。
    t = t0 + 帧序号/fps（fps 缺省退化为帧序号）。
    """
    from PIL import Image
    paths = list_frames(frames) if isinstance(frames, str) else list(frames)
    if not paths:
        raise RuntimeError(f"无帧可跟踪：{frames}")
    pts = []
    for i, p in enumerate(paths):
        im = Image.open(p).convert("RGB")
        if band:
            im = im.crop((0, int(band[0]), im.width, int(band[1])))
        c = _mask_centroid(im, color, tol, thresh)
        if c is None:
            continue
        cx, cy = c
        cy += int(band[0]) if band else 0
        t = (t0 + i / float(fps)) if fps else float(i)
        pts.append((t, cx, cy))
    if out_csv:
        os.makedirs(os.path.dirname(os.path.abspath(out_csv)), exist_ok=True)
        with open(out_csv, "w", newline="", encoding="utf-8") as f:
            wcsv = csv.writer(f)
            wcsv.writerow(["t", "x", "y"])
            wcsv.writerows(pts)
    return pts


def speed_ratio(track, split=0.5):
    """前后半程平均速率比（时间中点 split 分割；减速叙事应 >1，参考 ≥1.6）。"""
    if len(track) < 3:
        raise ValueError("track 太短（至少 3 点）")
    t_a, t_z = track[0][0], track[-1][0]
    t_cut = t_a + (t_z - t_a) * float(split)

    def avg(seg):
        dist = sum(math.dist(seg[i - 1][1:], seg[i][1:]) for i in range(1, len(seg)))
        dt = seg[-1][0] - seg[0][0]
        return dist / dt if dt > 1e-9 else 0.0

    first = [p for p in track if p[0] <= t_cut]
    second = [p for p in track if p[0] > t_cut]
    if len(first) < 2 or len(second) < 2:
        raise ValueError("分割点两侧样本不足")
    v2 = avg(second)
    if v2 < 1e-9:
        return float("inf")
    return avg(first) / v2


def motion_energy(frames_dir):
    """相邻帧平均绝对灰度差（0-255 标度）——帧间运动强度，切页处为尖峰。"""
    from PIL import Image, ImageChops
    prev, out = None, []
    for p in list_frames(frames_dir):
        im = Image.open(p).convert("L").resize((320, 180))
        if prev is not None:
            h = ImageChops.difference(prev, im).histogram()
            out.append(sum(i * c for i, c in enumerate(h)) / (320 * 180))
        prev = im
    return out


# ---------------- CLI ----------------
def _parse_kv_dict(s):
    out = {}
    for part in (s or "").split(","):
        part = part.strip()
        if not part:
            continue
        k, v = part.split(":", 1)
        out[int(k)] = float(v)
    return out


def main():
    ap = argparse.ArgumentParser(description="渲染 QA + 运动验证（QA 闭环工具化）",
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("render", help="交付 pptx → 渲染副本 → CreateVideo 出 mp4")
    p.add_argument("pptx")
    p.add_argument("--out-dir", default=None)
    p.add_argument("--fps", type=int, default=30)
    p.add_argument("--hold", type=int, default=2, help="终帧保持秒数（AdvanceTime 加项）")
    p.add_argument("--durations", default=None, help="页时长覆盖 1:11.1,2:9.25（缺省 COM 读回）")
    p.add_argument("--morph", default=None, help="Morph 额外窗口 6:2（页号:秒）")
    p.add_argument("--timeout", type=float, default=600)
    p.add_argument("--reuse", action="store_true", help="mp4 已存在则跳过渲染")

    p = sub.add_parser("extract", help="mp4 → 帧序列")
    p.add_argument("mp4")
    p.add_argument("--out-dir", required=True)
    p.add_argument("--fps", type=float, default=None)
    p.add_argument("--start", type=float, default=None)
    p.add_argument("--end", type=float, default=None)

    p = sub.add_parser("montage", help="帧目录 → n 帧均匀采样蒙太奇")
    p.add_argument("frames_dir")
    p.add_argument("--out", required=True)
    p.add_argument("-n", type=int, default=16)
    p.add_argument("--cols", type=int, default=4)
    p.add_argument("--t0", type=float, default=0.0, help="页窗起点（标签= t-t0 即页内时刻）")
    p.add_argument("--t1", type=float, default=None, help="页窗终点（窗口均分模式）")
    p.add_argument("--fps", type=float, default=None, help="帧率模式（无 t1 时）")
    p.add_argument("--prefix", default="")
    p.add_argument("--font", default=None)

    p = sub.add_parser("track", help="频带内质心逐帧跟踪")
    p.add_argument("frames_dir")
    p.add_argument("--band", default=None, help="y 频带 y0:y1（像素行）")
    p.add_argument("--color", default=None, help="色带 R,G,B（目标形状填色）")
    p.add_argument("--tol", type=int, default=60)
    p.add_argument("--thresh", type=float, default=None, help="无 color 时亮度阈值")
    p.add_argument("--fps", type=float, default=None)
    p.add_argument("--t0", type=float, default=0.0)
    p.add_argument("--csv", dest="out_csv", default=None)

    p = sub.add_parser("speed", help="轨迹前后半程速率比（减速叙事断言）")
    p.add_argument("track_csv")
    p.add_argument("--split", type=float, default=0.5)

    a = ap.parse_args()
    if a.cmd == "render":
        if a.reuse and a.out_dir and os.path.exists(os.path.join(a.out_dir, "render.mp4")):
            print("reuse existing mp4")
            return
        mp4, durs = render(a.pptx, fps=a.fps, hold=a.hold, out_dir=a.out_dir,
                           durations=_parse_kv_dict(a.durations),
                           morph_extra=_parse_kv_dict(a.morph), timeout=a.timeout)
        print("mp4:", mp4)
        for si, d in sorted(durs.items()):
            print(f"  slide {si}: anim={d:.2f}s advance={int(math.ceil(d)) + a.hold}s")
    elif a.cmd == "extract":
        print("frames:", extract(a.mp4, a.out_dir, fps=a.fps, start=a.start, end=a.end))
    elif a.cmd == "montage":
        print("montage:", montage(a.frames_dir, a.out, n=a.n, cols=a.cols, t0=a.t0,
                                  t1=a.t1, fps=a.fps, prefix=a.prefix, font=a.font))
    elif a.cmd == "track":
        band = tuple(int(v) for v in a.band.split(":")) if a.band else None
        color = tuple(int(v) for v in a.color.split(",")) if a.color else None
        pts = track_centroid(a.frames_dir, band=band, color=color, tol=a.tol,
                             thresh=a.thresh, fps=a.fps, t0=a.t0, out_csv=a.out_csv)
        print(f"points: {len(pts)}" + (f" -> {a.out_csv}" if a.out_csv else ""))
        for row in pts[:3]:
            print("  %.3f  (%.1f, %.1f)" % row)
    elif a.cmd == "speed":
        with open(a.track_csv, encoding="utf-8") as f:
            rows = [r for r in csv.reader(f) if r and r[0] != "t"]
        track = [(float(t), float(x), float(y)) for t, x, y in rows]
        r = speed_ratio(track, split=a.split)
        print(f"speed_ratio(first/second) = {r:.2f}  (减速叙事参考阈值 ≥1.6)")


if __name__ == "__main__":
    main()
