# -*- coding: utf-8 -*-
"""tavLst XML 注入器：给既有 pptx 的动画时间线注入阻尼弹簧或自定义 tav 缓动。

覆盖 COM 不可达的第 2 层能力（播放器时间线）。方法论：COM 造基础效果 → 本工具
对该页 click 组注入 `<p:anim>`（tavLst 采样）→ 三关验证，任一失败自动回滚
（不产出注入文件，原文件从不改动）。

内置三致命坑处理：
  1. tav tm 是千分比（0–100000），不是毫秒——越界直接报错拒绝；
  2. 值必须是 `<p:strVal>` 公式串——字面量自动加括号、#ppt_x/# 变量原样保留，
     永远不写 `<p:fltVal>`；
  3. cTn id 全局唯一——扫描全部 slide XML 取 max+1 连续分配。

三关验证（①②默认跑，--no-verify 跳过；③ --frames 可选）：
  ① 注入文件只读 Open 不触发修复（页数与原文件一致、主序列计数不减少）；
  ② CreateVideo 能出片；
  ③ 帧抽取 + 帧间运动能量分析（qa_motion.motion_energy）。

用法（Python）:
    inject("deck.pptx", slide_idx=3, shape_name_or_spid="d1", effect_index=0,
           spec={"freq": 3, "decay": 0.7, "amplitude": 0.15},
           out_path="deck_inj.pptx")
    spec 另支持 "spring"（默认参数）与自定义 tav："0=#ppt_x;50000=(0.6);100000=(0.4)"
    或 [(0, "#ppt_x"), (50000, "(0.6)"), ...]

用法（CLI）:
    python inject_easing.py deck.pptx --slide 3 --shape d1 --spring freq=3,decay=0.7 --out deck_inj.pptx
    python inject_easing.py deck.pptx --slide 3 --shape 12 --tav "0=#ppt_x;50000=(0.6);100000=(0.4)"
"""
import argparse
import math
import os
import re
import shutil
import sys
import time
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

MTRUE, MFALSE = -1, 0

# 默认 spring 参数 = 已验证口径（A=0.15, W=19.0, DECAY=2.6, T0=600ms, DUR=2600ms, N=30）
SPRING_DEFAULTS = {
    "amplitude": 0.15,     # 振幅（attr 归一化坐标分数，0.15 ≈ 0.15×页宽）
    "freq": 19.0,          # 角频率（cos 相位）
    "decay": 2.6,          # 阻尼系数（e^-decay·u，u∈[0,1]）
    "center": 0.5,         # 平衡位置（0.5 = 页面中央）
    "n": 30,               # tav 采样点数
    "t0": 600,             # 注入 par 起始延迟（ms，相对时间线起点，绝对语义）
    "dur": 2600,           # anim 时长（ms）
    "attr": "ppt_x",       # 作用属性：ppt_x/ppt_y/ppt_w/ppt_h/ppt_r/style.color
}


# ---------------- 三致命坑处理 ----------------
def _formula(v):
    """致命坑 2：值必须是 fbml 公式串——字面量加括号；#引用/$表达式原样保留。"""
    s = str(v).strip()
    if s.startswith("#") or s.startswith("$") or (s.startswith("(") and s.endswith(")")):
        return s
    return "(" + s + ")"


def _fmt(v):
    s = f"{v:.5f}".rstrip("0").rstrip(".")
    return s if s not in ("", "-") else "0"


def normalize_tavs(pairs, attr):
    """致命坑 1+2：tm 千分比校验（0–100000，拒绝毫秒）、公式串化、排序去重、补 tm=0 锚点。"""
    out = {}
    for tm, val in pairs:
        tm = int(round(float(tm)))
        if not (0 <= tm <= 100000):
            raise ValueError(f"tav tm 越界：{tm}（tm 是千分比 0–100000，不是毫秒——致命坑 1）")
        out[tm] = _formula(val)
    if 0 not in out:
        out[0] = "#" + attr          # 补当前位置引用锚点（#ppt_x/#ppt_y/…）
    return [(tm, out[tm]) for tm in sorted(out)]


def build_spec(spec, attr="ppt_x", t0=None, dur=None):
    """spec → (params, tavs)。spring：dict（可只给部分键）或 "spring"；
    自定义 tav：[(tm, 值)...] / "tm=值;tm=值"。attr/t0/dur 覆盖 spring 默认。"""
    if isinstance(spec, dict) or (isinstance(spec, str) and spec.strip().lower() == "spring"):
        p = dict(SPRING_DEFAULTS)
        if isinstance(spec, dict):
            p.update(spec)
            p.pop("type", None)
        if attr:
            p["attr"] = attr
        if t0 is not None:
            p["t0"] = t0
        if dur is not None:
            p["dur"] = dur
        n = max(int(p["n"]), 2)
        pairs = [(0, "#" + p["attr"])]
        for i in range(1, n + 1):
            u = i / float(n)
            off = p["amplitude"] * math.exp(-p["decay"] * u) * math.cos(p["freq"] * u)
            pairs.append((int(round(100000.0 * i / n)), _fmt(p["center"] + off)))
        pairs[-1] = (100000, pairs[-1][1])          # 终点精确收在 100000
        return p, normalize_tavs(pairs, p["attr"])
    pairs = _parse_tav_pairs(spec)
    p = {"attr": attr or "ppt_x",
         "t0": int(t0) if t0 is not None else 0,
         "dur": int(dur) if dur is not None else 1000}
    return p, normalize_tavs(pairs, p["attr"])


def _parse_tav_pairs(spec):
    if isinstance(spec, str):
        items = []
        for part in spec.split(";"):
            part = part.strip()
            if not part:
                continue
            if "=" not in part:
                raise ValueError(f"tav 段需 tm=值 形式：{part}")
            tm, val = part.split("=", 1)
            items.append((tm, val.strip()))
        if not items:
            raise ValueError("自定义 tav 为空")
        return items
    return list(spec)


# ---------------- XML 构造与定位 ----------------
def anim_par_xml(tavs, spid, attr, t0, dur, id0):
    """注入 par（致命坑 3 在此落实：id0 = 全局 max+1，连续占两个 id）。"""
    body = "".join(f'<p:tav tm="{tm}"><p:val><p:strVal val="{f}"/></p:val></p:tav>'
                   for tm, f in tavs)
    return (f'<p:par><p:cTn id="{int(id0)}" fill="hold"><p:stCondLst>'
            f'<p:cond delay="{int(t0)}"/></p:stCondLst><p:childTnLst>'
            f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            f'<p:cTn id="{int(id0) + 1}" dur="{int(dur)}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{int(spid)}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst>{body}</p:tavLst></p:anim></p:childTnLst></p:cTn></p:par>')


def resolve_spid(slide_xml, name_or_spid):
    """形状 Name → spid（精确匹配优先，唯一子串兜底）；纯数字按 spid 直通。"""
    s = str(name_or_spid).strip()
    if s.isdigit():
        return int(s)
    hits_exact, hits_sub = [], []
    for tag in re.findall(r"<p:cNvPr\b[^>]*>", slide_xml):
        mid = re.search(r'\bid="(\d+)"', tag)
        mname = re.search(r'\bname="([^"]*)"', tag)
        if not (mid and mname):
            continue
        name, spid = mname.group(1), int(mid.group(1))
        if name == s:
            hits_exact.append(spid)
        elif s.lower() in name.lower():
            hits_sub.append(spid)
    if len(hits_exact) == 1:
        return hits_exact[0]
    if len(hits_exact) > 1:
        raise ValueError(f"形状名不唯一：{s} 命中 spid {hits_exact}（请改用数字 spid）")
    if len(hits_sub) == 1:
        return hits_sub[0]
    raise ValueError(f"找不到形状：{name_or_spid}（exact={hits_exact}, sub={hits_sub}）")


def _max_tn_id(files):
    """致命坑 3：扫描全部 slide XML 的 cTn id 取全局 max。"""
    mx = 0
    for name, data in files.items():
        if re.match(r"ppt/slides/slide\d+\.xml$", name):
            for s in re.findall(r'<p:cTn id="(\d+)"', data.decode("utf-8")):
                mx = max(mx, int(s))
    return mx


def _click_group_content(x):
    """定位主序列 click 组 childTnLst 的内容区间 (lo, hi)。"""
    m = re.search(r'<p:par><p:cTn id="\d+" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                  r'</p:stCondLst><p:childTnLst>', x)
    if m:
        tag_end = m.end()
    else:                                  # 兜底：从 mainSeq 找
        i = x.find('nodeType="mainSeq"')
        j = x.find("<p:childTnLst>", i) if i >= 0 else -1
        if j < 0:
            raise ValueError("未找到主序列 click 组——该页可能没有动画；"
                             "请先用 COM 建基础效果（本工具只做第 2 层注入）")
        tag_end = j + len("<p:childTnLst>")
    depth = 1
    for m2 in re.finditer(r"<(/?)p:childTnLst>", x[tag_end:]):
        depth += -1 if m2.group(1) else 1
        if depth == 0:
            return tag_end, tag_end + m2.start()
    raise ValueError("childTnLst 未闭合（XML 异常）")


def _par_spans(x, lo, hi):
    """[lo,hi) 内顶层 <p:par>…</p:par> 的 (start, end)（深度平衡扫描）。"""
    spans, depth, start = [], 0, None
    for m in re.finditer(r"<(/?)p:par(?=[\s/>])[^>]*?(/?)>", x[lo:hi]):
        if m.group(2):                     # <p:par/>
            if depth == 0:
                spans.append((lo + m.start(), lo + m.end()))
            continue
        if not m.group(1):
            if depth == 0:
                start = lo + m.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0:
                spans.append((start, lo + m.end()))
    return spans


def _insert_par(x, effect_index, par_xml):
    """把注入 par 插为 click 组 childTnLst 内第 effect_index 个效果 par 的前驱
    （同组兄弟，共时间线起点；起始时刻由 par 的 cond delay 绝对语义控制）。
    effect_index=-1 表示追加到组尾。"""
    lo, hi = _click_group_content(x)
    spans = _par_spans(x, lo, hi)
    if effect_index is None or effect_index == -1 or not spans:
        pos = hi
    else:
        if not (0 <= int(effect_index) < len(spans)):
            raise ValueError(f"effect_index={effect_index} 越界：该 click 组共 {len(spans)} 个效果 par")
        pos = spans[int(effect_index)][0]
    return x[:pos] + par_xml + x[pos:]


# ---------------- zip 读写 ----------------
def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def write_zip(path, files):
    """OPC 重打包：[Content_Types].xml 必须是第一个条目、ZIP_DEFLATED、无目录条目。"""
    names = sorted(files)
    ordered = ([n for n in names if n == "[Content_Types].xml"] +
               [n for n in names if n != "[Content_Types].xml"])
    if os.path.exists(path):
        os.remove(path)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ordered:
            z.writestr(n, files[n])        # 只写文件条目，不写目录条目


# ---------------- 三关验证 ----------------
def _kill_pp():
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)


def _boot():
    try:
        import common as K
    except ImportError as e:
        raise RuntimeError("需要同目录 common.py 且已安装 pywin32（pip install pywin32）") from e
    return K


def _probe(a, path, slide_idx):
    """只读打开读回 (页数, 主序列效果数)。触发修复/损坏会在此抛 COM 错误。"""
    pres = a.Presentations.Open(path, MTRUE, MFALSE, MFALSE)
    try:
        n_slides = int(pres.Slides.Count)
        ms = int(pres.Slides.Item(slide_idx).TimeLine.MainSequence.Count)
        return n_slides, ms
    finally:
        pres.Close()


def _create_video(path, mp4, timeout):
    """非只读打开并 CreateVideo（与渲染验证管线同口径）。返回 mp4 路径。"""
    K = _boot()
    a = K.boot()
    pres = a.Presentations.Open(path, MFALSE, MFALSE, MFALSE)
    try:
        if os.path.exists(mp4):
            os.remove(mp4)
        pres.CreateVideo(mp4, True, 5)
        t0 = time.time()
        while True:
            st = pres.CreateVideoStatus
            if st == 3:
                return mp4
            if st not in (0, 1, 2):
                raise RuntimeError(f"CreateVideo 失败 status={st}")
            if time.time() - t0 > timeout:
                raise RuntimeError("CreateVideo 超时")
            time.sleep(2)
    finally:
        pres.Close()


def _run_gates(orig, tmp, slide_idx, do_frames, motion_eps, timeout):
    """①重开不修复 ②CreateVideo 出片 ③（可选）帧运动能量。失败删 tmp 抛错=回滚。"""
    K = _boot()
    a = K.boot()
    report = {"gate1": None, "gate2": None, "gate3": None}
    mp4 = tmp + ".verify.mp4"
    try:
        b_slides, b_ms = _probe(a, orig, slide_idx)          # 原文件基线
        i_slides, i_ms = _probe(a, tmp, slide_idx)           # 关①
        if i_slides != b_slides:
            raise RuntimeError(f"关①失败：页数变化 {b_slides}→{i_slides}（疑似修复丢弃内容）")
        if i_ms < b_ms:
            raise RuntimeError(f"关①失败：主序列效果数减少 {b_ms}→{i_ms}（疑似修复丢弃效果）")
        report["gate1"] = f"OK（slides={i_slides}, mainSeq={i_ms} vs 基线 {b_ms}）"
        _create_video(tmp, mp4, timeout)                     # 关②
        report["gate2"] = f"OK（{os.path.getsize(mp4)} bytes）"
        if do_frames:                                        # 关③
            import qa_motion as q
            d = q.probe_dur(mp4) or 8.0
            fdir = mp4 + ".frames"
            try:
                q.extract(mp4, fdir, fps=min(10.0, 24.0 / max(d, 0.1)))
                e = q.motion_energy(fdir)
                peak = max(e) if e else 0.0
                if motion_eps is not None and peak < motion_eps:
                    raise RuntimeError(f"关③失败：峰值帧间运动能量 {peak:.2f} < 阈值 {motion_eps}"
                                       "（疑似静默不播）")
                report["gate3"] = f"OK（峰值帧间运动能量 {peak:.2f}，帧数 {len(e) + 1}）"
            finally:
                shutil.rmtree(fdir, ignore_errors=True)
        return report
    finally:
        K.quit_app()
        _kill_pp()
        if os.path.exists(mp4):
            try:
                os.remove(mp4)
            except OSError:
                pass


# ---------------- 主入口 ----------------
def inject(pptx_path, slide_idx, shape_name_or_spid, effect_index=0, spec="spring",
           out_path=None, attr=None, t0=None, dur=None,
           verify=True, frames=False, motion_eps=None, timeout=600):
    """对第 slide_idx 页注入 spring/自定义 tav 缓动，产出 out_path。

    原 pptx 不改动；先写 <out>.tmp → 三关验证 → 通过才落位 out_path（失败即回滚）。
    返回报告 dict。
    """
    pptx_path = os.path.abspath(pptx_path)
    out_path = os.path.abspath(out_path or
                               os.path.splitext(pptx_path)[0] + "_injected.pptx")
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(pptx_path)
    slide_name = f"ppt/slides/slide{int(slide_idx)}.xml"
    files = read_zip(pptx_path)
    if slide_name not in files:
        raise ValueError(f"无此页 slide{slide_idx}（或该包无 ppt/slides/slide{slide_idx}.xml）")
    x = files[slide_name].decode("utf-8")

    spid = resolve_spid(x, shape_name_or_spid)
    params, tavs = build_spec(spec, attr=attr, t0=t0, dur=dur)
    id0 = _max_tn_id(files) + 1                     # 致命坑 3：全局 max+1
    par = anim_par_xml(tavs, spid, params["attr"], params["t0"], params["dur"], id0)
    files[slide_name] = _insert_par(x, effect_index, par).encode("utf-8")

    tmp = out_path + ".tmp"
    write_zip(tmp, files)                           # Content_Types 第一 / DEFLATED / 无目录条目
    report = {"out": out_path, "spid": spid, "cTn_ids": [id0, id0 + 1],
              "attr": params["attr"], "t0_ms": params["t0"], "dur_ms": params["dur"],
              "n_tav": len(tavs), "gates": {}}
    try:
        if verify:
            report["gates"] = _run_gates(pptx_path, tmp, int(slide_idx),
                                         frames, motion_eps, timeout)
        os.replace(tmp, out_path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)                          # 回滚：不产出注入文件
        raise
    return report


def main():
    ap = argparse.ArgumentParser(description="tavLst XML 注入器（spring/自定义 tav）",
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--slide", type=int, required=True, help="页号（1 起）")
    ap.add_argument("--shape", required=True, help="形状 Name 或数字 spid")
    ap.add_argument("--spring", default=None,
                    help="spring 参数逗号分隔：freq=3,decay=0.7,amplitude=0.15,"
                         "center=0.5,n=30,t0=600,dur=2600,attr=ppt_x")
    ap.add_argument("--tav", default=None,
                    help='自定义 tav（替代 --spring）："0=#ppt_x;50000=(0.6);100000=(0.4)"')
    ap.add_argument("--effect-index", type=int, default=0,
                    help="插入为 click 组内第 N 个效果 par 的前驱（0 起；-1=组尾）")
    ap.add_argument("--attr", default=None, help="作用属性（默认 ppt_x）")
    ap.add_argument("--t0", type=float, default=None, help="注入 par 起始延迟 ms")
    ap.add_argument("--dur", type=float, default=None, help="anim 时长 ms")
    ap.add_argument("--out", default=None)
    ap.add_argument("--no-verify", action="store_true", help="跳过三关验证（不推荐）")
    ap.add_argument("--frames", action="store_true", help="跑第三关：帧运动能量分析")
    ap.add_argument("--motion-eps", type=float, default=None,
                    help="第三关峰值运动能量断言阈值（如 2.0）")
    a = ap.parse_args()

    spec = a.tav if a.tav is not None else "spring"
    if a.spring is not None:
        if a.tav is not None:
            ap.error("--spring 与 --tav 只能二选一")
        spec = {}
        for part in a.spring.split(","):
            part = part.strip()
            if not part:
                continue
            k, v = part.split("=", 1)
            k = k.strip()
            if k == "attr":
                spec[k] = v.strip()
            elif k == "n":
                spec[k] = int(v)
            else:
                spec[k] = float(v)
    rep = inject(a.pptx, a.slide, a.shape, effect_index=a.effect_index, spec=spec,
                 out_path=a.out, attr=a.attr, t0=a.t0, dur=a.dur,
                 verify=not a.no_verify, frames=a.frames, motion_eps=a.motion_eps)
    print(f"injected -> {rep['out']}")
    print(f"  spid={rep['spid']} attr={rep['attr']} t0={rep['t0_ms']}ms "
          f"dur={rep['dur_ms']}ms tav={rep['n_tav']} pts cTn={rep['cTn_ids']}")
    for g in ("gate1", "gate2", "gate3"):
        if rep["gates"].get(g):
            print(f"  {g}: {rep['gates'][g]}")


if __name__ == "__main__":
    main()
