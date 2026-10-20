# PowerPoint 原生动画技术文档（COM 自动化完整参考）

> 来源：某 VLA 答辩项目实测沉淀 · 基于 Office 16.0（版本号 16.0，非 365）+ pywin32 实测
> 配套：`../scripts/common.py`（构建框架）· `capabilities.md`（能力矩阵：MsoAnimEffect 0–154 全量与切换/图表/SmartArt/媒体/声音族，capability3 审计版）
> 本文所有语义均经 **交付文件 XML 解包 + 实机 F5 + CreateVideo 渲染** 三重验证，非文档推断。

---

## 0. 三层能力模型（先建立世界观）

```
┌─ 第 1 层：COM API（pywin32）────────── MsoAnimEffect 全量 155 个 id（0–154）：
│   普通形状可达 1–82 ∪ 86–149（含 64 条预设路径）；83/84/85 媒体专用；Timing 全参数、
│   MotionEffect.Path、Morph、交互/书签触发、图表/SmartArt 建图、效果音效、切换族      │
├─ 第 2 层：播放器时间线（OOXML）──────── 比 COM 宽：<p:tavLst> 任意采样缓动
│   （弹簧/过冲/任意一维轨迹）、<p:iterate> 自定义 tmPct、<p:bldChart> 图表分步 build、
│   Pulse/ColorPulse（emph presetID 35/36）——COM 只读或无枚举，写入靠 XML 注入        │
└─ 第 3 层：确认不可行 ──────────────── 3D 模型动画（Add3DModel 拒绝合法 GLB）、
    墨迹回放、SmartArt 逐个/逐级 build（COM 拒 + XML 打开即拒）、沿路径自动转向
```
**核心心法：COM 做不到 ≠ PowerPoint 做不到。** 第 1 层不够时，用 COM 造基础效果 → 解包注入第 2 层 XML → 三关验证（见 §6）。

## 1. 环境

- Windows + Office 16.0 桌面版；`pip install pywin32` 后 `win32com.client.gencache.EnsureDispatch("PowerPoint.Application")`。
- 探测某机器是否可用：PowerShell `New-Object -ComObject PowerPoint.Application`（WPS 的 `Kwpp.Application` 是另一套，API 子集不保证）。
- **进程卫生**：COM 崩溃会遗留 POWERPNT 进程并锁文件。收尾必须 `app.Quit()`；异常后 `taskkill /IM POWERPNT /F` + 删 `~$` 锁文件。用非只读方式 Open 前确认无残留实例。
- 准备 ffmpeg（CreateVideo 出片后的抽帧分析用）。

## 2. 最小示例（common.py）

```python
import sys; sys.path.insert(0, r"<skill目录>\scripts")   # 即本 skill 的 scripts 目录
import common as K
deck = K.Deck()                     # 交付版：不自动翻页；Deck(advance_for_video=True) 供渲染定片长
s = deck.slide()                    # 960×540pt 画布，浅底
card = s.card(100, 100, 200, 80, text="hello", size=14)
s.fade(card, T=0.5)                 # T = 绝对开始时刻（秒）
dot = s.dot(700, 300, 14)
s.path(dot, pts=[(x,y),...], T=1.2, dur=1.0, ease=(0.0, 1.0))
deck.save(r"out.pptx"); deck.render_video(r"out.mp4"); deck.close()
```

## 3. 调度模型（最重要的语义，踩过的最大坑）

- **主序列只有单个自动开始的 click 组**；组内每个效果 `<p:cond delay>` 是**相对整条时间线起点的绝对时刻**——`WithPrevious`/`AfterPrevious` 皆然。COM 写什么值就按什么值播。
- ⚠️ **历史事故**：common.py 曾按"相对上一效果的增量"写 delay，整页 11.3s 编排被压缩进前 2.4s 乱序齐发（用户实测"错位 + 元素最后才出现"）。且当时 CreateVideo QA 渲染的是单独构建的导出版、从未渲染交付文件本身，故 QA 未拦截。**教训双条：delay 一律写绝对时刻；QA 必须直连交付文件渲染。**
- 因此 `_sched` 的契约：效果必须**按开始时间升序添加**（逆序直接抛错）；`WithPrevious` + 绝对 delay 实现同刻并发；交叉淡出 = 前一效果 Exit 与后一效果 Enter 重叠 0.15–0.3s。
- 换片策略：交付版 `AdvanceOnTime=False`（进页自动播、播完停终帧、单击才翻页）；渲染版 True + `AdvanceTime = 页时长+3`，让 CreateVideo 按窗口定片长。

## 4. 效果系统

### 4.1 常用 msoAnimEffect（详见 capabilities.md 全表；枚举全量 0–154，普通形状可达 1–82 ∪ 86–149，83/84/85 为媒体专用效果）
Appear=1, FlyIn=2, Fade=10, Wipe=22, Zoom=23, FadedZoom=48, GrowShrink=59, Spin=61, 任意路径=149（Path 会被整体覆写）。83/84/85 拒绝。

### 4.2 通用封装（common.py `fx`）
```python
s.fx(shape, eff, T=绝对秒, dur=0.5, direction=DIR_L, exit_=False,
     repeat=0, accel=None, decel=None, smooth=None)
```
- `smooth=True` 时先写 SmoothStart/SmoothEnd=msoTrue，再用 Accelerate/Decelerate（0~1 浮点）覆写精确曲线——**布尔在前、数值在后，数值生效**。
- 位移类效果默认 smooth=True + (0.2, 0.5)；fade 类保持线性便于交叉淡出对拍。
- 语义包装：`wipe(shape, T, dir_="left", ease=(0.2,0.5))`（曲线描画=Wipe from left）、`fade`、`pop`（FadedZoom 柔和弹出）、`fly`、`appear`（硬切，仅脉冲圈等"骤现"叙事用）。
- **Repeat 的坑**：`Timing.RepeatCount=n` 使实际播放时长 ×n；**BounceEnd 的回弹时长不计入 Duration**——换片/排片时要另加 ~1s 余量。`AutoReverse=-1` 同理时长 ×2。
- **BounceEnd 生效条件**：只对 Float 类入场（msoAnimEffect 30 等）产生位移回弹；Fade/FadedZoom/Zoom 等原位类入场加 BounceEnd 无可见过冲（且 FadedZoom 会多 0.5s 起播延迟）——要弹跳感选 Float 类形状入场。

### 4.3 文本动画
- **按词/按字母的 COM 直达（capability3 新增，优先用这条）**：建效果后写遗留 API
  `shape.AnimationSettings.TextUnitEffect = 1 / 2`（ppAnimateByWord / ppAnimateByCharacter），
  直接在效果 cTn 生成现代 `<p:iterate type="wd|lt"><p:tmPct val="100000"/>`。实测按字母渲染
  为逐字阶跃（前沿 16 跳），整框对照为 1 跳。坑：遗留写会改写该效果的时长细节——设完
  回手复核 `Timing.Duration`。`EffectInformation.TextUnitEffect` 仍只读。
- XML 注入通道（要自定义 tmPct 时用）：`<p:iterate type="wd"><p:tmPct val="10000"/></p:iterate>`
  插入效果 cTn（stCondLst 之后、childTnLst 之前）。
- 按段落 build：COM 建效果后自动写 `<p:bldP spid grpId/>`，补 `build="p"` 属性即可。
- UI 独占强调 Pulse/ColorPulse：COM 无枚举 id；建 GrowShrink(59) 后把该效果 cTn 的
  `presetID="6"` 改 `35`/`36`（presetClass 保持 emph）即播放（35=Pulse、36=ColorPulse，
  名称对应为 UI 顺序推断）。

## 5. 运动路径

- 语法：`"M x y L x y … E"`；坐标 = **相对形状原始中心的位移，x 除以页宽、y 除以页高**（960×540 时除 960/540）。
- **绝对式**：每条路径效果都从形状原始位置出发、结束位置持久保留（放映后形状停在路径终点）。多段接续：后续路径的 `M` 写在上一条终点偏移上，即视觉无缝。
- 匀速保证：折线拐点会导致沿路径变速，用 `resample(pts, n)` 按弧长重采样等长段。
- **减速叙事**：`ease=(0.0, 1.0)`（纯 Decelerate）= 整程 ease-out。实测逐帧跟踪：前半程均速/后半程均速 = 4.25:1——"先冲后挪"的衰竭感可直接量化验证。
- **沿路径转向不可行**（无 orient-to-path）；需要时同刻叠加 Spin 人工模拟。

## 6. XML 注入层（高级缓动：弹簧/过冲/任意一维轨迹）

COM 造基础效果 → 解包 pptx → 在该效果 par 的 childTnLst 注入 → 重打包。以阻尼弹簧为例（位移沿 ppt_x）：

```xml
<p:anim calcmode="lin" valueType="num">
  <p:cBhvr additive="base">
    <p:cTn id="{全局唯一}" dur="{dur*1000}" fill="hold"/>
    <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
    <p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>
  </p:cBhvr>
  <p:tavLst>
    <p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>
    <p:tav tm="{t_i*100000}"><p:val><p:strVal val="({表达式})"/></p:val></p:tav>
    …
    <p:tav tm="100000"><p:val><p:strVal val="({终值表达式})"/></p:val></p:tav>
  </p:tavLst>
</p:anim>
```

**三个致命坑**（错一个 = 动画静默不播）：
1. `tav tm` 是**千分比**（0–100000），不是毫秒；
2. 值必须是 `<p:strVal>` 的公式串（字面量加括号 `(0.53)`、当前位置引用 `#ppt_x`）——`<p:fltVal>` 写了也不被执行；
3. `cTn id` 全局唯一，注入前扫描既有 id 取 max+1。

`attrName` 可用：`ppt_x / ppt_y / ppt_w / ppt_h / ppt_r / style.color` 等。表达式语言（fbml）支持 `$`（0-1 进度）、`#ppt_x` 引用、算术、`min/max/abs/sqrt/if` 等函数——`sqrt($)` 即 ease-out。

**注入三关验证**（任一失败回滚）：① `Presentations.Open` 不触发修复；② COM 读回 MainSequence 计数正常；③ CreateVideo 出片且逐帧轨迹符合设计（如弹簧 8 次过零、幅值衰减）。

## 7. Morph 跨页变形

- `slide.SlideShowTransition.EntryEffect = 3954/3955/3956`（Morph ByObject/Word/Char，本机 typelib 实测值）；落盘 `<p159:morph option="byObject"/>` 自带 fallback。
- 跨页配对：形状**同名**自动匹配；强制配对用 `"!!名字"` 前缀。
- Morph 时长由 `p14:dur` 控制（默认 2000ms），换片时间另计——排片按两页动画 + Morph 时长通盘算。
- CreateVideo 能渲染出连续中间形变（已验证卡片跨页滑移+变色），所以 Morph 页同样可无头 QA。

### 7.1 幻灯片切换族（capability3 新增：Morph 之外全部 COM 直达）

```python
tr = slide.SlideShowTransition
tr.EntryEffect = 3855        # push 上推；经典：wipe 2817–2820、cover 1281–1288、fade 3849、cut 257…
tr.EntryEffect = 3867        # ripple 居中（2010+）：reveal 3894/3895、vortex 3863–3866、
                             # glitter 3872–3879、honeycomb 3898、cube 3914–3917、morph 3954–3956
tr.Duration = 1.25           # 现代切换写 p14:dur 毫秒；经典切换被量化为 spd（0.75s→"med"）
tr.SoundEffect.ImportFromFile(r"x.wav")   # 切换声音：落盘 ✅，CreateVideo 不渲染（实机验证）
```
PpEntryEffect 全集 192 个成员见探针产物 `enum_dump.json`。落盘形态：经典写 `<p:push dir="u"/>`、
现代写 `<p14:reveal/>`/`<p14:ripple/>`（COM 自动带 p14 命名空间，无需手工 mc:AlternateContent）。
CreateVideo 渲染换片过程，帧差尖峰可断言；切换占用窗口用 `qa_motion.render(morph_extra=...)` 另计。

## 8. 交互触发（问答/互动页用）

```python
seq = slide.TimeLine.InteractiveSequences.Add(1)
seq.AddTriggerEffect(shape, msoAnimEffectFade, 4, trigger_shape)  # 4 = msoAnimTriggerOnPageClick 语境下的对象点击
```
落盘 `evt="onClick"` 交互序列。注意 CreateVideo **不会**触发点击——交互效果只能实机放映验证。

### 8.1 媒体书签触发（capability3 翻案：COM 全链路可达）

```python
mb = video_shape.MediaFormat.MediaBookmarks
mb.Add(1.5, "b1")                                  # 签名 = Add(Position秒, Name)，顺序勿反
seq = slide.TimeLine.InteractiveSequences.Add(1)
seq.AddTriggerEffect(shape, msoAnimEffectAppear, 5, video_shape, "b1")   # 5 = OnMediaBookmark
```
书签落 `<p14:bmk name="b1" time="1"/>`（Position 会被取整到秒），触发效果落 `evt="onMediaBookmark"` 交互序列，重开不修复。CreateVideo 同样**不会**触发书签——实机验证。

## 9. 布局与 z 序的坑（全部实测）

- **Group 的 z 序插入语义**：`Shapes.Range(...).Group()` 把组插到**组内最高成员**的 z 位。组外后建的形状会盖住组内元素；且 pywin32 下组内成员 ZOrder 无法越过组外形状。**修法（唯一可靠）**：先分组，后创建需要置顶的元素（创建序即 z 序）。
- **文本框自动扩高**：设置文本后 PowerPoint 自动扩高——12pt 文本 nominal 16pt 框 → 实高 ~29.1pt（垂直居中，向两侧各溢 ~6.5pt）。布局断言必须用**双层**：标称框（设计格点）+ 有效字形带（中心 ±8.5pt@12-13pt、±17pt@28pt），两层都查相交。
- 其余：`AutoSize=None` + Margin 全 0（common.py 已封装）；自由曲线 `BuildFreeform` 只支持直线段节点（画曲线用密集采样点）；图片插入用 `AddPicture` 等比缩放防拉伸。

## 10. QA 方法学（怎么证明动画真的在播、播对了）

1. **必须渲染交付文件本身**——渲染单独构建的"导出版"是最典型的 QA 失误（导出版时序正确，掩盖了交付版调度 bug，曾致事故漏检）。
2. `CreateVideo(mp4, UseTimingsAndNarrations=True, fps)` + 轮询 `CreateVideoStatus`（3=done）；渲染副本只改换片时间、动画时间线不动。
3. ffmpeg/imageio 抽帧 → 两类分析：
   - **运动验证**：频带内质心逐帧跟踪 → 位置-时间序列 → 速率比（如减速叙事要求前后半程 ≥1.6，实测 4.25 通过）；
   - **终帧验证**：终帧 vs 设计布局叠加，关键元素齐全、数值正确。
4. **16 帧密距蒙太奇**（帧号+页内秒数标签在画面下方条带）→ 交 visual-judge 子代理做"评委视角"审查：视觉叠压 + 教学清晰度（每页问：看完能否说出这是什么改进、为什么有效；哪些术语会卡壳；因果链哪一步跳跃最大）。
5. **断言内建在 build 脚本里**（assert 直接抛错中断出片）：数字口径（步数、指标值、占比等，按项目自行声明）、几何断言（bbox 双层不相交、路径终点偏差、距离 ≥N px）。

## 11. 性能边界（实测）

单页形状 100/300/600（效果 200/600/1200）：构建 2.8/9.3/22.4s；**打开恒 ~0.04s**；CreateVideo 3–10/12–48/43–98s。**建议：单页 ≤300 效果为舒适区，≤600 可用**；实测真实编排单页最多用到 82 个效果，余量充足。

## 12. 确认不可行清单

3D 模型动画（`Add3DModel` 在 OM 中存在但拒绝合法 glTF GLB——16.0 无 3D 支持；效果 id 151–154 对普通形状"接受"但为静默空操作）；墨迹手写回放（无 COM 入口）；**SmartArt 逐个/逐级 build 动画**（`AddEffect` Level=12–26 全拒；XML `<a:bldDiagram>` 补丁三种规范形态打开即硬失败——SmartArt 只有整体动画 bldAsOne 可用）；**内置音效名库**（`SoundEffect.Name="applause"` 拒绝，只能 ImportFromFile 导 wav）；沿路径自动转向（无 orient-to-path 概念）。~~视频书签触发~~已翻案为可行（见 §8.1）。

## 13. 文件索引与 QA 工具范式

本 skill 收录：

| 文件 | 职责 |
|---|---|
| `../scripts/common.py` | 构建框架：Slide/Deck 封装、`_sched` 绝对调度、fx/wipe/fade/pop/fly/path |
| `capabilities.md` | 能力矩阵（A–G 区原 87 预设探针 + H/I 区 capability3 补录 × COM 写法 × 验证状态 × 坑） |

以下三套 QA 工具未随 skill 收录脚本（都很薄，新项目按范式重建即可）：
- **渲染 QA**：复制交付 pptx → 只改换片时间（`AdvanceOnTime=True` + `AdvanceTime=ceil(页时长)+2`，保证终帧 ≥2s 阅读）→ `CreateVideo` 30fps → ffmpeg 抽帧 → 16 帧蒙太奇 + 终帧保持核对；
- **COM 读回审计**：只读 Open 交付文件读回 MainSequence——效果数=构建声明、逐效果时刻表、OnPageClick 残留=0、同形状多段路径接续 M 一致性、路径终点世界坐标断言；
- **XML 时序复核**：解包 pptx 比对——clickEffect 节点=0、click 组唯一且带 onBegin 自动开始条件、每个效果 par 的有效开始时刻与设计绝对时刻差 ≤1ms。

## 14. 口径红线（通用实践）

把交付口径（关键数字、比例、指标值）写成构建脚本内的 assert，任何后续改动破坏口径即中断出片——数字红线不靠人眼复核，靠构建期强制。

## 15. 图表 / SmartArt / 媒体 / 声音 / 动画刷（capability3 新增）

**图表动画**：`Shapes.AddChart2(-1, 51, ...)`（51=柱状簇）；`AddEffect(chart, eff, 0, trigger)` 自动写
`<p:bldGraphic spid grpId><p:bldAsOne/>`（整体动画）。按系列/类别：COM `AddEffect` Level 参数（7–11）
与遗留 `ChartUnitEffect` 全部拒绝，唯一通道 = XML 补丁 `<p:bldAsOne/>` →
`<p:bldSub><a:bldChart bld="series|category" animBg="0"/></p:bldSub>`——重开/另存稳定保持，
但 CreateVideo 渲染与整体无差别，分步效果仅实机可验（无头 QA 视为不可证）。
⚠️ `AddChart2` 每次拉起 EXCEL.EXE 数据网格且**同会话第二次调用被"数据网格已打开"阻塞**：
每演示文稿只加一图；异常收尾除 POWERPNT 外还要 `taskkill /F /IM EXCEL.EXE`
（标题含"Microsoft PowerPoint 中的图表"的实例）。

**SmartArt 动画**：`dyn = win32com.client.Dispatch("PowerPoint.Application")` 后
`Shapes.AddSmartArt(dyn.SmartArtLayouts.Item(i), ...)`（SmartArtLayouts 需动态绑定获取）；
节点填字 `shape.SmartArt.Nodes.Item(i).TextFrame2.TextRange.Text = ...`；整体动画同图表
（bldAsOne 自动写）。逐个/逐级不可行（见 §12）。

**媒体动画**：`Shapes.AddMediaObject2(file, LinkToFile=msoFalse, SaveWithDocument=msoTrue, x, y, w, h)`。
播放控制三效果（媒体形状专用，普通形状拒绝）：`AddEffect(video, 83/84/85, 0, 2)` =
MediaPlay/Pause/Stop。自动播放：`AnimationSettings.PlaySettings.PlayOnEntry = msoTrue`。
媒体属性 `shape.MediaFormat`：Volume/TrimStart/TrimEnd/FadeIn/OutDuration/Muted 可写。
CreateVideo 渲染内嵌媒体：画面会动、媒体音频进 mp4 音轨——无头管线可验证播放。
书签触发见 §8.1。

**效果声音**：`effect.EffectInformation.SoundEffect.ImportFromFile(r"x.wav")`
（方法名是 **ImportFromFile**，没有 Import）；wav 嵌入 ppt/media/，Type=2（ppSoundFile）。
**CreateVideo 渲染声音**：实测效果时间窗 mean −28dB / max −24dB（静音段 −91dB），
可用 `ffmpeg volumedetect` 断言。内置音效名（Name="applause"）拒绝——只能导入文件。

**动画刷 COM 等价**：`src.PickupAnimation()` → `dst.ApplyAnimation()`（注意拼写 Pickup），
含路径效果整体复制，落主序列。预设运动路径家族（msoAnimEffect 86–149 共 64 条）COM 直达，
`Behaviors(1).MotionEffect.Path` 可读出贝塞尔串、可覆写（见 capabilities.md H6）。

## 16. QA 管线新口径（capability3）

- CreateVideo 默认竖向分辨率 720 → 16:9 页出 **1280×720** 帧；像素坐标按
  `k = frame_width / 960` 换算（勿写死 0.75）。
- 触发语义再确认：全部效果用 trigger=WithPrevious(2) + 绝对 `TriggerDelayTime`（common.py
  口径）；若写 trigger=AfterPrevious(3)，PowerPoint 会在效果 par 外再包一层"累计前置时长"
  的 wrapper delay，实际开始 = 累计 + 你写的 delay——capability3 探针首版即因此整页错位，
  再次印证铁律 1。
- 区域帧差断言模板：给目标形状独占区域 → 抽帧裁剪 → 逐帧灰度差峰值/活动帧窗断言
  （强调/切换/媒体/文本测试均用此法）。
- 探针脚本与机器可读结果（能力复测入口）：项目内
  `ppt/assets/make_native_anim/_qa/capability3/`——`enum_dump.json`（MsoAnimEffect 0–154、
  PpEntryEffect 192 成员等全枚举）、`preset_map.json`（效果 id→presetClass/presetID 全映射）、
  `t03–t11_*_result.json`（分项实测）。
