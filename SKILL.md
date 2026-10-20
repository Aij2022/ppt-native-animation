---
name: ppt-native-animation
description: PowerPoint 原生动画 COM 自动化构建与 QA 方法论。当用户要求给 PPT/幻灯片/.pptx/演示文稿加动画、做原生动画、动画演示、形状动画、路径动画、放映动画，询问 PowerPoint 动画能做什么/怎么做/支持哪些动画效果，或要在 Windows Office/WPS 环境用脚本自动化生成带动画的演示文稿时使用。覆盖对象动画（进入/强调/退出/路径/文本逐字逐词）、图表与 SmartArt 动画、媒体播放与书签触发、音效、幻灯片切换族（含 Morph）；也适用于排查脚本生成的动画**放映错位/元素不显示/时序异常**等问题（见 §8 QA 闭环与审计范式）。Windows + 桌面版 PowerPoint 环境专用；Mac/Keynote 与纯 GIF 方案见文内降级指引。
---

# PowerPoint 原生动画 COM 自动化

用 pywin32 驱动桌面版 PowerPoint，构建"进页自动播放"的原生动画演示文稿。所有语义经交付文件 XML 解包 + 实机放映 + CreateVideo 渲染三重验证。本文只列铁律与速查；细节一律查 `references/guide.md`（下称"指南"，按节引用）与 `references/capabilities.md`（下称"能力矩阵"）。

## 1. 适用判断

- 职责边界：从零建 deck / 版式排版走 presentations:pptx（其构建栈不支持动画）；**本 skill 专管给页面加原生动画**。
- 命中条件：Windows + 桌面版 PowerPoint（Office 16.0 实测）+ `pip install pywin32`。先探测：PowerShell `New-Object -ComObject PowerPoint.Application`。WPS（`Kwpp.Application`）是另一套 API 子集，不保证兼容，勿默认可用。
- 无 PowerPoint 环境时降级，按优先级：
  1. 动画可预渲染 → 用 ffmpeg/代码生成 GIF 或 MP4，作为静态媒体嵌入幻灯片；
  2. 必须原生动画 → 直接改写 .pptx 的 OOXML `<p:timing>` 时间树，但没有 PowerPoint 校验与渲染闭环（本 skill 的 QA 手段全部失效），须向用户明示风险；
  3. Mac/Keynote：本 skill 不适用，走 Keynote AppleScript 或手写 OOXML。另注：python-pptx 官方确认无动画 API 且无计划（GitHub issue scanny/python-pptx#1106），勿试图走 python-pptx 做原生动画。
- 进程卫生：COM 崩溃会遗留 POWERPNT 进程并锁文件。收尾必须 `app.Quit()`；异常后 `taskkill /IM POWERPNT /F` 并删 `~$` 锁文件；非只读 Open 前确认无残留实例；用 AddChart2 的场合连 EXCEL.EXE 数据网格一起清（见 §9）。（指南 §1/§15）

## 2. 快速开始

框架在 `scripts/common.py`（画布 960×540pt，形状工厂 card/text/rect/dot/ring/line/freeform/picture/group/chip/pill，效果封装 fx/wipe/fade/pop/fly/appear/path；源码内注释即语义规范）。把本 skill 的 scripts 目录加入 sys.path 后：

```python
import sys; sys.path.insert(0, r"<本skill目录>\scripts")
import common as K
deck = K.Deck()                     # 交付版：不自动翻页；Deck(advance_for_video=True) 供渲染定片长
s = deck.slide()                    # 960×540pt 画布，浅底
card = s.card(100, 100, 200, 80, text="hello", size=14)
dot = s.dot(700, 300, 14)
s.fade(card, T=0.5)                 # T = 绝对开始时刻（秒）
s.path(dot, pts=[(700, 300), (500, 300), (400, 200)], T=1.2, dur=1.0, ease=(0.0, 1.0))
deck.save(r"out.pptx"); deck.render_video(r"out.mp4"); deck.close()
```

## 3. 调度铁律（每条带原因）

1. **delay 一律写绝对时刻 T（秒），禁止相对增量。** 原因：主序列是单个自动开始的 click 组，组内每个效果 `<p:cond delay>` 是相对整条时间线起点的绝对时刻，WithPrevious/AfterPrevious 皆然；按增量写会把整页编排压缩进前几秒乱序齐发（来源项目头号事故，整页 11.3s 被压进 2.4s）。（指南 §3）
2. **效果按开始时间升序添加。** 原因：时间线组装契约，`_sched` 对逆序直接抛错；同刻并发用 WithPrevious + 相同绝对 T；交叉淡出 = 前一效果 Exit 与后一效果 Enter 重叠 0.15–0.3s。
3. **交付版关闭自动翻页**（Deck 默认 AdvanceOnTime=False）：进页自动播、播完停终帧、单击才翻页。渲染版才开 True + AdvanceTime=页时长+3，让 CreateVideo 按窗口定片长。
4. **QA 必须渲染交付文件本身，而非单独构建的"导出副本"。** 原因：导出版时序正确会掩盖交付版调度 bug——上述头号事故正是因此漏检。渲染副本只许改换片时间，动画时间线不动。
5. **复杂场景先写页级编排表**（效果/时刻/时长/缓动/触发一张表），人工过目后再翻译成构建代码——先表后码可评审、可复用、易回滚。
6. **时长记账三坑**：RepeatCount=n 实际播放 ×n；AutoReverse=-1 时长 ×2；BounceEnd 回弹不计入 Duration（另加 ~1s 换片余量）。（指南 §4.2）

## 4. 效果与缓动速查

- 常用 msoAnimEffect：Appear=1、FlyIn=2、Fade=10、Wipe=22、Zoom=23、Bounce=26、FadedZoom=48、GrowShrink=59、Spin=61、Transparency=62、ChangeFontColor=56/54、任意路径=149（Path 会被整体覆写）。枚举全量 0–154（capability3 typelib 实测）：普通形状可达 **1–82 ∪ 86–149**（86–149=64 条预设路径，`MotionEffect.Path` 可读可覆写）；**83/84/85=MediaPlay/Pause/Stop，仅媒体形状可用**（普通形状拒绝）；150 需书签；151–154=3D（本机不可用）。（能力矩阵 A1/H6/H9）
- 缓动叠加规则：`smooth=True` 先写 SmoothStart/SmoothEnd=msoTrue，再用 Accelerate/Decelerate（0~1 浮点）覆写精确曲线——**布尔在前、数值在后，数值生效**。位移类（wipe/fly/pop/path）默认 smooth=True+(0.2,0.5)；fade 类保持线性（accel=decel=0）便于交叉淡出对拍。
- 语义封装：`wipe(shape, T, dir_="left")`（曲线描画=Wipe from left）、`pop`=FadedZoom 柔和弹出、`fly(from_=...)`、`appear`（硬切，仅"骤现"叙事用）、`path`（见 §5）。
- 颜色渐变 emphasis 走 `Behaviors(1).ColorEffect.From/To.RGB`（EffectParameters.Color1 在 pywin32 下不可用）；**按词/按字母走遗留 API 直达：`shape.AnimationSettings.TextUnitEffect = 1/2`（写后复核 Duration）**，XML 注入仅在要自定义 tmPct 时用；按段落 build 见指南 §4.3。
- 幻灯片切换族：`SlideShowTransition.EntryEffect` 直达全部 192 个 PpEntryEffect 成员——经典（push/wipe/cover…）+ 2010+ 现代（ripple/reveal/vortex…）+ Morph=3954/3955/3956（ByObject/Word/Char）；`Duration` 现代
  切换写 p14:dur 毫秒、经典量化 spd；`SoundEffect.ImportFromFile(wav)` 可挂切换/效果音效。（指南 §7.1/§15）
- Morph 跨页变形：同名形状自动配对，`"!!名字"` 前缀强制配对；时长 `p14:dur` 默认 2000ms，换片排片时另计。（指南 §7）
- 点击触发：`TimeLine.InteractiveSequences.Add(1)` + `AddTriggerEffect(shape, eff, 4, trigger_shape)`；**媒体书签触发 = 先 `MediaFormat.MediaBookmarks.Add(秒, "b1")` 再 `AddTriggerEffect(shape, eff, 5, video, "b1")`**（参数顺序勿反）。CreateVideo **不会**触发点击/书签，只能实机放映验证。（指南 §8/§8.1）
- 图表/SmartArt：`AddChart2`/`AddSmartArt` 建图 + 整体动画 COM 直达；图表按系列/类别仅 XML 补丁（bldChart，渲染不受 CreateVideo 区分）；SmartArt 逐级不可行。图表/SmartArt/媒体/声音/动画刷细节一律查指南 §15。

## 5. 运动路径

- 语法 `"M x y L x y … E"`（M/L/C 贝塞尔/Z 均支持）；坐标 = **相对形状原始中心的位移，x 除以页宽、y 除以页高**（960×540 时除 960/540），C 控制点同相对坐标。（指南 §5、能力矩阵 B1）
- **绝对式语义**：每条路径都从形状原始位置出发，终点持久保留（放映后停在终点）。多段接续：后续路径的 `M` 写在上一条终点偏移上——common.py 的 `path()` 直接把上一条终点作为 pts 首点即可。
- 匀速保证：折线拐点导致沿路径变速，用 `resample(pts, n)` 按弧长重采样等长段（common.py 已带）。
- **减速叙事**：`ease=(0.0, 1.0)`（纯 Decelerate）= 整程 ease-out；实测前后半程速率比 4.25:1，可被质心跟踪量化验证。
- 沿路径转向不可行（无 orient-to-path）；需要时同刻叠加 Spin 人工模拟。

## 6. XML 注入三致命坑

流程：COM 造基础效果 → 解包 pptx → 在该效果 par 的 childTnLst 注入 `<p:anim>`（tavLst 采样）→ 重打包。用于弹簧/过冲/任意一维轨迹与按词动画。完整 XML 模板见指南 §6。

三个致命坑（错一个 = 动画静默不播）：
1. `tav tm` 是**千分比**（0–100000），不是毫秒；
2. 值必须是 `<p:strVal>` 公式串（字面量加括号 `(0.53)`、引用 `#ppt_x`）——`<p:fltVal>` 写了也不执行；
3. `cTn id` 全局唯一，注入前扫描既有 id 取 max+1。

表达式语言（fbml）支持 `$`（0–1 进度）、`#ppt_x` 引用、算术、min/max/abs/sqrt/if；`sqrt($)` 即 ease-out。attrName 可用 ppt_x/ppt_y/ppt_w/ppt_h/ppt_r/style.color。

**注入三关验证**（任一失败回滚该注入）：① `Presentations.Open` 不触发修复；② COM 读回 MainSequence 计数正常；③ CreateVideo 出片且逐帧轨迹符合设计（如弹簧 8 次过零、幅值衰减）。

现成注入器：`scripts/inject_easing.py`（spring/自定义 tav，内置三致命坑处理与三关验证）。

## 7. 布局陷阱

- **Group z 序插入语义**：`Shapes.Range(...).Group()` 把组插到组内最高成员的 z 位；组外后建的形状会盖住组内元素，且 pywin32 下组内成员 ZOrder 无法越过组外形状。**唯一可靠修法：先分组，后创建需要置顶的元素**（创建序即 z 序）。
- **文本框自动扩高**：设置文本后 PowerPoint 自动扩高——12pt 文本 nominal 16pt 框 → 实高 ~29.1pt（垂直居中，两侧各溢 ~6.5pt）。布局断言必须**双层**：标称框（设计格点）+ 有效字形带（中心 ±8.5pt@12–13pt、±17pt@28pt），两层都查相交。
- 其余：`AutoSize=None` + Margin 全 0（common.py 已封装）；`BuildFreeform` 只支持直线段节点（曲线用密集采样点）；`AddPicture` 等比缩放防拉伸。（指南 §9）

## 8. QA 闭环（怎么证明动画在播、播对了）

1. **渲染交付文件本身**：复制交付 pptx 为渲染副本，只改换片时间（`AdvanceOnTime=True` + `AdvanceTime=ceil(页时长)+2`，保证终帧 ≥2s 阅读窗口），动画时间线不动。
2. `CreateVideo(mp4, UseTimingsAndNarrations=True, fps)` + 轮询 `CreateVideoStatus`（3=done，0/1/2=进行中，超时抛错）。
3. ffmpeg 抽帧 → 两类分析：**运动验证**（频带内质心逐帧跟踪 → 位置-时间序列 → 速率比断言，如减速叙事前后半程 ≥1.6）；**终帧验证**（终帧 vs 设计布局叠加，关键元素齐全、数值正确）。
4. **16 帧密距蒙太奇**（帧号+页内秒数标签画在画面下方条带，不遮挡内容）→ 交 visual-judge 子代理做"评委视角"教学审查：每页问——看完能否说出这是什么改进、为什么有效；哪些术语会卡壳；因果链哪一步跳跃最大。
   以上渲染/抽帧/蒙太奇/质心跟踪已工具化为 `scripts/qa_motion.py`。
5. **断言内建在构建脚本**（assert 直接抛错中断出片）：数字口径（按项目声明）+ 几何断言（bbox 双层不相交、路径终点偏差、最小间距）。
6. **渲染前秒级预检**：解包比对 XML 时间树与请求效果（效果数、逐效果绝对时刻），在昂贵渲染前拦截调度类事故。
7. 另两道审计范式（薄脚本，按需重建）：**COM 读回审计**（效果数=构建声明、逐效果时刻表、OnPageClick=0、同形状多段路径接续 M 一致、终点世界坐标）；**XML 复核**（解包比对 clickEffect=0、click 组唯一且带 onBegin 自动开始、有效开始时刻与设计值差 ≤1ms）。（指南 §10、§13）

## 9. 性能与不可行

- 单页效果数实测：100/300/600 形状（200/600/1200 效果）→ 构建 2.8/9.3/22.4s，打开恒 ~0.04s，CreateVideo 3–10/12–48/43–98s。**建议：单页 ≤300 效果为舒适区，≤600 可用**；>600 构建时间陡增且渲染收益低。
- 进程卫生补充（capability3）：`AddChart2` 会拉起 EXCEL.EXE 数据网格且同会话第二次调用被"数据网格已打开"阻塞——异常收尾除 POWERPNT 外还要 `taskkill /F /IM EXCEL.EXE`（标题含"Microsoft PowerPoint 中的图表"）；CreateVideo 默认出 **1280×720** 帧，帧分析坐标按 `帧宽/960` 换算。
- 不可行清单（勿尝试）：3D 模型动画（Add3DModel 拒绝合法 GLB）、墨迹手写回放（无 COM 入口）、**SmartArt 逐个/逐级 build 动画**（COM Level 全拒 + bldDiagram XML 打开即拒；只有整体动画）、**内置音效名库**（SoundEffect.Name="applause" 拒绝，只能 ImportFromFile 导 wav）、沿路径自动转向（无 orient-to-path）。~~视频书签触发~~已翻案可行：`MediaBookmarks.Add(秒,名)` + `AddTriggerEffect(...,5,video,"b1")` 全链路落盘（指南 §8.1）。
- QA 边界（capability3）：CreateVideo 渲染**效果音效**（音频进 mp4，volumedetect 可断言）与**内嵌媒体播放**（画面+音频），但不渲染**切换声音**、不触发**点击/书签**交互、不区分图表 **bldChart 分步 build**——这些只能实机验证。（指南 §16）

## 10. 深入阅读

- `references/guide.md`：§0 三层能力模型 / §1 环境与进程卫生 / §3 调度语义与事故复盘 / §4 效果系统与文本动画 / §5 运动路径 / §6 XML 注入 / §7 Morph 与 §7.1 切换族 / §8 交互触发与 §8.1 书签触发 / §9 布局与 z 序 / §10 QA 方法学 / §11 性能 / §13 QA 工具范式 / §15 图表·SmartArt·媒体·声音·动画刷 / §16 QA 新口径。
- `references/capabilities.md`：能力矩阵——A 效果广度 / B 路径与节奏 / C XML 注入 / D Morph / E 交互 / F 性能 / G 不可行（原 87 预设探针），H 区 17 行 capability3 新增实测（切换族/按字母/Pulse/预设路径/图表 build/SmartArt/媒体 83-85/书签/音效/动画刷）+ I 区新确认不可行，每行含 COM 写法、验证状态、关键坑。
- `scripts/common.py`：全部封装源码，Docstring 与 `_sched` 注释即调度语义规范。
- 能力复测探针：项目内 `ppt/assets/make_native_anim/_qa/capability3/`（枚举全量 dump、preset 全映射、t03–t11 分项实测脚本与结果 JSON）。
