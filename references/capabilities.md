# PowerPoint 原生动画能力边界报告（本机 Office 16.0 + pywin32）

> 注：本文件全文收录自来源项目（某 VLA 答辩项目）的实测能力探针报告。下文出现的探针/演示 deck/蒙太奇等路径均为来源项目内路径，仅作示例，本 skill 未收录这些文件；能力矩阵本身与机器无关，可直接参照。
>
> **capability3 全集审计补录（2026-10-20）**：以本机 typelib 全量枚举为准做了二次审计与补测，新增 H 区（切换族/文本按字母/图表 build/SmartArt/媒体/声音/动画刷等）与 I 区（新确认不可行），并就地修订 A1/A2/G2 三行（原文保留、以〔capability3 修订〕标注）。探针脚本与机器可读结果：某项目 `ppt/assets/make_native_anim/_qa/capability3/`（`enum_dump.json`＝MsoAnimEffect/PpEntryEffect 等全枚举、`preset_map.json`＝1–154 每个效果 id→presetClass/presetID 映射、`t03–t11_*_result.json`＝分项实测）。

- 探针：`make_native_anim/capability_probe.py`（可重跑；`--full` 跑满 600 效果、`--only C,D` 选探区）
- 机器可读结果：`make_native_anim/_qa/capability/probe_results.json`
- 演示 deck：`ppt/assets/_anim_capability_showcase.pptx`（10 技巧/11 页，进页自动播放）
- 演示蒙太奇：`make_native_anim/_qa/capability/cap_mont_p{1..10}.png`，全片 `full.mp4`
- 验证闭环：COM 写入 → 保存 → （必要时 zip 注入）→ `Presentations.Open` 不触发修复 → `CreateVideo` 30fps → ffmpeg 抽帧 → 质心轨迹/暗像素分析

**一句话总结：除"沿路径转向、3D 模型、墨迹回放"等 365 专属能力外，PowerPoint 播放器的时间线远比 COM API 宽——COM 直达 87 种 preset + BounceEnd/AutoReverse/Repeat/Morph/点击触发，且 `<p:iterate>` 按词与 `<p:tavLst>` 任意采样缓动（弹簧/弹跳/过冲）可经 XML 注入稳定播放，等于"任意一维物理轨迹"都做得到。**

## 能力矩阵（验证状态三档：✅ COM/播放实测通过 · ⚠️ 仅文档推断未验证 · ❌ 确认不可行）

| # | 能力 | COM 写法 | 状态 | 关键坑 |
|---|------|----------|------|--------|
| A1 | 效果 preset 广度 | `MainSequence.AddEffect(shape, id, 0, trigger)`，id=msoAnimEffect | ✅ 1–90 中 **87 个可用**（仅 83/84/85 拒绝）；含 Fade/Zoom/Bounce(26)/Split/Wipe/RiseUp/Swing/Teeter/Wave 等 | 部分 preset 仅对特定形状类型有意义。〔capability3 修订〕枚举全量实为 **0–154**：1–82 进/强/退、**83/84/85=MediaPlay/Pause/Stop（媒体专用：普通形状拒绝、媒体形状可用）**、**86–149=64 条预设路径（普通形状全部可用，见 B7）**、150=书签播放（需书签）、151–154=3D（需 365 3D 模型）。普通形状 COM 可达面 = 1–82 ∪ 86–149 共 146 条；落盘 presetClass 实测：entr 53 / emph 33（含 152–154 映射 emph 37–39）/ path 64。 |
| A2 | 按词/按字母文字动画 | `EffectInformation.TextUnitEffect` **只读**；写 = XML 注入 `<p:iterate type="wd"><p:tmPct val="10000"/></p:iterate>` 插到效果 cTn 内（stCondLst 之后、childTnLst 之前） | ✅ 注入后播放验证（暗像素带 138→573px 逐词扩张） | 必须针对文本形状的效果 cTn；id 全局唯一。〔capability3 修订〕**有更简的 COM 直达：遗留 `shape.AnimationSettings.TextUnitEffect = 1/2`（by word/by character）可写**，直接生成现代 `<p:iterate type="wd|lt"><p:tmPct val="100000"/>`（见 H3），XML 注入仅在需要自定义 tmPct 时保留 |
| A3 | 按段落分级 build | 单效果 + XML `<p:bldLst><p:bldP spid grpId build="p"/>` | ✅ showcase P2 播放验证 | COM 建效果后自动写 `<p:bldP spid grpId/>`，补 `build="p"` 属性即可 |
| A4 | 颜色渐变 emphasis（红→绿） | `AddEffect(shp, 56/54, 0, 2)` + `Behaviors(1).ColorEffect.From/To.RGB` | ✅ showcase P3 | `EffectParameters.Color1` 在 pywin32 下不可用，走 Behaviors |
| A5 | 透明度 emphasis | `AddEffect(shp, 62, 0, 2)`（Transparency） | ✅ 可加（透明度数值经 EffectParameters/Behavior 视 preset 而定） | 数值调节的 COM 口径不稳定，建议只做"渐隐层" |
| A6 | 陀螺旋+放大组合 | 同形状两条 emphasis 同刻起：`fx(shp,61,…)`+`fx(shp,59,…)`+`ScaleEffect.ByX/By` | ✅ showcase P4 | 无 |
| B1 | 自定义路径语法 | `Behaviors(1).MotionEffect.Path = "M 0 0 C … E"` | ✅ **M/L/C（贝塞尔）/Z（闭合）+ 相对坐标**全部接受（坐标=相对原始中心，x/页宽、y/页高） | C 控制点也是相对坐标 |
| B2 | autoReverse 去而复返 | `Timing.AutoReverse = -1` | ✅ 播放验证（时长翻倍，质心往返） | 排片时按 2×时长预留 |
| B3 | repeat + smooth | `Timing.RepeatCount = n`（可与小数？整数为稳）+ SmoothStart/End | ✅ showcase P6 | Repeat 使**实际播放时长=×n**，换片时间要按倍数预留 |
| B4 | 沿路径转向 | —— | ❌ 无 orient-to-path 概念；只能同刻叠加 `Spin`/`animRot` 人工模拟 | |
| B5 | 原生落地弹跳 | `Timing.BounceEnd = -1`；`Timing.BounceEndIntensity = 0~1` | ✅ COM 可写+播放验证：0.4 → 单次小回弹（~10px），0.9 → 5 次回弹阻尼弹跳 | Bounce 时长不被 `prev_end` 记账，换片时间需另加 ~1s；**只对 Float 类入场（msoAnimEffect 30 等）产生位移回弹，Fade/FadedZoom/Zoom 等原位类入场无可见过冲**（实测：Zoom 高度单调无 >8px 过冲，FadedZoom+BounceEnd 仅多 0.5s 起播延迟）——要弹跳感选 Float 类目标 |
| B6 | 蜂群（多形状同径错峰） | 同一路径字符串 + `TriggerDelayTime` 错峰 | ✅ showcase P6 | 无 |
| C1 | **XML 注入自定义缓动（弹簧/过冲/任意一维轨迹）** | COM 造基础效果 → 解包 → 效果 par 的 childTnLst 里注入 `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn dur fill="hold"/><p:tgtEl><p:spTgt spid/><p:attrNameLst><p:attrName>ppt_x</…><p:tavLst><p:tav tm="千分比0-100000"><p:val><p:strVal val="(0.53)"/#ppt_x 公式/>` → 重打包 | ✅ **打开不触发修复**；CreateVideo 质心轨迹 8 次过零、幅值衰减 92%（阻尼弹簧） | ① `tm` 是**千分比**不是毫秒；② 值必须 `<p:strVal>` 公式串（字面量加括号 `(0.53)`、引用 `#ppt_x`），`<p:fltVal>` **无效**；③ attrName 可用 `ppt_x/ppt_y/ppt_w/ppt_h/ppt_r/style.color` 等；④ 注入后 COM MainSequence 仍计数（可在 COM 侧核对页效果数） |
| C2 | 打开不触发修复判据 | `Presentations.Open(injected)` 成功 + MainSequence 计数正常 + CreateVideo 出片 | ✅ 三关全过 | 任一关失败 = 回滚该注入 |
| D1 | Morph 换片 | `SlideShowTransition.EntryEffect = 3954/3955/3956`（MorphByObject/Word/Char，本机 typelib 实测值） | ✅ COM 写入 → 落盘 `<p159:morph option="byObject"/>`（自带 fallback）→ **CreateVideo 渲染出连续中间形变**（卡片质心 230→720pt 滑移+尺寸/颜色混合） | 跨页形状须同名；普通名按"复制匹配"启发式；用 `"!!名字"` 前缀强制配对 |
| D2 | Morph 进页自动播放 | Morph 换片 + `AdvanceOnTime` 组合 | ✅ showcase P9/P10 | Morph 时长由 `p14:dur`（默认 2000ms）控制，换片时间另计 |
| E1 | 单对象点击触发 | `TimeLine.InteractiveSequences.Add(1)` + `seq.AddTriggerEffect(shape, eff, 4, triggerShape)` | ✅ COM 可写，XML 落 `evt="onClick"` 交互序列 | CreateVideo **不会**触发点击——只对实机放映问答有效 |
| E2 | 超链接/跳页 | `Shape.ActionSettings(ppMouseClick).Hyperlink.Address / Action=ppAction*` | ✅ 可写，`<a:hlinkClick>` 落盘 | |
| F1 | 单页效果数性能 | — | ✅ 实测 100/300/600 形状（200/600/1200 效果）：构建 2.8/9.3/22.4s；**打开恒 ~0.04s**；CreateVideo 3–10/12–48/43–98s（两次测量区间，受机器负载影响） | 建议：单页 **≤300 效果**为舒适区；≤600 可用；>600 构建时间陡增且渲染收益低 |
| G1 | 3D 模型动画 | —— | ❌ 需 Microsoft 365 模型插入，本机 16.0 无对象模型 | |
| G2 | 视频书签触发 | `Timing.TriggerBookmark` 属性存在 | ✅ **〔capability3 翻案〕COM 全链路可达**：`MediaFormat.MediaBookmarks.Add(Position秒, Name)` 创建书签（落 `<p14:bmk name time/>`），`InteractiveSequences.AddTriggerEffect(shape, eff, 5, videoShape, "b1")`（5=msoAnimTriggerOnMediaBookmark）落 `evt="onMediaBookmark"` 交互序列，重开不修复。CreateVideo **不会**触发书签（同点击触发，仅实机可验） | 原判"无 COM 打书签入口"有误；`Add` 参数顺序是 **(Position, Name)**——(Name, Position) 会因 pywin32 类型转换报 invalid literal |
| G3 | 墨迹手写回放 | —— | ❌ 依赖笔迹输入，无 COM 入口 | |

## H 区：capability3 全集审计新增（2026-10-20 实测补录）

| # | 能力 | COM 写法 | 状态 | 关键坑 |
|---|------|----------|------|--------|
| H1 | 幻灯片切换族·经典成员 | `slide.SlideShowTransition.EntryEffect = PpEntryEffect`（push=3855、wipe=2817–2820、cover=1281–1288、dissolve=1537 等 192 个枚举成员） | ✅ 写入→落盘 `<p:push dir="u"/>` 等→CreateVideo 渲染换片可见（帧差尖峰） | 经典切换的 `Duration` 被量化为 spd（0.75s→"med"）；毫秒级时长见 H2 |
| H2 | 幻灯片切换族·2010+ 现代成员 | 同上，枚举直达：reveal=3894/3895、**ripple=3867–3871**、vortex=3863–3866、glitter=3872–3879、honeycomb=3898、ferrisWheel=3899/3900、cube=3914–3917、morph=3954–3956 等 | ✅ 落盘 `<p14:reveal/>`、`<p14:ripple/>`，现代切换 `Duration` 写 `p14:dur` 毫秒（1.25→1250）；重开不修复、渲染可见 | 无需手工 mc:AlternateContent/命名空间注册——COM 全部代写 |
| H3 | 按字母文本动画（COM 直达） | 建效果后 `shape.AnimationSettings.TextUnitEffect = 2`（by character；1=by word） | ✅ 落盘 `<p:iterate type="lt"><p:tmPct val="100000"/>`；渲染前沿 16 帧逐字阶跃（整框对照=1 次跳变）；重开不修复、其他效果时序不受牵连 | 遗留写会**改写该效果时长细节**（实测 animEffect 内层 dur 被重写），设完 TextUnitEffect 后复核一遍 `Timing.Duration`；`EffectInformation.TextUnitEffect` 仍只读 |
| H4 | UI 独占强调：Pulse/ColorPulse | COM 无枚举 id。做法：建 GrowShrink(59) 后解包把该效果 cTn 的 `presetID="6"` 改 `35`（Pulse）/`36`（ColorPulse），presetClass 保持 `emph` | ✅ 35/36 补丁后重开不修复、按排定时刻播放（瞬时脉冲帧差） | 名称↔presetID 对应按 UI 画廊顺序推断（35=Pulse、36=ColorPulse）为 ⚠️ 推断；播放本身 ✅。OOXML 基础集 emph presetID 仅 1–34（+3D 37–39），35/36 属 2010+ 扩展位 |
| H5 | 强调族逐效果渲染验证 | `AddEffect(shape, id, 0, 2)`：Teeter=80、Wave=82、Flicker=76、Desaturate=74 | ✅ 四者 CreateVideo 渲染均见区域运动（峰值 2.6–69 灰度帧差，时刻与排定一致） | StyleEmphasis=79 在无文字形状上无可见变化（它动画字体样式）——对文字框使用 |
| H6 | 预设运动路径家族（64 条） | `AddEffect(shape, 86–149)`（如 Circle=86、SineWave=125、SpiralRight=131、BounceRight=139） | ✅ 全部接受；`Behaviors(1).MotionEffect.Path` 可读出贝塞尔路径串（M/C 相对坐标），可覆写为自定义路径 | 预设路径坐标同自定义路径口径（相对原始中心，x/页宽、y/页高） |
| H7 | 图表 build 按系列/类别 | COM：`AddEffect` Level 参数（7–11）❌ 拒绝、遗留 `AnimationSettings.ChartUnitEffect` ❌ 拒绝；**可行 = XML 补丁**：level=0 建效果自动写 `<p:bldGraphic spid grpId><p:bldAsOne/>`，替换为 `<p:bldSub><a:bldChart bld="series|category" animBg="0"/></p:bldSub>` | ⚠️ 补丁重开/另存 roundtrip 保持（COM 接受该形态），但 **CreateVideo 渲染与 asOne 无差别**——分步播放仅实机可验 | `AddChart2` 每次会拉起 EXCEL.EXE 数据网格且**同会话第二次 AddChart2 被"数据网格已打开"阻塞**（每演示文稿一图，异常后连 EXCEL 一起清，标题含"Microsoft PowerPoint 中的图表"） |
| H8 | SmartArt 建图 + 整体动画 | `Shapes.AddSmartArt(Application.SmartArtLayouts.Item(i), ...)`（pywin32 下 SmartArtLayouts 需动态绑定 `win32com.client.Dispatch` 取）+ `AddEffect(sa, eff, 0, trigger)` | ✅ 建图/填节点文本/整体动画（自动落 `<p:bldGraphic><p:bldAsOne/>`） | 逐个/逐级 build 不可行，见 I2 |
| H9 | 媒体播放/暂停/停止效果 | `AddMediaObject2(file, LinkToFile=msoFalse, SaveWithDocument=msoTrue, ...)` 后 `AddEffect(mediaShape, 83/84/85, 0, 2)`（MediaPlay/Pause/Stop） | ✅ 媒体形状上全部接受、落 mediacall presetClass；渲染可见视频真实播放（testsrc 图案区域帧差 25 帧）+ mp4 含媒体音频流 | 普通形状上 83/84/85 **拒绝**（A1 旧"拒绝"结论即源于此）；插入视频时 PowerPoint 自动生成"点击播放" mediacall 交互序列 |
| H10 | 媒体自动播放与播放设置 | `shape.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue`；LoopUntilStopped/HideWhileNotPlaying/RewindMovie/StopAfterSlides 可读 | ✅ | 与 83 效果二选一即可 |
| H11 | MediaFormat 媒体属性 | `shape.MediaFormat.Volume/TrimStart/TrimEnd/FadeInDuration/FadeOutDuration/Muted/StartPoint/EndPoint/SetDisplayPictureFromFile` | ✅ Volume=0.6 写入回读一致 | 成员表见 `t08_result.json`；`MediaBookmarks` 见 G2 |
| H12 | CreateVideo 渲染内嵌媒体 | `CreateVideo(mp4, True, fps)` 照常 | ✅ 媒体画面动 + **音频进入 mp4 音轨**（ffprobe 有 audio 流） | 无头 QA 管线可直接验证媒体播放 |
| H13 | 效果伴随音效 | `effect.EffectInformation.SoundEffect.ImportFromFile(wav_path)`（**方法名是 ImportFromFile，无 Import**） | ✅ wav 嵌包（ppt/media/audio1.wav，Type=2 ppSoundFile）；**CreateVideo 渲染有声**：效果时间窗 mean −28dB / max −24dB，静音段 −91dB | 内置音效名（`SoundEffect.Name="applause"`）❌ 拒绝——本机无内置音效库，只能导入文件 |
| H14 | 切换伴随音效 | `slide.SlideShowTransition.SoundEffect.ImportFromFile(wav)` | ⚠️ 落盘 ✅（`<p:transition><p:sndAc><p:stSnd><p:snd r:embed/>`，与效果音效共用媒体部件）；本次 CreateVideo 渲染该时间窗无声——切换音效仅实机可验 | |
| H15 | 动画刷 COM 等价 | `srcShape.PickupAnimation()` → `dstShape.ApplyAnimation()` | ✅ 动画（含路径效果）复制到目标形状主序列 | 注意大小写是 Pickup（非 PickUp） |
| H16 | 交互序列与主序列/默认序列并存 | 主序列 + `InteractiveSequences.Add()` + 插入视频自动生成的点击 mediacall 序列同页共存 | ✅ 三序列同页落盘、重开不修复、渲染主序列正常播放 | CreateVideo 只播主序列（交互序列等待实机触发） |
| H17 | 3D 模型插入（3D 动画前置条件复核） | `Shapes.Add3DModel(glb, ...)`（OM 成员存在） | ❌ 传入手工构造的合法最小 glTF 2.0 GLB 仍拒绝（0x80004005）——本机 16.0 无 3D 模型支持，G1 维持不可行 | 效果 id 151–154 在普通形状上"接受"但为静默空操作，勿被 AddEffect 不报错误导 |

## I 区：capability3 新确认不可行

| # | 能力 | 结论 | 原因（实测） |
|---|------|------|--------------|
| I1 | 图表 build 经 COM 直达 | ❌ | `AddEffect` Level=7–11（chart 值）对任意 effectId 一律 "Invalid request"；遗留 `AnimationSettings.ChartUnitEffect` 写入拒绝。唯一通道是 H7 的 XML 补丁（且渲染验证受限） |
| I2 | SmartArt 逐个/逐级 build 动画 | ❌ | `AddEffect` Level=12–26（diagram 值）一律拒绝；XML 补丁 `<p:bldSub><a:bldDiagram bld="one|lvlOne"/></p:bldSub>`（含 uiExpand 变体）**打开即硬失败**（非修复，COM Open 直接抛错）。SmartArt 仅支持整体动画（bldAsOne） |
| I3 | 内置音效名库 | ❌ | `SoundEffect.Name = "applause"` 拒绝（"Sorry, cann…"，本机未装内置音效库）；只能 `ImportFromFile` |
| I4 | 3D 模型动画 | ❌（证据升级） | 原判"无对象模型"不准确：`Add3DModel` 在 OM 中存在，但传入合法 GLB 拒绝执行——内容/许可层面不可用，结论不变 |


## 复现（来源项目内路径，仅示例）
```bash
cd ppt/assets/make_native_anim
python capability_probe.py            # 全探区（F 至 300）
python capability_probe.py --full     # F 至 600
python capability_probe.py --only C,D # 选探区
python build_showcase.py              # 重建演示 deck（含 XML 注入后处理）
python capability_qa.py --render      # 渲染 + 每技巧蒙太奇 + 轨迹验证
```

## 对答辩动画的升级建议（Top 3，来源项目示例，思路可平移）
1. **物理落点（BounceEnd / tavLst 弹簧）**：P1 统计卡三行数字、P4 判定公式、P7 AUROC 大字入场改 `BounceEnd=0.6` 或 5% 过冲弹簧——"数据砸下来落地"的确定感比 ease-out 强得多；BounceEnd 是 COM 直达（一行属性），弹簧走 `apply_xml_injections` 注入管线（build_showcase 里已有现成实现，可平移进 build_all）。
2. **Morph 跨页连续变形（EntryEffect=3954）**：P2（现象）→P6（成因）用同名 EE 卡跨页滑移接续，"同一块证据从现象页飞进机制页"；或 P7 结语条从 P1 统计卡 Morph 而来，强化"三页互证"的叙事闭环。注意换片时间按 `p14:dur` 另计。
3. **按词浮现 + 颜色渐变 emphasis**：P7 结语条"退化有结构 → 可弃判 · 可接手"改按词浮现（XML iterate），P1 风险曲线越过 τ 的瞬间把 y 轴"风险分（探针输出）"做一次 ChangeFontColor 红→强调色渐变——零新增形状，纯文字层升级。
