# ppt-native-animation

PowerPoint **原生动画** COM 自动化构建与 QA 的 Agent Skill（Windows + 桌面版 PowerPoint + pywin32）。

用脚本驱动桌面版 PowerPoint，构建"进页自动播放"的原生动画演示文稿。所有语义均经**交付文件 XML 解包 + 实机放映 + CreateVideo 渲染**三重验证，而非仅凭 API 文档推断。

## 能做什么

- 对象动画：进入 / 强调 / 退出 / 任意运动路径 / 文本逐字·逐词
- 图表（含按系列/类别 build）与整体 SmartArt 动画
- 媒体播放与书签触发、点击触发交互、效果/切换音效
- 幻灯片切换族全量 192 个 PpEntryEffect（含 Morph 跨页变形）
- XML 注入自定义缓动（弹簧/过冲/任意一维轨迹）
- QA 闭环：CreateVideo 渲染 + ffmpeg 抽帧质心跟踪 + 终帧叠加验证 + 16 帧蒙太奇

同时覆盖排查脚本生成动画的**放映错位 / 元素不显示 / 时序异常**等问题。

## 环境要求

- Windows + 桌面版 PowerPoint（Office 16.0 实测；WPS 不保证兼容）
- Python 3 + `pip install pywin32`
- QA 渲染链路另需 ffmpeg

无 PowerPoint 环境时的降级方案（GIF/MP4 预渲染、手写 OOXML、Mac/Keynote）见 SKILL.md §1。

## 安装

将本目录复制到 Agent 的 skills 目录即可，例如：

```
~/.agents/skills/ppt-native-animation/
```

## 目录结构

```
SKILL.md                  # 铁律与速查（入口，先读这个）
references/guide.md       # 逐节深入：调度语义/XML 注入/Morph/QA 方法学/图表媒体等
references/capabilities.md# 能力矩阵：每项能力的 COM 写法、验证状态、关键坑
scripts/common.py         # 构建框架：画布/形状工厂/效果封装/调度契约
scripts/inject_easing.py  # XML 缓动注入器（弹簧/自定义 tav，内置三致命坑处理）
scripts/qa_motion.py      # QA 工具：渲染/抽帧/蒙太奇/质心跟踪
```

## 快速开始

```python
import sys; sys.path.insert(0, r"<本skill目录>\scripts")
import common as K
deck = K.Deck()                     # 交付版：不自动翻页
s = deck.slide()                    # 960×540pt 画布
card = s.card(100, 100, 200, 80, text="hello", size=14)
dot = s.dot(700, 300, 14)
s.fade(card, T=0.5)                 # T = 绝对开始时刻（秒）
s.path(dot, pts=[(700, 300), (500, 300), (400, 200)], T=1.2, dur=1.0, ease=(0.0, 1.0))
deck.save(r"out.pptx"); deck.render_video(r"out.mp4"); deck.close()
```

核心铁律一句话版：**delay 一律写绝对时刻 T（秒），禁止相对增量**——主序列组内每个效果的 `<p:cond delay>` 是相对整条时间线起点的绝对时刻，按增量写会把整页编排压缩进前几秒乱序齐发。
