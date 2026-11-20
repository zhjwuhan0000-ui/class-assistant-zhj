# dsh-video-course-agent · 视频课程知识地图 Agent

一个纯 **DeepSeek Harness（DSH）插件**：实时检测网页端 B 站视频课程 → 按关键词自动学科分类 → 抓取视频简介 / 字幕 / 弹幕 / 评论做知识总结 → 输出 Word 知识地图与学习手册。基于 DSH 的 Cordis 插件框架从零实现，无需修改 DSH 核心。

## 功能

- **实时检测 & 静默去重**：定时读取 Windows 前台窗口标题，识别浏览器 / 视频客户端中的视频；新视频自动生成，已处理过的保持静默。
- **学科分类**：由 LLM 按标题 + 关键词判断学科（中文 ≤8 字）；旧学科追加更新，新学科自动新建手册。
- **内容总结 & Word 输出**：按标题反查 B 站，抓取简介 / 标签 / 字幕（如有）/ 弹幕（≤150 条）/ 热评（≤20 条），提炼核心知识点与思维导图，输出 Word（.docx）。思维导图在 Word 里渲染为原生分级大纲。

## 安装

```bash
# 方式一：本地目录安装（开发调试用，链接 checkout）
dsh plugin add "<本仓库目录绝对路径>"

# 方式二：GitHub Release 预构建包
dsh plugin add "https://github.com/<你>/dsh-video-course-agent/releases/download/v1.0.0/dsh-video-course-agent-1.0.0.tgz"
```

## 配置项（`cordis.patch.yml` 中该行的 `config`）

| 配置 | 默认 | 说明 |
| --- | --- | --- |
| `enabled` | `true` | 总开关 |
| `pollIntervalMs` | `15000` | 轮询间隔（毫秒） |
| `stablePolls` | `2` | 连续 N 次观察到同一新标题才处理 |
| `outputDir` | `null` | 输出目录，`null` → `<用户目录>/课程知识地图` |
| `provider` / `model` | `null` | LLM 提供方 / 模型，`null` → 跟随 DSH 默认模型 |
| `maxTokens` | `4096` | 生成知识地图的最大 token |
| `titleMinLen` | `4` | 标题最小长度 |
| `fetchContent` | `true` | 是否抓取 B 站内容（失败自动降级为仅标题总结） |

## 提供的工具

| 工具 | 作用 |
| --- | --- |
| `course_map_status` | 查看 Agent 状态：是否启用、输出目录、已分类学科与视频数、最近观察标题 |
| `course_map_scan` | 立即执行一次前台窗口观察与知识地图生成（可传 `title` 指定视频标题） |

## 输出目录结构

```
<输出目录>/
├── 课程总览.docx            # 全局知识地图
├── <学科A>/学习手册.docx     # 该学科所有视频的知识总结
├── <学科B>/学习手册.docx
└── .state/                  # 内部状态（state.json、manuals/*.md 源、poll.ps1）
```

## 技术要点 / 坑

- 改 `lib/index.js` 后必须重启 DSH 才生效（DSH 对插件模块做缓存）。
- B 站 API 需带浏览器 UA + Referer（否则 `-412`）；通过 `/x/frontend/finger/spi` 取 `buvid3` 提升稳定性。
- 评论匿名接口封顶 3 条；AI 总结接口需登录不可用。
- 多数 B 站视频无机器字幕（`subtitle` 为空），字幕尽力抓取、失败静默降级。
- 前台窗口必须是浏览器 / 视频客户端才能识别；后台播放（窗口不在前台）属预期限制。
- Word 生成依赖本机 Python（优先复用 DSH 捆绑运行时 `~/.dsh/dsh-runtimes/*/dependencies/python/python.exe`）。

## 许可证

MIT
