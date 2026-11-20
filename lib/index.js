/**
 * dsh-video-course-agent — 视频课程知识地图 Agent
 *
 * 一个基于 DeepSeek Harness（Cordis）插件框架实现的 Agent：
 * 定时读取 Windows 前台窗口标题，识别浏览器里的 B 站视频课程，
 * 按学科分类，抓取视频简介 / 字幕 / 弹幕 / 评论做知识总结，
 * 输出 Markdown 源 + Word（.docx）学习手册与课程总览。
 *
 * 框架入口（Cordis 插件三要素）：
 *   - export const name     插件名
 *   - export const inject   声明要用的 DSH 服务（llm / tools）
 *   - export function apply(ctx, config)  插件逻辑入口
 */

import { execFile, spawnSync } from 'node:child_process'
import { randomUUID } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync, renameSync, readdirSync, appendFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join, dirname } from 'node:path'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'

const execFileP = promisify(execFile)

export const name = 'dsh-video-course-agent'
export const inject = ['llm', 'tools']

const DEFAULTS = {
  enabled: true,            // 总开关
  pollIntervalMs: 15000,    // 轮询间隔（毫秒）
  stablePolls: 2,           // 连续 N 次观察到同一新标题才处理（防标签切换抖动）
  outputDir: null,          // 输出目录，null → <用户目录>/课程知识地图
  provider: null,           // LLM 提供方，null → 跟随 DSH 默认模型
  model: null,              // 模型名，null → 跟随 DSH 默认模型
  maxTokens: 4096,          // 生成知识地图的最大 token
  titleMinLen: 4,           // 标题最小长度（过滤空/短标题）
  fetchContent: true,       // 是否抓取 B 站简介/字幕/弹幕/评论（失败自动降级为仅标题总结）
}

// 浏览器 / 视频客户端进程名（统一转小写再去 .exe）
const BROWSER_PROCESSES = new Set([
  'chrome', 'msedge', 'edge', 'firefox', 'iexplore', 'opera', 'brave', 'chromium',
  'vivaldi', 'yandex', 'librewolf', 'waterfox',
  '360chrome', '360se', '360browser', 'qqbrowser', 'sogouexplorer', 'maxthon', 'ucbrowser', 'quark', '夸克',
  'bilibili', '哔哩哔哩', 'douyin', '抖音', 'aweme', 'kuaishou', '快手',
  'xigua', '西瓜视频', 'youku', '优酷', 'iqiyi', '爱奇艺', 'pps',
  'qqlive', 'tencentvideo', '腾讯视频', 'mgtv', '芒果tv', '芒果TV',
  'migu', '咪咕视频', 'weishi', '微视', 'sohuvideo', '搜狐视频', 'letv', '乐视',
  'potplayer', 'vlc', 'mpv', 'kmplayer', 'mpc-hc', 'mpc-be',
])

// 标题末尾要剥离的平台/浏览器后缀（循环剥离）
const TITLE_SUFFIXES = [
  ' - Google Chrome', ' - Microsoft Edge', ' - Mozilla Firefox',
  ' - Chromium', ' - Opera', ' - Brave', ' - Internet Explorer',
  ' - 谷歌浏览器', ' - 微软 Edge', ' - 火狐浏览器', ' - 360极速浏览器',
  '_哔哩哔哩_bilibili', '_哔哩哔哩', ' - 哔哩哔哩', ' - bilibili', '-bilibili', ' - Bilibili',
  ' - YouTube', ' - 优酷', ' - 爱奇艺', ' - 腾讯视频', ' - 芒果TV', ' - 抖音', ' - 快手', ' - 夸克', ' - Quark',
]

// 前台窗口标题读取脚本：结果写入 UTF-8 文件，避免 stdout 中文编码问题。
const POLL_PS1 = `param([string]$OutFile)
Add-Type -ErrorAction SilentlyContinue @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class FGWin {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
}
"@
$h = [FGWin]::GetForegroundWindow()
$sb = New-Object System.Text.StringBuilder 512
[void][FGWin]::GetWindowText($h, $sb, 512)
$procId = 0
[void][FGWin]::GetWindowThreadProcessId($h, [ref]$procId)
$pname = (Get-Process -Id $procId -ErrorAction SilentlyContinue).ProcessName
$nl = [Environment]::NewLine
$text = "TITLE=" + $sb.ToString() + $nl + "PROC=" + $pname
[System.IO.File]::WriteAllText($OutFile, $text, (New-Object System.Text.UTF8Encoding($false)))
`

// ---------- 配置与通用工具 ----------
function resolveConfig(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
    throw new Error('dsh-video-course-agent: 配置必须是对象')
  }
  const config = { ...DEFAULTS }
  for (const [key, value] of Object.entries(raw)) {
    if (value === undefined) continue
    if (!(key in DEFAULTS)) throw new Error(`dsh-video-course-agent: 未知配置项 "${key}"`)
    config[key] = value
  }
  for (const key of ['pollIntervalMs', 'stablePolls', 'maxTokens', 'titleMinLen']) {
    const v = config[key]
    if (typeof v !== 'number' || !Number.isFinite(v) || v <= 0) throw new Error(`dsh-video-course-agent: ${key} 必须是正数`)
  }
  if (typeof config.enabled !== 'boolean') throw new Error('dsh-video-course-agent: enabled 必须是布尔值')
  if (typeof config.fetchContent !== 'boolean') throw new Error('dsh-video-course-agent: fetchContent 必须是布尔值')
  for (const key of ['provider', 'model']) {
    if (config[key] !== null && (typeof config[key] !== 'string' || config[key].trim() === '')) {
      throw new Error(`dsh-video-course-agent: ${key} 必须是 null 或非空字符串`)
    }
  }
  config.outputDir = config.outputDir ?? join(homedir(), '课程知识地图')
  return config
}

function ensureDir(p) { if (!existsSync(p)) mkdirSync(p, { recursive: true }) }

function atomicWrite(file, content) {
  ensureDir(dirname(file))
  const tmp = `${file}.tmp.${process.pid}`
  writeFileSync(tmp, content, 'utf8')
  renameSync(tmp, file)
}

function loadState(p) {
  try {
    const j = JSON.parse(readFileSync(p, 'utf8'))
    return j && typeof j === 'object' && !Array.isArray(j) ? j : {}
  } catch { return {} }
}

function cleanTitle(raw) {
  let t = String(raw || '').trim()
  let changed = true
  while (changed) {
    changed = false
    for (const suf of TITLE_SUFFIXES) {
      if (t.endsWith(suf)) { t = t.slice(0, -suf.length).trim(); changed = true }
    }
  }
  return t
}

function sanitizeSubject(s) {
  return String(s || '').trim().replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').slice(0, 60) || '未分类'
}

function mermaidSafe(s) {
  return String(s || '').replace(/[()[\]{}|<>'"`\\/]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60)
}

function parseJson(text) {
  if (!text) return null
  const t = String(text).trim()
  try { return JSON.parse(t) } catch { /* 继续 */ }
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/)
  if (fence) { try { return JSON.parse(fence[1].trim()) } catch { /* 继续 */ } }
  const m = t.match(/\{[\s\S]*\}/)
  if (!m) return null
  const candidate = m[0]
  try { return JSON.parse(candidate) } catch { /* 继续 */ }
  try {
    const fixed = candidate
      .replace(/,\s*([}\]])/g, '$1')
      .replace(/([{,]\s*)([A-Za-z_][A-Za-z0-9_]*)(\s*:)/g, '$1"$2"$3')
      .replace(/:\s*'([^']*)'/g, ':"$1"')
    return JSON.parse(fixed)
  } catch { return null }
}

// ---------- LLM ----------
function resolveModel(ctx, config) {
  if (config.provider && config.model) return { provider: config.provider, model: config.model }
  try {
    const sel = ctx.get?.('agentDefaultModel')?.currentSelection?.()
    if (sel && typeof sel.provider === 'string' && typeof sel.model === 'string') {
      return { provider: sel.provider, model: sel.model }
    }
  } catch { /* 忽略 */ }
  try {
    const providers = ctx.llm?.listProviders?.() ?? []
    for (const p of providers) {
      const mid = p.models?.[0]?.id ?? p.defaultModel ?? p.model
      if (p.id && mid) return { provider: p.id, model: mid }
    }
  } catch { /* 忽略 */ }
  return { provider: 'deepseek-official', model: 'deepseek-v4-pro' }
}

async function callLlm(ctx, config, { system, messages, maxTokens }) {
  const model = resolveModel(ctx, config)
  const ac = new AbortController()
  const timer = setTimeout(() => ac.abort(new Error('dsh-video-course-agent: LLM 调用超时')), 120000)
  try {
    const stream = ctx.llm.stream({
      provider: model.provider,
      model: model.model,
      system,
      messages,
      maxTokens: maxTokens ?? config.maxTokens,
      signal: ac.signal,
    })
    let text = ''
    for await (const chunk of stream) {
      if (chunk?.type === 'text-delta' && typeof chunk.text === 'string') text += chunk.text
    }
    return text
  } finally {
    clearTimeout(timer)
  }
}

const userMsg = (text) => ({ role: 'user', id: randomUUID(), content: [{ type: 'text', text }] })

// ---------- B 站内容抓取 ----------
const BILI_UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

let cachedBuvid3 = null
async function getBuvid3() {
  if (cachedBuvid3 !== null) return cachedBuvid3
  try {
    const r = await fetch('https://api.bilibili.com/x/frontend/finger/spi', { headers: { 'User-Agent': BILI_UA } })
    const j = await r.json()
    cachedBuvid3 = j?.data?.b_3 || ''
  } catch { cachedBuvid3 = '' }
  return cachedBuvid3
}

async function biliFetch(url, { json = true, timeoutMs = 10000, retries = 2 } = {}) {
  const buvid3 = await getBuvid3()
  let lastErr = null
  for (let attempt = 0; attempt <= retries; attempt++) {
    const ac = new AbortController()
    const timer = setTimeout(() => ac.abort(new Error('timeout')), timeoutMs)
    try {
      const res = await fetch(url, {
        headers: {
          'User-Agent': BILI_UA,
          'Referer': 'https://www.bilibili.com/',
          'Accept': 'application/json, text/plain, */*',
          ...(buvid3 ? { Cookie: 'buvid3=' + buvid3 } : {}),
        },
        signal: ac.signal,
      })
      if (res.status === 412) throw new Error('HTTP 412')
      if (!res.ok) throw new Error('HTTP ' + res.status)
      const text = await res.text()
      return json ? JSON.parse(text) : text
    } catch (e) {
      lastErr = e
      if (attempt < retries) await sleep(1500 * (attempt + 1))
    } finally {
      clearTimeout(timer)
    }
  }
  throw lastErr
}

const stripHtml = (s) => String(s || '').replace(/<[^>]+>/g, '').trim()

function decodeXml(s) {
  return String(s || '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, ' ')
}

const normTitle = (s) => String(s || '').toLowerCase().replace(/[\s\p{P}\p{S}]+/gu, '')

function titleSimilar(a, b) {
  const na = normTitle(a); const nb = normTitle(b)
  if (!na || !nb) return 0
  if (na.includes(nb) || nb.includes(na)) return 1
  const bigrams = (s) => { const m = new Set(); for (let i = 0; i < s.length - 1; i++) m.add(s.slice(i, i + 2)); return m }
  const A = bigrams(na); const B = bigrams(nb)
  if (!A.size || !B.size) return 0
  let inter = 0
  for (const x of A) if (B.has(x)) inter++
  return (2 * inter) / (A.size + B.size)
}

async function searchBili(title) {
  const url = 'https://api.bilibili.com/x/web-interface/search/type?search_type=video&keyword=' + encodeURIComponent(title)
  const data = await biliFetch(url)
  if (data?.code !== 0) throw new Error('bili search code ' + data?.code)
  const results = (data?.data?.result ?? []).filter((r) => r?.bvid)
  if (!results.length) return null
  const top = results[0]
  const clean = stripHtml(top.title)
  if (titleSimilar(title, clean) < 0.5) return null
  return { bvid: top.bvid, aid: top.aid, title: clean, desc: stripHtml(top.description || ''), tags: String(top.tag || '').split(',').map((s) => s.trim()).filter(Boolean) }
}

async function viewBili(bvid) {
  const data = await biliFetch('https://api.bilibili.com/x/web-interface/view?bvid=' + encodeURIComponent(bvid))
  if (data?.code !== 0) throw new Error('bili view code ' + data?.code)
  const d = data?.data ?? {}
  return { aid: d.aid, cid: d.cid ?? d.pages?.[0]?.cid, desc: d.desc || '', owner: d.owner?.name || '', stat: d.stat || {}, subtitle: d.subtitle?.list ?? [] }
}

async function fetchDanmaku(cid) {
  if (!cid) return []
  const xml = await biliFetch('https://comment.bilibili.com/' + cid + '.xml', { json: false })
  const items = []
  const re = /<d[^>]*>([\s\S]*?)<\/d>/g
  let m
  while ((m = re.exec(xml)) !== null) {
    const t = decodeXml(m[1]).trim()
    if (!t) continue
    if (t.startsWith('[') && t.endsWith(']')) continue // 高级弹幕 JSON
    items.push(t)
  }
  const seen = new Set()
  return items.filter((t) => { if (seen.has(t)) return false; seen.add(t); return true }).slice(0, 150)
}

async function fetchComments(aid) {
  if (!aid) return []
  const data = await biliFetch('https://api.bilibili.com/x/v2/reply?type=1&oid=' + aid + '&sort=2&ps=20&pn=1')
  if (data?.code !== 0) throw new Error('bili reply code ' + data?.code)
  return (data?.data?.replies ?? []).map((r) => String(r?.content?.message || '').trim()).filter(Boolean).slice(0, 20)
}

async function fetchSubtitle(view) {
  const list = view?.subtitle ?? []
  if (!list.length) return ''
  let u = list[0]?.subtitle_url
  if (!u) return ''
  if (u.startsWith('//')) u = 'https:' + u
  const data = await biliFetch(u, { json: true })
  return (data?.body ?? []).map((b) => String(b?.content || '').trim()).filter(Boolean).join(' ')
}

async function fetchBiliContent(title) {
  const s = await searchBili(title)
  if (!s) return null
  let view = null
  try { view = await viewBili(s.bvid) } catch { /* 忽略 */ }
  const cid = view?.cid
  const aid = view?.aid || s.aid
  const [danmaku, comments] = await Promise.allSettled([
    cid ? fetchDanmaku(cid) : Promise.resolve([]),
    aid ? fetchComments(aid) : Promise.resolve([]),
  ])
  let transcript = ''
  try { transcript = await fetchSubtitle(view) } catch { /* 忽略 */ }
  return {
    bvid: s.bvid, aid, cid,
    desc: String(view?.desc || s.desc || '').trim(),
    tags: s.tags || [], owner: view?.owner || '', stat: view?.stat || {},
    transcript,
    danmaku: danmaku.status === 'fulfilled' ? danmaku.value : [],
    comments: comments.status === 'fulfilled' ? comments.value : [],
  }
}

// ---------- 分类与总结 ----------
async function classify(ctx, config, title, existingSubjects) {
  const system = '你是视频内容分类助手。判断一个「浏览器当前标题」是否值得整理成知识地图，并归类到学科。'
  const user = `浏览器当前标题：${JSON.stringify(title)}
已有学科列表：${JSON.stringify(existingSubjects)}

只返回一个合法 JSON 对象（不要任何多余文字、不要 markdown 代码块、不要尾逗号），格式：
{"isCourse":true,"subject":"学科名","title":"清洗后的视频标题","keywords":["关键词1","关键词2"]}

规则：
- isCourse 填 true 的「有知识含量」内容：课程、教程、教学、科普、纪录片、开箱测评、知识分享、技能讲解、历史/文化/科技解读等。
- isCourse 填 false 的「无知识含量」内容：纯音乐、纯搞笑段子、无解说游戏实况、直播聊天、新闻播报、普通网页、搜索页、社交页、邮箱、工作软件等。
- 若该视频属于已有学科之一，subject 必须与列表里某个已有名字完全一致；否则给一个简洁的新学科名（中文，≤8 字）。
- title 为去掉浏览器与平台后缀后的纯视频标题。
- keywords 为 3~6 个核心知识点关键词。`
  let text = await callLlm(ctx, config, { system, messages: [userMsg(user)], maxTokens: 700 })
  let cls = parseJson(text)
  if (!cls) {
    const retry = user + '\n\n（注意：上一次你的回复无法被解析。请严格只输出一个合法 JSON 对象本身，前后不要任何解释文字、不要 markdown 代码块、不要尾逗号。）'
    text = await callLlm(ctx, config, { system, messages: [userMsg(retry)], maxTokens: 700 })
    cls = parseJson(text)
  }
  return cls
}

async function generate(ctx, config, { title, subject, keywords, content }) {
  const system = '你是学习笔记整理专家，为视频课程生成「知识地图」（Markdown + Mermaid 思维导图）。'
  const safeTitle = mermaidSafe(title)
  let srcBlock = ''
  if (content) {
    const parts = []
    if (content.desc) parts.push('【视频简介】\n' + content.desc.slice(0, 800))
    if (content.tags?.length) parts.push('【视频标签】\n' + content.tags.join('、'))
    if (content.transcript) parts.push('【字幕/讲解文本（节选）】\n' + content.transcript.slice(0, 1500))
    if (content.danmaku?.length) parts.push('【弹幕（节选 ' + content.danmaku.length + ' 条）】\n' + content.danmaku.slice(0, 120).join('\n'))
    if (content.comments?.length) parts.push('【评论区热评（节选 ' + content.comments.length + ' 条）】\n' + content.comments.slice(0, 20).join('\n'))
    if (parts.length) srcBlock = '\n\n已抓取到以下视频真实内容/弹幕/评论，请基于它们总结（不要凭空编造）：\n\n' + parts.join('\n\n')
  }
  const user = `视频标题：${title}
学科：${subject}
关键词：${keywords.join('、')}${srcBlock}

只返回 Markdown（不要任何解释、不要首尾多余说明），结构如下：

## ${title}

> 一句话简介（概括本视频核心内容）

### 核心知识点
- **主题一**
  - 要点
  - 要点

### 弹幕与评论讨论焦点
- 观众关注/讨论的重点（1~3 条）

### 思维导图
\`\`\`mermaid
mindmap
  root((${safeTitle}))
    ...
\`\`\`

要求：${content ? '优先基于抓取到的简介/字幕/弹幕/评论提炼真实知识点，不要编造。' : '基于标题与关键词合理推断、通用化。'}全部用中文；知识点 8~20 条、分 2~3 层；mermaid 用 mindmap 语法、2~3 层、节点文字避免括号等特殊字符。`
  return await callLlm(ctx, config, { system, messages: [userMsg(user)], maxTokens: config.maxTokens })
}

// ---------- Word 转换 ----------
const PLUGIN_DIR = join(dirname(fileURLToPath(import.meta.url)), '..')
const MD2DOCX_PATH = join(PLUGIN_DIR, 'scripts', 'md2docx.py')

function findPython() {
  const cands = []
  const runtimesDir = join(homedir(), '.dsh', 'dsh-runtimes')
  try {
    for (const rt of readdirSync(runtimesDir)) cands.push(join(runtimesDir, rt, 'dependencies', 'python', 'python.exe'))
  } catch { /* 忽略 */ }
  cands.push('python', 'python3')
  for (const c of cands) {
    try { if (spawnSync(c, ['--version'], { timeout: 8000, windowsHide: true }).status === 0) return c } catch { /* 忽略 */ }
  }
  return 'python'
}

// ---------- 插件入口 ----------
let debugLog = () => {}

export function apply(ctx, rawConfig = {}) {
  const config = resolveConfig(rawConfig)
  ensureDir(config.outputDir)
  const stateDir = join(config.outputDir, '.state')
  ensureDir(stateDir)
  const statePath = join(stateDir, 'state.json')
  const ps1Path = join(stateDir, 'poll.ps1')
  const pollOutPath = join(stateDir, 'poll-out.txt')
  const manualsDir = join(stateDir, 'manuals')
  ensureDir(manualsDir)
  const pythonPath = findPython()
  writeFileSync(ps1Path, POLL_PS1, 'utf8')

  const state = loadState(statePath)
  if (!state.subjects || typeof state.subjects !== 'object') state.subjects = {}
  if (!state.seen || typeof state.seen !== 'object') state.seen = {}
  if (typeof state.lastTitle !== 'string') state.lastTitle = ''
  if (typeof state.lastRunAt !== 'number') state.lastRunAt = 0

  const saveState = () => atomicWrite(statePath, JSON.stringify(state, null, 2) + '\n')

  const logPath = join(stateDir, 'agent.log')
  debugLog = (msg) => { try { appendFileSync(logPath, `[${new Date().toISOString()}] ${msg}\n`, 'utf8') } catch { /* 忽略 */ } }

  const subjectDir = (subject) => join(config.outputDir, subject)
  const manualMdPath = (subject) => join(manualsDir, `${subject}.md`)
  const manualDocxPath = (subject) => join(subjectDir(subject), '学习手册.docx')
  const indexMdPath = () => join(manualsDir, 'index.md')

  function convertToDocx(mdPath, docxPath) {
    ensureDir(dirname(docxPath))
    const r = spawnSync(pythonPath, [MD2DOCX_PATH, mdPath, docxPath], { timeout: 60000, windowsHide: true, encoding: 'utf8' })
    if (r.status !== 0) throw new Error(`md2docx 转换失败（${pythonPath}）：${String(r.stderr || r.stdout || '').trim().slice(0, 500)}`)
  }

  function createSubjectManual(subject) {
    const header = `# ${subject} 学习手册\n\n<!-- OVERVIEW_START -->\n## 学科总览\n\`\`\`mermaid\nmindmap\n  root((${mermaidSafe(subject)}))\n\`\`\`\n<!-- OVERVIEW_END -->\n`
    atomicWrite(manualMdPath(subject), header)
  }

  function buildOverviewBlock(subject) {
    const videos = state.subjects[subject]?.videos ?? []
    const lines = ['```mermaid', 'mindmap', `  root((${mermaidSafe(subject)}))`]
    for (const v of videos) lines.push(`    ${mermaidSafe(v)}`)
    lines.push('```')
    return `<!-- OVERVIEW_START -->\n## 学科总览\n${lines.join('\n')}\n<!-- OVERVIEW_END -->`
  }

  function regenerateOverview(subject) {
    const p = manualMdPath(subject)
    const text = readFileSync(p, 'utf8')
    const start = text.indexOf('<!-- OVERVIEW_START -->')
    const end = text.indexOf('<!-- OVERVIEW_END -->')
    if (start === -1 || end === -1) return
    const merged = text.slice(0, start) + buildOverviewBlock(subject) + text.slice(end + '<!-- OVERVIEW_END -->'.length)
    atomicWrite(p, merged)
  }

  function appendVideo(subject, title, md) {
    const p = manualMdPath(subject)
    const text = readFileSync(p, 'utf8').replace(/\s+$/, '') + `\n\n---\n\n${md}\n`
    atomicWrite(p, text)
  }

  function convertSubject(subject) {
    convertToDocx(manualMdPath(subject), manualDocxPath(subject))
  }

  function regenerateIndex() {
    const subjects = Object.keys(state.subjects)
    const lines = ['# 视频课程知识地图', '', `> 自动生成 · 最后更新 ${new Date().toLocaleString('zh-CN', { hour12: false })}`, '', '## 学科总览', '```mermaid', 'mindmap', '  root((课程知识地图))']
    for (const s of subjects) lines.push(`    ${mermaidSafe(s)}`)
    lines.push('```', '', '## 学科列表')
    if (subjects.length === 0) lines.push('（暂无，开始观看视频课程后会自动生成）')
    for (const s of subjects) {
      const n = state.subjects[s]?.videos?.length ?? 0
      lines.push(`- ${s} — ${n} 个视频 → 「${s}/学习手册.docx」`)
    }
    lines.push('')
    atomicWrite(indexMdPath(), lines.join('\n'))
    convertToDocx(indexMdPath(), join(config.outputDir, '课程总览.docx'))
  }

  // 启动时：为已有学科重新生成 Word（幂等）
  try {
    for (const subject of Object.keys(state.subjects)) convertSubject(subject)
    if (Object.keys(state.subjects).length > 0) regenerateIndex()
  } catch (error) {
    console.error('[dsh-video-course-agent] 启动时生成 Word 失败：', error)
  }

  let pending = null   // { title, count } 稳定性去抖
  let running = false  // 处理中互斥

  async function pollForeground() {
    try {
      await execFileP('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps1Path, '-OutFile', pollOutPath], { timeout: 8000, windowsHide: true })
      const txt = readFileSync(pollOutPath, 'utf8').replace(/^\uFEFF/, '')
      const title = (txt.match(/^TITLE=(.*)$/m)?.[1] ?? '').trim()
      const proc = (txt.match(/^PROC=(.*)$/m)?.[1] ?? '').trim()
      return { title, proc }
    } catch { return null }
  }

  function isBrowserProc(proc) {
    return BROWSER_PROCESSES.has(String(proc || '').toLowerCase().replace(/\.exe$/i, ''))
  }

  function pruneSeen() {
    const keys = Object.keys(state.seen)
    if (keys.length <= 500) return
    const sorted = keys.sort((a, b) => (state.seen[a] ?? 0) - (state.seen[b] ?? 0))
    for (const k of sorted.slice(0, keys.length - 500)) delete state.seen[k]
  }

  async function processTitle(title) {
    if (running) return
    running = true
    try {
      const existing = Object.keys(state.subjects)
      const cls = await classify(ctx, config, title, existing)
      debugLog('classify 结果 cls=' + JSON.stringify(cls))
      if (!cls || cls.isCourse === false) {
        state.seen[title] = Date.now()
        state.lastTitle = title
        pruneSeen(); saveState()
        return
      }
      const subject = sanitizeSubject(cls.subject)
      const clean = String(cls.title || title).trim() || title
      const keywords = (Array.isArray(cls.keywords) ? cls.keywords : []).map(String).filter(Boolean)
      let content = null
      if (config.fetchContent) {
        try {
          content = await fetchBiliContent(clean)
          if (content?.tags?.length) for (const t of content.tags) if (t && !keywords.includes(t) && keywords.length < 10) keywords.push(t)
        } catch { /* 静默降级为仅标题总结 */ }
      }
      const md = await generate(ctx, config, { title: clean, subject, keywords, content })
      const isNew = !state.subjects[subject]
      if (isNew) { state.subjects[subject] = { createdAt: Date.now(), videos: [] }; createSubjectManual(subject) }
      appendVideo(subject, clean, md)
      if (!state.subjects[subject].videos.includes(clean)) state.subjects[subject].videos.push(clean)
      regenerateOverview(subject)
      convertSubject(subject)
      regenerateIndex()
      state.seen[title] = Date.now()
      state.lastTitle = title
      state.lastRunAt = Date.now()
      pruneSeen(); saveState()
      console.log(`[dsh-video-course-agent] ${isNew ? '新学科' : '更新学科'}「${subject}」← ${clean}`)
    } catch (error) {
      debugLog('processTitle 异常: ' + (error?.message || error))
      console.error('[dsh-video-course-agent] 处理失败（不影响后续轮询）：', error)
    } finally {
      running = false
    }
  }

  async function observeOnce() {
    if (!config.enabled) return
    const fg = await pollForeground()
    if (!fg) return
    if (!isBrowserProc(fg.proc)) return
    const title = cleanTitle(fg.title)
    if (title.length < config.titleMinLen) return
    if (/哔哩哔哩|bilibili/i.test(title)) return // 搜索页/分类页/首页
    if (Object.hasOwn(state.seen, title)) { pending = null; return }
    if (!pending || pending.title !== title) pending = { title, count: 1 }
    else pending.count += 1
    if (pending.count >= config.stablePolls) {
      const t = pending.title
      pending = null
      await processTitle(t)
    }
  }

  // 状态查看工具
  ctx.effect(() => ctx.tools.register({
    name: 'course_map_status',
    description: '查看「视频课程知识地图 Agent」的运行状态：是否启用、输出目录、已分类学科与视频数、最近观察到的标题。',
    parameters: { type: 'object', properties: {} },
    output: {
      schema: {
        type: 'object', additionalProperties: false,
        properties: {
          ok: { type: 'boolean' }, enabled: { type: 'boolean' },
          outputDir: { type: 'string' }, lastTitle: { type: 'string' },
          lastRunAt: { type: 'integer' },
          subjects: { type: 'array', items: { type: 'object', additionalProperties: false, properties: { subject: { type: 'string' }, videos: { type: 'integer' } }, required: ['subject', 'videos'] } },
          message: { type: 'string' },
        },
        required: ['ok'],
      },
      render: (_a, v) => [{ type: 'text', text: v.message ?? '' }],
    },
    async execute() {
      const subjects = Object.entries(state.subjects).map(([subject, d]) => ({ subject, videos: d.videos?.length ?? 0 }))
      const message = [
        '视频课程知识地图 Agent 状态：',
        `- 启用：${config.enabled}`,
        `- 输出目录：${config.outputDir}`,
        `- 学科数：${subjects.length}`,
        ...subjects.map((s) => `  · ${s.subject}（${s.videos} 个视频）`),
        `- 最近观察：${state.lastTitle || '（无）'}`,
      ].join('\n')
      return { ok: true, enabled: config.enabled, outputDir: config.outputDir, lastTitle: state.lastTitle, lastRunAt: state.lastRunAt, subjects, message }
    },
  }), 'dsh-video-course-agent: status tool')

  // 手动扫描工具
  ctx.effect(() => ctx.tools.register({
    name: 'course_map_scan',
    description: '立即执行一次前台窗口观察与知识地图生成（绕过轮询间隔，便于测试或手动触发）；可传 title 直接指定要处理的视频标题。',
    parameters: {
      type: 'object',
      properties: {
        title: { type: 'string', description: '可选：直接指定要处理的视频标题（跳过前台窗口读取，便于测试/手动输入）' },
      },
    },
    output: {
      schema: {
        type: 'object', additionalProperties: false,
        properties: { ok: { type: 'boolean' }, message: { type: 'string' } },
        required: ['ok'],
      },
      render: (_a, v) => [{ type: 'text', text: v.message ?? '' }],
    },
    async execute(args) {
      if (!config.enabled) return { ok: false, message: 'Agent 未启用（config.enabled=false）' }
      let title = null
      if (args && typeof args.title === 'string' && args.title.trim()) {
        title = cleanTitle(args.title)
      } else {
        const fg = await pollForeground()
        if (!fg) return { ok: false, message: '无法读取前台窗口' }
        if (!isBrowserProc(fg.proc)) return { ok: false, message: `前台窗口不是浏览器/视频客户端（进程 ${fg.proc || '未知'}）：${fg.title}` }
        title = cleanTitle(fg.title)
      }
      let waited = 0
      while (running && waited < 90000) { await sleep(1000); waited += 1000 }
      await processTitle(title)
      return { ok: true, message: `已处理「${title}」，已生成 Word 文档。输出目录：${config.outputDir}` }
    },
  }), 'dsh-video-course-agent: scan tool')

  // 轮询定时器
  ctx.effect(() => {
    const id = setInterval(() => { observeOnce().catch((e) => console.error('[dsh-video-course-agent] 观察异常：', e)) }, config.pollIntervalMs)
    return () => clearInterval(id)
  }, 'dsh-video-course-agent: poll timer')

  console.log(`[dsh-video-course-agent] 已启动：每 ${config.pollIntervalMs}ms 轮询前台窗口，输出目录 ${config.outputDir}`)
}
