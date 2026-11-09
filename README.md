# 公文格式与校对（cookie-gongwen-formatter）

把 Word 文档自动改成**党政机关公文格式（GB/T 9704-2012）**，并主动校对文稿内容（错别字、语病、标点、空格、日期格式等）。

一个可分享的 **Agent 技能包**（Skill），任何支持 Skill 的 Agent（如 WorkBuddy / Claude / Codex 等）安装后即可调用。同时也可作为独立命令行工具直接使用。

## 特性

- ✅ **零第三方依赖**：核心脚本只用 Python 标准库（`zipfile` + `xml`），不依赖 python-docx。
- ✅ **Word / WPS 双适配**：格式改写与校对完全不依赖 Office；`.doc` 转换和页码映射可选走 Word 或 WPS COM，自动探测、优雅降级。
- ✅ **文种自动分流**：自动识别通知 / 报告 / 请示 / 函 / 纪要 / 决定 / 意见 / 通报 / 批复 / 讲话稿 / 总结 / 台账。
- ✅ **表格保护**：改页边距时等比例缩放表格宽度，避免一页变两页、横向溢出；表格类材料不碰格式。
- ✅ **字体降级**：目标电脑缺标准公文字体时自动改用系统自带替代字体（方正小标宋→华文中宋等）。
- ✅ **主动校对**：小错（半角标点、残留空格、日期格式）自动修，大错（错别字、语病、搭配）列清单待确认。
- ✅ **精确汇报**：交付报告一张表，位置精确到「第几段 + 第几页第几行」。

## 快速开始

### 作为 Agent 技能包

不同 Agent 的安装方式：

- **WorkBuddy**：把整个 `cookie-gongwen-formatter` 目录放进 `~/.workbuddy/skills/`。
- **Claude Code**：放进项目的 `.claude/skills/` 或全局 skills 目录。
- **OpenAI Codex**：把 `AGENTS.md` 的内容放进项目的 `AGENTS.md`，脚本放 `scripts/` 下。
- **Gemini / Grok 等**：把 `INSTRUCTIONS.md` 里的「角色指令」整段复制进该 Agent 的系统提示词/自定义指令。

装好后，说「用公文技能处理这份文件」并上传/给出 `.docx` 文件路径即可。

### 作为命令行工具

需要 Python 3.7+（无需安装任何第三方包）：

```bash
# 格式改写（自动判断类型）
python scripts/format_gongwen.py 输入.docx

# 指定类型：gongwen / speech / summary / table
python scripts/format_gongwen.py 输入.docx --mode speech

# 校对（小错自动修 + 大错列清单）
python scripts/proofread.py 输入.docx --fix

# 页码映射（可选，需 Word/WPS）
python scripts/map_pages.py 成品.docx
```

## 目录结构

```
cookie-gongwen-formatter/
├── SKILL.md                      # 技能说明（WorkBuddy / Claude Code 读取的入口）
├── AGENTS.md                     # 技能说明（Codex 及 AGENTS.md 规范读取的入口）
├── INSTRUCTIONS.md               # 万能指令版（Gemini / Grok / 任意 Agent 当提示词用）
├── README.md                     # 本文件（给人看）
├── LICENSE                       # MIT
├── references/
│   └── gongwen-format.md         # GB/T 9704-2012 详细规范、高频错字表、文种结构
└── scripts/
    ├── docx_lite.py              # 零依赖 .docx 读写核心
    ├── format_gongwen.py         # 格式改写主脚本
    ├── proofread.py              # 校对脚本
    └── map_pages.py              # 页码映射（可选增强）
```

## 多 Agent 兼容（不冲突）

本技能是「多 Agent 通用」包，**核心脚本对所有 Agent 完全通用**（零第三方依赖的纯 Python 标准库代码）。不同 Agent 只是「如何读取说明」的方式不同，所以提供了三个入口文件，各 Agent 各取所需：

| 你用什么 Agent | 读哪个入口 |
|---|---|
| WorkBuddy | `SKILL.md` |
| Claude Code | `SKILL.md` |
| OpenAI Codex / AGENTS.md 规范 | `AGENTS.md` |
| Gemini / Grok / 任意贴提示词即用的助手 | `INSTRUCTIONS.md`（整段当系统提示词复制进去） |

三份入口描述的是**同一套流程和脚本**，不会产生冲突。统一的触发句式是「用公文技能处理 <文件>」。

## 可选依赖（仅增强能力）

| 能力 | 可选依赖 | 没有时的表现 |
|---|---|---|
| `.doc` 转 `.docx` | Word / WPS（COM） | 提示用户手动另存为 .docx |
| 页码/行号映射 | pywin32 + Word/WPS | 降级为「第几段 + 原文」定位 |

核心的格式改写与校对**永远可用**，与是否安装 Office 无关。

## 支持的文件类型

- 正式公文（通知/报告/请示/函/纪要/决定/意见/通报/批复）
- 讲话稿 / 发言 / 致辞
- 总结 / 汇报 / 体会 / 述职
- 表格 / 台账（只校对、不改格式）

## 字体说明

标准公文字体（方正小标宋简体、仿宋_GB2312、楷体_GB2312）是商业字体，技能不会去下载安装。检测到目标电脑缺失时自动降级：

| 标准字体 | 降级替代 |
|---|---|
| 方正小标宋简体 | 华文中宋 |
| 仿宋_GB2312 | 仿宋 |
| 楷体_GB2312 | 楷体 |

降级信息会写进改动明细，让用户知道文件在别的电脑上也能正常显示。

## License

MIT
