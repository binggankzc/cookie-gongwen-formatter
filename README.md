# 公文格式与校对

cookie-gongwen-formatter 是用于 Word 公文排版和文字校对的 Agent 技能，也可作为命令行工具使用。版本 1.1.2，MIT License，作者 cookie。

支持正式公文、讲话稿、总结汇报和表格台账。核心脚本只使用 Python 标准库，建议 Python 3.9+。不需要安装 python-docx。

## 安装和使用

Codex：把整个 `cookie-gongwen-formatter` 文件夹放到 `~/.codex/skills/`，下一轮对话使用。也可把本仓库链接交给 Codex，要求安装。

Claude Code：放到适用的 `.claude/skills/` 目录。WorkBuddy：放到 `~/.workbuddy/skills/`。其他支持自定义指令的助手可参考 `INSTRUCTIONS.md`，并确保能访问脚本。

安装后说“用公文技能处理这份文件”，附 `.docx` 文件或路径。

## 命令行

```bash
# 只读识别，不生成文件
python scripts/format_gongwen.py 输入.docx --inspect

# 格式处理，另存新文件
python scripts/format_gongwen.py 输入.docx 输出.docx --mode gongwen

# 其他模式：auto、speech、summary、table
# table 模式只提取文本，不生成格式版

# 只读规则检查
python scripts/proofread.py 输出.docx

# 修复明确的中文标点、空格和有效日期，另存 _校对.docx
python scripts/proofread.py 输出.docx --fix

# Windows 下可选的页码行号映射
python scripts/map_pages.py 输出.docx

# 回归测试，使用合成数据
python -m unittest discover -s tests -v
```

## 1.1.1 修复

保护数字序号、小数、英文、网址、邮箱、路径和文件名，避免将“1.”改成“1。”。日期转换验证真实日历日期，跳过技术字符串。保留普通段落的局部强调，复杂段落和表格内文字只提示。

文种优先依据标题末尾，再看结尾惯用语。“意见建议报告”识别为报告。增加 `--inspect`，避免为了查看类型而先生成文件。

macOS 从字体文件读取真实字体名称；替代字体必须确实存在。清理正文段落继承的缩进和分页设置，移除与指定字体冲突的主题字体引用。保留红色文件头、图形和域段落。

## 能力范围

脚本检查的是预设规则，不能完整判断语法、事实和数据口径；仍需助手结合上下文复核。未命中规则不代表没有错误。

版式参数参考 GB/T 9704-2012 和常见材料样式，但自动生成不保证整篇文件完整达标。复杂红头、版记、机关署名、表格及分页需要人工或渲染检查。

Windows 通过注册表检测字体，macOS 读取已安装字体文件。缺少标准字体时会使用已检测到的替代字体并说明；没有可用中文字体时停止排版。不会自动下载商业字体。

`.doc` 自动转换和精确页码定位依赖 Windows COM。Mac 上建议直接提供 `.docx`，用段号定位问题。

处理原稿会另存输出，不覆盖原稿。脚本未实现联网发送；通过 AI 助手使用时，文档内容的处理范围仍取决于该助手与服务配置。不要将真实文稿、账号配置、凭据或测试输出提交到公开仓库。

完成后成品 `.docx` 会直接呈现在对话中作为可点击交付物，而非只给磁盘路径；同时附简短说明（文种、关键改动、待确认问题）。

## 文件结构

`SKILL.md` 是完整技能入口；`AGENTS.md` 和 `INSTRUCTIONS.md` 提供其他助手入口。`scripts/` 包含格式、校对、轻量 DOCX 读写及可选页码映射；`references/` 包含版式参考；`tests/` 包含回归测试。
