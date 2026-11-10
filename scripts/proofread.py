# -*- coding: utf-8 -*-
"""
公文/文稿自动校对脚本 —— 零第三方依赖版。

依赖：仅 Python 标准库（docx_lite.py 提供 .docx 读写）。

用法：python proofread.py <输入.docx> [--fix]
- 不传 --fix：只检查，输出问题清单（不改文件）
- 传 --fix：明确的中文标点、首尾空格和有效日期可自动修复，
           大错（错别字/语病/用词/口径）只列清单，由用户确认后再改。
- 输出到原文件同目录，文件名加 _校对 后缀；原文件不动。
"""
import sys
import re
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_lite import Document, set_para_text, strip_para_spaces

# ---------- 错别字/用词 大错表（需人工确认，不自动改） ----------
# 注意：“截止”不在此表里：它是子串匹配会误报（“截止日期”“截止目前”
# 都是正确用法），见下面的 _jiezhi_problems()，只在接具体日期时才提示。
TYPO_MAP = [
    ("帐号", "账号", "错别字"),
    ("帐户", "账户", "错别字"),
    ("其它", "其他", "用词不规范"),
    ("既使", "即使", "错别字"),
    ("再接再励", "再接再厉", "错别字"),
    ("布署", "部署", "错别字"),
    ("安祥", "安详", "错别字"),
    ("好象", "好像", "错别字"),
    ("撤消", "撤销", "错别字"),
    ("复盖", "覆盖", "错别字"),
    ("成份", "成分", "用词不规范"),
    ("座落", "坐落", "错别字"),
    ("份量", "分量", "错别字"),
    ("辩别", "辨别", "错别字"),
    ("渲泄", "宣泄", "错别字"),
    ("精萃", "精粹", "错别字"),
    ("藉此", "借此", "用词不规范"),
]

# 常见语病模式（正则，命中即列清单）
GRAMMAR_PATTERNS = [
    (r'水平得到了完善', "搭配不当（「水平」用「提高」，不用「完善」）"),
    (r'进一步加强了.{0,6}的力度和深度', "搭配冗余（力度/深度可精简）"),
    (r'大概.{0,4}左右', "语义重复（「大概」与「左右」重复）"),
    (r'凯旋而归', "语义重复（凯旋已含「归」）"),
    (r'涉及到', "语义重复（「涉及」不含「到」，除非特指）"),
    (r'付诸于', "语义重复（「付诸」即含「于」）"),
    (r'目的是为了', "语义重复"),
]

# “通过…使…” / “经过…使…” 单独处理：只有“使”后面直接跟动词、
# 确实缺主语时才提示；“通过学习使大家提高了认识”里“大家”就是
# “使”的主语，不应误报。见 _shi_missing_subject()。
_SHI_PATTERNS = [
    (r'通过.{0,10}使', "句式杂糅（「通过…使…」缺主语）"),
    (r'经过.{0,10}使', "句式杂糅（「经过…使…」缺主语）"),
]

# “使”后直接跟这些动词时，基本可以判定缺主语（无名词主语）。
_SHI_MISSING_SUBJECT_VERBS = (
    "提高了", "提高", "取得了", "取得", "实现", "增强", "促进",
    "导致", "造成", "形成", "达到", "完成", "提升", "改善",
    "推动", "加强", "扩大", "缩小", "得到", "获得",
)


def _shi_missing_subject(text, match):
    """判断“通过…使…”是否真的缺主语。

    “使”后直接跟动词（“使提高了效率”）→ 缺主语，提示；
    “使”后跟名词/代词主语（“使大家提高了认识”）→ 不提示。"""
    after = text[match.end():].lstrip(" \t，,")
    if not after:
        return True
    return after.startswith(_SHI_MISSING_SUBJECT_VERBS)


def _jiezhi_problems(text):
    """“截止”只在接具体日期/时间时才提示改用“截至”。

    不提示的正确用法：“截止日期”“截止时间”“截止目前”“截止现在”
    （作名词/时间状语）。提示的用法：“截止2026年11月9日……”
    “截止到2026年……”“截止本月底……”（本应用“截至”）。"""
    problems = []
    for match in re.finditer(r'截止', text):
        after = text[match.end():]
        if after.startswith(("日期", "时间", "目前", "现在", "时", "为止")):
            continue
        if re.match(
            r'\s*(?:到|于)?\s*(?:\d{4}年|\d{4}[-./]\d{1,2}|\d{1,2}月\d{1,2}日'
            r'|本(?:月|年|周|月底)|下(?:月|年|周)|上(?:月|年|周)'
            r'|今年|去年|明年)',
            after,
        ):
            problems.append("「截止」→ 疑似「截至」（接时间应用「截至」）")
    return problems


def _technical_spans(text):
    """Protect ASCII technical tokens and quoted code before editing punctuation."""
    patterns = [r'`[^`]*`', r'(?:https?://|www\.)[^\s，。；！？<>]+',
                r'(?:[A-Za-z]:[\\/]|/)[^\s，。；！？<>]+',
                r'[A-Za-z0-9_@./\\:+?=&%#~!$*()\-]*[A-Za-z@][A-Za-z0-9_@./\\:+?=&%#~!$*()\-]*']
    return [(m.start(), m.end()) for pattern in patterns for m in re.finditer(pattern, text)]


def _inside(index, spans):
    return any(start <= index < end for start, end in spans)


def half_to_full(text):
    """半角标点转全角（中文语境）。句点仅在非数字语境转句号。

    Preserve technical tokens, paragraph numbering, decimals and English text.
    """
    spans = _technical_spans(text)
    spans += [(m.start(), m.end()) for m in re.finditer(r'^\s*\d+[.)、]|\d+(?:\.\d+)+', text)]
    simple = {',': '，', ';': '；', ':': '：', '?': '？', '!': '！'}
    out = list(text)
    cjk = lambda ch: bool(ch and re.match(r'[\u3400-\u9fff，。；：！？、）》」』]', ch))
    # Convert only matched parentheses whose interior contains Chinese.
    for match in re.finditer(r'\(([^()]*)\)', text):
        if re.search(r'[\u3400-\u9fff]', match.group(1)) and not any(
                _inside(i, spans) for i in (match.start(), match.end() - 1)):
            out[match.start()], out[match.end() - 1] = '（', '）'
    for i, ch in enumerate(text):
        if _inside(i, spans):
            continue
        prev = out[i - 1] if i else ''
        nxt = out[i + 1] if i + 1 < len(text) else ''
        if ch in simple and (cjk(prev) or cjk(nxt)):
            out[i] = simple[ch]
        elif ch == '.' and cjk(prev) and (not nxt or cjk(nxt) or nxt.isspace()):
            out[i] = '。'
    result = ''.join(out)
    return result, result != text


def fix_dates(text):
    """统一日期格式：2024.1.15 / 2024-1-15 → 2024年1月15日。

    Only valid standalone dates are changed; paths, URLs and identifiers stay intact.
    """
    spans = _technical_spans(text)
    def repl(match):
        if any(_inside(i, spans) for i in range(match.start(), match.end())):
            return match.group(0)
        y, mo, d = map(int, (match.group(1), match.group(3), match.group(4)))
        try:
            date(y, mo, d)
        except ValueError:
            return match.group(0)
        return f"{y}年{mo}月{d}日"
    result = re.sub(r'(?<![A-Za-z0-9_./\\-])(\d{4})([.\-/])(\d{1,2})\2(\d{1,2})(?![A-Za-z0-9_./\\-])', repl, text)
    return result, result != text


def check_and_report_detailed(text):
    """检查一段文本，返回 (小错自动修后的文本, [(severity, 问题), ...])。

    severity 按规则类型分，不按问题文案里的字眼分：
    major = 错别字/用词/语病（需人工确认，不自动改）；
    minor = 日期格式/半角标点/首尾空格/标点重复（可自动修或仅提示）。"""
    problems = []  # list of (severity, message)
    t = text

    # --- 大错：错别字/用词 ---
    for wrong, right, kind in TYPO_MAP:
        if wrong in t:
            problems.append(("major", f"「{wrong}」→ 疑似「{right}」（{kind}）"))

    # --- 大错：“截止”只在接具体日期时提示 ---
    for prob in _jiezhi_problems(t):
        problems.append(("major", prob))

    # --- 大错：语病 ---
    for pat, desc in GRAMMAR_PATTERNS:
        m = re.search(pat, t)
        if m:
            problems.append(("major", f"「{m.group(0)}」— {desc}"))

    # --- 大错：“通过/经过…使…”只在确实缺主语时提示 ---
    for pat, desc in _SHI_PATTERNS:
        m = re.search(pat, t)
        if m and _shi_missing_subject(t, m):
            problems.append(("major", f"「{m.group(0)}」— {desc}"))

    # --- 小错：日期格式 ---
    fixed, changed = fix_dates(t)
    if changed:
        problems.append(("minor", "日期格式可自动统一为「年月日」"))
        t = fixed
    # --- 小错：半角标点 ---
    fixed, changed = half_to_full(t)
    if changed:
        problems.append(("minor", "中文语境半角标点可自动转全角"))
        t = fixed

    # --- 小错：段首/段尾残留空格 ---
    if t != t.strip(" \t\u3000"):
        problems.append(("minor", "段落首尾残留空格（将自动清理）"))

    # --- 小错：标点重复检测 ---
    for dup in ["。。", "，，", "、、", "；；", "！！", "？？"]:
        if dup in t:
            problems.append(("minor", f"标点重复「{dup}」"))

    return t, problems


def check_and_report(text):
    """检查一段文本，返回 (小错自动修后的文本, 问题清单列表)。"""
    t, detailed = check_and_report_detailed(text)
    return t, [msg for _severity, msg in detailed]


def main():
    args = sys.argv[1:]
    if not args:
        print("用法：python proofread.py <输入.docx> [--fix]")
        sys.exit(1)

    src = Path(args[0])
    do_fix = "--fix" in args
    if not src.exists():
        print(f"[错误] 文件不存在：{src}")
        sys.exit(1)
    if src.suffix.lower() == ".doc":
        print("[错误] 请先用 format_gongwen.py 或 Word/WPS 把 .doc 转成 .docx")
        sys.exit(1)

    doc = Document(str(src))
    paragraphs = doc.paragraphs()

    all_problems = []  # list of (severity, formatted_message)
    small_fixed_count = 0
    para_no = 0

    print("[范围] 仅规则检查；复杂语病、事实和数据口径仍需结合上下文复核。")
    for p in paragraphs:
        t = Document.para_text(p)
        if not t.strip():
            continue
        para_no += 1
        new_t, detailed = check_and_report_detailed(t)
        for severity, prob in detailed:
            all_problems.append((severity, f"第{para_no}段「{t.strip()[:20]}…」：{prob}"))

        # Avoid rewriting fields, hyperlinks, drawings or revision markup automatically.
        complex_tags = {"fldChar", "instrText", "fldSimple", "hyperlink", "drawing", "pict", "ins", "del", "tab", "br"}
        complex_para = any(e.tag.rsplit('}', 1)[-1] in complex_tags for e in p.iter())
        if do_fix and not complex_para:
            if new_t != t:
                set_para_text(p, new_t)
                small_fixed_count += 1
            if strip_para_spaces(p):
                small_fixed_count += 1

        elif do_fix and complex_para and new_t != t:
            all_problems.append(("major", f"第{para_no}段：含域、超链接或复杂内容，仅提示，不自动改写"))

    # Table text is reviewed without changing table structure or cell formatting.
    for table_no, rows in enumerate(doc.table_texts(), 1):
        for row_no, row in enumerate(rows, 1):
            for col_no, text in enumerate(row, 1):
                _, detailed = check_and_report_detailed(text)
                for severity, problem in detailed:
                    all_problems.append((severity, f"表{table_no}行{row_no}列{col_no}：{problem}〔表格仅提示〕"))

    if do_fix:
        dst = src.with_name(src.stem + "_校对.docx")
        doc.save(str(dst))
        print(f"[完成] 小错已自动修复并保存：{dst}")
        print(f"[小错自动修复] {small_fixed_count} 处")
    else:
        dst = src

    # 大错/小错按规则类型分，不按问题文案里的字眼分（旧逻辑按“自动”
    # 二字分流，会把“标点重复”等小错划进大错）。表格里的日期问题
    # 同样按规则类型分，即使它只提示不自动修。
    big = [msg for severity, msg in all_problems if severity == "major"]
    small = [msg for severity, msg in all_problems if severity == "minor"]

    print(f"\n【大错清单（需你确认后修改）】共 {len(big)} 处")
    for x in big:
        print("  -", x)
    print(f"\n【小错（已自动修复/可自动修复）】共 {len(small)} 处")
    for x in small:
        print("  -", x)

    if not big and not small:
        print("\n（未发现明显问题）")


if __name__ == "__main__":
    main()
