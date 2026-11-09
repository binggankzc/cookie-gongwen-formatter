# -*- coding: utf-8 -*-
"""
公文/文稿自动校对脚本 —— 零第三方依赖版。

依赖：仅 Python 标准库（docx_lite.py 提供 .docx 读写）。

用法：python proofread.py <输入.docx> [--fix]
- 不传 --fix：只检查，输出问题清单（不改文件）
- 传 --fix：小错（标点全角、残留空格、日期格式）自动修好，
           大错（错别字/语病/用词/口径）只列清单，由用户确认后再改。
- 输出到原文件同目录，文件名加 _校对 后缀；原文件不动。
"""
import sys
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_lite import Document, set_para_text, strip_para_spaces

# ---------- 错别字/用词 大错表（需人工确认，不自动改） ----------
TYPO_MAP = [
    ("帐号", "账号", "错别字"),
    ("帐户", "账户", "错别字"),
    ("截止", None, "用词（接时间应用「截至」）"),
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
    (r'通过.{0,10}使', "句式杂糅（「通过…使…」缺主语）"),
    (r'经过.{0,10}使', "句式杂糅（「经过…使…」缺主语）"),
]


def half_to_full(text):
    """半角标点转全角（中文语境）。句点仅在非数字语境转句号。"""
    simple = {
        ',': '，', ';': '；', ':': '：', '?': '？', '!': '！',
        '(': '（', ')': '）',
    }
    changed = False
    out = []
    n = len(text)
    for i, ch in enumerate(text):
        prev = text[i - 1] if i > 0 else ''
        nxt = text[i + 1] if i < n - 1 else ''
        if ch in simple:
            out.append(simple[ch])
            changed = True
        elif ch == '.':
            if prev.isdigit() and nxt.isdigit():
                out.append('.')
            else:
                out.append('。')
                changed = True
        else:
            out.append(ch)
    return "".join(out), changed


def fix_dates(text):
    """统一日期格式：2024.1.15 / 2024-1-15 → 2024年1月15日。"""
    def repl(m):
        y, mo, d = m.group(1), m.group(2), m.group(3)
        return f"{y}年{int(mo)}月{int(d)}日"
    new, n = re.subn(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', repl, text)
    return new, n > 0


def check_and_report(text):
    """检查一段文本，返回 (小错自动修后的文本, 问题清单列表)。"""
    problems = []
    t = text

    # --- 大错：错别字/用词 ---
    for wrong, right, kind in TYPO_MAP:
        if wrong in t:
            problems.append(f"「{wrong}」→ 疑似「{right if right else '截至'}」（{kind}）")

    # --- 大错：语病 ---
    for pat, desc in GRAMMAR_PATTERNS:
        m = re.search(pat, t)
        if m:
            problems.append(f"「{m.group(0)}」— {desc}")

    # --- 小错：半角标点 ---
    fixed, changed = half_to_full(t)
    if changed:
        problems.append("半角标点已自动转全角")
        t = fixed

    # --- 小错：日期格式 ---
    fixed, changed = fix_dates(t)
    if changed:
        problems.append("日期格式已自动统一为「年月日」")
        t = fixed

    # --- 小错：段首/段尾残留空格 ---
    if t != t.strip(" \t\u3000"):
        problems.append("段落首尾残留空格（将自动清理）")

    # --- 标点重复检测 ---
    for dup in ["。。", "，，", "、、", "；；", "！！", "？？"]:
        if dup in t:
            problems.append(f"标点重复「{dup}」")

    return t, problems


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

    all_problems = []
    small_fixed_count = 0
    para_no = 0

    for p in paragraphs:
        t = Document.para_text(p)
        if not t.strip():
            continue
        para_no += 1
        new_t, problems = check_and_report(t)
        for prob in problems:
            all_problems.append(f"第{para_no}段「{t.strip()[:20]}…」：{prob}")

        if do_fix:
            if new_t != t:
                set_para_text(p, new_t)
                small_fixed_count += 1
            if strip_para_spaces(p):
                small_fixed_count += 1

    if do_fix:
        dst = src.with_name(src.stem + "_校对.docx")
        doc.save(str(dst))
        print(f"[完成] 小错已自动修复并保存：{dst}")
        print(f"[小错自动修复] {small_fixed_count} 处")
    else:
        dst = src

    big = [x for x in all_problems if "自动" not in x and "将自动" not in x]
    small = [x for x in all_problems if "自动" in x]

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
