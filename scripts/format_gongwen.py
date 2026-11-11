# -*- coding: utf-8 -*-
"""
公文/文档格式改写脚本（GB/T 9704-2012 + 常见文种分流）—— 零第三方依赖版。

依赖：仅 Python 标准库（docx_lite.py 提供 .docx 读写）。
若环境恰好有 python-docx，本脚本也不依赖它；.doc 转换走 Word/WPS COM（可选）。

用法：python format_gongwen.py <输入.docx|.doc> [输出.docx] [--mode 类型]
- mode 可选：auto(默认，自动判断) / gongwen / speech / summary / table
- 只处理 w:p 段落，遇到 w:tbl 直接跳过（避免表格变形）
- 输出默认：原文件名_公文格式.docx
"""
import sys
import re
import struct
import subprocess
from functools import lru_cache
from pathlib import Path

# 自研零依赖 docx 核心（同目录）
sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_lite import (
    Document, set_para_format, set_para_text, strip_para_spaces,
    set_page_margins, get_content_width_twips, scale_table_widths,
    add_page_number, cm_to_twips, _w, find_child,
)

# 字体名
FONT_BIAOTI = "方正小标宋简体"   # 标题
FONT_FANGSONG = "仿宋_GB2312"    # 正文
FONT_HEITI = "黑体"              # 一级标题
FONT_KAITI = "楷体_GB2312"       # 二级标题

# 字体降级映射：标准字体 → 系统自带替代（目标电脑缺字体时用）
FONT_FALLBACK = {
    "方正小标宋简体": "华文中宋",   # 标题
    "仿宋_GB2312": "仿宋",          # 正文
    "楷体_GB2312": "楷体",          # 二级标题
}

# 字号（磅）
SIZE_TITLE = 22.0   # 二号
SIZE_BODY = 16.0    # 三号

# 行距（固定值，统一 30 磅）
LINE_PT = 30.0


def _font_family_names(path):
    """Read SFNT/TTC family names using the standard library, including Chinese aliases."""
    names = set()
    try:
        with path.open('rb') as stream:
            header = stream.read(12)
            if header[:4] == b'ttcf':
                count = struct.unpack('>I', header[8:12])[0]
                if count > 1024:
                    return names
                offsets = struct.unpack('>' + 'I' * count, stream.read(4 * count))
            else:
                offsets = (0,)
            for offset in offsets:
                stream.seek(offset + 4)
                table_count = struct.unpack('>H', stream.read(2))[0]
                if table_count > 512:
                    continue
                stream.seek(offset + 12)
                records = [stream.read(16) for _ in range(table_count)]
                for record in records:
                    tag, _, start, length = struct.unpack('>4sIII', record)
                    if tag != b'name' or length > 4 * 1024 * 1024:
                        continue
                    stream.seek(start)
                    data = stream.read(length)
                    _, count, strings = struct.unpack('>HHH', data[:6])
                    for index in range(count):
                        record = data[6 + index * 12:18 + index * 12]
                        platform, encoding, language, name_id, size, pos = struct.unpack('>HHHHHH', record)
                        if name_id not in (1, 16):
                            continue
                        raw = data[strings + pos:strings + pos + size]
                        codec = 'utf-16-be' if platform in (0, 3) else 'mac_roman'
                        name = raw.decode(codec, errors='ignore').strip('\x00 ')
                        if name:
                            names.add(name)
    except (OSError, ValueError, struct.error):
        pass
    return names


def _parse_fc_list_output(text):
    """解析 `fc-list` 的输出，返回字体家族名集合。

    每行形如：
        /path/to/font.ttf: Family Name,Family Name Bold:style=Regular
    冒号后的第二个字段是逗号分隔的家族名；文件 stem 也一并加入
    （与 Windows 分支同时记录注册表名和文件名的做法一致）。
    无法解析的行会被跳过，不抛异常。
    """
    fonts = set()
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split(":", 2)
        if len(parts) < 2:
            continue
        path, families = parts[0].strip(), parts[1]
        for name in families.split(","):
            name = name.strip()
            if name:
                fonts.add(name)
        stem = Path(path).stem
        if stem:
            fonts.add(stem)
    return fonts


def _fonts_from_font_dirs(directories):
    """fontconfig 不可用时的退路：直接扫描常见字体目录。

    复用 _font_family_names() 读取字体文件的家族名（含中文别名），
    避免在没有 fc-list 的 Linux 环境上直接返回空集。
    """
    fonts = set()
    for directory in directories:
        if not directory.exists():
            continue
        for path in directory.rglob("*"):
            if path.suffix.lower() in (".ttf", ".otf", ".ttc", ".otc"):
                fonts.update(_font_family_names(path))
    return fonts


@lru_cache(maxsize=1)
def get_installed_fonts():
    """读取系统已安装字体列表（macOS 扫字体目录；Linux 用 fc-list；Windows 注册表 HKLM + HKCU）。
    返回集合（含字体家族名和字体文件名）。失败返回空集。"""
    if sys.platform.startswith("linux"):
        try:
            result = subprocess.run(
                ["fc-list"], capture_output=True, text=True,
                errors="replace", timeout=30,
            )
            if result.returncode == 0 and result.stdout.strip():
                return _parse_fc_list_output(result.stdout)
        except (OSError, subprocess.SubprocessError):
            pass
        # fontconfig 不可用或无输出时，退回直接扫描常见字体目录。
        return _fonts_from_font_dirs((
            Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
            Path.home() / ".fonts", Path.home() / ".local/share/fonts",
        ))
    if sys.platform == "darwin":
        fonts = set()
        for directory in (Path('/System/Library/Fonts'), Path('/Library/Fonts'), Path.home() / 'Library/Fonts'):
            if directory.exists():
                for path in directory.rglob('*'):
                    if path.suffix.lower() in ('.ttf', '.otf', '.ttc', '.otc'):
                        fonts.update(_font_family_names(path))
        return fonts
    try:
        import winreg
        fonts = set()
        paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ]
        for hive, subkey in paths:
            try:
                key = winreg.OpenKey(hive, subkey)
                try:
                    i = 0
                    while True:
                        name, value, _type = winreg.EnumValue(key, i)
                        clean = re.sub(r'\s*\(.*?\)\s*$', '', name)
                        clean = re.sub(r'\.(ttf|ttc|otf)$', '', clean, flags=re.IGNORECASE)
                        fonts.add(clean.strip())
                        fname = Path(str(value)).stem
                        fonts.add(fname)
                        i += 1
                except OSError:
                    pass
            except OSError:
                pass
        return fonts
    except Exception:
        return set()


def font_available(font_name, installed):
    """判断字体是否可用：精确匹配优先；避免短名字串反向证明长名可用。"""
    if font_name in installed:
        return True
    core = re.sub(r'[_0-9]', '', font_name)
    if not core:
        return False
    for f in installed:
        fk = re.sub(r'[_0-9]', '', f)
        if fk == core:
            if '_' in font_name and '_' not in f and 'GB' not in f:
                continue
            return True
    return False


def resolve_fonts():
    """检测已装字体，返回 (字体映射 dict, 降级说明 dict)。"""
    installed = get_installed_fonts()
    choices = {
        # 每组最后一个是 Linux 常见中文字体（Noto CJK）兜底：只有前面
        # 的标准/常见字体都不存在时才会用到，且仍会记入 fallback_notes 说明。
        # 若连兜底也没有，下面的循环仍会 raise RuntimeError 停止改写。
        'title': (FONT_BIAOTI, '华文中宋', 'Songti SC', '宋体-简', '宋体', 'Arial Unicode MS', 'Noto Serif CJK SC'),
        'body': (FONT_FANGSONG, '仿宋', 'STFangsong', 'Songti SC', '宋体-简', '宋体', 'Arial Unicode MS', 'Noto Serif CJK SC'),
        'heiti': (FONT_HEITI, '黑体-简', 'Heiti SC', 'STHeiti', 'Arial Unicode MS', 'Noto Sans CJK SC'),
        'kaiti': (FONT_KAITI, '楷体', 'Kaiti SC', '楷体-简', 'STKaiti', 'Songti SC', '宋体-简', 'Arial Unicode MS', 'Noto Serif CJK SC'),
    }
    font_map, fallback_notes = {}, {}
    for key, candidates in choices.items():
        selected = next((font for font in candidates if font_available(font, installed)), None)
        if selected is None:
            raise RuntimeError('未检测到可用中文字体，停止格式改写；请安装字体或仅做内容校对。')
        font_map[key] = selected
        if selected != candidates[0]:
            fallback_notes[candidates[0]] = selected
    return font_map, fallback_notes


_CN_NUM = "一二三四五六七八九十百千零两"


def classify(text):
    """判断段落属于哪类：h1 / h2 / h3 / h4 / body。

    层级（GB/T 9704-2012 惯用写法）：
      h1: 「一、」「一百、」…（中文数字 + 顿号）
      h2: 「（一）」「(一)」…（全角或半角括号 + 中文数字）
      h3: 「1.」「2.」…（阿拉伯数字 + 点）
      h4: 「（1）」「(1)」…（全角或半角括号 + 1-2 位阿拉伯数字）
    h4 只认 1-2 位数字：行首出现「(2026)」这类年份括号时不误判为标题。
    括号序号不在行首时（如「根据(一)款规定」）仍是 body。
    """
    if re.match(r'^[' + _CN_NUM + r']+、', text):
        return "h1"
    if re.match(r'^[（(][' + _CN_NUM + r']+[)）]', text):
        return "h2"
    if re.match(r'^\d+[\.、]', text):
        return "h3"
    if re.match(r'^[（(]\d{1,2}[)）]', text):
        return "h4"
    return "body"


def is_salutation(text):
    """判断是否为称谓/主送机关行（顶格、以冒号结尾的短行）。"""
    t = text.strip()
    if not t.endswith('：'):
        return False
    if len(t) > 40:
        return False
    return True


def detect_doc_type(text):
    """识别文种：返回文种名或 None（仅针对正式公文）。"""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return None
    # The title suffix outranks keywords quoted in the title or body.
    title = lines[0].rstrip('。！？： ')
    match = re.search(r'(通知|报告|请示|函|纪要|决定|意见|通报|批复)$', title)
    if match:
        return match.group(1)
    # Without a conclusive title, use the closing formula before body references.
    ending = '\n'.join(lines[-2:])
    for kind, pattern in [('请示', r'妥否[，,]请批示|请予审批'), ('报告', r'特此报告|以上报告.*审阅'),
                          ('通知', r'特此通知'), ('函', r'特此函告|敬请函复'), ('批复', r'此复|特此批复')]:
        if re.search(pattern, ending):
            return kind
    patterns = [
        ("通知", ["的通知", "现将有关事项通知如下", "特此通知"]),
        ("报告", ["的报告", "现将有关情况报告如下", "特此报告", "以上报告"]),
        ("请示", ["的请示", "妥否，请批示", "请予审批", "请示如下"]),
        ("函", ["的函", "特此函告", "敬请函复", "函复如下", "特此函商"]),
        ("纪要", ["会议纪要", "纪要"]),
        ("决定", ["的决定", "经研究，决定如下", "本决定自"]),
        ("意见", ["的意见", "现提出如下意见"]),
        ("通报", ["的通报", "现将有关情况通报如下", "特此通报"]),
        ("批复", ["的批复", "收悉", "此复", "特此批复"]),
    ]
    for name, kws in patterns:
        for kw in kws:
            if kw in text:
                return name
    return None


_SPEECH_KEYWORDS = ["讲话", "发言", "致辞", "演讲", "动员", "辅导"]
_SUMMARY_KEYWORDS = ["总结", "汇报", "体会", "心得", "述职", "情况报告", "自查"]


def detect_category_with_warning(doc):
    """判断文件大类，返回 (category, warning_or_None)。

    优先级：标题末尾的文种（通知/决定/请示/报告等，见 detect_doc_type）
    高于标题关键词。关键词（“动员”“汇报”等）只在没有明确文种时
    才作为启发式依据，且此时置信度低，会返回一条“疑似…，请确认”
    提示，由调用方打印，而不是静默误判。
    """
    paragraphs = doc.paragraphs()
    texts = [Document.para_text(p) for p in paragraphs]
    text_all = "\n".join(texts)
    n_tables = len(doc.tables())
    n_nonempty_p = sum(1 for t in texts if t.strip())

    if n_tables >= 1 and n_nonempty_p <= n_tables * 3:
        return "table", None

    # 标题可能多行（主标题 + 副标题「XX讲话稿」），看前 3 行非空文本
    head_lines = [t.strip() for t in texts if t.strip()][:3]
    title_joined = " ".join(head_lines)
    speech_kw = [k for k in _SPEECH_KEYWORDS if k in title_joined]
    summary_kw = [k for k in _SUMMARY_KEYWORDS if k in title_joined]

    # 1. 标题末尾的文种优先：例如“关于动员部署工作的通知”是通知，
    #    不是讲话稿；“关于汇报工作的通知”是通知，不是总结。
    doc_type = detect_doc_type(text_all)
    if doc_type is not None:
        if speech_kw or summary_kw:
            kw = "、".join(speech_kw + summary_kw)
            return "gongwen", (
                f"标题含「{kw}」等词，但文种为「{doc_type}」，"
                "已按正式公文处理，请确认")
        return "gongwen", None

    # 2. 无明确文种时，才用标题关键词启发式判断，置信度低，需提示。
    if speech_kw:
        return "speech", (
            f"疑似 speech：仅凭标题关键词「{speech_kw[0]}」判断，"
            "未见明确文种，请确认")
    if summary_kw:
        return "summary", (
            f"疑似 summary：仅凭标题关键词「{summary_kw[0]}」判断，"
            "未见明确文种，请确认")
    if re.search(r'^[^\n]{0,30}：$', text_all, re.M):
        return "gongwen", None
    return "summary", "未识别到明确文种或类型关键词，已按 summary 处理，请确认"


def detect_category(doc):
    """判断文件大类：gongwen / speech / summary / table。"""
    category, _warning = detect_category_with_warning(doc)
    return category


def convert_doc_to_docx(src):
    """把 .doc 转成 .docx（走 Word/WPS COM）。返回 .docx 路径，失败返回 None。"""
    try:
        import win32com.client
        import pythoncom
        import shutil
    except ImportError:
        return None
    # 中文文件名会触发 WPS 拦截，复制成 ASCII 名
    tmp_doc = src.with_name("__conv_tmp.doc")
    tmp_docx = src.with_name("__conv_tmp.docx")
    shutil.copy2(src, tmp_doc)
    try:
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.DisplayAlerts = 0
            d = word.Documents.Open(str(tmp_doc), ReadOnly=True)
            d.SaveAs(str(tmp_docx), FileFormat=12)  # 12 = wdFormatXMLDocument (.docx)
            d.Close(False)
            if tmp_docx.exists():
                return tmp_docx
        finally:
            try:
                word.Quit()
            except Exception:
                pass
    except Exception:
        return None
    finally:
        try:
            tmp_doc.unlink(missing_ok=True)
        except Exception:
            pass
    return None


def _is_red_color(value):
    """判断一个 w:color 值是否为红色（红分量高且明显高于绿/蓝）。"""
    if not re.fullmatch(r'[0-9A-Fa-f]{6}', value or ''):
        return False
    red, green, blue = (int(value[i:i+2], 16) for i in (0, 2, 4))
    return red > 150 and red > green * 1.4 and red > blue * 1.4


def _protected_paragraph(p):
    """整段保护的段落：含图形/图片，或红色文字占段落主体。

    注意：域代码（fldChar/instrText/fldSimple，如超链接、页码、目录）不整段保护，
    因为设置字体/字号/缩进/行距这些 run 属性和段落属性不会破坏域结构；
    含超链接的正文段仍应排版。文本改写（校对）环节才对域代码做保护（见 proofread.py）。

    红色保护只在红色文字占段落文本一半及以上时才触发：单个红色强调词
    不应让整段跳过排版；整段都是红色（红头/文件头）时才原样保留。
    """
    for element in p.iter():
        if element.tag.rsplit('}', 1)[-1] in ('drawing', 'pict', 'object'):
            return True
    total_len = 0
    red_len = 0
    for run in p.iter(_w('r')):
        text = "".join(t.text or "" for t in run.iter(_w('t')))
        if not text:
            continue
        total_len += len(text)
        rpr = find_child(run, 'rPr')
        if rpr is None:
            continue
        color = find_child(rpr, 'color')
        if color is not None and _is_red_color(color.get(_w('val'), '')):
            red_len += len(text)
    return total_len > 0 and red_len * 2 >= total_len


def _normalize_layout(doc, p, title=False, kind='body'):
    """Override inherited paragraph layout without deleting styles or body emphasis."""
    set_para_format(doc, p, indent_chars=0 if title else 2, space_before=0, space_after=0,
                    paging={key: '0' for key in ('widowControl', 'keepNext', 'keepLines', 'pageBreakBefore')},
                    linebreak={'kinsoku': '1', 'overflowPunct': '1', 'wordWrap': '0'})
    props = doc.pPr(p)
    ind = find_child(props, 'ind')
    for name in ('left', 'right', 'leftChars', 'rightChars', 'hanging', 'hangingChars'):
        if name in ('left', 'right', 'leftChars', 'rightChars'):
            ind.set(_w(name), '0')
        else:
            ind.attrib.pop(_w(name), None)
    # Explicitly suppress inherited bold, while keeping direct body emphasis.
    for run in p.iter(_w('r')):
        rpr = doc.rPr(run)
        if not title and kind == 'body' and find_child(rpr, 'b') is None:
            import xml.etree.ElementTree as ET
            ET.SubElement(rpr, _w('b')).set(_w('val'), '0')


def parse_args(argv):
    """解析命令行参数（argparse，标准库）。

    支持 --mode gongwen 和 --mode=gongwen 两种写法；--mode 缺值、
    模式非法或出现未知参数时，argparse 直接给出用法并以 exit 2
    退出，不抛 traceback。"""
    import argparse
    parser = argparse.ArgumentParser(
        prog="format_gongwen.py",
        description="公文/文档格式改写脚本（GB/T 9704-2012 + 常见文种分流）",
    )
    parser.add_argument("input", help="输入文件（.docx 或 .doc）")
    parser.add_argument("output", nargs="?", default=None,
                        help="输出文件（默认：原文件名_公文格式.docx）")
    parser.add_argument("--mode", default="auto",
                        choices=["auto", "gongwen", "speech", "summary", "table"],
                        help="处理模式（默认 auto，自动判断）")
    parser.add_argument("--inspect", action="store_true",
                        help="只读识别，不生成或改写文件")
    return parser.parse_args(argv)


def _main_impl():
    args = parse_args(sys.argv[1:])
    mode = args.mode
    inspect_only = args.inspect

    src = Path(args.input)
    original_src = src
    dst = Path(args.output) if args.output else None

    if not src.exists():
        print(f"[错误] 文件不存在：{src}")
        sys.exit(1)

    # .doc → 先转 .docx（零依赖脚本本身不处理 .doc，走 COM 可选转换）
    converted_tmp = None
    if src.suffix.lower() == ".doc":
        tmp_docx = convert_doc_to_docx(src)
        if tmp_docx is None:
            print("[错误] .doc 转 .docx 失败（本机无 Word/WPS 或 COM 不可用）。")
            print("       请先用 Word/WPS 另存为 .docx 后再处理。")
            sys.exit(1)
        converted_tmp = tmp_docx
        src = tmp_docx

    if dst is None:
        # 默认输出名按原文件（.doc 转来的也按原 .doc 的文件名），
        # 不能按临时转换文件 __conv_tmp.docx 的 stem。
        dst = original_src.with_name(original_src.stem + "_公文格式.docx")

    if dst.resolve() == original_src.resolve() or dst.resolve() == src.resolve():
        print("[错误] 输出路径不能与原文件相同")
        sys.exit(1)
    doc = Document(str(src))
    original_tables = doc.table_texts()

    # 判断大类（auto 时才自动判断；已指定 --mode 时直接用指定值）
    if mode == "auto":
        category, category_warning = detect_category_with_warning(doc)
    else:
        category, category_warning = mode, None
    print(f"[类型] {category}")
    if category_warning:
        print(f"[提示] {category_warning}")
    if inspect_only:
        print(f"[文种] {detect_doc_type(doc.full_text()) or '未明确'}")
        print("[只读检查] 未生成或改写文件")
        return

    # 表格类：不改格式，只输出文本供校对
    if category == "table":
        print("[说明] 表格类材料，不修改格式，仅输出内容供校对")
        print("[正文文本]")
        for p in doc.paragraphs():
            t = Document.para_text(p).strip()
            if t:
                print(t)
        print("[表格文本]")
        for rows in doc.table_texts():
            for row in rows:
                print(" | ".join(row))
        print("[未生成新文件，原文件未改动]")
        sys.exit(0)

    # 字体检测 + 降级
    font_map, fallback_notes = resolve_fonts()
    F_BODY = font_map["body"]
    F_TITLE = font_map["title"]
    F_HEITI = font_map["heiti"]
    F_KAITI = font_map["kaiti"]

    # 按大类定版式参数
    if category == "gongwen":
        margin = (3.7, 3.5, 2.8, 2.6)
    else:  # speech / summary
        margin = (2.8, 2.8, 2.6, 2.6)

    # 记旧正文宽度
    old_width = get_content_width_twips(doc)

    # 记录旧页边距描述
    old_sec_desc = _describe_margins(doc)

    # 改页边距 + 缩放表格
    set_page_margins(doc, margin[0], margin[1], margin[2], margin[3])
    new_width = get_content_width_twips(doc)

    if old_width and new_width and doc.tables():
        n = scale_table_widths(doc, old_width, new_width)
        if n:
            print(f"[表格] 已等比例缩放 {n} 个表格的列宽")
        else:
            print("[表格] 表格为自动/百分比宽度，无需缩放")

    para_changes = []

    # 标题块识别
    paragraphs = doc.paragraphs()
    title_indices = []
    for i, p in enumerate(paragraphs):
        t = Document.para_text(p).strip()
        if not t:
            continue
        if _protected_paragraph(p):
            continue
        if ((not title_indices and len(t) <= 80 and not re.search(r'[。；，！？：.!?:]$', t))
                or (title_indices and (t.startswith('——') or re.match(r'^[（(]\d{4}年.*[)）]$', t)))):
            title_indices.append(i)
        else:
            break
    date_re = re.compile(r'^[（(].*?[)）]$')
    subtitle_re = re.compile(r'^——')

    # 非空段落序号映射
    para_no = {}
    _n = 0
    for p in paragraphs:
        if Document.para_text(p).strip():
            _n += 1
            para_no[id(p)] = _n

    # 遍历所有顶层段落（enumerate 一次取下标：原来每段调两次
    # paragraphs.index(p)，是 O(n²)；index 还会在重复段落上找错位置）。
    for idx, p in enumerate(paragraphs):
        text = Document.para_text(p).strip()
        if not text:
            continue
        is_title = idx in title_indices
        is_date = bool(date_re.match(text))
        is_subtitle = bool(subtitle_re.match(text))

        _no = para_no.get(id(p), 0)
        if _protected_paragraph(p):
            para_changes.append({"idx": _no, "text": text, "what": "红色文件头或图形段落：原样保留"})
            continue
        kind = classify(text)
        _normalize_layout(doc, p, title=is_title, kind=kind)

        if is_title:
            strip_para_spaces(p)
            if is_date:
                set_para_format(doc, p, font_ea=F_KAITI, font_ascii=F_KAITI,
                                size_pt=SIZE_BODY, bold=False, align="center", line_pt=LINE_PT, indent_chars=0)
                para_changes.append({"idx": _no, "text": text, "what": "日期行：楷体三号、居中"})
            elif is_subtitle:
                set_para_format(doc, p, font_ea=F_KAITI, font_ascii=F_KAITI,
                                size_pt=SIZE_BODY, bold=False, align="center", line_pt=LINE_PT, indent_chars=0)
                para_changes.append({"idx": _no, "text": text, "what": "副标题：楷体三号、居中"})
            elif idx == title_indices[0]:
                set_para_format(doc, p, font_ea=F_TITLE, font_ascii=F_TITLE,
                                size_pt=SIZE_TITLE, bold=False, align="center", line_pt=LINE_PT, indent_chars=0)
                para_changes.append({"idx": _no, "text": text, "what": "主标题：小标宋二号、居中"})
            else:
                set_para_format(doc, p, font_ea=F_TITLE, font_ascii=F_TITLE,
                                size_pt=SIZE_TITLE, bold=False, align="center", line_pt=LINE_PT, indent_chars=0)
                para_changes.append({"idx": _no, "text": text, "what": "标题续行：小标宋二号、居中"})
        else:
            if is_salutation(text):
                strip_para_spaces(p)
                set_para_format(doc, p, font_ea=F_BODY, font_ascii=F_BODY,
                                size_pt=SIZE_BODY, align="justify", line_pt=LINE_PT, indent_chars=0)
                para_changes.append({"idx": _no, "text": text, "what": "称谓：顶格（不缩进）"})
                continue
            kind = classify(text)
            if kind == "h1":
                set_para_format(doc, p, font_ea=F_HEITI, font_ascii=F_BODY,
                                size_pt=SIZE_BODY, bold=False, align="justify", line_pt=LINE_PT)
                para_changes.append({"idx": _no, "text": text, "what": "一级标题：三号黑体"})
            elif kind == "h2":
                set_para_format(doc, p, font_ea=F_KAITI, font_ascii=F_BODY,
                                size_pt=SIZE_BODY, bold=False, align="justify", line_pt=LINE_PT)
                para_changes.append({"idx": _no, "text": text, "what": "二级标题：三号楷体"})
            elif kind == "h3":
                set_para_format(doc, p, font_ea=F_BODY, font_ascii=F_BODY,
                                size_pt=SIZE_BODY, bold=True, align="justify", line_pt=LINE_PT)
                para_changes.append({"idx": _no, "text": text, "what": "三级标题：三号仿宋加粗"})
            elif kind == "h4":
                set_para_format(doc, p, font_ea=F_BODY, font_ascii=F_BODY,
                                size_pt=SIZE_BODY, bold=False, align="justify", line_pt=LINE_PT)
                para_changes.append({"idx": _no, "text": text, "what": "四级标题：三号仿宋"})
            else:
                set_para_format(doc, p, font_ea=F_BODY, font_ascii=F_BODY,
                                size_pt=SIZE_BODY, align="justify", line_pt=LINE_PT)
                para_changes.append({"idx": _no, "text": text, "what": "正文：三号仿宋、首行缩进2字符、行距30磅"})

    # 页码
    page_added = add_page_number(doc)

    doc.save(str(dst))
    # Verify that width changes did not alter row/column content.
    checked = Document(str(dst))
    if checked.table_texts() != original_tables:
        # 校验失败：删掉已保存的坏输出文件，别留在磁盘上误导人。
        try:
            Path(dst).unlink()
        except OSError:
            pass
        raise RuntimeError('输出表格内容或行列结构校验失败，已删除该格式版，请勿使用')

    full_text = "\n".join(Document.para_text(p) for p in doc.paragraphs() if Document.para_text(p).strip())
    doc_type = detect_doc_type(full_text)

    print(f"[完成] 已保存：{dst}")
    print(f"[类型] {category}")
    print(f"[文种] {doc_type if doc_type else ('无（' + category + '类）')}")
    print("[改动摘要]")
    print(f"页边距：{old_sec_desc} → 上{margin[0]:.2f}/下{margin[1]:.2f}/左{margin[2]:.2f}/右{margin[3]:.2f}cm")
    if page_added:
        print("页码：已在页脚居中插入「— 1 —」")
    else:
        print("页码：页脚已有内容或未改动")
    if old_width and new_width and doc.tables():
        print("表格：已按新页宽等比例缩放列宽")
    if fallback_notes:
        print("[字体降级] 使用已检测到的替代字体，不代表完整符合标准公文字体要求。")
        for std, fall in fallback_notes.items():
            print(f"  缺「{std}」→ 已用「{fall}」替代")
    print("[逐条改动]")
    for c in para_changes:
        t = c["text"]
        if len(t) > 20:
            t = t[:20] + "…"
        print(f"第{c['idx']}段 | {t} | {c['what']}")
    print("[正文文本]")
    for p in doc.paragraphs():
        t = Document.para_text(p).strip()
        if t:
            print(t)

    # 清理临时转换文件
    if converted_tmp and converted_tmp.exists():
        try:
            converted_tmp.unlink()
        except Exception:
            pass


def main():
    """入口：调用 _main_impl，并在任何退出/报错路径上都清理 .doc 转换
    产生的临时文件 __conv_tmp.docx。

    _main_impl 内部只在成功路径末尾清理；中途 raise / sys.exit 时
    临时文件会泄漏到原文件目录，这里用 try/finally 兜底。
    Windows COM 转换本身无法在 Linux 上验证（见 M4，未确认）。"""
    argv = sys.argv[1:]
    tmp_path = None
    for arg in argv:
        if not arg.startswith("-") and Path(arg).suffix.lower() == ".doc":
            tmp_path = Path(arg).with_name("__conv_tmp.docx")
            break
    try:
        _main_impl()
    finally:
        if tmp_path is not None and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass


def _describe_margins(doc):
    """读取当前页边距描述（供改动摘要）。"""
    from docx_lite import _w
    sp = None
    for c in doc.body:
        if c.tag == _w("sectPr"):
            sp = c
            break
    if sp is None:
        return "未知"
    from docx_lite import find_child, twips_to_cm
    pgMar = find_child(sp, "pgMar")
    if pgMar is None:
        return "未知"
    top = twips_to_cm(int(pgMar.get(_w("top"), 0)))
    bottom = twips_to_cm(int(pgMar.get(_w("bottom"), 0)))
    left = twips_to_cm(int(pgMar.get(_w("left"), 0)))
    right = twips_to_cm(int(pgMar.get(_w("right"), 0)))
    return "上%.2f/下%.2f/左%.2f/右%.2fcm" % (top, bottom, left, right)


if __name__ == "__main__":
    main()
