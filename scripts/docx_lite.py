# -*- coding: utf-8 -*-
"""
docx_lite.py —— 零第三方依赖的 .docx 读写核心（仅用标准库 zipfile + xml）。

为什么存在：本技能要发布成可分享工具，别的 Agent 运行环境可能没装 python-docx。
.doc x 本质是「zip 打包的一堆 XML」，用标准库即可完成公文格式改写所需的全部操作，
从而去掉 python-docx 硬依赖。

支持的能力（只做公文格式化要用的，够用即可）：
  - 打开/保存 .docx（保留其它所有部件原样，只改 word/document.xml 与页脚）
  - 读段落文本、表格文本
  - 写段落：字体(中西文)、字号、加粗、对齐、行距(固定值)、首行缩进(字符)、
    分页/换行设置(孤行控制/段中不分页/与下段同页/段前分页/中文首尾/标点溢出/西文换行)
  - 页边距、纸张
  - 表格等比例缩放（tblW / tblGrid / tcW）
  - 页脚居中页码（PAGE 域）
  - 段落文本写回（改 run 文本）

设计约定：
  - 命名空间前缀 w: 对应 wordprocessingml；r: 对应 relationships；wp: 对应 drawing。
  - 所有元素操作走 ElementTree，元素用 {namespace}tag 完整名。
  - 遇异常抛出 DocxError，调用方自行处理降级。

作者：cookie。可随技能包一起分发，不依赖任何第三方库。
"""
import re
import shutil
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# 命名空间
# ---------------------------------------------------------------------------
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{%s}" % NS["w"]


def _register_namespaces():
    """注册 OOXML 命名空间前缀，保证 ElementTree 输出标准前缀 w:/r:/wp: 等。"""
    import xml.etree.ElementTree as ET
    ET.register_namespace("w", NS["w"])
    ET.register_namespace("r", NS["r"])
    ET.register_namespace("wp", NS["wp"])
    ET.register_namespace("m", NS["m"])


def _w(tag):
    """w:tbl -> {ns}tbl"""
    return "{%s}%s" % (NS["w"], tag)


def _extract_namespaces(doc_xml_bytes):
    """从 document.xml 原始字节提取根元素上声明的所有 xmlns:* 前缀→URI 映射。
    用正则提取（不依赖解析，因为可能有 ET 无法解析的边界情况）。"""
    import re
    text = doc_xml_bytes.decode("utf-8", errors="replace")
    # 定位根元素标签：从 <w:document（或任意 <X:document）开始到对应的第一个 >
    m_start = re.search(r"<(\w+):document\b", text)
    if m_start:
        i = m_start.start()
        j = text.find(">", i)
        root_tag = text[i:j + 1] if j != -1 else text[i:]
    else:
        # 兜底：找第一个元素标签（跳过 <?xml ...?> 声明）
        m = re.search(r"<\w+:", text)
        if m:
            i = m.start()
            j = text.find(">", i)
            root_tag = text[i:j + 1] if j != -1 else text[i:]
        else:
            root_tag = text
    ns_map = {}
    for prefix, uri in re.findall(r'xmlns:(\w+)\s*=\s*"([^"]*)"', root_tag):
        ns_map[prefix] = uri
    return ns_map


def _register_all_namespaces(ns_map):
    """把提取到的所有命名空间前缀注册到 ElementTree，供序列化时保留。
    跳过 ET 临时前缀 ns0/ns1/...，避免与标准前缀 w: 等冲突。"""
    import xml.etree.ElementTree as ET
    import re
    for prefix, uri in ns_map.items():
        if re.match(r'^ns\d+$', prefix):
            continue
        try:
            ET.register_namespace(prefix, uri)
        except Exception:
            pass


def _restore_root_namespaces(xml_bytes, ns_map):
    """在序列化后的 document.xml 里，把缺失的 xmlns:* 声明手动补回根元素。

    ElementTree 序列化会丢弃「未被元素/属性实际引用」的前缀声明
    （如 w14/w15/wp14 只出现在 mc:Ignorable="w14 w15 wp14" 字符串里），
    导致 mc:Ignorable 引用未声明前缀 → invalid token。这里用字符串方式补回。

    安全规则：
    - 跳过 ET 临时前缀 ns0/ns1/...（这些是 ElementTree 在未注册时的默认前缀，
      不应原样保留，否则会与标准前缀 w: 等产生同 URI 重复声明）。
    - 若某 URI 已被当前根标签里其它前缀声明，则不再重复补（避免 duplicate attribute）。
    """
    import re
    text = xml_bytes.decode("utf-8", errors="replace")
    # 定位根元素开标签（<w:document ... 到第一个 >）
    m = re.search(r"<(\w+):document\b[^>]*>", text)
    if not m:
        m = re.search(r"<\w+:[^>]*>", text)
    if not m:
        return xml_bytes
    root_open = m.group(0)
    # 收集当前根标签里已有的前缀和 URI
    existing_prefixes = set(re.findall(r'xmlns:(\w+)\s*=', root_open))
    existing_uris = set(re.findall(r'xmlns:\w+\s*=\s*"([^"]*)"', root_open))
    # 需要补的声明
    to_add = []
    for prefix, uri in ns_map.items():
        # 跳过 ET 临时前缀 ns0/ns1/...
        if re.match(r'^ns\d+$', prefix):
            continue
        if not uri:
            continue
        # 前缀已存在，或 URI 已被其它前缀声明，则跳过（避免重复）
        if prefix in existing_prefixes:
            continue
        if uri in existing_uris:
            continue
        to_add.append('xmlns:%s="%s"' % (prefix, uri))
        existing_prefixes.add(prefix)
        existing_uris.add(uri)
    # 额外确保 xml 命名空间已声明（xml:space 属性需要）
    xml_uri = "http://www.w3.org/XML/1998/namespace"
    if "xml" not in existing_prefixes and xml_uri not in existing_uris:
        to_add.append('xmlns:xml="%s"' % xml_uri)
    if not to_add:
        return xml_bytes
    # 在根标签的 <...:document 之后、第一个属性之前插入（或追加到 > 之前）
    new_root = root_open[:-1] + " " + " ".join(to_add) + ">"
    return text.replace(root_open, new_root, 1).encode("utf-8")


# xml:space 属性的正确命名空间（XML 标准命名空间，不是 w 命名空间）
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# ---------------------------------------------------------------------------
# 异常
# ---------------------------------------------------------------------------
class DocxError(Exception):
    pass


# ---------------------------------------------------------------------------
# 单位换算
# ---------------------------------------------------------------------------
def cm_to_twips(cm):
    """厘米 -> twip（1 cm = 567 twips，实际用 567 四舍五入）。"""
    return int(round(float(cm) * 567))


def twips_to_cm(tw):
    return float(tw) / 567.0


def pt_to_half(pt):
    """磅 -> 半磅（w:sz 的单位是半磅）。"""
    return int(round(float(pt) * 2))


def pt_to_twips(pt):
    """磅 -> twip（1 pt = 20 twip）。"""
    return int(round(float(pt) * 20))


# ---------------------------------------------------------------------------
# 小工具：子元素查找
# ---------------------------------------------------------------------------
def find_child(elem, tag):
    """在 elem 的直属子元素里找第一个 tag（本地名，不带命名空间前缀）。"""
    q = _w(tag)
    for c in elem:
        if c.tag == q:
            return c
    return None


def get_or_add_child(elem, tag):
    """找或创建直属子元素（按 w: 命名空间）。"""
    c = find_child(elem, tag)
    if c is None:
        import xml.etree.ElementTree as ET
        c = ET.SubElement(elem, _w(tag))
    return c


def set_attr(elem, tag, name, value):
    """设置 elem 直属子元素 tag 的 name 属性；子元素不存在则创建。"""
    c = get_or_add_child(elem, tag)
    c.set(_w(name), str(value))
    return c


# ---------------------------------------------------------------------------
# 文档对象
# ---------------------------------------------------------------------------
class Document:
    """轻量 docx 文档：读入 zip，解析 document.xml，可改后保存。"""

    def __init__(self, path):
        self.path = Path(path)
        if not self.path.exists():
            raise DocxError("文件不存在：%s" % self.path)
        self._read()

    def close(self):
        """释放底层 zip 句柄（读完后应调用，避免 Windows 下文件被锁）。"""
        try:
            self._zip.close()
        except Exception:
            pass

    # ---- 读取 ----
    def _read(self):
        import xml.etree.ElementTree as ET
        self._zip = zipfile.ZipFile(str(self.path), "r")
        names = self._zip.namelist()
        self._parts = {}  # name -> bytes（非 document.xml 的部件原样缓存）
        doc_xml = None
        for n in names:
            data = self._zip.read(n)
            if n == "word/document.xml":
                doc_xml = data
            else:
                self._parts[n] = data
        if doc_xml is None:
            raise DocxError("不是有效的 docx（缺少 word/document.xml）")
        # 从原始 XML 提取所有命名空间声明，保存供 save() 时注册/恢复回去，
        # 避免丢失 w14/w15/wp14/mc/o/v 等前缀（否则 mc:Ignorable 引用未声明前缀导致 invalid token）
        self._root_ns = _extract_namespaces(doc_xml)
        self._tree = ET.fromstring(doc_xml)
        self.body = self._tree.find(_w("body"))
        if self.body is None:
            raise DocxError("document.xml 缺 body")

    # ---- 顶层块级元素（w:p / w:tbl，按文档顺序） ----
    def blocks(self):
        """按文档顺序返回 body 直属的 w:p / w:tbl 元素列表。"""
        return [c for c in self.body if c.tag in (_w("p"), _w("tbl"))]

    def paragraphs(self):
        """返回所有顶层 w:p 元素（不含表格内段落）。"""
        return [c for c in self.body if c.tag == _w("p")]

    def tables(self):
        """返回所有顶层 w:tbl 元素。"""
        return [c for c in self.body if c.tag == _w("tbl")]

    # ---- 文本提取 ----
    @staticmethod
    def para_text(p):
        """提取一个 w:p 的纯文本（遍历所有 w:t）。"""
        return "".join(t.text or "" for t in p.iter(_w("t")))

    def full_text(self):
        """全文文本（段落 + 表格内容拼接，段落间换行）。"""
        out = []
        for c in self.body:
            if c.tag == _w("p"):
                out.append(self.para_text(c))
            elif c.tag == _w("tbl"):
                for tc in c.iter(_w("tc")):
                    for p in tc.findall(_w("p")):
                        out.append(self.para_text(p))
        return "\n".join(out)

    def table_texts(self):
        """返回每个表格的行列表（每行是 cell 文本列表）。"""
        result = []
        for tbl in self.tables():
            rows = []
            for tr in tbl.findall(_w("tr")):
                cells = []
                for tc in tr.findall(_w("tc")):
                    # 单元格内所有段落文本拼接
                    txt = "".join(self.para_text(p) for p in tc.findall(_w("p")))
                    cells.append(txt.strip())
                rows.append(cells)
            result.append(rows)
        return result

    # ---- 段落属性 ----
    def pPr(self, p):
        """返回段落的 w:pPr（不存在则创建）。"""
        props = get_or_add_child(p, "pPr")
        p.remove(props)
        p.insert(0, props)
        return props

    def rPr(self, run):
        """返回 run 的 w:rPr（不存在则创建）。"""
        props = get_or_add_child(run, "rPr")
        run.remove(props)
        run.insert(0, props)
        return props

    # ---- 保存 ----
    def save(self, path):
        """保存文档：document.xml 用改后的树，其余部件原样写回。"""
        import xml.etree.ElementTree as ET
        # 注册命名空间前缀：先注册从原文件提取的全部前缀（保留 w14/w15/mc/o/v 等），
        # 再补充标准前缀兜底，保证输出用标准前缀而不是 ns0:/ns1:
        _register_all_namespaces(self._root_ns)
        _register_namespaces()
        dst = Path(path)
        # 用临时文件写，避免覆盖失败
        tmp = dst.with_suffix(dst.suffix + ".tmp")
        try:
            with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as z:
                # 先写 document.xml
                xml_bytes = ET.tostring(self._tree, encoding="utf-8", xml_declaration=True)
                # 手动补回缺失的命名空间声明（ElementTree 会丢弃「未被元素实际引用」
                # 的前缀声明，如 w14/w15/wp14 只出现在 mc:Ignorable 字符串里），
                # 否则 mc:Ignorable="w14 w15 wp14" 引用未声明前缀会报 invalid token。
                xml_bytes = _restore_root_namespaces(xml_bytes, self._root_ns)
                z.writestr("word/document.xml", xml_bytes)
                # 其余部件原样写回
                for n, data in self._parts.items():
                    z.writestr(n, data)
            shutil.move(str(tmp), str(dst))
        finally:
            if tmp.exists():
                tmp.unlink()
        self._zip.close()


# ---------------------------------------------------------------------------
# 段落格式写操作
# ---------------------------------------------------------------------------
def set_para_format(doc, p, *, font_ea=None, font_ascii=None, size_pt=None,
                    bold=None, align=None, line_pt=None, indent_chars=None,
                    space_before=None, space_after=None,
                    paging=None, linebreak=None):
    """给段落 p 设置格式。所有参数 None 表示不修改该项。

    参数：
      font_ea / font_ascii : 中文字体 / 西文字体名
      size_pt : 字号（磅）
      bold : True/False
      align : "center"/"right"/"left"/"justify"/"both"（传 None 不改）
      line_pt : 行距（磅，固定值 EXACTLY）
      indent_chars : 首行缩进字符数（0 表示清除）
      space_before / space_after : 段前/段后（磅）
      paging : dict，分页四项（widowControl/keepNext/keepLines/pageBreakBefore -> "0"/"1"）
      linebreak : dict，换行三项（kinsoku/overflowPunct 存在即开，wordWrap "0"/"1"）
    """
    import xml.etree.ElementTree as ET
    pPr = doc.pPr(p)

    # 对齐 jc
    if align is not None:
        _align = {"center": "center", "right": "right", "left": "left",
                  "justify": "both", "both": "both"}.get(align, align)
        jc = get_or_add_child(pPr, "jc")
        jc.set(_w("val"), _align)

    # 间距 spacing（行距 + 段前段后）
    if line_pt is not None or space_before is not None or space_after is not None:
        spacing = get_or_add_child(pPr, "spacing")
        if line_pt is not None:
            spacing.set(_w("line"), str(pt_to_twips(line_pt)))
            spacing.set(_w("lineRule"), "exact")  # 固定值
        if space_before is not None:
            spacing.set(_w("before"), str(pt_to_twips(space_before)))
        if space_after is not None:
            spacing.set(_w("after"), str(pt_to_twips(space_after)))

    # 缩进 ind
    if indent_chars is not None:
        if indent_chars == 0:
            ind = get_or_add_child(pPr, "ind")
            ind.set(_w("firstLine"), "0")
            ind.set(_w("firstLineChars"), "0")
            ind.attrib.pop(_w("hanging"), None)
            ind.attrib.pop(_w("hangingChars"), None)
        else:
            ind = get_or_add_child(pPr, "ind")
            ind.set(_w("firstLineChars"), str(indent_chars * 100))
            if size_pt is not None:
                ind.set(_w("firstLine"), str(pt_to_twips(size_pt) * indent_chars))

    # 分页四项
    if paging is not None:
        for k, v in paging.items():
            set_attr(pPr, k, "val", v)

    # 换行三项
    if linebreak is not None:
        for k, v in linebreak.items():
            e = get_or_add_child(pPr, k)
            if v is not None:
                e.set(_w("val"), v)

    # 字体/字号应用到每个 run
    if font_ea or font_ascii or size_pt is not None or bold is not None:
        for run in p.findall(_w("r")):
            set_run_format(doc, run, font_ea=font_ea, font_ascii=font_ascii,
                           size_pt=size_pt, bold=bold)


def set_run_format(doc, run, *, font_ea=None, font_ascii=None, size_pt=None,
                   bold=None):
    """设置单个 run 的字体/字号/加粗。"""
    rPr = doc.rPr(run)
    if font_ea or font_ascii:
        rFonts = get_or_add_child(rPr, "rFonts")
        if font_ascii:
            rFonts.set(_w("ascii"), font_ascii)
            rFonts.set(_w("hAnsi"), font_ascii)
            rFonts.set(_w("cs"), font_ascii)
        if font_ea:
            rFonts.set(_w("eastAsia"), font_ea)
        for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"):
            rFonts.attrib.pop(_w(key), None)
    if size_pt is not None:
        set_attr(rPr, "sz", "val", pt_to_half(size_pt))
        set_attr(rPr, "szCs", "val", pt_to_half(size_pt))
    if bold is not None:
        b = get_or_add_child(rPr, "b")
        b.set(_w("val"), "1" if bold else "0")


# ---------------------------------------------------------------------------
# 段落文本写操作（改 run 文本）
# ---------------------------------------------------------------------------
def set_para_text(p, text):
    """按文本差异替换段落文字，不合并原有 run。
    若段落没有 run，则新建一个 run。保留各文本节点及其 run 格式，避免校对吞掉局部强调。"""
    import xml.etree.ElementTree as ET
    from difflib import SequenceMatcher
    nodes = list(p.iter(_w("t")))
    if not nodes:
        run = ET.SubElement(p, _w("r"))
        nodes = [ET.SubElement(run, _w("t"))]
    old = ''.join(node.text or '' for node in nodes)
    owners = [index for index, node in enumerate(nodes) for _ in (node.text or '')]
    pieces = [[] for _ in nodes]
    for kind, a, b, c, d in SequenceMatcher(None, old, text, autojunk=False).get_opcodes():
        if kind == 'equal':
            for offset in range(b - a):
                pieces[owners[a + offset]].append(text[c + offset])
        elif kind in ('replace', 'insert'):
            owner = owners[a] if a < len(owners) else (owners[-1] if owners else 0)
            pieces[owner].append(text[c:d])
    for node, parts in zip(nodes, pieces):
        node.set(XML_SPACE, "preserve")
        node.text = ''.join(parts)


def strip_para_spaces(p):
    """清理段落首尾多余空格/制表符/全角空格。返回是否改动。"""
    text = "".join(t.text or "" for t in p.iter(_w("t")))
    stripped = text.strip(" \t\u3000")
    if stripped == text:
        return False
    set_para_text(p, stripped)
    return True


# ---------------------------------------------------------------------------
# 页面设置
# ---------------------------------------------------------------------------
def set_page_margins(doc, top_cm, bottom_cm, left_cm, right_cm,
                     page_w_cm=21.0, page_h_cm=29.7):
    """设置页面尺寸与页边距（单位 cm）。应用到所有 sectPr。"""
    # 找所有 sectPr（可能在 body 尾部，或段落 pPr 里）
    sect_prs = []
    # body 直属 sectPr
    for c in doc.body:
        if c.tag == _w("sectPr"):
            sect_prs.append(c)
    # 段落 pPr 里的 sectPr
    for p in doc.paragraphs():
        pPr = find_child(p, "pPr")
        if pPr is not None:
            sp = find_child(pPr, "sectPr")
            if sp is not None:
                sect_prs.append(sp)

    for sp in sect_prs:
        pgSz = get_or_add_child(sp, "pgSz")
        pgSz.set(_w("w"), str(cm_to_twips(page_w_cm)))
        pgSz.set(_w("h"), str(cm_to_twips(page_h_cm)))
        pgMar = get_or_add_child(sp, "pgMar")
        pgMar.set(_w("top"), str(cm_to_twips(top_cm)))
        pgMar.set(_w("bottom"), str(cm_to_twips(bottom_cm)))
        pgMar.set(_w("left"), str(cm_to_twips(left_cm)))
        pgMar.set(_w("right"), str(cm_to_twips(right_cm)))


def get_content_width_twips(doc):
    """返回正文可用宽度（twip）：页宽 − 左右边距。用于表格缩放比例计算。"""
    sp = None
    for c in doc.body:
        if c.tag == _w("sectPr"):
            sp = c
            break
    if sp is None:
        for p in doc.paragraphs():
            pPr = find_child(p, "pPr")
            if pPr is not None:
                sp = find_child(pPr, "sectPr")
                if sp is not None:
                    break
    if sp is None:
        return None
    pgSz = find_child(sp, "pgSz")
    pgMar = find_child(sp, "pgMar")
    if pgSz is None or pgMar is None:
        return None
    w = int(pgSz.get(_w("w"), 0))
    left = int(pgMar.get(_w("left"), 0))
    right = int(pgMar.get(_w("right"), 0))
    return w - left - right


# ---------------------------------------------------------------------------
# 表格宽度缩放
# ---------------------------------------------------------------------------
def scale_table_widths(doc, old_width, new_width):
    """等比例缩放所有表格宽度。返回缩放的表格数。"""
    if old_width <= 0 or new_width <= 0:
        return 0
    ratio = new_width / old_width
    if abs(ratio - 1.0) < 0.001:
        return 0
    scaled = 0
    for tbl in doc.tables():
        tblPr = get_or_add_child(tbl, "tblPr")
        # 总宽 tblW
        tblW = find_child(tblPr, "tblW")
        if tblW is not None:
            wtype = tblW.get(_w("type"))
            w = tblW.get(_w("w"))
            if wtype == "dxa" and w:
                tblW.set(_w("w"), str(int(round(int(w) * ratio))))
        # 列宽 tblGrid
        grid = find_child(tbl, "tblGrid")
        if grid is not None:
            for gc in grid.findall(_w("gridCol")):
                w = gc.get(_w("w"))
                if w:
                    gc.set(_w("w"), str(int(round(int(w) * ratio))))
        # 单元格 tcW
        for tc in tbl.iter(_w("tc")):
            tcPr = find_child(tc, "tcPr")
            if tcPr is not None:
                tcW = find_child(tcPr, "tcW")
                if tcW is not None:
                    wtype = tcW.get(_w("type"))
                    w = tcW.get(_w("w"))
                    if wtype == "dxa" and w:
                        tcW.set(_w("w"), str(int(round(int(w) * ratio))))
        scaled += 1
    return scaled


# ---------------------------------------------------------------------------
# 页脚页码
# ---------------------------------------------------------------------------
def add_page_number(doc, footer_file="word/footer1.xml", font="宋体", size_pt=14):
    """在页脚居中插入页码「— 1 —」（PAGE 域）。返回是否成功。

    简化策略：直接改写/创建一个 footer XML 部件，并在 sectPr 里引用它。
    若正文已存在页脚引用且内容非空，则跳过（避免破坏版记）。
    """
    import xml.etree.ElementTree as ET
    _register_namespaces()
    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", rels_ns)

    def _footer_relationship():
        """返回 (rel_id, target) —— document.xml.rels 里已有的 footer 关系。
        没有则返回 (None, None)。"""
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in doc._parts:
            return None, None
        try:
            rels = ET.fromstring(doc._parts[rels_name])
        except Exception:
            return None, None
        for rel in rels:
            if rel.get("Type", "").endswith("/footer"):
                return rel.get("Id"), rel.get("Target")
        return None, None

    def _footer_has_content(target):
        """判断 footer 关系指向的部件是否已有非空内容（有则说明可能已有页码/版记）。"""
        if not target:
            return False
        # target 形如 "footer1.xml"，对应部件 "word/footer1.xml"
        for n in doc._parts:
            base = n.split("/")[-1]
            if base == target and n.startswith("word/"):
                try:
                    f = ET.fromstring(doc._parts[n])
                    text = "".join(t.text or "" for t in f.iter(_w("t"))).strip()
                    if text:
                        return True
                except Exception:
                    pass
        return False

    # 已有 footer 关系且部件内容非空 → 跳过（避免覆盖已有页码/版记）
    rel_id, target = _footer_relationship()
    if _footer_has_content(target):
        return False

    # 确定 footer 部件名与关系 id：
    #   若已有 footer 关系（但部件空/缺失），复用其 target 和 id；
    #   否则新建一个 footer 部件和关系。
    if target and target.endswith(".xml"):
        footer_part_name = "word/" + target
        use_rel_id = rel_id
    else:
        footer_part_name = footer_file
        use_rel_id = None  # 需要新建关系

    # 构造 footer XML（root 不手动指定 nsmap，交给 _register_namespaces 统一用 w: 前缀）
    footer = ET.Element(_w("ftr"))
    p = ET.SubElement(footer, _w("p"))
    pPr = ET.SubElement(p, _w("pPr"))
    jc = ET.SubElement(pPr, _w("jc"))
    jc.set(_w("val"), "center")

    def _mk_run(text=None, fldChar=None, instrText=None):
        r = ET.SubElement(p, _w("r"))
        rPr = ET.SubElement(r, _w("rPr"))
        rFonts = ET.SubElement(rPr, _w("rFonts"))
        rFonts.set(_w("ascii"), font)
        rFonts.set(_w("hAnsi"), font)
        rFonts.set(_w("eastAsia"), font)
        sz = ET.SubElement(rPr, _w("sz"))
        sz.set(_w("val"), str(pt_to_half(size_pt)))
        szCs = ET.SubElement(rPr, _w("szCs"))
        szCs.set(_w("val"), str(pt_to_half(size_pt)))
        if text is not None:
            t = ET.SubElement(r, _w("t"))
            t.set(XML_SPACE, "preserve")
            t.text = text
        if fldChar is not None:
            f = ET.SubElement(r, _w("fldChar"))
            f.set(_w("fldCharType"), fldChar)
        if instrText is not None:
            it = ET.SubElement(r, _w("instrText"))
            it.set(XML_SPACE, "preserve")
            it.text = instrText
        return r

    _mk_run(text="— ")
    _mk_run(fldChar="begin")
    _mk_run(instrText=" PAGE ")
    _mk_run(fldChar="separate")
    _mk_run(text="1")
    _mk_run(fldChar="end")
    _mk_run(text=" —")

    footer_bytes = ET.tostring(footer, encoding="utf-8", xml_declaration=True)
    doc._parts[footer_part_name] = footer_bytes

    # 写 document.xml.rels（新建 footer 关系或复用已有 id）
    rels_name = "word/_rels/document.xml.rels"
    footer_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    if use_rel_id is not None:
        # 复用已有关系：确保其 Id 和 Target 正确（Target 可能已正确指向空 footer）
        if rels_name in doc._parts:
            rels = ET.fromstring(doc._parts[rels_name])
            for rel in rels:
                if rel.get("Type", "").endswith("/footer"):
                    rel.set("Target", footer_part_name.split("/")[-1])
            doc._parts[rels_name] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
        # 沿用 use_rel_id
        new_rel_id = use_rel_id
    else:
        # 新建关系，生成唯一 id
        if rels_name in doc._parts:
            rels = ET.fromstring(doc._parts[rels_name])
        else:
            rels = ET.Element("{%s}Relationships" % rels_ns)
        # 找最大 rId 编号，生成下一个
        max_n = 0
        for rel in rels:
            rid = rel.get("Id", "")
            m = re.match(r"rId(\d+)", rid)
            if m:
                max_n = max(max_n, int(m.group(1)))
        new_rel_id = "rId%d" % (max_n + 1)
        rel = ET.SubElement(rels, "{%s}Relationship" % rels_ns)
        rel.set("Id", new_rel_id)
        rel.set("Type", footer_rel_type)
        rel.set("Target", footer_part_name.split("/")[-1])
        doc._parts[rels_name] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    # 在 sectPr 里引用 footerReference
    sp = None
    for c in doc.body:
        if c.tag == _w("sectPr"):
            sp = c
            break
    if sp is None:
        return False
    ref = get_or_add_child(sp, "footerReference")
    ref.set(_w("type"), "default")
    ref.set("{%s}id" % NS["r"], new_rel_id)
    # 需要在 sectPr 上声明 r: 命名空间
    rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ET.register_namespace("r", rns)

    return True
