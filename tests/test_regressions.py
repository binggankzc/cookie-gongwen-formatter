"""Run with Python 3.9+: python -m unittest discover -s tests -v."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(os.environ.get('COOKIE_SKILL_TEST_ROOT', Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(ROOT / 'scripts'))
import docx_lite as core
import format_gongwen as fmt
import map_pages as maps
import proofread as proof
W = core.W


def fixture(path, title='关于开展学习的通知', table=False):
    body = ET.Element(W + 'body')
    for text in [title, '各部门：', '请于2026-11-09参加学习,谢谢!', '一、学习安排', '（一）材料准备', '1.每项得0.5分。']:
        p = ET.SubElement(body, W + 'p')
        if text.startswith('一、'):
            props = ET.SubElement(p, W + 'pPr')
            ET.SubElement(props, W + 'keepNext')
            ET.SubElement(props, W + 'ind', {W + 'left': '500', W + 'hanging': '100'})
        r = ET.SubElement(p, W + 'r')
        ET.SubElement(r, W + 't').text = text
    if table:
        tbl = ET.SubElement(body, W + 'tbl')
        pr = ET.SubElement(tbl, W + 'tblPr')
        ET.SubElement(pr, W + 'tblW', {W + 'type': 'dxa', W + 'w': '6000'})
        grid = ET.SubElement(tbl, W + 'tblGrid')
        for _ in range(2):
            ET.SubElement(grid, W + 'gridCol', {W + 'w': '3000'})
        tr = ET.SubElement(tbl, W + 'tr')
        for text in ['项目', '再接再励']:
            tc = ET.SubElement(tr, W + 'tc')
            pp = ET.SubElement(tc, W + 'p')
            rr = ET.SubElement(pp, W + 'r')
            ET.SubElement(rr, W + 't').text = text
    sec = ET.SubElement(body, W + 'sectPr')
    ET.SubElement(sec, W + 'pgSz', {W + 'w': '11906', W + 'h': '16838'})
    ET.SubElement(sec, W + 'pgMar', {W + 'top': '1440', W + 'bottom': '1440', W + 'left': '1440', W + 'right': '1440'})
    doc = ET.Element(W + 'document')
    doc.append(body)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', ET.tostring(doc))
        z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def make_docx(path, texts, sect=True):
    """Build a minimal docx with the given paragraph texts.

    sect=True adds a body sectPr (A4, 2.54cm margins); sect=False
    leaves the document with no sectPr at all (M8 empty-doc case)."""
    body = ET.Element(W + 'body')
    for text in texts:
        p = ET.SubElement(body, W + 'p')
        r = ET.SubElement(p, W + 'r')
        ET.SubElement(r, W + 't').text = text
    if sect:
        sec = ET.SubElement(body, W + 'sectPr')
        ET.SubElement(sec, W + 'pgSz', {W + 'w': '11906', W + 'h': '16838'})
        ET.SubElement(sec, W + 'pgMar', {W + 'top': '1440', W + 'bottom': '1440', W + 'left': '1440', W + 'right': '1440'})
    doc = ET.Element(W + 'document')
    doc.append(body)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', ET.tostring(doc))
        z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


class ProofreadingTests(unittest.TestCase):
    def test_numbering_and_decimals(self):
        for text in ['1.承揽项目，每项得5分。', '  14.每项得0.5分。', '版本1.2.3，增长3.14%。']:
            self.assertEqual(proof.check_and_report(text)[0], text)

    def test_technical_text_and_english(self):
        for text in ['English text, hello!', '网址https://example.com/a?x=1&y=2', '邮箱foo.bar@example.com', '文件report.docx，路径C:\\docs\\a.docx', '代码`print("hi")`', '日期见https://example.com/2026-11-09/report.docx']:
            self.assertEqual(proof.check_and_report(text)[0], text)

    def test_chinese_punctuation(self):
        self.assertEqual(proof.check_and_report('你好,请检查(材料)!')[0], '你好，请检查（材料）！')

    def test_valid_dates_only(self):
        self.assertEqual(proof.check_and_report('日期2026.11.09')[0], '日期2026年11月9日')
        for text in ['日期2026-02-30', '日期2026.13.09', '编号A2026-11-09B', '文件2026-11-09.docx']:
            self.assertEqual(proof.check_and_report(text)[0], text)

    def test_no_false_past_tense_in_read_only_report(self):
        _, issues = proof.check_and_report('你好,谢谢!')
        self.assertFalse(any('已自动' in issue for issue in issues))

    def test_typo_is_suggestion_only(self):
        result, issues = proof.check_and_report('再接再励')
        self.assertEqual(result, '再接再励')
        self.assertTrue(any('再接再厉' in issue for issue in issues))

    # M5: “截止”只在接具体日期时才提示。
    def test_jiezhi_not_flagged_for_deadline_noun(self):
        for text in ['报名截止日期为2026年11月9日。', '截止目前，已完成三项工作。']:
            _, issues = proof.check_and_report(text)
            self.assertFalse(any('截至' in i for i in issues), text)

    def test_jiezhi_flagged_for_date(self):
        _, issues = proof.check_and_report('截止2026年11月9日，已完成三项工作。')
        self.assertTrue(any('截至' in i for i in issues))
        _, issues = proof.check_and_report('截止到2026年11月9日，已完成三项工作。')
        self.assertTrue(any('截至' in i for i in issues))

    # M5: “通过…使…”只在确实缺主语时才提示。
    def test_tongguo_shi_with_subject_not_flagged(self):
        _, issues = proof.check_and_report('通过学习使大家提高了认识。')
        self.assertFalse(any('句式杂糅' in i for i in issues))

    def test_tongguo_shi_missing_subject_flagged(self):
        _, issues = proof.check_and_report('通过加强管理使提高了效率。')
        self.assertTrue(any('句式杂糅' in i for i in issues))

    # M5: 大错/小错按规则类型分，不按文案里的“自动”二字分。
    def test_severity_is_by_rule_type(self):
        _, detailed = proof.check_and_report_detailed('你好。。')
        self.assertTrue(any(s == 'minor' and '标点重复' in m for s, m in detailed))
        _, detailed = proof.check_and_report_detailed('再接再励')
        self.assertTrue(any(s == 'major' for s, _ in detailed))

    def test_preserve_runs_when_replacing_text(self):
        p = ET.Element(W + 'p')
        for text, bold in [('你好,', True), ('请检查!', False)]:
            r = ET.SubElement(p, W + 'r')
            pr = ET.SubElement(r, W + 'rPr')
            ET.SubElement(pr, W + 'b', {W + 'val': '1' if bold else '0'})
            ET.SubElement(r, W + 't').text = text
        core.set_para_text(p, '你好，请检查！')
        runs = p.findall(W + 'r')
        self.assertEqual([r.find(W + 't').text for r in runs], ['你好，', '请检查！'])
        self.assertEqual(runs[0].find(W + 'rPr/' + W + 'b').get(W + 'val'), '1')

    def test_complex_hyperlink_is_not_rewritten(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'link.docx'
            fixture(src)
            with zipfile.ZipFile(src) as z:
                parts = {name: z.read(name) for name in z.namelist()}
            tree = ET.fromstring(parts['word/document.xml'])
            para = tree.find(W + 'body').findall(W + 'p')[2]
            link = ET.SubElement(para, W + 'hyperlink')
            run = ET.SubElement(link, W + 'r')
            ET.SubElement(run, W + 't').text = '链接'
            before = ET.tostring(para)
            parts['word/document.xml'] = ET.tostring(tree)
            with zipfile.ZipFile(src, 'w') as z:
                for name, data in parts.items():
                    z.writestr(name, data)
            result = subprocess.run([sys.executable, str(ROOT/'scripts/proofread.py'), str(src), '--fix'], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            output = core.Document(src.with_name('link_校对.docx'))
            self.assertEqual(ET.tostring(output.paragraphs()[2]), before)



class ClassificationTests(unittest.TestCase):
    def test_title_outranks_body(self):
        self.assertEqual(fmt.detect_doc_type('关于考核办法的意见建议报告\n现提出如下意见'), '报告')
        self.assertEqual(fmt.detect_doc_type('关于转发调查报告的通知\n特此报告'), '通知')
        self.assertEqual(fmt.detect_doc_type('关于调整意见的请示\n请予审批'), '请示')

    def test_closing_formula(self):
        self.assertEqual(fmt.detect_doc_type('年度材料\n情况如下\n以上报告，请审阅。'), '报告')

    def _category_of(self, tmp, title, body="特此通知。"):
        path = Path(tmp) / "m1.docx"
        make_docx(path, [title, body])
        return core.Document(str(path))

    # M1: 标题末尾的文种优先于标题关键词。
    def test_doc_type_outranks_dongyuan_keyword(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = self._category_of(tmp, "关于动员部署工作的通知")
            self.assertEqual(fmt.detect_category(doc), "gongwen")

    def test_doc_type_outranks_huibao_keyword(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = self._category_of(tmp, "关于汇报工作的通知")
            self.assertEqual(fmt.detect_category(doc), "gongwen")

    def test_keyword_only_category_warns(self):
        # 只有关键词、没有明确文种时，仍按关键词判断，但要给出提示。
        with tempfile.TemporaryDirectory() as tmp:
            doc = self._category_of(tmp, "动员大会讲话", "同志们：")
            cat, warning = fmt.detect_category_with_warning(doc)
            self.assertEqual(cat, "speech")
            self.assertIn("请确认", warning)

    # L6: 半角括号、中文数字「一百、」、四级标题都要识别；行首以外不误判。
    def test_half_width_paren_chinese_numeral_is_h2(self):
        self.assertEqual(fmt.classify('(一)材料准备'), 'h2')
        self.assertEqual(fmt.classify('（一）材料准备'), 'h2')

    def test_hundred_numeral_is_h1(self):
        self.assertEqual(fmt.classify('一百、总体要求'), 'h1')
        self.assertEqual(fmt.classify('一、总体要求'), 'h1')

    def test_level4_paren_arabic_is_h4(self):
        self.assertEqual(fmt.classify('(1)具体要求如下'), 'h4')
        self.assertEqual(fmt.classify('（1）具体要求如下'), 'h4')

    def test_paren_number_not_at_line_start_stays_body(self):
        self.assertEqual(fmt.classify('根据(一)款规定执行。'), 'body')
        self.assertEqual(fmt.classify('2026年(1)月1日起实施。'), 'body')
        self.assertEqual(fmt.classify('(2026)年工作要点'), 'body')


class FontAndLayoutTests(unittest.TestCase):
    def test_mac_font_name_reader(self):
        if sys.platform != 'darwin':
            self.skipTest('只在 macOS 上验证字体名读取，非 macOS 平台跳过')
        names = fmt._font_family_names(Path('/Library/Fonts/Arial Unicode.ttf'))
        self.assertIn('Arial Unicode MS', names)

    def test_fallback_must_exist(self):
        with patch.object(fmt, 'get_installed_fonts', return_value={'Songti SC', 'Heiti SC'}):
            fonts, notes = fmt.resolve_fonts()
            self.assertTrue(set(fonts.values()) <= {'Songti SC', 'Heiti SC'})
            self.assertEqual(len(notes), 4)
        with patch.object(fmt, 'get_installed_fonts', return_value=set()):
            with self.assertRaises(RuntimeError):
                fmt.resolve_fonts()

    def test_linux_fc_list_parsing(self):
        sample = (
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc: Noto Serif CJK SC,Noto Serif CJK SC Regular:style=Regular\n"
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf: DejaVu Sans:style=Book\n"
            "not-a-font-line\n"
            "\n"
        )
        fonts = fmt._parse_fc_list_output(sample)
        self.assertIn('Noto Serif CJK SC', fonts)
        self.assertIn('DejaVu Sans', fonts)
        self.assertIn('NotoSerifCJK-Regular', fonts)  # 文件 stem 也应被记录
        self.assertNotIn('not-a-font-line', fonts)

    def test_linux_font_detection_uses_fc_list(self):
        sample = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc: Noto Serif CJK SC:style=Regular\n"
        with patch.object(fmt.sys, 'platform', 'linux'), \
             patch.object(fmt.subprocess, 'run') as mock_run:
            mock_run.return_value = subprocess.CompletedProcess(
                ['fc-list'], 0, stdout=sample, stderr='')
            fmt.get_installed_fonts.cache_clear()
            try:
                self.assertIn('Noto Serif CJK SC', fmt.get_installed_fonts())
            finally:
                fmt.get_installed_fonts.cache_clear()

    def test_linux_fontconfig_missing_returns_empty(self):
        # fc-list 不存在时不得崩溃：退回目录扫描，本测试用空扫描代替以免扫盘。
        with patch.object(fmt.sys, 'platform', 'linux'), \
             patch.object(fmt.subprocess, 'run', side_effect=FileNotFoundError), \
             patch.object(fmt, '_fonts_from_font_dirs', return_value=set()):
            fmt.get_installed_fonts.cache_clear()
            try:
                self.assertEqual(fmt.get_installed_fonts(), set())
            finally:
                fmt.get_installed_fonts.cache_clear()

    def test_noto_cjk_is_accepted_fallback(self):
        # 只有 Noto CJK 可用时也应能完成字体解析（Linux 兜底），且仍记降级说明。
        with patch.object(fmt, 'get_installed_fonts',
                          return_value={'Noto Serif CJK SC', 'Noto Sans CJK SC'}):
            fonts, notes = fmt.resolve_fonts()
            self.assertEqual(fonts['title'], 'Noto Serif CJK SC')
            self.assertEqual(fonts['body'], 'Noto Serif CJK SC')
            self.assertEqual(fonts['heiti'], 'Noto Sans CJK SC')
            self.assertEqual(fonts['kaiti'], 'Noto Serif CJK SC')
            self.assertEqual(len(notes), 4)

    def test_theme_override_and_property_order(self):
        class Doc:
            rPr = core.Document.rPr
        run = ET.Element(W + 'r')
        ET.SubElement(run, W + 't').text = '正文'
        core.set_run_format(Doc(), run, font_ea='Songti SC', font_ascii='Songti SC', size_pt=16)
        fonts = run.find(W + 'rPr/' + W + 'rFonts')
        fonts.set(W + 'eastAsiaTheme', 'majorEastAsia')
        core.set_run_format(Doc(), run, font_ea='Songti SC')
        self.assertNotIn(W + 'eastAsiaTheme', fonts.attrib)
        self.assertEqual(run[0].tag, W + 'rPr')

    def test_redhead_is_protected(self):
        p = ET.fromstring(f'<w:p xmlns:w="{core.NS["w"]}"><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>示例机关文件</w:t></w:r></w:p>')
        self.assertTrue(fmt._protected_paragraph(p))

    # M7: 单个红色词不保护整段；红色占多数时才保护。
    def _red_para(self, runs):
        p = ET.Element(W + 'p')
        for text, red in runs:
            r = ET.SubElement(p, W + 'r')
            if red:
                pr = ET.SubElement(r, W + 'rPr')
                ET.SubElement(pr, W + 'color', {W + 'val': 'FF0000'})
            ET.SubElement(r, W + 't').text = text
        return p

    def test_single_red_word_does_not_protect(self):
        p = self._red_para([('这是正文段落，其中只有', False), ('重要', True), ('两个字标红，其余都是正常文字。', False)])
        self.assertFalse(fmt._protected_paragraph(p))

    def test_majority_red_protects(self):
        p = self._red_para([('示例机关文件', True), ('通知', True), ('。', False)])
        self.assertTrue(fmt._protected_paragraph(p))

    # M2: 超链接里的 run 也要拿到字体/字号。
    def test_hyperlink_runs_get_font_and_size(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'link.docx'
            fixture(src)
            doc = core.Document(str(src))
            p = doc.paragraphs()[0]
            link = ET.SubElement(p, W + 'hyperlink')
            r = ET.SubElement(link, W + 'r')
            ET.SubElement(r, W + 't').text = '链接文字'
            core.set_para_format(doc, p, font_ea='宋体', font_ascii='宋体', size_pt=16)
            fonts = r.find(W + 'rPr/' + W + 'rFonts')
            self.assertIsNotNone(fonts)
            self.assertEqual(fonts.get(W + 'ascii'), '宋体')
            self.assertEqual(r.find(W + 'rPr/' + W + 'sz').get(W + 'val'), '32')

    # M3: 按字符数缩进时，清除残留的 firstLine（twips）。
    def test_indent_chars_clears_first_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'ind.docx'
            fixture(src)
            doc = core.Document(str(src))
            p = doc.paragraphs()[0]
            pPr = core.get_or_add_child(p, 'pPr')
            ind = core.get_or_add_child(pPr, 'ind')
            ind.set(W + 'firstLine', '720')
            fmt._normalize_layout(doc, p)
            ind = p.find(W + 'pPr/' + W + 'ind')
            self.assertNotIn(W + 'firstLine', ind.attrib)
            self.assertEqual(ind.get(W + 'firstLineChars'), '200')


class WorkflowTests(unittest.TestCase):
    def run_script(self, script, *args):
        result = subprocess.run([sys.executable, str(ROOT / 'scripts' / script), *map(str, args)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout

    def test_real_workflow_preserves_source_and_numbering(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'input.docx', Path(tmp) / 'formatted.docx'
            fixture(src)
            original = src.read_bytes()
            self.run_script('format_gongwen.py', src, out, '--mode', 'gongwen')
            self.run_script('proofread.py', out, '--fix')
            final = out.with_name('formatted_校对.docx')
            doc = core.Document(final)
            text = doc.full_text()
            self.assertIn('1.每项得0.5分。', text)
            self.assertIn('2026年11月9日参加学习，谢谢！', text)
            self.assertEqual(src.read_bytes(), original)
            with zipfile.ZipFile(final) as z:
                for name in z.namelist():
                    if name.endswith(('.xml', '.rels')):
                        ET.fromstring(z.read(name))
            heading = doc.paragraphs()[3]
            props = heading.find(W + 'pPr')
            self.assertEqual(props.find(W + 'keepNext').get(W + 'val'), '0')
            self.assertEqual(props.find(W + 'ind').get(W + 'left'), '0')
            self.assertNotIn(W + 'hanging', props.find(W + 'ind').attrib)

    def test_inspect_does_not_create_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            fixture(src, title='关于考核办法的意见建议报告')
            before = src.read_bytes()
            output = self.run_script('format_gongwen.py', src, '--inspect')
            self.assertIn('[文种] 报告', output)
            self.assertEqual(list(Path(tmp).iterdir()), [src])
            self.assertEqual(src.read_bytes(), before)

    def test_speech_and_table(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'speech.docx', Path(tmp) / 'speech_out.docx'
            fixture(src, title='学习交流会上的讲话', table=True)
            original = core.Document(src).table_texts()
            self.run_script('format_gongwen.py', src, out, '--mode', 'speech')
            self.assertEqual(core.Document(out).table_texts(), original)
            before = out.read_bytes()
            output = self.run_script('proofread.py', out, '--fix')
            self.assertIn('表1行1列2', output)
            self.assertEqual(out.read_bytes(), before)
            self.assertEqual(core.Document(out.with_name('speech_out_校对.docx')).table_texts(), original)
            no_output = Path(tmp) / 'table_out.docx'
            self.run_script('format_gongwen.py', src, no_output, '--mode', 'table')
            self.assertFalse(no_output.exists())

    def test_refuses_overwriting_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            fixture(src)
            before = src.read_bytes()
            result = subprocess.run([sys.executable, str(ROOT/'scripts/format_gongwen.py'), str(src), str(src)], capture_output=True)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(src.read_bytes(), before)

    def test_footer_content_type_override(self):
        # 端到端：format_gongwen --mode gongwen 新建页脚后，输出 zip 的
        # [Content_Types].xml 必须含 footer Override，且 footer 部件与 rels 引用齐全。
        ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'input.docx', Path(tmp) / 'formatted.docx'
            fixture(src)
            self.run_script('format_gongwen.py', src, out, '--mode', 'gongwen')
            with zipfile.ZipFile(out) as z:
                self.assertIn('word/footer1.xml', z.namelist())
                self.assertIn('[Content_Types].xml', z.namelist())
                types = ET.fromstring(z.read('[Content_Types].xml'))
                overrides = {o.get('PartName'): o.get('ContentType')
                             for o in types.findall('{%s}Override' % ct_ns)}
                self.assertEqual(
                    overrides.get('/word/footer1.xml'),
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
                rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
                self.assertTrue(any(r.get('Type', '').endswith('/footer') for r in rels))

    def test_footer_override_is_not_duplicated(self):
        # 对同一文档重复调用 add_page_number 不得产生重复的 Override 声明。
        ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            fixture(src)
            doc = core.Document(str(src))
            self.assertTrue(core.add_page_number(doc))
            self.assertFalse(core.add_page_number(doc))  # 已有内容，第二次应跳过
            doc.save(str(src))
            with zipfile.ZipFile(src) as z:
                types = ET.fromstring(z.read('[Content_Types].xml'))
                parts = [o.get('PartName') for o in types.findall('{%s}Override' % ct_ns)]
                self.assertEqual(parts.count('/word/footer1.xml'), 1)

    # M6: argparse 解析 --mode 的两种写法；缺值/未知参数友好报错（exit 2，无 traceback）。
    def test_mode_equals_form(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'input.docx', Path(tmp) / 'formatted.docx'
            fixture(src)
            self.run_script('format_gongwen.py', src, out, '--mode=gongwen')
            self.assertTrue(out.exists())

    def test_mode_missing_value_is_friendly(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            fixture(src)
            result = subprocess.run(
                [sys.executable, str(ROOT / 'scripts/format_gongwen.py'), str(src), '--mode'],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            self.assertNotIn('Traceback', result.stderr)
            self.assertIn('usage', result.stderr.lower())

    def test_unknown_arg_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            fixture(src)
            result = subprocess.run(
                [sys.executable, str(ROOT / 'scripts/format_gongwen.py'), str(src), '--bogus'],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            self.assertNotIn('Traceback', result.stderr)

    # L7: dst==src 给友好错误，不抛 ValueError traceback。
    def test_dst_same_as_src_is_friendly(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            fixture(src)
            result = subprocess.run(
                [sys.executable, str(ROOT / 'scripts/format_gongwen.py'), str(src), str(src)],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 1)
            self.assertNotIn('Traceback', result.stderr)
            self.assertIn('[错误] 输出路径不能与原文件相同', result.stdout)

    # M4: .doc 转换的默认输出名按原 .doc 的 stem，临时文件失败时也要清理。
    # Windows COM 转换本身无法在 Linux 上验证（未确认），这里用 mock 替代。
    def _fake_convert(self, tmp):
        def fake(src):
            tmp_docx = Path(tmp) / '__conv_tmp.docx'
            fixture(tmp_docx)
            return tmp_docx
        return fake

    def test_doc_conversion_default_output_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / '报告.doc'
            src.write_bytes(b'')
            with patch.object(fmt, 'convert_doc_to_docx', side_effect=self._fake_convert(tmp)), \
                 patch.object(sys, 'argv', ['format_gongwen.py', str(src)]):
                fmt.main()
            self.assertTrue((Path(tmp) / '报告_公文格式.docx').exists())
            self.assertFalse((Path(tmp) / '__conv_tmp.docx').exists())
            self.assertFalse((Path(tmp) / '__conv_tmp_公文格式.docx').exists())

    def test_doc_conversion_temp_cleaned_on_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / '报告.doc'
            src.write_bytes(b'')
            with patch.object(fmt, 'convert_doc_to_docx', side_effect=self._fake_convert(tmp)), \
                 patch.object(fmt, 'resolve_fonts', side_effect=RuntimeError('boom')), \
                 patch.object(sys, 'argv', ['format_gongwen.py', str(src)]):
                with self.assertRaises(RuntimeError):
                    fmt.main()
            self.assertFalse((Path(tmp) / '__conv_tmp.docx').exists())
            self.assertFalse((Path(tmp) / '报告_公文格式.docx').exists())

    # M8: 多分节文档边距/页脚应用到所有 sectPr；无 sectPr 的空文档也能处理。
    def test_multi_section_margins_and_footer(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'multi.docx'
            fixture(src)
            doc = core.Document(str(src))
            p = doc.paragraphs()[0]
            pPr = core.get_or_add_child(p, 'pPr')
            sp2 = ET.SubElement(pPr, W + 'sectPr')
            ET.SubElement(sp2, W + 'pgSz', {W + 'w': '11906', W + 'h': '16838'})
            ET.SubElement(sp2, W + 'pgMar', {W + 'top': '1440', W + 'bottom': '1440',
                                             W + 'left': '1440', W + 'right': '1440'})
            core.set_page_margins(doc, 3.7, 3.5, 2.8, 2.6)
            for sp in core._all_sect_prs(doc):
                self.assertEqual(len(core._all_sect_prs(doc)), 2)
                self.assertEqual(sp.find(W + 'pgMar').get(W + 'top'), str(core.cm_to_twips(3.7)))
            self.assertTrue(core.add_page_number(doc))
            for sp in core._all_sect_prs(doc):
                self.assertIsNotNone(sp.find(W + 'footerReference'))

    def test_no_section_doc_gets_margins_and_footer(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'empty.docx'
            make_docx(src, ['只有一段文字。'], sect=False)
            doc = core.Document(str(src))
            self.assertEqual(core._all_sect_prs(doc), [])
            core.set_page_margins(doc, 3.7, 3.5, 2.8, 2.6)
            self.assertEqual(len(core._all_sect_prs(doc)), 1)
            self.assertEqual(
                core._all_sect_prs(doc)[0].find(W + 'pgMar').get(W + 'top'),
                str(core.cm_to_twips(3.7)))
            self.assertTrue(core.add_page_number(doc))
            self.assertIsNotNone(core._all_sect_prs(doc)[0].find(W + 'footerReference'))
            doc.save(str(src))
            self.assertEqual(core.Document(str(src)).full_text().strip(), '只有一段文字。')

    # L4: map_pages 里不再有 _cleanup_zombie（不强杀 Word/WPS）。
    def test_cleanup_zombie_is_gone(self):
        self.assertFalse(hasattr(maps, '_cleanup_zombie'))

    # L1: 500 段长文档跑通，段落文本按顺序完整保留（enumerate 取下标，
    # 不再用 paragraphs.index(p)）；含大量重复段落时也不错位。
    def test_long_document_500_paragraphs(self):
        import time
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'long.docx', Path(tmp) / 'long_out.docx'
            texts = ['关于开展学习的通知', '各部门：', '一、学习安排'] + [
                '这是重复的正文段落，用于验证长文档处理。'] * 497
            make_docx(src, texts)
            start = time.perf_counter()
            self.run_script('format_gongwen.py', src, out, '--mode', 'gongwen')
            elapsed = time.perf_counter() - start
            self.assertLess(elapsed, 30)
            got = [core.Document.para_text(p) for p in core.Document(str(out)).paragraphs()]
            self.assertEqual(got, texts)

    # L3: 表格宽度四舍五入后没变就不计数；真改了才计数。
    def test_scale_counts_only_changed_tables(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'tbl.docx'
            fixture(src, table=True)
            doc = core.Document(str(src))
            tbl = doc.tables()[0]
            tblPr = core.get_or_add_child(tbl, 'tblPr')
            tblW = core.find_child(tblPr, 'tblW')
            tblW.set(W + 'w', '2')
            grid = core.find_child(tbl, 'tblGrid')
            for gc in grid.findall(W + 'gridCol'):
                gc.set(W + 'w', '1')
            # 比例 1.1：2 -> 2、1 -> 1，四舍五入后都没变，计数应为 0。
            self.assertEqual(core.scale_table_widths(doc, 1000, 1100), 0)
            # 比例 2：宽度确实变了，计数应为 1。
            self.assertEqual(core.scale_table_widths(doc, 1000, 2000), 1)

    # L3: 表格校验失败时，已保存的输出文件要删掉，不留坏文件。
    def test_table_validation_failure_removes_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'input.docx', Path(tmp) / 'formatted.docx'
            fixture(src, table=True)
            real_document = fmt.Document
            calls = {'n': 0}

            def fake_document(path):
                doc = real_document(path)
                calls['n'] += 1
                if calls['n'] >= 2:  # 保存后的回读校验那一次
                    doc.table_texts = lambda: [['__mismatch__']]
                return doc
            fake_document.para_text = real_document.para_text

            with patch.object(fmt, 'Document', fake_document), \
                 patch.object(sys, 'argv', ['format_gongwen.py', str(src), str(out), '--mode', 'gongwen']):
                with self.assertRaises(RuntimeError):
                    fmt.main()
            self.assertFalse(out.exists())


if __name__ == '__main__':
    unittest.main()
