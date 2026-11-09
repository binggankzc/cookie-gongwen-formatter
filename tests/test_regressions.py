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


class FontAndLayoutTests(unittest.TestCase):
    def test_mac_font_name_reader(self):
        if sys.platform == 'darwin':
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


if __name__ == '__main__':
    unittest.main()
