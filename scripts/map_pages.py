# -*- coding: utf-8 -*-
"""
段落 → 页码/行号映射脚本（尽力而为，失败优雅降级）。

依赖：仅 Python 标准库。页码映射需要 pywin32 + Word/WPS COM（可选增强）。

Word / WPS 双适配：
  - 依次探测 Word.Application 和 WPS（Kwps.Application），哪个可用用哪个；
  - WPS 隐藏模式易触发 RPC 崩溃，故不设 Visible=False；
  - 中文文件名会触发 WPS 拦截，复制成 ASCII 名再打开。

用法：python map_pages.py <成品.docx>
输出：每段一行「第N段 | 第X页第Y行起 | 段落开头文字」；COM 不可用时降级输出段落序号+原文。
"""
import sys
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))


# 注意：这里曾经有一个 _cleanup_zombie()，用 taskkill /F 强杀
# winword.exe / wps.exe 等进程，全文件没有任何调用点，已删除。
# 不要恢复它：它一旦被调用会杀掉用户正在用的 Word/WPS。
def _try_word_progids():
    """返回可用的 Word/WPS COM ProgID 列表（按优先级）。"""
    # WPS 通常兼容 Word.Application，但若机器装了真 Word 优先用 Word
    return ["Word.Application", "Kwps.Application"]


def _dispatch_available():
    """探测可用的 Word/WPS COM，返回 (progid, dispatch) 或 (None, None)。"""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return None, None
    for progid in _try_word_progids():
        try:
            word = win32com.client.DispatchEx(progid)
            # 探活：读 Version（不设 Visible，避免 WPS RPC 崩溃）
            _ = word.Version
            return progid, word
        except Exception:
            try:
                word.Quit()
            except Exception:
                pass
            continue
    return None, None


def main():
    if len(sys.argv) < 2:
        print("用法：python map_pages.py <成品.docx>")
        sys.exit(1)
    src = Path(sys.argv[1])
    if not src.exists():
        print(f"[错误] 文件不存在：{src}")
        sys.exit(1)

    # 先试 COM（Word/WPS）
    progid, word = _dispatch_available()
    doc = None
    tmp = None
    if word is not None:
        # 中文文件名会触发 WPS 拦截，复制成 ASCII 名再打开
        tmp = src.with_name("__map_tmp.docx")
        shutil.copy2(src, tmp)
        try:
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(tmp), ReadOnly=True)
            paras = doc.Paragraphs
            total = paras.Count
            print(f"[页码映射] 引擎：{progid}")
            n = 0
            for i in range(1, total + 1):
                p = paras(i)
                txt = p.Range.Text.strip()
                if not txt:
                    continue
                n += 1
                try:
                    page = p.Range.Information(3)
                except Exception:
                    page = -1
                try:
                    line = p.Range.Information(10)
                except Exception:
                    line = -1
                t = txt[:24] + "…" if len(txt) > 24 else txt
                print(f"第{n}段 | 第{page}页第{line}行起 | {t}")
        except Exception as e:
            print(f"[降级] 页码映射不可用（{type(e).__name__}），改用段落序号+原文定位。")
            _fallback_text(src)
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                word.Quit()
            except Exception:
                pass
            if tmp is not None:
                try:
                    tmp.unlink(missing_ok=True)
                except Exception:
                    pass
    else:
        print("[降级] 本机无 Word/WPS（或 COM 不可用），改用段落序号+原文定位。")
        _fallback_text(src)


def _fallback_text(src):
    """降级：用 docx_lite 读段落序号 + 原文。"""
    from docx_lite import Document
    d = Document(str(src))
    n = 0
    for p in d.paragraphs():
        t = Document.para_text(p).strip()
        if not t:
            continue
        n += 1
        tt = t[:24] + "…" if len(t) > 24 else t
        print(f"第{n}段 | (页码未知) | {tt}")


if __name__ == "__main__":
    main()
