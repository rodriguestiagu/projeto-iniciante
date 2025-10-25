
"""
Fluxo:
  1) Converter .docx/.xlsx/.xlsm -> .pdf (somente na pasta do script; sem subpastas).
  2) Apagar originais .docx/.xlsx/.xlsm após conversão bem-sucedida.
  3) Mesclar todos os PDFs por data (mais recente -> mais antigo), excluindo o arquivo final.
  4) Gravar 'arquivo_final.pdf' e, em seguida, apagar os PDFs mesclados (mantendo apenas o final).

Requisitos:
  - Windows com Microsoft Office (Word/Excel) instalado.
  - pip install pywin32 PyPDF2

Observações:
  - Trabalha apenas na pasta raiz (sem subpastas).
  - Ignora arquivos temporários do Office/OneDrive que começam com "~$".
"""

from __future__ import annotations

import sys
import time
import logging
import traceback
from pathlib import Path

# ---------- Mesclagem de PDF (apenas PyPDF2, para evitar aviso do Pylance) ----------
from PyPDF2 import PdfMerger
logging.getLogger("PyPDF2").setLevel(logging.ERROR)  # silencia avisos de PDFs "imperfeitos"

# ---------- Automação Office (COM) ----------
import pythoncom
import win32com.client
from pywintypes import com_error


# ---------------------------- Conversão: Word ----------------------------

def convert_docx_to_pdf(word_app, src: Path, dst: Path, retries: int = 3, delay: float = 1.0) -> None:
    """
    Converte .docx para .pdf usando Word COM.
    - Usa retries para mitigar (-2147418111) "A chamada foi rejeitada pelo chamado."
    - FileFormat=17 => wdFormatPDF
    """
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            doc = word_app.Documents.Open(
                str(src),
                ReadOnly=True,
                ConfirmConversions=False,
                Visible=False
            )
            try:
                doc.SaveAs2(str(dst), FileFormat=17)  # 17 = wdFormatPDF
            finally:
                doc.Close(False)
            return  # sucesso
        except com_error as e:
            last_exc = e
            # RPC_E_CALL_REJECTED
            if getattr(e, "hresult", None) == -2147418111 and attempt < retries:
                time.sleep(delay * attempt)  # backoff crescente
                continue
            raise
        except Exception as e:
            last_exc = e
            if attempt < retries:
                time.sleep(delay * attempt)
                continue
            raise
    if last_exc:
        raise last_exc


# ---------------------------- Conversão: Excel (.xlsx e .xlsm) ----------------------------

def convert_excel_to_pdf(excel_app, src: Path, dst: Path, retries: int = 3, delay: float = 1.0) -> None:
    """
    Converte .xlsx/.xlsm para .pdf usando Excel COM.
    - Workbook.ExportAsFixedFormat(0, ...) => 0 = xlTypePDF
    - Usa retries para mitigar (-2147418111).
    """
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            wb = excel_app.Workbooks.Open(str(src), ReadOnly=True)
            try:
                wb.ExportAsFixedFormat(0, str(dst))  # 0 = xlTypePDF
            finally:
                wb.Close(False)
            return  # sucesso
        except com_error as e:
            last_exc = e
            if getattr(e, "hresult", None) == -2147418111 and attempt < retries:
                time.sleep(delay * attempt)
                continue
            raise
        except Exception as e:
            last_exc = e
            if attempt < retries:
                time.sleep(delay * attempt)
                continue
            raise
    if last_exc:
        raise last_exc


# ---------------------------- Programa principal ----------------------------

def main() -> int:
    base = Path(__file__).resolve().parent  # pasta do script
    pasta = base
    saida = base / "arquivo_final.pdf"

    # Inicializa COM
    pythoncom.CoInitialize()
    word = None
    excel = None

    try:
        # Instancia Word e Excel (uma única vez)
        try:
            word = win32com.client.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone
        except Exception:
            print("Aviso: Word indisponível; .docx não serão convertidos.")

        try:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception:
            print("Aviso: Excel indisponível; .xlsx/.xlsm não serão convertidos.")

        # Lista SOMENTE arquivos da raiz (sem subpastas)
        itens = [p for p in pasta.iterdir() if p.is_file()]

        # Filtra DOCX/XLSX/XLSM e ignora temporários (~$)
        docxs = [p for p in itens if p.suffix.lower() == ".docx" and not p.name.startswith("~$")]
        xlsxs = [p for p in itens if p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]
        xlsms = [p for p in itens if p.suffix.lower() == ".xlsm" and not p.name.startswith("~$")]

        # 1) Converter DOCX -> PDF
        if word and docxs:
            for src in sorted(docxs, key=lambda p: p.name.lower()):
                dst = src.with_suffix(".pdf")
                try:
                    convert_docx_to_pdf(word, src, dst)
                    # Remove o original após sucesso
                    try:
                        src.unlink()
                    except Exception as e:
                        print(f"Aviso: não foi possível apagar {src.name}: {e}")
                    print(f"OK Word -> {dst.name}")
                except Exception as e:
                    print(f"ERRO DOCX {src.name}: {e}")
                    traceback.print_exc(limit=1)

        # 2) Converter XLSX -> PDF
        if excel and xlsxs:
            for src in sorted(xlsxs, key=lambda p: p.name.lower()):
                dst = src.with_suffix(".pdf")
                try:
                    convert_excel_to_pdf(excel, src, dst)
                    try:
                        src.unlink()
                    except Exception as e:
                        print(f"Aviso: não foi possível apagar {src.name}: {e}")
                    print(f"OK Excel -> {dst.name}")
                except Exception as e:
                    print(f"ERRO XLSX {src.name}: {e}")
                    traceback.print_exc(limit=1)

        # 3) Converter XLSM -> PDF
        if excel and xlsms:
            for src in sorted(xlsms, key=lambda p: p.name.lower()):
                dst = src.with_suffix(".pdf")
                try:
                    convert_excel_to_pdf(excel, src, dst)
                    try:
                        src.unlink()
                    except Exception as e:
                        print(f"Aviso: não foi possível apagar {src.name}: {e}")
                    print(f"OK Excel (XLSM) -> {dst.name}")
                except Exception as e:
                    print(f"ERRO XLSM {src.name}: {e}")
                    traceback.print_exc(limit=1)

        # 4) Monta a lista de PDFs para mesclar (exceto o final)
        pdfs = [p for p in pasta.glob("*.pdf") if p.name.lower() != saida.name.lower()]
        if not pdfs:
            print("Nenhum PDF para mesclar.")
            return 0

        # Ordena por data de modificação (mais novo -> mais antigo)
        pdfs_ordenados = sorted(pdfs, key=lambda p: p.stat().st_mtime, reverse=True)

        # 5) Mescla
        merger = PdfMerger(strict=False)
        try:
            for pdf in pdfs_ordenados:
                try:
                    merger.append(str(pdf))
                except Exception as e:
                    print(f"Aviso: falha ao anexar {pdf.name}: {e}")
            with open(saida, "wb") as f_out:
                merger.write(f_out)
        finally:
            try:
                merger.close()
            except Exception:
                pass

        print(f"PDF final criado em: {saida}")

        # 6) Apagar os PDFs que foram mesclados (manter apenas o final)
        erros = False
        for pdf in pdfs_ordenados:
            try:
                pdf.unlink()
            except Exception as e:
                erros = True
                print(f"Aviso: não foi possível apagar {pdf.name}: {e}")
        if not erros:
            print("PDFs de origem removidos com sucesso.")

        return 0

    finally:
        # Fechar apps do Office e COM
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())
