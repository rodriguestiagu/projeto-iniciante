
from pathlib import Path
from PyPDF2 import PdfMerger

base = Path(__file__).resolve().parent   # pasta onde está o mesclar.py
pasta = base                              # PDFs na mesma pasta do script
saida = base / "arquivo_final.pdf"

merger = PdfMerger(strict=False)

# Lista PDFs da pasta raiz, exceto o arquivo final, e ordena por mtime (mais recente primeiro)
pdfs = [
    pdf for pdf in pasta.glob("*.pdf")
    if pdf.name.lower() != saida.name.lower()
]

# Ordenação por data de modificação (mtime): mais novo -> mais antigo
pdfs_ordenados = sorted(pdfs, key=lambda p: p.stat().st_mtime, reverse=True)

for pdf in pdfs_ordenados:
    merger.append(pdf)

with open(saida, "wb") as f_out:
    merger.write(f_out)

merger.close()
print("PDF final criado em:", saida)
