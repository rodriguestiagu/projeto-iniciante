
from pathlib import Path
from PyPDF2 import PdfMerger

base = Path(__file__).resolve().parent  # pasta onde está o mesclar.py
pasta = base                             # PDFs na mesma pasta do script
saida = base / "arquivo_final.pdf"

merger = PdfMerger(strict=False)

for pdf in sorted(pasta.glob("*.pdf")):
    if pdf.name != "arquivo_final.pdf":  # evita incluir o PDF final se já existir
        merger.append(pdf)

with open(saida, "wb") as f_out:
    merger.write(f_out)

merger.close()
print("PDF final criado em:", saida)

