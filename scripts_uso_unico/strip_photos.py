"""Gera uma copia leve do Warranty Report.docx removendo as fotos inline do
corpo do documento (mantém o watermark do cabecalho, que fica em word/media/
e é uma parte separada). O arquivo original nao e alterado."""

import re
import sys
import zipfile
from pathlib import Path

SRC = Path(r"c:\Automacao\Warranty Report.docx")
DST = Path(r"c:\Automacao\Warranty Report (leve).docx")

DRAWING_RE = re.compile(rb"<w:drawing\b.*?</w:drawing>", re.DOTALL)
IMAGE_REL_RE = re.compile(
    rb'<Relationship [^>]*Type="[^"]*?/relationships/image"[^>]*Target="/media/[^"]*"[^>]*/>'
)


def main():
    with zipfile.ZipFile(SRC, "r") as zin:
        names = zin.namelist()
        document_xml = zin.read("word/document.xml")
        rels_xml = zin.read("word/_rels/document.xml.rels")

        drawings_removed = len(DRAWING_RE.findall(document_xml))
        document_xml = DRAWING_RE.sub(b"", document_xml)

        rels_removed = len(IMAGE_REL_RE.findall(rels_xml))
        rels_xml = IMAGE_REL_RE.sub(b"", rels_xml)

        skipped = 0
        with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name == "word/document.xml":
                    zout.writestr(item, document_xml)
                elif name == "word/_rels/document.xml.rels":
                    zout.writestr(item, rels_xml)
                elif name.startswith("media/"):
                    skipped += 1
                    continue
                else:
                    zout.writestr(item, zin.read(name))

    print(f"drawings removidos do corpo: {drawings_removed}")
    print(f"relacionamentos de imagem removidos: {rels_removed}")
    print(f"arquivos de midia (fotos) descartados: {skipped}")
    print(f"salvo em: {DST}")


if __name__ == "__main__":
    main()
