"""Reduz o Warranty Report (leve).docx a um unico bloco de exemplo (o primeiro
registro), mantendo cabecalho/rodape/marca d'agua/formatacao. Esse arquivo
passa a ser o modelo-semente usado pelo app: o primeiro bloco serve apenas
como referencia de formatacao para clonar novos blocos e e removido do
resultado final antes de salvar."""

from pathlib import Path

import docx
from docx.oxml.ns import qn

SRC = Path(r"c:\Automacao\Warranty Report (leve).docx")
DST = Path(r"c:\Automacao\warranty_app\assets\Warranty Report_template.docx")


def main():
    doc = docx.Document(SRC)
    body = doc.element.body

    heading_indices = []
    children = list(body.iterchildren())
    for idx, el in enumerate(children):
        if el.tag == qn("w:p"):
            p = docx.text.paragraph.Paragraph(el, doc)
            if "Warranty Report" in p.text:
                heading_indices.append(idx)

    if len(heading_indices) < 2:
        raise SystemExit("Esperava encontrar pelo menos 2 blocos de exemplo no documento.")

    cut_start = heading_indices[1]
    sect_pr = body.find(qn("w:sectPr"))
    cut_end = list(body.iterchildren()).index(sect_pr)

    to_remove = children[cut_start:cut_end]
    for el in to_remove:
        body.remove(el)

    DST.parent.mkdir(parents=True, exist_ok=True)
    doc.save(DST)
    print(f"Blocos removidos: {len(to_remove)}")
    print(f"Modelo-semente salvo em: {DST}")


if __name__ == "__main__":
    main()
