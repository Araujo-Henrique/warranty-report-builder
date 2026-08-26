"""Ajusta o posicionamento da imagem de cabecalho/rodape no modelo-semente
para que ela ocupe a pagina inteira (borda a borda), em vez de ficar
ancorada relativa a margem (o que deixava uma folga em branco antes dela)."""

import re
import zipfile
from pathlib import Path

TEMPLATE = Path(r"c:\Automacao\warranty_app\assets\Warranty Report_template.docx")

OLD_STYLE_FRAGMENT = (
    "position:absolute;left:0pt;margin-left:-73.15pt;margin-top:-72.4pt;"
    "height:792pt;width:613.5pt;mso-position-horizontal-relative:margin;"
    "mso-position-vertical-relative:margin;"
)
NEW_STYLE_FRAGMENT = (
    "position:absolute;left:0pt;top:0pt;margin-left:0pt;margin-top:0pt;"
    "height:792pt;width:612pt;mso-position-horizontal-relative:page;"
    "mso-position-vertical-relative:page;"
)


def main():
    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        header_xml = zin.read("word/header1.xml").decode("utf-8")

        if OLD_STYLE_FRAGMENT not in header_xml:
            raise SystemExit("Fragmento de estilo esperado nao encontrado em header1.xml")

        new_header_xml = header_xml.replace(OLD_STYLE_FRAGMENT, NEW_STYLE_FRAGMENT)

        tmp_path = TEMPLATE.with_suffix(".tmp.docx")
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/header1.xml":
                    zout.writestr(item, new_header_xml.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))

    tmp_path.replace(TEMPLATE)
    print("header1.xml atualizado.")


if __name__ == "__main__":
    main()
