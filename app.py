"""Interface grafica do Warranty Report Builder.

Fluxo:
 1) "Enviar arquivo" -> escolhe a planilha (.xlsx) com os chamados de garantia.
    O app filtra as colunas e monta um Warranty Report.docx novo, contendo
    somente os registros dessa planilha, com a formatacao/marca d'agua do
    modelo.
 2) "Fazer download" -> salva uma copia desse relatorio onde o usuario quiser.
"""

import os
import shutil
import sys
import traceback
from pathlib import Path


def _configure_tcl_tk():
    """Aponta TCL_LIBRARY/TK_LIBRARY para os arquivos extraidos e empacotados
    junto do executavel, ja que esta build do Python guarda o Tcl/Tk numa
    zipfs virtual que o PyInstaller nao consegue empacotar automaticamente."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    tcl_dir = base / "_tcl_data"
    tk_dir = base / "_tk_data"
    if tcl_dir.exists():
        os.environ["TCL_LIBRARY"] = str(tcl_dir)
    if tk_dir.exists():
        os.environ["TK_LIBRARY"] = str(tk_dir)


_configure_tcl_tk()

import tkinter as tk
from tkinter import filedialog, messagebox

import report_engine as engine

WINDOW_TITLE = "Warranty Report Builder"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("520x320")
        self.resizable(False, False)

        self.last_report_path = None

        tk.Label(
            self,
            text="Warranty Report Builder",
            font=("Segoe UI", 16, "bold"),
        ).pack(pady=(20, 4))

        tk.Label(
            self,
            text="Envie a planilha de garantia para gerar o relatorio.",
            font=("Segoe UI", 10),
        ).pack(pady=(0, 20))

        self.send_button = tk.Button(
            self,
            text="Enviar arquivo",
            font=("Segoe UI", 11),
            width=24,
            height=2,
            command=self.on_send_file,
        )
        self.send_button.pack(pady=8)

        self.download_button = tk.Button(
            self,
            text="Fazer download",
            font=("Segoe UI", 11),
            width=24,
            height=2,
            state=tk.DISABLED,
            command=self.on_download,
        )
        self.download_button.pack(pady=8)

        self.status_label = tk.Label(
            self,
            text="Nenhum arquivo processado ainda.",
            font=("Segoe UI", 9),
            fg="#444444",
            wraplength=460,
            justify="center",
        )
        self.status_label.pack(pady=(20, 0))

    def on_send_file(self):
        xlsx_path = filedialog.askopenfilename(
            title="Selecione a planilha de garantia",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if not xlsx_path:
            return

        self.send_button.config(state=tk.DISABLED)
        self.status_label.config(text="Processando...")
        self.update_idletasks()

        try:
            added, duplicated, report_file = engine.process_xlsx_into_report(xlsx_path)
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror(WINDOW_TITLE, f"Erro ao processar o arquivo:\n{exc}")
            self.status_label.config(text="Falha ao processar o arquivo.")
            self.send_button.config(state=tk.NORMAL)
            return

        self.last_report_path = report_file
        self.download_button.config(state=tk.NORMAL)
        self.send_button.config(state=tk.NORMAL)

        msg = f"Relatorio gerado com {len(added)} registro(s)."
        if duplicated:
            msg += f" {len(duplicated)} repetido(s) na planilha foram ignorados."
        self.status_label.config(text=msg)

    def on_download(self):
        if not self.last_report_path:
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar relatorio como",
            defaultextension=".docx",
            initialfile="Warranty Report.docx",
            filetypes=[("Documento Word", "*.docx")],
        )
        if not destino:
            return
        try:
            shutil.copy(self.last_report_path, destino)
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror(WINDOW_TITLE, f"Erro ao salvar o arquivo:\n{exc}")
            return
        self.status_label.config(text=f"Relatorio salvo em: {destino}")


if __name__ == "__main__":
    App().mainloop()
