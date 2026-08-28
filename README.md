# Warranty Report Builder

Aplicativo desktop (Windows) que converte uma planilha de chamados de garantia em um relatório Word formatado, pronto para envio.

## Como usar

1. Abra `WarrantyReportBuilder.exe`.
2. Clique em **Enviar arquivo** e selecione a planilha `xlsx` com os chamados de garantia.
3. O app gera um `Warranty Report.docx` com um bloco formatado por chamado.
4. Clique em **Fazer download** e escolha onde salvar o relatório.

## Formato esperado da planilha

A planilha `.xlsx` DEVE conter estas colunas na primeira linha:

- `ID CHAMADO`
- `CODIGO CLIENTE`
- `CLIENTE ORIGEM`
- `ITEM`
- `QUANTIDADE`
- `SUBTIPO`
- `DOT`

## Estrutura do projeto

```
warranty_app/
├── app.py                      # Interface gráfica (tkinter)
├── report_engine.py            # Lógica de leitura da planilha e geração do relatório
├── assets/
│   └── Warranty Report_template.docx   # Modelo-semente (formatação/marca d'água)
├── scripts_uso_unico/           # Scripts que geraram os assets acima (fora do fluxo de produção)
├── Warranty 2026_anonimizado.xlsx      # Planilha de exemplo para testes
├── WarrantyReportBuilder.spec  # Receita de empacotamento do PyInstaller
├── requirements.txt            # Dependências de runtime
├── requirements-dev.txt        # Dependências de build (inclui pyinstaller)
└── dist/                       # Executável gerado (após o build)
```

A planilha `Warranty 2026_anonimizado.xlsx` segue a estrutura de colunas esperada pelo app e serve como exemplo para testes locais, com dados anonimizados.

## Stack

- Python 3.14
- [tkinter](https://docs.python.org/3/library/tkinter.html) — interface gráfica (biblioteca padrão)
- [openpyxl](https://openpyxl.readthedocs.io/) — leitura da planilha `.xlsx`
- [python-docx](https://python-docx.readthedocs.io/) — geração do relatório `.docx`
- [PyInstaller](https://pyinstaller.org/) — empacotamento em `.exe`


## Setup do ambiente de desenvolvimento

1. Crie o ambiente virtual:
   ```powershell
   python -m venv venv
   ```
2. Ative:
   ```powershell
   .\venv\Scripts\Activate.ps1
   ```
3. Instale as dependências:
   ```powershell
   python -m pip install -r requirements.txt
   ```
4. Rode o app:
   ```powershell
   python app.py
   ```

## Gerando o executável

1. Com a venv ativada, instale as dependências de build:
   ```powershell
   python -m pip install -r requirements-dev.txt
   ```
2. Gere o executável:
   ```powershell
   pyinstaller WarrantyReportBuilder.spec
   ```
3. O arquivo final fica em `dist\WarrantyReportBuilder.exe` — é autossuficiente, pode ser distribuído sozinho (não precisa de Python nem das libs instaladas na máquina do usuário final).


## Notas de arquitetura
 
- Um arquivo `.docx` é internamente um `.zip` com XML. O `python-docx` expõe uma API amigável sobre esse XML.
- `report_engine.py` não monta o relatório do zero: ele parte de um bloco de exemplo já formatado em `assets/Warranty Report_template.docx` (fonte, tabelas, marca d'água) e usa `deepcopy` nos elementos XML pra clonar um bloco novo por chamado. O bloco de exemplo original é removido do resultado final.
- O app é empacotado em modo *onefile* pelo PyInstaller: tudo (assets, Tcl/Tk) fica embutido em um único `.exe`, que se autoextrai para uma pasta temporária a cada execução (`sys._MEIPASS`).

## Scripts de uso único (`scripts_uso_unico/`)

Não fazem parte do app em produção — não são chamados por `app.py` nem `report_engine.py`. Foram usados uma única vez para preparar os artefatos em `assets/` e as pastas `_tcl_data/`/`_tk_data/`. Mantidos como referência caso seja necessário refazer algum desses processos no futuro.

- `strip_photos.py` — remove fotos do corpo de um `.docx` (mantendo o watermark do cabeçalho), gerando uma versão leve.
- `build_seed_template.py` — reduz um `.docx` a um único bloco de exemplo, virando o modelo-semente usado pelo app (`assets/Warranty Report_template.docx`).
- `fix_header_image.py` — ajusta o posicionamento da imagem de marca d'água no XML do cabeçalho do template para ocupar a página inteira.
- `extract_tcltk.py` — extrai os arquivos internos do Tcl/Tk (necessário porque esta instalação de Python guarda o Tcl/Tk numa zipfs virtual que o PyInstaller não empacota sozinho) para as pastas `_tcl_data/`/`_tk_data/`.
- `generate_devlog_docx.py` — gera um diário de desenvolvimento em `.docx` a partir de conteúdo HTML. Sem relação com a lógica do app.
