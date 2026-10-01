import os
import re
import threading
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Cm, Pt, RGBColor
import pdfplumber

# --- ATIVAÇÃO DE ALTA RESOLUÇÃO (DPI AWARENESS) ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

CURSOS_PERMITIDOS = [
    "CLASSE ESPECIAL D.I.",
    "ENSINO FUND.1/5 ANO-SERIE",
    "EDUC INFANTIL",
]


# --- FUNÇÕES DE PROCESSAMENTO DO WORD E PDF ---


def aplicar_cor_fundo_celula(celula, hex_color="E0E0E0"):
    """Aplica cor de fundo (shading) numa célula da tabela do Word."""
    tcPr = celula._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def formatar_paragrafo_celula(
    p, texto, alignment, bold=False, color=None, font_size=12
):
    """Aplica texto e formatação compacta numa célula."""
    p.text = texto
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0

    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(font_size)
        run.bold = bold
        if color:
            run.font.color.rgb = color


def formatar_titulo_turma(curso, seriacao, letra_turma):
    """Gera o título do cabeçalho da turma conforme as regras."""
    curso_upper = curso.upper().strip()
    if "CLASSE ESPECIAL D.I." in curso_upper:
        return f"D.I. - Turma {letra_turma}"
    else:
        return f"{seriacao} - Turma {letra_turma}"


def calcular_larguras_colunas(colunas_extras, largura_total_cm=18.0):
    """Retorna as larguras originais padrão para as colunas da tabela."""
    w_num = 1.0   # Largura original N.º
    w_nome = 11.0 # Largura Nome do Estudante

    qtd_extras = len(colunas_extras)
    if qtd_extras > 0:
        espaco_restante = largura_total_cm - (w_num + w_nome)
        w_col_extra = max(1.5, espaco_restante / qtd_extras)
        w_extras = [w_col_extra] * qtd_extras
    else:
        w_nome = largura_total_cm - w_num
        w_extras = []

    larguras_finais = [Cm(w_num), Cm(w_nome)] + [Cm(w) for w in w_extras]
    return larguras_finais


def extrair_dados_pdf(caminho_pdf, log_callback=None, progress_callback=None):
    """Extrai turmas e alunos do PDF filtrando os cursos permitidos."""
    turmas = {}
    turma_atual = None
    processar_turma_atual = False

    nome_arquivo = os.path.basename(caminho_pdf)
    if log_callback:
        log_callback(f"Lendo {nome_arquivo}")

    with pdfplumber.open(caminho_pdf) as pdf:
        total_paginas = len(pdf.pages)
        for idx, pagina in enumerate(pdf.pages):
            if progress_callback and total_paginas > 0:
                progress_callback((idx + 1) / total_paginas)

            texto = pagina.extract_text()
            if not texto:
                continue

            linhas = texto.split("\n")
            for linha in linhas:
                linha_str = linha.strip()

                if (
                    "Curso:" in linha_str
                    or "Seriação:" in linha_str
                    or "Turma:" in linha_str
                ):
                    match_turma = re.search(
                        r"Curso:\s*(.*?)\s+Seriação:\s*(.*?)(?:\s+Turno:\s*(.*?))?\s+Turma:\s*(.*)",
                        linha_str,
                        re.IGNORECASE,
                    )
                    if match_turma:
                        curso = match_turma.group(1).strip()
                        seriacao = match_turma.group(2).strip()
                        letra_turma = match_turma.group(4).strip()

                        curso_permitido = any(
                            c in curso.upper() for c in CURSOS_PERMITIDOS
                        )

                        if curso_permitido:
                            processar_turma_atual = True
                            turma_atual = formatar_titulo_turma(
                                curso, seriacao, letra_turma
                            )
                            if turma_atual not in turmas:
                                turmas[turma_atual] = {}
                        else:
                            processar_turma_atual = False
                            turma_atual = None
                    continue

                if processar_turma_atual and turma_atual:
                    match_aluno = re.match(
                        r"^\s*(\d+)\s+\d+\s+(.+?)\s+\d{2}/\d{2}/\d{4}.*?\b(Matriculado|Transferido|Remanejado)\b",
                        linha_str,
                        re.IGNORECASE,
                    )

                    if match_aluno:
                        num_chamada = int(match_aluno.group(1))
                        nome_aluno = match_aluno.group(2).strip()
                        situacao = match_aluno.group(3).strip().capitalize()

                        turmas[turma_atual][num_chamada] = {
                            "nome": nome_aluno,
                            "situacao": situacao,
                        }

    return turmas


def ler_dados_word_existente(caminho_word):
    """Lê um ficheiro Word existente, extrai os dados e detecta as colunas extras."""
    if not os.path.exists(caminho_word):
        return {}, []

    doc = Document(caminho_word)
    turmas_existentes = {}
    colunas_extras = []

    if doc.tables:
        primeira_tabela = doc.tables[0]
        num_cols = len(primeira_tabela.rows[0].cells)
        if num_cols > 2:
            for c_idx in range(2, num_cols):
                colunas_extras.append(primeira_tabela.rows[0].cells[c_idx].text.strip())

    tabela_idx = 0
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto:
            turma_atual = texto
            if turma_atual not in turmas_existentes:
                turmas_existentes[turma_atual] = {}

            if tabela_idx < len(doc.tables):
                tabela = doc.tables[tabela_idx]
                for row in tabela.rows[1:]:
                    if not row.cells:
                        continue
                    n_text = row.cells[0].text.strip()
                    nome_text = row.cells[1].text.strip()

                    situacao = "Matriculado"

                    if len(row.cells) >= 3:
                        for cell_extra in row.cells[2:]:
                            ass_text = cell_extra.text.strip().lower()
                            if "transferido" in ass_text:
                                situacao = "Transferido"
                                break
                            elif "remanejado" in ass_text:
                                situacao = "Remanejado"
                                break

                    if "transferido" in nome_text.lower():
                        situacao = "Transferido"
                        nome_text = re.sub(
                            r"\s*-\s*transferido",
                            "",
                            nome_text,
                            flags=re.IGNORECASE,
                        ).strip()
                    elif "remanejado" in nome_text.lower():
                        situacao = "Remanejado"
                        nome_text = re.sub(
                            r"\s*-\s*remanejado",
                            "",
                            nome_text,
                            flags=re.IGNORECASE,
                        ).strip()

                    if n_text.isdigit():
                        num = int(n_text)
                        turmas_existentes[turma_atual][num] = {
                            "nome": nome_text,
                            "situacao": situacao,
                        }
                tabela_idx += 1

    return turmas_existentes, colunas_extras


def gerar_documento_word(
    turmas_dados,
    caminho_saida,
    colunas_extras=None,
    log_callback=None,
    progress_callback=None,
):
    """Gera o ficheiro Word mantendo a largura de coluna fixa."""
    if colunas_extras is None:
        colunas_extras = []

    doc = Document()

    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    primeira_turma = True
    num_cols = 2 + len(colunas_extras)
    titulos = ["N.º", "Nome do Estudante"] + colunas_extras
    larguras = calcular_larguras_colunas(colunas_extras, largura_total_cm=18.0)

    if log_callback:
        log_callback("Gerando tabelas no Word...")

    turmas_chaves = sorted(turmas_dados.keys())
    total_turmas = len(turmas_chaves)

    for idx_turma, turma_titulo in enumerate(turmas_chaves):
        if progress_callback and total_turmas > 0:
            progress_callback((idx_turma + 1) / total_turmas)

        alunos = turmas_dados[turma_titulo]
        if not alunos:
            continue

        if not primeira_turma:
            doc.add_page_break()
        primeira_turma = False

        p_titulo = doc.add_paragraph()
        p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_titulo.paragraph_format.space_after = Pt(6)
        run_t = p_titulo.add_run(turma_titulo)
        run_t.bold = True
        run_t.font.size = Pt(14)
        run_t.font.name = "Arial"

        tabela = doc.add_table(rows=1, cols=num_cols)
        tabela.style = "Table Grid"
        tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabela.autofit = False

        for i, col in enumerate(tabela.columns):
            col.width = larguras[i]

        hdr_cells = tabela.rows[0].cells
        for i, titulo in enumerate(titulos):
            hdr_cells[i].width = larguras[i]
            aplicar_cor_fundo_celula(hdr_cells[i], "2F4F4F")

            p = hdr_cells[i].paragraphs[0]
            align = (
                WD_ALIGN_PARAGRAPH.CENTER
                if i != 1
                else WD_ALIGN_PARAGRAPH.LEFT
            )
            formatar_paragrafo_celula(
                p,
                titulo,
                align,
                bold=True,
                color=RGBColor(255, 255, 255),
                font_size=12,
            )

        for num in sorted(alunos.keys()):
            aluno = alunos[num]
            row_cells = tabela.add_row().cells

            for i in range(num_cols):
                row_cells[i].width = larguras[i]
                row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            situacao_str = aluno["situacao"].capitalize()
            is_especial = situacao_str in ["Transferido", "Remanejado"]

            # Coluna 0: N.º
            formatar_paragrafo_celula(
                row_cells[0].paragraphs[0],
                str(num),
                WD_ALIGN_PARAGRAPH.CENTER,
                bold=True,
                font_size=12,
            )

            if is_especial:
                formatar_paragrafo_celula(
                    row_cells[1].paragraphs[0],
                    aluno["nome"],
                    WD_ALIGN_PARAGRAPH.LEFT,
                    bold=False,
                    font_size=12,
                )

                if len(colunas_extras) >= 2:
                    celula_mesclada_extras = row_cells[2].merge(row_cells[-1])
                    celula_mesclada_extras.width = sum(larguras[2:])
                    formatar_paragrafo_celula(
                        celula_mesclada_extras.paragraphs[0],
                        situacao_str,
                        WD_ALIGN_PARAGRAPH.CENTER,
                        bold=True,
                        font_size=12,
                    )
                elif len(colunas_extras) == 1:
                    formatar_paragrafo_celula(
                        row_cells[2].paragraphs[0],
                        situacao_str,
                        WD_ALIGN_PARAGRAPH.CENTER,
                        bold=True,
                        font_size=12,
                    )
                else:
                    formatar_paragrafo_celula(
                        row_cells[1].paragraphs[0],
                        f"{aluno['nome']} - {situacao_str}",
                        WD_ALIGN_PARAGRAPH.LEFT,
                        bold=False,
                        font_size=12,
                    )

                for cell in row_cells:
                    aplicar_cor_fundo_celula(cell, "E0E0E0")
            else:
                formatar_paragrafo_celula(
                    row_cells[1].paragraphs[0],
                    aluno["nome"],
                    WD_ALIGN_PARAGRAPH.LEFT,
                    bold=False,
                    font_size=12,
                )

                for col_idx in range(len(colunas_extras)):
                    cell_extra = row_cells[2 + col_idx]
                    formatar_paragrafo_celula(
                        cell_extra.paragraphs[0],
                        "",
                        WD_ALIGN_PARAGRAPH.CENTER,
                        bold=False,
                        font_size=12,
                    )

    doc.save(caminho_saida)
    if log_callback:
        log_callback("Documento salvo com sucesso!")


# --- INTERFACE WIZARD COM ABAS LATERAIS E RESUMO ---


class SideTabWizardApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Listas de Alunos")

        self._centralizar_janela(1150, 720)
        self.root.minsize(1000, 620)

        # Dados da Aplicação
        self.opcao_modo = tk.StringVar(value="1")
        self.caminho_word = ""
        self.colunas_extras = []
        self.caminhos_pdfs = []

        self.passo_atual = 1
        self.max_passo_alcancado = 1

        self.passos_info = [
            (1, "Modo de Operação"),
            (2, "Arquivo Word"),
            (3, "Layout da Tabela"),
            (4, "Seleção de PDFs"),
            (5, "Execução e Log"),
        ]

        self._configurar_estilos()
        self._construir_layout_principal()
        self._criar_passos_conteudo()
        self._exibir_passo(1)
        self._atualizar_resumo()

    def _centralizar_janela(self, largura, altura):
        self.root.update_idletasks()
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()

        pos_x = max(0, (largura_tela // 2) - (largura // 2))
        pos_y = max(0, (altura_tela // 2) - (altura // 2))

        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    def _configurar_estilos(self):
        self.style = ttk.Style()

        self.style.configure("SidebarTitle.TLabel", font=("TkDefaultFont", 10, "bold"))
        self.style.configure("Header.TLabel", font=("TkDefaultFont", 12, "bold"))
        self.style.configure("SubHeader.TLabel", font=("TkDefaultFont", 9))
        self.style.configure("SummaryHeader.TLabel", font=("TkDefaultFont", 10, "bold"))
        self.style.configure("SummaryLabel.TLabel", font=("TkDefaultFont", 9, "bold"))
        self.style.configure("SummaryVal.TLabel", font=("TkDefaultFont", 9), wraplength=200)

        self.style.configure("StepItem.TLabel", font=("TkDefaultFont", 9), padding=(8, 6))
        self.style.configure("StepItemActive.TLabel", font=("TkDefaultFont", 9, "bold"), padding=(8, 6))

    def _construir_layout_principal(self):
        self.container = ttk.Frame(self.root)
        self.container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 1. SIDEBAR ESQUERDA
        self.sidebar_left = ttk.Frame(self.container, width=230)
        self.sidebar_left.pack(side=tk.LEFT, fill=tk.Y, expand=False, padx=(5, 5))
        self.sidebar_left.pack_propagate(False)

        lbl_sidebar_title = ttk.Label(
            self.sidebar_left, text="Passos do Processo", style="SidebarTitle.TLabel"
        )
        lbl_sidebar_title.pack(anchor=tk.W, padx=10, pady=(15, 10))

        ttk.Separator(self.sidebar_left, orient="horizontal").pack(fill=tk.X, padx=5, pady=(0, 10))

        self.labels_sidebar = {}
        for num, nome in self.passos_info:
            lbl = ttk.Label(
                self.sidebar_left, text=f"  {num}. {nome}", style="StepItem.TLabel", cursor="hand2"
            )
            lbl.pack(fill=tk.X, padx=5, pady=2)
            lbl.bind("<Button-1>", lambda e, p=num: self._clique_aba_lateral(p))
            self.labels_sidebar[num] = lbl

        btn_reiniciar = ttk.Button(self.sidebar_left, text="⟲ Reiniciar Tudo", command=self.reiniciar)
        btn_reiniciar.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=10)

        ttk.Separator(self.container, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=5)

        # 2. PAINEL DE RESUMO À DIREITA
        self.summary_right = ttk.Frame(self.container, width=250)
        self.summary_right.pack(side=tk.RIGHT, fill=tk.Y, expand=False, padx=(5, 5))
        self.summary_right.pack_propagate(False)

        ttk.Separator(self.container, orient="vertical").pack(side=tk.RIGHT, fill=tk.Y, padx=5)

        lbl_summary_title = ttk.Label(
            self.summary_right, text="Resumo das Escolhas", style="SummaryHeader.TLabel"
        )
        lbl_summary_title.pack(anchor=tk.W, padx=10, pady=(15, 10))

        ttk.Separator(self.summary_right, orient="horizontal").pack(fill=tk.X, padx=5, pady=(0, 10))

        ttk.Label(self.summary_right, text="Modo:", style="SummaryLabel.TLabel").pack(anchor=tk.W, padx=10)
        self.sum_val_modo = ttk.Label(self.summary_right, text="-", style="SummaryVal.TLabel")
        self.sum_val_modo.pack(anchor=tk.W, padx=10, pady=(0, 10))

        ttk.Label(self.summary_right, text="Arquivo Word:", style="SummaryLabel.TLabel").pack(anchor=tk.W, padx=10)
        self.sum_val_word = ttk.Label(self.summary_right, text="Não selecionado", style="SummaryVal.TLabel")
        self.sum_val_word.pack(anchor=tk.W, padx=10, pady=(0, 10))

        ttk.Label(self.summary_right, text="Layout Tabela:", style="SummaryLabel.TLabel").pack(anchor=tk.W, padx=10)
        self.sum_val_layout = ttk.Label(self.summary_right, text="-", style="SummaryVal.TLabel")
        self.sum_val_layout.pack(anchor=tk.W, padx=10, pady=(0, 10))

        ttk.Label(self.summary_right, text="Arquivos PDF:", style="SummaryLabel.TLabel").pack(anchor=tk.W, padx=10)
        self.sum_val_pdfs = ttk.Label(self.summary_right, text="Nenhum PDF selecionado", style="SummaryVal.TLabel")
        self.sum_val_pdfs.pack(anchor=tk.W, padx=10, pady=(0, 10))

        # 3. ÁREA DE CONTEÚDO CENTRAL
        self.center_content = ttk.Frame(self.container, padding="15")
        self.center_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.frames_passos = {}

    def _criar_passos_conteudo(self):
        # Passo 1: Modo
        f1 = ttk.Frame(self.center_content)
        ttk.Label(f1, text="Passo 1: Modo de Operação", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(
            f1,
            text="Escolha se deseja criar um novo arquivo Word do zero ou atualizar um existente.",
            style="SubHeader.TLabel",
        ).pack(anchor=tk.W, pady=(2, 15))

        f1_card = ttk.LabelFrame(f1, text=" Opções ", padding="15")
        f1_card.pack(fill=tk.X, pady=10)

        ttk.Radiobutton(
            f1_card,
            text="Criar um NOVO arquivo Word",
            variable=self.opcao_modo,
            value="1",
            command=self._atualizar_resumo,
        ).pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(
            f1_card,
            text="ATUALIZAR um arquivo Word existente",
            variable=self.opcao_modo,
            value="2",
            command=self._atualizar_resumo,
        ).pack(anchor=tk.W, pady=5)

        self._criar_bar_navegacao(f1, btn_proximo_cmd=lambda: self._avancar_passo(1))
        self.frames_passos[1] = f1

        # Passo 2: Arquivo Word
        f2 = ttk.Frame(self.center_content)
        self.lbl_p2_titulo = ttk.Label(f2, text="Passo 2: Arquivo Word", style="Header.TLabel")
        self.lbl_p2_titulo.pack(anchor=tk.W)
        self.lbl_p2_desc = ttk.Label(f2, text="Selecione o destino do arquivo Word.", style="SubHeader.TLabel")
        self.lbl_p2_desc.pack(anchor=tk.W, pady=(2, 15))

        f2_card = ttk.LabelFrame(f2, text=" Seleção de Arquivo ", padding="15")
        f2_card.pack(fill=tk.X, pady=10)

        self.lbl_word_selected = ttk.Label(f2_card, text="Nenhum arquivo selecionado.", wraplength=550)
        self.lbl_word_selected.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(f2_card, text="Buscar Arquivo...", command=self._selecionar_word).pack(side=tk.RIGHT)

        self._criar_bar_navegacao(
            f2,
            btn_voltar_cmd=lambda: self._voltar_passo(2),
            btn_proximo_cmd=lambda: self._avancar_passo(2),
        )
        self.frames_passos[2] = f2

        # Passo 3: Layout Tabela
        f3 = ttk.Frame(self.center_content)
        ttk.Label(f3, text="Passo 3: Layout da Tabela", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(
            f3,
            text="As colunas 'N.º' e 'Nome do Estudante' são fixas.\nAdicione ou remova colunas adicionais para a sua tabela abaixo:",
            style="SubHeader.TLabel",
        ).pack(anchor=tk.W, pady=(2, 15))

        f3_card = ttk.LabelFrame(f3, text=" Adicionar Nova Coluna ", padding="15")
        f3_card.pack(fill=tk.X, pady=(0, 10))

        frame_input_col = ttk.Frame(f3_card)
        frame_input_col.pack(fill=tk.X)

        ttk.Label(frame_input_col, text="Título da Coluna:").pack(side=tk.LEFT, padx=(0, 5))
        self.entry_nova_coluna = ttk.Entry(frame_input_col)
        self.entry_nova_coluna.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(
            frame_input_col,
            text="➕ Adicionar Coluna",
            command=self._adicionar_coluna_extra,
        ).pack(side=tk.RIGHT)

        f3_list_card = ttk.LabelFrame(f3, text=" Colunas Adicionais Configuradas ", padding="15")
        f3_list_card.pack(fill=tk.BOTH, expand=True, pady=5)

        self.listbox_colunas = tk.Listbox(
            f3_list_card, height=5, font=("TkDefaultFont", 9), selectmode=tk.SINGLE
        )
        self.listbox_colunas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        btn_box_cols = ttk.Frame(f3_list_card)
        btn_box_cols.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(
            btn_box_cols,
            text="❌ Remover Coluna",
            command=self._remover_coluna_extra,
        ).pack(anchor=tk.N)

        self._criar_bar_navegacao(
            f3,
            btn_voltar_cmd=lambda: self._voltar_passo(3),
            btn_proximo_cmd=lambda: self._avancar_passo(3),
        )
        self.frames_passos[3] = f3

        # Passo 4: Seleção PDFs
        f4 = ttk.Frame(self.center_content)
        ttk.Label(f4, text="Passo 4: Seleção de PDFs", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(
            f4,
            text="Escolha um ou mais arquivos PDF com os dados das turmas.",
            style="SubHeader.TLabel",
        ).pack(anchor=tk.W, pady=(2, 15))

        f4_card = ttk.LabelFrame(f4, text=" Arquivos Selecionados ", padding="15")
        f4_card.pack(fill=tk.BOTH, expand=True, pady=10)

        btn_box = ttk.Frame(f4_card)
        btn_box.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(btn_box, text="➕ Adicionar PDFs...", command=self._selecionar_pdfs).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        ttk.Button(
            btn_box,
            text="❌ Remover Selecionado",
            command=self._remover_pdf_selecionado,
        ).pack(side=tk.LEFT)

        self.listbox_pdfs = tk.Listbox(
            f4_card, height=8, font=("TkDefaultFont", 9), selectmode=tk.SINGLE
        )
        self.listbox_pdfs.pack(fill=tk.BOTH, expand=True)

        self._criar_bar_navegacao(
            f4,
            btn_voltar_cmd=lambda: self._voltar_passo(4),
            btn_proximo_cmd=lambda: self._avancar_passo(4),
        )
        self.frames_passos[4] = f4

        # Passo 5: Execução
        f5 = ttk.Frame(self.center_content)
        ttk.Label(f5, text="Passo 5: Processamento e Execução", style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(
            f5,
            text="Clique no botão para iniciar a extração e geração do documento.",
            style="SubHeader.TLabel",
        ).pack(anchor=tk.W, pady=(2, 15))

        self.btn_gerar = ttk.Button(
            f5,
            text="⚡ Gerar Documento",
            command=self._executar_processamento,
        )
        self.btn_gerar.pack(fill=tk.X, ipady=5, pady=(0, 10))

        frame_prog = ttk.Frame(f5)
        frame_prog.pack(fill=tk.X, pady=5)

        self.progress = ttk.Progressbar(frame_prog, mode="determinate", maximum=100, value=0)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.lbl_porcentagem = ttk.Label(
            frame_prog, text="0%", font=("TkDefaultFont", 9, "bold"), width=5, anchor=tk.E
        )
        self.lbl_porcentagem.pack(side=tk.RIGHT)

        ttk.Label(f5, text="Execução:", font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W, pady=(10, 2))
        self.txt_log = scrolledtext.ScrolledText(f5, height=12, state="disabled", font=("Consolas", 9))
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        self._criar_bar_navegacao(
            f5,
            btn_voltar_cmd=lambda: self._voltar_passo(5),
            btn_proximo_txt=None,
        )
        self.frames_passos[5] = f5

    def _criar_bar_navegacao(
        self,
        parent_frame,
        btn_voltar_cmd=None,
        btn_proximo_cmd=None,
        btn_proximo_txt="Avançar ➔",
    ):
        nav_frame = ttk.Frame(parent_frame)
        nav_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 0))

        if btn_voltar_cmd:
            ttk.Button(nav_frame, text="Voltar", command=btn_voltar_cmd).pack(side=tk.LEFT)

        if btn_proximo_cmd and btn_proximo_txt:
            ttk.Button(nav_frame, text=btn_proximo_txt, command=btn_proximo_cmd).pack(side=tk.RIGHT)

    # --- GERENCIAMENTO DAS COLUNAS EXTRAS ---

    def _adicionar_coluna_extra(self):
        titulo = self.entry_nova_coluna.get().strip()
        if not titulo:
            messagebox.showwarning("Atenção", "Digite o título da coluna antes de adicionar.")
            return
        if titulo in self.colunas_extras:
            messagebox.showwarning("Atenção", "Esta coluna já foi adicionada.")
            return

        self.colunas_extras.append(titulo)
        self.entry_nova_coluna.delete(0, tk.END)
        self._atualizar_listbox_colunas()
        self._atualizar_resumo()

    def _remover_coluna_extra(self):
        selecao = self.listbox_colunas.curselection()
        if not selecao:
            messagebox.showwarning("Atenção", "Selecione uma coluna na lista para remover.")
            return

        index = selecao[0]
        del self.colunas_extras[index]
        self._atualizar_listbox_colunas()
        self._atualizar_resumo()

    def _atualizar_listbox_colunas(self):
        self.listbox_colunas.delete(0, tk.END)
        for col in self.colunas_extras:
            self.listbox_colunas.insert(tk.END, col)

    # --- NAVEGAÇÃO E LÓGICA DO WIZARD ---

    def _exibir_passo(self, passo_num):
        if passo_num == 3 and self.opcao_modo.get() == "2":
            if self.passo_atual < 3:
                passo_num = 4
            else:
                passo_num = 2

        self.passo_atual = passo_num
        if passo_num > self.max_passo_alcancado:
            self.max_passo_alcancado = passo_num

        for f in self.frames_passos.values():
            f.pack_forget()

        if passo_num == 2:
            if self.opcao_modo.get() == "1":
                self.lbl_p2_titulo.config(text="Passo 2: Onde salvar o NOVO arquivo Word?")
                self.lbl_p2_desc.config(text="Escolha o local e nome para salvar o documento gerado.")
            else:
                self.lbl_p2_titulo.config(text="Passo 2: Seleção do arquivo Word EXISTENTE")
                self.lbl_p2_desc.config(
                    text="Escolha o arquivo .docx que deseja atualizar com novos dados."
                )

        self.frames_passos[passo_num].pack(fill=tk.BOTH, expand=True)
        self._atualizar_sidebar_visual()

    def _atualizar_sidebar_visual(self):
        modo = self.opcao_modo.get()

        for num, lbl in self.labels_sidebar.items():
            nome_passo = self.passos_info[num - 1][1]

            if num == 3 and modo == "2":
                lbl.config(style="StepItem.TLabel", text=f"  {num}. {nome_passo} (Inativo)")
                continue

            if num == self.passo_atual:
                lbl.config(style="StepItemActive.TLabel", text=f"▶ {num}. {nome_passo}")
            elif num <= self.max_passo_alcancado:
                lbl.config(style="StepItem.TLabel", text=f"✓ {num}. {nome_passo}")
            else:
                lbl.config(style="StepItem.TLabel", text=f"  {num}. {nome_passo}")

    def _clique_aba_lateral(self, passo_num):
        if passo_num <= self.max_passo_alcancado:
            self._exibir_passo(passo_num)

    def _avancar_passo(self, passo_origem):
        if passo_origem == 2 and not self.caminho_word:
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo Word antes de avançar.")
            return

        if passo_origem == 4 and not self.caminhos_pdfs:
            messagebox.showwarning(
                "Atenção", "Por favor, selecione pelo menos um arquivo PDF antes de avançar."
            )
            return

        self._exibir_passo(passo_origem + 1)

    def _voltar_passo(self, passo_origem):
        self._exibir_passo(passo_origem - 1)

    def _selecionar_word(self):
        modo = self.opcao_modo.get()
        if modo == "1":
            caminho = filedialog.asksaveasfilename(
                title="Onde deseja salvar o novo arquivo Word?",
                defaultextension=".docx",
                initialfile="Lista_de_Alunos.docx",
                filetypes=[("Documento Word", "*.docx")],
            )
        else:
            caminho = filedialog.askopenfilename(
                title="Selecione o arquivo Word existente",
                filetypes=[("Documento Word", "*.docx")],
            )

        if caminho:
            self.caminho_word = caminho
            self.lbl_word_selected.config(text=caminho)
            self._atualizar_resumo()

    def _selecionar_pdfs(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecione um ou mais arquivos PDF",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )

        if caminhos:
            for path in caminhos:
                if path not in self.caminhos_pdfs:
                    self.caminhos_pdfs.append(path)

            self._atualizar_listbox_pdfs()
            self._atualizar_resumo()

    def _remover_pdf_selecionado(self):
        selecao = self.listbox_pdfs.curselection()
        if not selecao:
            messagebox.showwarning("Atenção", "Selecione um PDF na lista para remover.")
            return

        index = selecao[0]
        del self.caminhos_pdfs[index]
        self._atualizar_listbox_pdfs()
        self._atualizar_resumo()

    def _atualizar_listbox_pdfs(self):
        self.listbox_pdfs.delete(0, tk.END)
        for path in self.caminhos_pdfs:
            self.listbox_pdfs.insert(tk.END, os.path.basename(path))

    def _atualizar_resumo(self):
        modo = self.opcao_modo.get()
        if modo == "1":
            self.sum_val_modo.config(text="Criar Novo Documento")
        else:
            self.sum_val_modo.config(text="Atualizar Existente")

        if self.caminho_word:
            self.sum_val_word.config(text=os.path.basename(self.caminho_word))
        else:
            self.sum_val_word.config(text="Não selecionado")

        if modo == "2":
            self.sum_val_layout.config(text="Extraído do Word existente")
        else:
            qtd_extras = len(self.colunas_extras)
            if qtd_extras == 0:
                self.sum_val_layout.config(text="2 Colunas (N.º, Nome)")
            else:
                cols_txt = ", ".join(self.colunas_extras)
                self.sum_val_layout.config(text=f"{2 + qtd_extras} Colunas (N.º, Nome, {cols_txt})")

        total_pdfs = len(self.caminhos_pdfs)
        if total_pdfs == 0:
            self.sum_val_pdfs.config(text="Nenhum PDF selecionado")
        else:
            self.sum_val_pdfs.config(text=f"{total_pdfs} arquivo(s) selecionado(s)")

    # --- EXECUÇÃO DO PROCESSAMENTO ---

    def _log(self, mensagem):
        self.txt_log.config(state="normal")
        self.txt_log.insert(tk.END, mensagem + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")

    def _set_progresso(self, valor):
        v_int = int(round(valor))
        self.root.after(0, lambda v=v_int: self._atualizar_interface_progresso(v))

    def _atualizar_interface_progresso(self, v):
        self.progress.configure(value=v)
        self.lbl_porcentagem.config(text=f"{v}%")

    def _executar_processamento(self):
        if not self.caminho_word:
            messagebox.showwarning("Erro", "Arquivo Word não selecionado!")
            return
        if not self.caminhos_pdfs:
            messagebox.showwarning("Erro", "Nenhum arquivo PDF selecionado!")
            return

        self.btn_gerar.config(state="disabled")
        self._set_progresso(0)

        threading.Thread(target=self._worker_processamento, daemon=True).start()

    def _worker_processamento(self):
        try:
            dados_consolidados = {}
            modo = self.opcao_modo.get()

            inicio_pdf = 10 if modo == "2" else 0
            fim_pdf = 70

            if modo == "2":
                self._log("Lendo dados do arquivo Word existente...")
                self._set_progresso(5)
                dados_consolidados, self.colunas_extras = ler_dados_word_existente(self.caminho_word)
                self._set_progresso(10)

            total_pdfs = len(self.caminhos_pdfs)
            for i_pdf, caminho_pdf in enumerate(self.caminhos_pdfs):

                def pdf_prog_cb(fatia_pag, index=i_pdf):
                    prog_pdf = (index + fatia_pag) / total_pdfs
                    val = inicio_pdf + prog_pdf * (fim_pdf - inicio_pdf)
                    self._set_progresso(val)

                novos_dados = extrair_dados_pdf(
                    caminho_pdf,
                    log_callback=self._log,
                    progress_callback=pdf_prog_cb,
                )
                for turma, alunos in novos_dados.items():
                    if turma not in dados_consolidados:
                        dados_consolidados[turma] = {}
                    for num, info in alunos.items():
                        dados_consolidados[turma][num] = info

            self._set_progresso(70)

            def word_prog_cb(fatia_turma):
                val = 70 + fatia_turma * 30
                self._set_progresso(val)

            gerar_documento_word(
                dados_consolidados,
                self.caminho_word,
                colunas_extras=self.colunas_extras,
                log_callback=self._log,
                progress_callback=word_prog_cb,
            )

            self._set_progresso(100)

            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Sucesso",
                    f"Processo concluído com sucesso!\nSalvo em: {self.caminho_word}",
                ),
            )
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Erro", f"Ocorreu um erro no processamento:\n{str(e)}"
                ),
            )
        finally:
            self.root.after(0, self._finalizar_execucao)

    def _finalizar_execucao(self):
        self.btn_gerar.config(state="normal")

    def reiniciar(self):
        self.opcao_modo.set("1")
        self.caminho_word = ""
        self.colunas_extras = []
        self.caminhos_pdfs = []
        self.entry_nova_coluna.delete(0, tk.END)

        self.lbl_word_selected.config(text="Nenhum arquivo selecionado.")
        self._atualizar_listbox_colunas()
        self._atualizar_listbox_pdfs()

        self.passo_atual = 1
        self.max_passo_alcancado = 1

        self._set_progresso(0)
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state="disabled")

        self._exibir_passo(1)
        self._atualizar_resumo()


if __name__ == "__main__":
    root = tk.Tk()
    app = SideTabWizardApp(root)
    root.mainloop()