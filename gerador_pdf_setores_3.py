# -*- coding: utf-8 -*-
"""
GERADOR DE PDF POR SETOR
========================
Tela simples:
    1) Selecionar o arquivo Excel
    2) Escolher a ABA (campanha) -> ou "TODAS as abas"
    3) Escolher a pasta de saida
    4) Clicar em GERAR PDFs

Gera 1 PDF para cada setor, mantendo TODAS as colunas e repetindo o
cabecalho no topo de cada pagina.

INSTALAR (so na primeira vez):
    pip install pandas openpyxl reportlab ftfy
    (o ftfy eh opcional: corrige acentos corrompidos da planilha, ex.: CLAUDIO)
RODAR:
    python gerador_pdf_setores.py
"""

import os
import re
import unicodedata
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pandas as pd
except ImportError:
    raise SystemExit("Falta 'pandas'. Rode:  pip install pandas openpyxl reportlab")

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
except ImportError:
    raise SystemExit("Falta 'reportlab'. Rode:  pip install reportlab")


# ftfy eh OPCIONAL: corrige acentos corrompidos (mojibake) vindos da planilha.
try:
    import ftfy
    def corrigir_texto(s):
        if s is None:
            return ""
        try:
            return ftfy.fix_text(str(s))
        except Exception:
            return str(s)
except ImportError:
    def corrigir_texto(s):
        return "" if s is None else str(s)


# ----------------------------- estilos -------------------------------
ESTILO_CELULA = ParagraphStyle("cel", fontName="Helvetica", fontSize=7,
                               leading=8.5, wordWrap="CJK")
ESTILO_CABECALHO = ParagraphStyle("cab", fontName="Helvetica-Bold", fontSize=7.5,
                                  leading=9, textColor=colors.white, wordWrap="CJK")
ESTILO_TITULO = ParagraphStyle("tit", fontName="Helvetica-Bold", fontSize=14,
                               leading=16, textColor=colors.HexColor("#1a1a1a"))
ESTILO_SUBTITULO = ParagraphStyle("sub", fontName="Helvetica", fontSize=9,
                                  leading=11, textColor=colors.HexColor("#666666"))

OPCAO_TODAS = "TODAS as abas (separa em subpastas)"


# --------------------------- utilidades ------------------------------
def normalizar(texto):
    t = str(texto).strip().lower()
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))


def limpar_nome_arquivo(texto):
    texto = str(texto).strip()
    texto = re.sub(r'[\\/:*?"<>|]', "_", texto)
    texto = re.sub(r"\s+", " ", texto)
    return (texto[:120] or "sem_setor")


def chave_setor(x):
    """Ordena 11,12,101,102 (numerico quando da, senao texto)."""
    s = str(x).strip()
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


def achar_coluna_setor(df):
    for c in df.columns:
        if normalizar(c) == "setor":
            return c
    for c in df.columns:
        if "setor" in normalizar(c):
            return c
    return None


def detectar_linha_cabecalho(xls, aba, max_busca=20):
    """Acha a linha onde fica o cabecalho: a primeira (dentro das primeiras
    linhas) que contenha uma celula com a palavra 'setor'. Resolve o caso de
    planilhas com titulo/linhas em branco acima do cabecalho real."""
    bruto = xls.parse(sheet_name=aba, header=None, nrows=max_busca, dtype=str)
    for i in range(len(bruto)):
        for v in bruto.iloc[i].tolist():
            if v is not None and "setor" in normalizar(v):
                return i
    return 0  # nao achou -> assume primeira linha (sera tratado depois)


def ler_aba(xls, aba):
    linha = detectar_linha_cabecalho(xls, aba)
    df = xls.parse(sheet_name=aba, header=linha, dtype=str)
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [corrigir_texto(c) for c in df.columns]
    return df


def calcular_larguras(colunas, df_setor, largura_util):
    pesos = []
    amostra = df_setor.head(200)
    for col in colunas:
        tam_cab = len(str(col))
        if len(amostra) > 0:
            m = amostra[col].astype(str).str.len().max()
            tam_dados = 0 if pd.isna(m) else int(m)
        else:
            tam_dados = 0
        pesos.append(min(max(tam_cab + 1, tam_dados, 5), 40))
    total = sum(pesos) or 1
    return [largura_util * (p / total) for p in pesos]


def gerar_pdf_setor(nome_setor, df_setor, colunas, caminho_pdf, origem, aba):
    doc = SimpleDocTemplate(
        caminho_pdf, pagesize=landscape(A4),
        leftMargin=1.0 * cm, rightMargin=1.0 * cm,
        topMargin=1.0 * cm, bottomMargin=1.0 * cm,
        title=f"Setor {nome_setor}")
    largura_util = doc.width

    elementos = [
        Paragraph(f"Setor: {nome_setor}", ESTILO_TITULO),
        Paragraph(f"{aba}  |  Origem: {origem}  |  {len(df_setor)} linha(s)",
                  ESTILO_SUBTITULO),
        Spacer(1, 0.3 * cm),
    ]

    cabecalho = [Paragraph(corrigir_texto(c), ESTILO_CABECALHO) for c in colunas]
    dados = [cabecalho]
    for _, linha in df_setor.iterrows():
        dados.append([Paragraph("" if pd.isna(v) else corrigir_texto(v), ESTILO_CELULA)
                      for v in linha[colunas].tolist()])

    larguras = calcular_larguras(colunas, df_setor, largura_util)
    tabela = Table(dados, colWidths=larguras, repeatRows=1)
    estilo = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2a4a")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ])
    for i in range(1, len(dados)):
        if i % 2 == 0:
            estilo.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f2f6fa"))
    tabela.setStyle(estilo)
    elementos.append(tabela)
    doc.build(elementos)


def processar_aba(xls, aba, pasta_destino, origem, callback):
    """Gera os PDFs de uma aba. callback(i, total, texto) atualiza a tela."""
    df = ler_aba(xls, aba)
    col = achar_coluna_setor(df)
    if col is None:
        return 0, f"aba '{aba}': sem coluna setor"
    df[col] = df[col].fillna("(sem setor)").astype(str).str.strip()
    df.loc[df[col] == "", col] = "(sem setor)"

    os.makedirs(pasta_destino, exist_ok=True)
    colunas = list(df.columns)
    setores = sorted(df[col].unique(), key=chave_setor)
    total = len(setores)
    usados = {}
    for i, setor in enumerate(setores, start=1):
        df_setor = df[df[col] == setor]
        nome = limpar_nome_arquivo(setor)
        if nome in usados:
            usados[nome] += 1
            nome = f"{nome}_{usados[nome]}"
        else:
            usados[nome] = 1
        caminho_pdf = os.path.join(pasta_destino, f"{nome}.pdf")
        callback(i, total, f"[{aba}] setor {setor} ({len(df_setor)} linhas)")
        gerar_pdf_setor(setor, df_setor, colunas, caminho_pdf, origem, aba)
    return total, None


# ------------------------------- tela --------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.caminho_excel = None
        self.pasta_saida = None
        self.xls = None

        root.title("Gerador de PDF por Setor")
        root.geometry("570x340")
        root.resizable(False, False)

        ttk.Label(root, text="Gerador de PDF por Setor",
                  font=("Segoe UI", 13, "bold")).pack(pady=(14, 2))
        ttk.Label(root, text="Gera 1 PDF para cada setor da planilha.",
                  foreground="#666").pack()

        # 1) Arquivo
        f1 = ttk.Frame(root); f1.pack(fill="x", padx=18, pady=(14, 4))
        ttk.Label(f1, text="1)  Arquivo Excel:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        l1 = ttk.Frame(f1); l1.pack(fill="x", pady=(2, 0))
        self.lbl_arquivo = ttk.Label(l1, text="Nenhum arquivo selecionado.",
                                     foreground="#888")
        self.lbl_arquivo.pack(side="left", fill="x", expand=True)
        ttk.Button(l1, text="Selecionar...",
                   command=self.selecionar_excel).pack(side="right")

        # 2) Aba
        f2 = ttk.Frame(root); f2.pack(fill="x", padx=18, pady=4)
        ttk.Label(f2, text="2)  Aba (campanha):",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.cb_aba = ttk.Combobox(f2, state="disabled", width=50)
        self.cb_aba.pack(anchor="w", pady=(2, 0))

        # 3) Pasta
        f3 = ttk.Frame(root); f3.pack(fill="x", padx=18, pady=4)
        ttk.Label(f3, text="3)  Pasta de saida:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        l3 = ttk.Frame(f3); l3.pack(fill="x", pady=(2, 0))
        self.lbl_saida = ttk.Label(l3, text="(padrao: subpasta ao lado do Excel)",
                                   foreground="#888")
        self.lbl_saida.pack(side="left", fill="x", expand=True)
        ttk.Button(l3, text="Escolher...",
                   command=self.escolher_pasta).pack(side="right")

        # 4) Gerar
        self.btn_gerar = ttk.Button(root, text="GERAR PDFs", command=self.iniciar)
        self.btn_gerar.pack(pady=(12, 6), ipadx=24, ipady=4)
        self.barra = ttk.Progressbar(root, mode="determinate", length=530)
        self.barra.pack(padx=18)
        self.lbl_status = ttk.Label(root, text="", foreground="#0b2a4a")
        self.lbl_status.pack(pady=(6, 0))

    def selecionar_excel(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if not caminho:
            return
        try:
            self.xls = pd.ExcelFile(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao abrir", str(e))
            return
        self.caminho_excel = caminho
        self.lbl_arquivo.config(text=os.path.basename(caminho), foreground="#000")
        valores = list(self.xls.sheet_names)
        if len(valores) > 1:
            valores = valores + [OPCAO_TODAS]
        self.cb_aba.config(state="readonly", values=valores)
        self.cb_aba.current(0)

    def escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Escolha onde salvar os PDFs")
        if pasta:
            self.pasta_saida = pasta
            self.lbl_saida.config(text=pasta, foreground="#000")

    def iniciar(self):
        if not self.caminho_excel:
            messagebox.showwarning("Atencao", "Selecione o arquivo Excel primeiro.")
            return
        if not self.cb_aba.get():
            messagebox.showwarning("Atencao", "Escolha a aba.")
            return
        self.btn_gerar.config(state="disabled")
        threading.Thread(target=self.gerar, daemon=True).start()

    def _cb(self, i, total, texto):
        self.barra.config(maximum=total, value=i)
        self.lbl_status.config(text=f"Gerando {i}/{total}: {texto}")
        self.root.update_idletasks()

    def gerar(self):
        try:
            origem = os.path.basename(self.caminho_excel)
            base = os.path.splitext(origem)[0]
            raiz = self.pasta_saida or os.path.join(
                os.path.dirname(self.caminho_excel), f"PDFs_{base}")

            escolha = self.cb_aba.get()
            abas = self.xls.sheet_names if escolha == OPCAO_TODAS else [escolha]

            total_geral = 0
            for aba in abas:
                destino = os.path.join(raiz, limpar_nome_arquivo(aba)) \
                    if escolha == OPCAO_TODAS else raiz
                qtd, erro = processar_aba(self.xls, aba, destino, origem, self._cb)
                if erro:
                    self.lbl_status.config(text="Aviso: " + erro)
                total_geral += qtd

            self.lbl_status.config(text=f"Concluido! {total_geral} PDF(s) em: {raiz}")
            messagebox.showinfo("Pronto!",
                                f"Foram gerados {total_geral} PDF(s).\n\nPasta:\n{raiz}")
            try:
                os.startfile(raiz)
            except Exception:
                pass
        except Exception:
            messagebox.showerror("Erro", traceback.format_exc())
        finally:
            self.btn_gerar.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
