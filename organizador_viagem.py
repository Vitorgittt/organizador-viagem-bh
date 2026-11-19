from datetime import datetime
import json, os, tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import requests

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# COLE AQUI A URL DO SEU FIREBASE (sem barra no final)
FIREBASE_URL = "https://viagem-bh-b2b87-default-rtdb.firebaseio.com/"
ENDPOINT_NUVEM = f"{FIREBASE_URL}/planejamento.json"
ARQUIVO_EXCEL = "Planejamento_Viagem_BH.xlsx"

# Paleta de Cores Moderna (Slate Dark Premium)
COR_FUNDO = "#0B0F19"
COR_CARD = "#131A2A"
COR_BORDA = "#1E293B"
COR_DESTAQUE = "#38BDF8"


class ViagemBHApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Nossa Viagem BH • 08 a 18 de Maio ☁️")
        self.geometry("1040x730")
        self.resizable(False, False)
        self.configure(fg_color=COR_FUNDO)

        self.dados = self.dados_padrao()
        self.criar_interface()
        self.sincronizar_da_nuvem()
        self.auto_atualizar()

    def dados_padrao(self):
        return {
            "gastos": [
                {
                    "desc": "Airbnb Lourdes (10 diárias)",
                    "valor": 3800.0,
                    "tipo": "Meiado (50/50)",
                },
                {
                    "desc": "Carro Compacto Automático (10 dias)",
                    "valor": 1400.0,
                    "tipo": "100% Ele",
                },
                {
                    "desc": "Passagem Aérea BSB - CNF",
                    "valor": 550.0,
                    "tipo": "100% Ele",
                },
                {
                    "desc": "Combustível e Estacionamentos",
                    "valor": 500.0,
                    "tipo": "100% Ele",
                },
                {
                    "desc": "Gastronomia (Jantares + Cafés)",
                    "valor": 2400.0,
                    "tipo": "100% Ele",
                },
                {
                    "desc": "Passeios (Inhotim + Ingressos)",
                    "valor": 650.0,
                    "tipo": "100% Ele",
                },
            ],
            "roteiro": [
                {
                    "dia": "Dia 1 • 08/05 (Sáb)",
                    "cat": "BH Centro-Sul",
                    "atividade": "Chegada + Jantar no Glouton",
                },
                {
                    "dia": "Dia 2 • 09/05 (Dom)",
                    "cat": "BH Centro-Sul",
                    "atividade": "Circuito Liberdade + Almoço no Ninita",
                },
                {
                    "dia": "Dia 3 • 10/05 (Seg)",
                    "cat": "BH Centro-Sul",
                    "atividade": "Mercado Central + Bistrô Taste-Vin",
                },
                {
                    "dia": "Dia 4 • 11/05 (Ter)",
                    "cat": "Bate-Volta",
                    "atividade": "Bate-volta romântico a Ouro Preto",
                },
                {
                    "dia": "Dia 5 • 12/05 (Qua)",
                    "cat": "Bate-Volta",
                    "atividade": "Instituto Inhotim (Entrada gratuita)",
                },
                {
                    "dia": "Dia 6 • 13/05 (Qui)",
                    "cat": "Pampulha / Outros",
                    "atividade": "Lagoa da Pampulha + Jantar no O Conde",
                },
                {
                    "dia": "Dia 7 • 14/05 (Sex)",
                    "cat": "BH Centro-Sul",
                    "atividade": "Noite cultural e drinks na Savassi",
                },
                {
                    "dia": "Dia 8 • 15/05 (Sáb)",
                    "cat": "Bate-Volta",
                    "atividade": "Passeio e almoço na Serra em Macacos",
                },
                {
                    "dia": "Dia 9 • 16/05 (Dom)",
                    "cat": "BH Centro-Sul",
                    "atividade": "Menu Degustação no Pacato + Mirante",
                },
                {
                    "dia": "Dia 10 • 17/05 (Seg)",
                    "cat": "BH Centro-Sul",
                    "atividade": "Jantar de despedida no Trintaeum",
                },
            ],
            "depositos": [],
            "malas": [
                {
                    "item": "CNH e Documento de Identidade",
                    "cat": "Documentos",
                    "feito": False,
                },
                {
                    "item": "Cartão de crédito físico (caução do carro)",
                    "cat": "Documentos",
                    "feito": False,
                },
                {
                    "item": "Roupas esporte fino + jaqueta leve para noite",
                    "cat": "Roupas",
                    "feito": False,
                },
                {
                    "item": "Tênis confortável (Ouro Preto e Inhotim)",
                    "cat": "Roupas",
                    "feito": False,
                },
                {
                    "item": "Kit higiene pessoal e hidratante labial",
                    "cat": "Higiene & Saúde",
                    "feito": False,
                },
                {
                    "item": "Carregador portátil (Power Bank) e cabos",
                    "cat": "Eletrônicos",
                    "feito": False,
                },
            ],
            "preparativos": [
                {
                    "item": "Comprar passagem aérea BSB - CNF",
                    "cat": "Transporte",
                    "feito": False,
                },
                {
                    "item": "Alugar carro compacto automático (Confins)",
                    "cat": "Transporte",
                    "feito": False,
                },
                {
                    "item": "Reservar apartamento no Lourdes (Airbnb)",
                    "cat": "Hospedagem",
                    "feito": False,
                },
                {
                    "item": "Reservar mesas: Glouton, Pacato e Trintaeum",
                    "cat": "Gastronomia",
                    "feito": False,
                },
                {
                    "item": "Emitir ingressos gratuitos Inhotim (12/05)",
                    "cat": "Passeios",
                    "feito": False,
                },
            ],
        }

    def sincronizar_da_nuvem(self):
        try:
            r = requests.get(ENDPOINT_NUVEM, timeout=5)
            if r.status_code == 200:
                nuvem = r.json()
                if nuvem:
                    self.dados["gastos"] = nuvem.get("gastos") or []
                    self.dados["roteiro"] = nuvem.get("roteiro") or []
                    self.dados["depositos"] = nuvem.get("depositos") or []
                    self.dados["malas"] = (
                        nuvem.get("malas") or self.dados["malas"]
                    )
                    self.dados["preparativos"] = (
                        nuvem.get("preparativos") or self.dados["preparativos"]
                    )
                else:
                    self.salvar_dados()
                self.lbl_status.configure(
                    text="● Nuvem Conectada", text_color="#34D399"
                )
        except Exception:
            self.lbl_status.configure(text="● Offline", text_color="#F87171")
        self.atualizar_tudo()

    def salvar_dados(self):
        try:
            requests.put(ENDPOINT_NUVEM, json=self.dados, timeout=5)
            self.lbl_status.configure(
                text="● Salvo na Nuvem", text_color="#34D399"
            )
        except Exception:
            messagebox.showerror(
                "Erro de Rede", "Não foi possível salvar na nuvem."
            )
        self.atualizar_tudo()

    def auto_atualizar(self):
        self.sincronizar_da_nuvem()
        self.after(5000, self.auto_atualizar)

    def criar_interface(self):
        header = ctk.CTkFrame(
            self, height=68, corner_radius=0, fg_color=COR_CARD
        )
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="✈️  Nossa Viagem • Belo Horizonte",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#F8FAFC",
        ).pack(side="left", padx=25, pady=18)

        ctk.CTkLabel(
            header,
            text="08 a 18 de Maio",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#94A3B8",
        ).pack(side="left", padx=(0, 15))

        self.lbl_status = ctk.CTkLabel(
            header,
            text="● Conectando...",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#FBBF24",
        )
        self.lbl_status.pack(side="left", padx=10)

        ctk.CTkButton(
            header,
            text="📊 Exportar Excel (.xlsx)",
            fg_color="#059669",
            hover_color="#047857",
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=self.exportar_excel,
        ).pack(side="right", padx=25, pady=18)

        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.pack(fill="x", padx=20, pady=(15, 10))
        self.lbl_total_viagem = self.criar_card(
            cards, "CUSTO TOTAL PLANEJADO", "R$ 0,00", "#FBBF24"
        )
        self.lbl_total_guardado = self.criar_card(
            cards, "JÁ GUARDADO NA CONTA", "R$ 0,00", "#34D399"
        )
        self.lbl_falta_juntar = self.criar_card(
            cards, "FALTA PARA A META", "R$ 0,00", "#38BDF8"
        )
        self.lbl_parte_dela = self.criar_card(
            cards, "PARTE DELA (MEIADO)", "R$ 0,00", "#C084FC"
        )

        self.tabview = ctk.CTkTabview(
            self,
            width=1000,
            height=520,
            fg_color=COR_CARD,
            segmented_button_fg_color=COR_FUNDO,
            segmented_button_selected_color="#2563EB",
            segmented_button_selected_hover_color="#1D4ED8",
            segmented_button_unselected_color=COR_FUNDO,
            corner_radius=14,
        )
        self.tabview.pack(padx=20, pady=5)

        self.tab_cofrinho = self.tabview.add("🏦 Cofrinho & Gráfico")
        self.tab_gastos = self.tabview.add("💰 Orçamento")
        self.tab_roteiro = self.tabview.add("🗺️ Roteiro")
        self.tab_malas = self.tabview.add("🧳 Malas & Docs")
        self.tab_prep = self.tabview.add("📋 Preparativos")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background=COR_FUNDO,
            foreground="#E2E8F0",
            fieldbackground=COR_FUNDO,
            rowheight=32,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background="#1E293B",
            foreground="#94A3B8",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", "#334155")],
            foreground=[("selected", "#FFFFFF")],
        )

        self.montar_aba_cofrinho()
        self.montar_aba_gastos()
        self.montar_aba_roteiro()
        self.montar_aba_malas()
        self.montar_aba_preparativos()

    def aplicar_cores_suaves(self, tree):
        # Paleta moderna: fundos escuros sutis com texto pastel legível
        tree.tag_configure("roxo", background="#1E1633", foreground="#D8B4FE")
        tree.tag_configure("azul", background="#12213A", foreground="#93C5FD")
        tree.tag_configure("verde", background="#0D2824", foreground="#6EE7B7")
        tree.tag_configure("ambar", background="#2A1D11", foreground="#FCD34D")
        tree.tag_configure("rosa", background="#2D1525", foreground="#F9A8D4")

    def criar_card(self, parent, titulo, valor, cor):
        card = ctk.CTkFrame(
            parent,
            width=230,
            height=88,
            corner_radius=12,
            fg_color=COR_CARD,
            border_width=1,
            border_color=COR_BORDA,
        )
        card.pack(side="left", expand=True, fill="both", padx=6)
        ctk.CTkLabel(
            card,
            text=titulo,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#64748B",
        ).pack(pady=(14, 2))
        lbl = ctk.CTkLabel(
            card,
            text=valor,
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=cor,
        )
        lbl.pack(pady=(0, 14))
        return lbl

    def montar_aba_cofrinho(self):
        form = ctk.CTkFrame(self.tab_cofrinho, fg_color="transparent")
        form.pack(fill="x", pady=10)
        self.entry_dep_data = ctk.CTkEntry(
            form, width=130, fg_color=COR_FUNDO, border_color=COR_BORDA
        )
        self.entry_dep_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_dep_data.pack(side="left", padx=5)
        self.entry_dep_desc = ctk.CTkEntry(
            form,
            placeholder_text="Descrição (ex: Depósito Novembro)",
            width=390,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_dep_desc.pack(side="left", padx=5)
        self.entry_dep_valor = ctk.CTkEntry(
            form,
            placeholder_text="Valor (R$)",
            width=140,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_dep_valor.pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="+ Guardar Dinheiro",
            fg_color="#059669",
            hover_color="#047857",
            width=150,
            command=self.adicionar_deposito,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=self.remover_deposito,
        ).pack(side="left", padx=5)

        content = ctk.CTkFrame(self.tab_cofrinho, fg_color="transparent")
        content.pack(fill="both", expand=True, pady=5)
        left = ctk.CTkFrame(content, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0, 12))
        self.tree_depositos = ttk.Treeview(
            left,
            columns=("data", "desc", "valor", "acumulado"),
            show="headings",
            height=12,
        )
        for col, txt, w, anc in [
            ("data", "Data", 95, "center"),
            ("desc", "Descrição", 200, "w"),
            ("valor", "Depósito", 110, "e"),
            ("acumulado", "Saldo Total", 120, "e"),
        ]:
            self.tree_depositos.heading(col, text=txt)
            self.tree_depositos.column(col, width=w, anchor=anc)
        self.tree_depositos.pack(fill="both", expand=True)

        right = ctk.CTkFrame(
            content,
            width=420,
            corner_radius=12,
            fg_color=COR_FUNDO,
            border_width=1,
            border_color=COR_BORDA,
        )
        right.pack(side="right", fill="both")
        right.pack_propagate(False)
        ctk.CTkLabel(
            right,
            text="📈 Evolução do Cofrinho rumo à Meta",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#E2E8F0",
        ).pack(pady=(12, 5))
        self.canvas_grafico = tk.Canvas(
            right, width=400, height=320, bg=COR_FUNDO, highlightthickness=0
        )
        self.canvas_grafico.pack(padx=10, pady=5)

    def montar_aba_gastos(self):
        form = ctk.CTkFrame(self.tab_gastos, fg_color="transparent")
        form.pack(fill="x", pady=10)
        self.entry_desc = ctk.CTkEntry(
            form,
            placeholder_text="Descrição do gasto...",
            width=390,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_desc.pack(side="left", padx=5)
        self.entry_valor = ctk.CTkEntry(
            form,
            placeholder_text="Valor (R$)",
            width=140,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_valor.pack(side="left", padx=5)
        self.combo_tipo = ctk.CTkComboBox(
            form,
            values=["Meiado (50/50)", "100% Ele", "100% Ela"],
            width=170,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.combo_tipo.pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="+ Adicionar",
            width=120,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=self.adicionar_gasto,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=self.remover_gasto,
        ).pack(side="left", padx=5)

        self.tree_gastos = ttk.Treeview(
            self.tab_gastos,
            columns=("desc", "tipo", "total", "ele", "ela"),
            show="headings",
            height=11,
        )
        for col, txt, w, anc in [
            ("desc", "Item / Experiência", 380, "w"),
            ("tipo", "Divisão", 150, "center"),
            ("total", "Valor Total", 135, "e"),
            ("ele", "Parte Dele", 135, "e"),
            ("ela", "Parte Dela", 135, "e"),
        ]:
            self.tree_gastos.heading(col, text=txt)
            self.tree_gastos.column(col, width=w, anchor=anc)
        self.aplicar_cores_suaves(self.tree_gastos)
        self.tree_gastos.pack(fill="both", expand=True, pady=5)

        barra = ctk.CTkFrame(self.tab_gastos, fg_color="transparent")
        barra.pack(fill="x", pady=(6, 0))
        ctk.CTkLabel(
            barra,
            text="🟣 Meiado (50/50)    🔵 100% Ele    🟢 100% Ela",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94A3B8",
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            barra,
            text="🗑️ Limpar Aba Completa",
            width=165,
            fg_color="#7F1D1D",
            hover_color="#991B1B",
            command=lambda: self.limpar_aba("gastos", "Orçamento"),
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            barra,
            text="Apagar Divisão Selecionada",
            width=195,
            fg_color="#1E293B",
            hover_color="#334155",
            command=lambda: self.limpar_por_categoria(
                "gastos", "tipo", self.combo_tipo.get()
            ),
        ).pack(side="right", padx=5)

    def montar_aba_roteiro(self):
        form = ctk.CTkFrame(self.tab_roteiro, fg_color="transparent")
        form.pack(fill="x", pady=10)
        self.entry_dia = ctk.CTkEntry(
            form,
            placeholder_text="Dia (ex: Dia 11 • 18/05)",
            width=190,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_dia.pack(side="left", padx=5)
        self.combo_rot_cat = ctk.CTkComboBox(
            form,
            values=["BH Centro-Sul", "Bate-Volta", "Pampulha / Outros"],
            width=170,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.combo_rot_cat.pack(side="left", padx=5)
        self.entry_ativ = ctk.CTkEntry(
            form,
            placeholder_text="Atividade ou Restaurante...",
            width=370,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_ativ.pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="+ Adicionar",
            width=120,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=self.adicionar_roteiro,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=self.remover_roteiro,
        ).pack(side="left", padx=5)

        self.tree_roteiro = ttk.Treeview(
            self.tab_roteiro,
            columns=("dia", "cat", "atividade"),
            show="headings",
            height=11,
        )
        self.tree_roteiro.heading("dia", text="Data / Dia")
        self.tree_roteiro.heading("cat", text="Região / Tipo")
        self.tree_roteiro.heading("atividade", text="Programação Planejada")
        self.tree_roteiro.column("dia", width=190)
        self.tree_roteiro.column("cat", width=170, anchor="center")
        self.tree_roteiro.column("atividade", width=575)
        self.aplicar_cores_suaves(self.tree_roteiro)
        self.tree_roteiro.pack(fill="both", expand=True, pady=5)

        barra = ctk.CTkFrame(self.tab_roteiro, fg_color="transparent")
        barra.pack(fill="x", pady=(6, 0))
        ctk.CTkLabel(
            barra,
            text="🔵 BH Centro-Sul    🟠 Bate-Volta    🟣 Pampulha / Outros",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94A3B8",
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            barra,
            text="🗑️ Limpar Aba Completa",
            width=165,
            fg_color="#7F1D1D",
            hover_color="#991B1B",
            command=lambda: self.limpar_aba("roteiro", "Roteiro"),
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            barra,
            text="Apagar Região Selecionada",
            width=195,
            fg_color="#1E293B",
            hover_color="#334155",
            command=lambda: self.limpar_por_categoria(
                "roteiro", "cat", self.combo_rot_cat.get()
            ),
        ).pack(side="right", padx=5)

    def montar_aba_malas(self):
        form = ctk.CTkFrame(self.tab_malas, fg_color="transparent")
        form.pack(fill="x", pady=8)
        self.entry_mala_item = ctk.CTkEntry(
            form,
            placeholder_text="Novo item de mala ou documento...",
            width=480,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_mala_item.pack(side="left", padx=5)
        self.combo_mala_cat = ctk.CTkComboBox(
            form,
            values=["Roupas", "Documentos", "Higiene & Saúde", "Eletrônicos"],
            width=180,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.combo_mala_cat.pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="+ Adicionar",
            width=120,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=lambda: self.adicionar_item_lista(
                "malas", self.entry_mala_item, self.combo_mala_cat
            ),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=lambda: self.remover_item_lista("malas", self.tree_malas),
        ).pack(side="left", padx=5)

        prog = ctk.CTkFrame(self.tab_malas, fg_color="transparent")
        prog.pack(fill="x", padx=5, pady=(0, 6))
        self.lbl_prog_malas = ctk.CTkLabel(
            prog,
            text="Malas Prontas: 0% (0/0)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#94A3B8",
        )
        self.lbl_prog_malas.pack(side="left", padx=(0, 12))
        self.barra_malas = ctk.CTkProgressBar(
            prog, width=740, progress_color="#38BDF8", fg_color=COR_FUNDO
        )
        self.barra_malas.set(0)
        self.barra_malas.pack(side="left", fill="x", expand=True)

        self.tree_malas = ttk.Treeview(
            self.tab_malas,
            columns=("status", "cat", "item"),
            show="headings",
            height=10,
        )
        self.tree_malas.heading("status", text="Status (Duplo clique)")
        self.tree_malas.heading("cat", text="Categoria")
        self.tree_malas.heading("item", text="Item da Mala / Documento")
        self.tree_malas.column("status", width=165, anchor="center")
        self.tree_malas.column("cat", width=165, anchor="center")
        self.tree_malas.column("item", width=605)
        self.aplicar_cores_suaves(self.tree_malas)
        self.tree_malas.pack(fill="both", expand=True)
        self.tree_malas.bind(
            "<Double-1>",
            lambda e: self.alternar_item_lista("malas", self.tree_malas),
        )

        barra = ctk.CTkFrame(self.tab_malas, fg_color="transparent")
        barra.pack(fill="x", pady=(6, 0))
        ctk.CTkLabel(
            barra,
            text="🔵 Roupas   🟣 Documentos   🌸 Higiene & Saúde   🟠 Eletrônicos   🟢 Concluído",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8",
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            barra,
            text="🗑️ Limpar Aba Completa",
            width=160,
            fg_color="#7F1D1D",
            hover_color="#991B1B",
            command=lambda: self.limpar_aba("malas", "Malas & Docs"),
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            barra,
            text="Apagar Categoria Selecionada",
            width=190,
            fg_color="#1E293B",
            hover_color="#334155",
            command=lambda: self.limpar_por_categoria(
                "malas", "cat", self.combo_mala_cat.get()
            ),
        ).pack(side="right", padx=5)

    def montar_aba_preparativos(self):
        form = ctk.CTkFrame(self.tab_prep, fg_color="transparent")
        form.pack(fill="x", pady=8)
        self.entry_prep_item = ctk.CTkEntry(
            form,
            placeholder_text="Nova reserva ou tarefa (ex: Reservar restaurante)...",
            width=480,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.entry_prep_item.pack(side="left", padx=5)
        self.combo_prep_cat = ctk.CTkComboBox(
            form,
            values=["Transporte", "Hospedagem", "Gastronomia", "Passeios"],
            width=180,
            fg_color=COR_FUNDO,
            border_color=COR_BORDA,
        )
        self.combo_prep_cat.pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="+ Adicionar",
            width=120,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=lambda: self.adicionar_item_lista(
                "preparativos", self.entry_prep_item, self.combo_prep_cat
            ),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=lambda: self.remover_item_lista(
                "preparativos", self.tree_prep
            ),
        ).pack(side="left", padx=5)

        prog = ctk.CTkFrame(self.tab_prep, fg_color="transparent")
        prog.pack(fill="x", padx=5, pady=(0, 6))
        self.lbl_prog_prep = ctk.CTkLabel(
            prog,
            text="Preparativos Feitos: 0% (0/0)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#94A3B8",
        )
        self.lbl_prog_prep.pack(side="left", padx=(0, 12))
        self.barra_prep = ctk.CTkProgressBar(
            prog, width=720, progress_color="#34D399", fg_color=COR_FUNDO
        )
        self.barra_prep.set(0)
        self.barra_prep.pack(side="left", fill="x", expand=True)

        self.tree_prep = ttk.Treeview(
            self.tab_prep,
            columns=("status", "cat", "item"),
            show="headings",
            height=10,
        )
        self.tree_prep.heading("status", text="Status (Duplo clique)")
        self.tree_prep.heading("cat", text="Categoria")
        self.tree_prep.heading("item", text="Reserva / Preparativo")
        self.tree_prep.column("status", width=165, anchor="center")
        self.tree_prep.column("cat", width=165, anchor="center")
        self.tree_prep.column("item", width=605)
        self.aplicar_cores_suaves(self.tree_prep)
        self.tree_prep.pack(fill="both", expand=True)
        self.tree_prep.bind(
            "<Double-1>",
            lambda e: self.alternar_item_lista("preparativos", self.tree_prep),
        )

        barra = ctk.CTkFrame(self.tab_prep, fg_color="transparent")
        barra.pack(fill="x", pady=(6, 0))
        ctk.CTkLabel(
            barra,
            text="🔵 Transporte   🟣 Hospedagem   🟠 Gastronomia   🌸 Passeios   🟢 Concluído",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8",
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            barra,
            text="🗑️ Limpar Aba Completa",
            width=160,
            fg_color="#7F1D1D",
            hover_color="#991B1B",
            command=lambda: self.limpar_aba("preparativos", "Preparativos"),
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            barra,
            text="Apagar Categoria Selecionada",
            width=190,
            fg_color="#1E293B",
            hover_color="#334155",
            command=lambda: self.limpar_por_categoria(
                "preparativos", "cat", self.combo_prep_cat.get()
            ),
        ).pack(side="right", padx=5)

    def limpar_por_categoria(self, chave_lista, campo_cat, valor_cat):
        if messagebox.askyesno(
            "Confirmar Exclusão",
            f"Deseja apagar TODOS os itens da categoria '{valor_cat}'?",
        ):
            self.dados[chave_lista] = [
                item
                for item in self.dados.get(chave_lista, [])
                if item.get(campo_cat) != valor_cat
            ]
            self.salvar_dados()

    def limpar_aba(self, chave_lista, nome_aba):
        if messagebox.askyesno(
            "Limpar Aba Completa",
            f"Tem certeza que deseja apagar TODOS os itens da aba '{nome_aba}'?\nEssa ação será sincronizada na nuvem.",
        ):
            self.dados[chave_lista] = []
            self.salvar_dados()

    def adicionar_deposito(self):
        data, desc = (
            self.entry_dep_data.get().strip(),
            self.entry_dep_desc.get().strip(),
        )
        try:
            valor = float(self.entry_dep_valor.get().replace(",", "."))
        except ValueError:
            return messagebox.showwarning(
                "Aviso", "Digite um valor numérico válido."
            )
        if not desc or valor <= 0:
            return
        self.dados["depositos"].append(
            {"data": data, "desc": desc, "valor": valor}
        )
        self.entry_dep_desc.delete(0, "end")
        self.entry_dep_valor.delete(0, "end")
        self.salvar_dados()

    def remover_deposito(self):
        sel = self.tree_depositos.selection()
        if sel:
            self.dados["depositos"].pop(self.tree_depositos.index(sel))
            self.salvar_dados()

    def adicionar_gasto(self):
        desc = self.entry_desc.get().strip()
        try:
            valor = float(self.entry_valor.get().replace(",", "."))
        except ValueError:
            return messagebox.showwarning(
                "Aviso", "Digite um valor numérico válido."
            )
        if not desc or valor <= 0:
            return
        self.dados["gastos"].append(
            {"desc": desc, "valor": valor, "tipo": self.combo_tipo.get()}
        )
        self.entry_desc.delete(0, "end")
        self.entry_valor.delete(0, "end")
        self.salvar_dados()

    def remover_gasto(self):
        sel = self.tree_gastos.selection()
        if sel:
            self.dados["gastos"].pop(self.tree_gastos.index(sel))
            self.salvar_dados()

    def adicionar_roteiro(self):
        dia, ativ = (
            self.entry_dia.get().strip(),
            self.entry_ativ.get().strip(),
        )
        if dia and ativ:
            self.dados["roteiro"].append(
                {
                    "dia": dia,
                    "cat": self.combo_rot_cat.get(),
                    "atividade": ativ,
                }
            )
            self.entry_dia.delete(0, "end")
            self.entry_ativ.delete(0, "end")
            self.salvar_dados()

    def remover_roteiro(self):
        sel = self.tree_roteiro.selection()
        if sel:
            self.dados["roteiro"].pop(self.tree_roteiro.index(sel))
            self.salvar_dados()

    def adicionar_item_lista(self, chave, entry_widget, combo_widget):
        item = entry_widget.get().strip()
        if item:
            self.dados[chave].append(
                {"item": item, "cat": combo_widget.get(), "feito": False}
            )
            entry_widget.delete(0, "end")
            self.salvar_dados()

    def alternar_item_lista(self, chave, tree_widget):
        sel = tree_widget.selection()
        if sel:
            idx = tree_widget.index(sel)
            self.dados[chave][idx]["feito"] = not self.dados[chave][idx][
                "feito"
            ]
            self.salvar_dados()

    def remover_item_lista(self, chave, tree_widget):
        sel = tree_widget.selection()
        if sel:
            self.dados[chave].pop(tree_widget.index(sel))
            self.salvar_dados()

    def formatar_moeda(self, v):
        return (
            f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    def desenhar_grafico(self, pontos, meta_total):
        c = self.canvas_grafico
        c.delete("all")
        w, h, pL, pR, pT, pB = 400, 320, 58, 22, 28, 42
        pW, pH = w - pL - pR, h - pT - pB
        c.create_line(pL, pT, pL, h - pB, fill="#1E293B", width=2)
        c.create_line(pL, h - pB, w - pR, h - pB, fill="#1E293B", width=2)

        max_v = max(meta_total, max([p[1] for p in pontos]) if pontos else 1000)
        if max_v == 0:
            max_v = 1000

        y_meta = (h - pB) - (meta_total / max_v) * pH
        c.create_line(
            pL, y_meta, w - pR, y_meta, fill="#FBBF24", dash=(4, 4), width=1.5
        )
        c.create_text(
            w - pR - 5,
            y_meta - 10,
            text=f"Meta: {self.formatar_moeda(meta_total)}",
            fill="#FBBF24",
            anchor="e",
            font=("Segoe UI", 8, "bold"),
        )
        c.create_text(
            pL - 8,
            h - pB,
            text="R$ 0",
            fill="#64748B",
            anchor="e",
            font=("Segoe UI", 8),
        )

        if len(pontos) == 1:
            c.create_oval(
                pL - 4, h - pB - 4, pL + 4, h - pB + 4, fill="#34D399", outline=""
            )
            c.create_text(
                w / 2,
                h / 2,
                text="Começando do zero!\nRegistre o 1º depósito ao lado.",
                fill="#64748B",
                font=("Segoe UI", 10),
                justify="center",
            )
            return

        coords = [
            (
                pL + (i / (len(pontos) - 1)) * pW,
                (h - pB) - (val / max_v) * pH,
                lbl,
                val,
            )
            for i, (lbl, val) in enumerate(pontos)
        ]
        for i in range(len(coords) - 1):
            c.create_line(
                coords[i][0],
                coords[i][1],
                coords[i + 1][0],
                coords[i + 1][1],
                fill="#34D399",
                width=3,
            )
        for i, (x, y, lbl, val) in enumerate(coords):
            c.create_oval(
                x - 4,
                y - 4,
                x + 4,
                y + 4,
                fill="#34D399",
                outline=COR_FUNDO,
                width=1.5,
            )
            if i == len(coords) - 1:
                c.create_text(
                    x,
                    y - 12,
                    text=self.formatar_moeda(val),
                    fill="#34D399",
                    font=("Segoe UI", 9, "bold"),
                )
            if i == 0 or i == len(coords) - 1:
                c.create_text(
                    x,
                    h - pB + 14,
                    text=lbl,
                    fill="#94A3B8",
                    font=("Segoe UI", 8),
                )

    def atualizar_tudo(self):
        for tree in (
            self.tree_gastos,
            self.tree_depositos,
            self.tree_roteiro,
            self.tree_malas,
            self.tree_prep,
        ):
            for item in tree.get_children():
                tree.delete(item)

        tot_g, tot_ela = 0.0, 0.0
        for g in self.dados["gastos"]:
            v, t = g["valor"], g["tipo"]
            tot_g += v
            p_ele = v / 2 if t == "Meiado (50/50)" else (v if t == "100% Ele" else 0.0)
            p_ela = v / 2 if t == "Meiado (50/50)" else (v if t == "100% Ela" else 0.0)
            tot_ela += p_ela
            cor = (
                "roxo"
                if t == "Meiado (50/50)"
                else ("azul" if t == "100% Ele" else "verde")
            )
            self.tree_gastos.insert(
                "",
                "end",
                values=(
                    g["desc"],
                    t,
                    self.formatar_moeda(v),
                    self.formatar_moeda(p_ele),
                    self.formatar_moeda(p_ela),
                ),
                tags=(cor,),
            )

        acum = 0.0
        pts = [("Início", 0.0)]
        for d in self.dados["depositos"]:
            acum += d["valor"]
            pts.append((d["data"][:5], acum))
            self.tree_depositos.insert(
                "",
                "end",
                values=(
                    d["data"],
                    d["desc"],
                    self.formatar_moeda(d["valor"]),
                    self.formatar_moeda(acum),
                ),
            )

        for r in self.dados["roteiro"]:
            cat_r = r.get("cat", "BH Centro-Sul")
            cor_r = (
                "ambar"
                if cat_r == "Bate-Volta"
                else ("roxo" if cat_r == "Pampulha / Outros" else "azul")
            )
            self.tree_roteiro.insert(
                "",
                "end",
                values=(r["dia"], cat_r, r["atividade"]),
                tags=(cor_r,),
            )

        mapa_malas = {
            "Roupas": "azul",
            "Documentos": "roxo",
            "Higiene & Saúde": "rosa",
            "Eletrônicos": "ambar",
        }
        self.preencher_checklist(
            "malas",
            self.tree_malas,
            self.barra_malas,
            self.lbl_prog_malas,
            "Malas Prontas",
            mapa_malas,
        )

        mapa_prep = {
            "Transporte": "azul",
            "Hospedagem": "roxo",
            "Gastronomia": "ambar",
            "Passeios": "rosa",
        }
        self.preencher_checklist(
            "preparativos",
            self.tree_prep,
            self.barra_prep,
            self.lbl_prog_prep,
            "Preparativos Feitos",
            mapa_prep,
        )

        falta = max(0.0, tot_g - acum)
        self.lbl_total_viagem.configure(text=self.formatar_moeda(tot_g))
        self.lbl_total_guardado.configure(text=self.formatar_moeda(acum))
        self.lbl_falta_juntar.configure(
            text=self.formatar_moeda(falta),
            text_color="#34D399" if falta == 0 else "#38BDF8",
        )
        self.lbl_parte_dela.configure(text=self.formatar_moeda(tot_ela))
        self.desenhar_grafico(pts, tot_g)

    def preencher_checklist(
        self, chave, tree, barra, lbl_prog, titulo_prog, mapa_cores
    ):
        feitos = 0
        lista = self.dados.get(chave, [])
        for item in lista:
            if item["feito"]:
                feitos += 1
            st = "✔  Concluído" if item["feito"] else "○  Pendente"
            cat = item.get("cat", "")
            cor = "verde" if item["feito"] else mapa_cores.get(cat, "azul")
            tree.insert("", "end", values=(st, cat, item["item"]), tags=(cor,))
        pct = (feitos / len(lista)) if lista else 0
        barra.set(pct)
        lbl_prog.configure(
            text=f"{titulo_prog}: {int(pct * 100)}% ({feitos}/{len(lista)})"
        )

    def exportar_excel(self):
        wb = Workbook()
        h_fill = PatternFill(
            start_color="0F172A", end_color="0F172A", fill_type="solid"
        )
        h_font = Font(color="FFFFFF", bold=True)
        m_fmt = "R$ #,##0.00"

        ws1 = wb.active
        ws1.title = "Orçamento Viagem BH"
        ws1.append(
            ["Item", "Divisão", "Valor Total", "Parte Dele", "Parte Dela"]
        )
        for c in range(1, 6):
            ws1.cell(1, c).fill = h_fill
            ws1.cell(1, c).font = h_font
            ws1.cell(1, c).alignment = Alignment(horizontal="center")

        for i, g in enumerate(self.dados["gastos"], start=2):
            ws1.cell(i, 1, g["desc"])
            ws1.cell(i, 2, g["tipo"])
            ws1.cell(i, 3, g["valor"]).number_format = m_fmt
            ws1.cell(
                i,
                4,
                f'=IF(B{i}="Meiado (50/50)",C{i}/2,IF(B{i}="100% Ele",C{i},0))',
            ).number_format = m_fmt
            ws1.cell(
                i,
                5,
                f'=IF(B{i}="Meiado (50/50)",C{i}/2,IF(B{i}="100% Ela",C{i},0))',
            ).number_format = m_fmt

        ult = len(self.dados["gastos"]) + 2
        ws1.cell(ult, 1, "TOTAL GERAL").font = Font(bold=True)
        for col, let in [(3, "C"), (4, "D"), (5, "E")]:
            cell = ws1.cell(ult, col, f"=SUM({let}2:{let}{ult-1})")
            cell.font = Font(bold=True)
            cell.number_format = m_fmt

        ws2 = wb.create_sheet("Cofrinho da Viagem")
        ws2.append(["Data", "Descrição", "Valor Depositado"])
        for c in range(1, 4):
            ws2.cell(1, c).fill = h_fill
            ws2.cell(1, c).font = h_font
        for i, d in enumerate(self.dados["depositos"], start=2):
            ws2.cell(i, 1, d["data"])
            ws2.cell(i, 2, d["desc"])
            ws2.cell(i, 3, d["valor"]).number_format = m_fmt
        l_dep = len(self.dados["depositos"]) + 2
        ws2.cell(l_dep, 2, "TOTAL GUARDADO").font = Font(bold=True)
        c_dep = ws2.cell(l_dep, 3, f"=SUM(C2:C{l_dep-1})")
        c_dep.font = Font(bold=True)
        c_dep.number_format = m_fmt

        ws3 = wb.create_sheet("Roteiro 08 a 18 Maio")
        ws3.append(["Data / Dia", "Região / Tipo", "Programação"])
        for c in range(1, 4):
            ws3.cell(1, c).fill = h_fill
            ws3.cell(1, c).font = h_font
        for r in self.dados["roteiro"]:
            ws3.append([r["dia"], r.get("cat", "BH Centro-Sul"), r["atividade"]])

        for aba_nome, chave_lista in [
            ("Checklist Malas e Docs", "malas"),
            ("Checklist Preparativos", "preparativos"),
        ]:
            ws = wb.create_sheet(aba_nome)
            ws.append(["Status", "Categoria", "Item"])
            for c in range(1, 4):
                ws.cell(1, c).fill = h_fill
                ws.cell(1, c).font = h_font
            for item in self.dados.get(chave_lista, []):
                ws.append([
                    "Concluído" if item["feito"] else "Pendente",
                    item["cat"],
                    item["item"],
                ])

        for ws in wb.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col.column_letter].width = (
                    max(len(str(c.value or "")) for c in col) + 6
                )

        wb.save(ARQUIVO_EXCEL)
        messagebox.showinfo(
            "Sucesso!", f"Planilha '{ARQUIVO_EXCEL}' gerada com 5 abas!"
        )


if __name__ == "__main__":
    ViagemBHApp().mainloop()