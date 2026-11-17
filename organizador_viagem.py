from datetime import datetime
import json, os, tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import requests

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# COLE AQUI A URL DO SEU FIREBASE (sem barra no final)
FIREBASE_URL = "https://viagem-bh-b2b87-default-rtdb.firebaseio.com/"
ENDPOINT_NUVEM = f"{FIREBASE_URL}/planejamento.json"
ARQUIVO_EXCEL = "Planejamento_Viagem_BH.xlsx"


class ViagemBHApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Nossa Viagem BH • Sincronizado na Nuvem ☁️")
        self.geometry("980x680")
        self.resizable(False, False)
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
                    "atividade": "Chegada + Jantar no Glouton",
                },
                {
                    "dia": "Dia 2 • 09/05 (Dom)",
                    "atividade": "Circuito Liberdade + Almoço no Ninita",
                },
                {
                    "dia": "Dia 3 • 10/05 (Seg)",
                    "atividade": "Mercado Central + Bistrô Taste-Vin",
                },
                {
                    "dia": "Dia 4 • 11/05 (Ter)",
                    "atividade": "Bate-volta romântico a Ouro Preto",
                },
                {
                    "dia": "Dia 5 • 12/05 (Qua)",
                    "atividade": "Instituto Inhotim (Entrada gratuita)",
                },
                {
                    "dia": "Dia 6 • 13/05 (Qui)",
                    "atividade": "Lagoa da Pampulha + Jantar no O Conde",
                },
                {
                    "dia": "Dia 7 • 14/05 (Sex)",
                    "atividade": "Noite cultural e drinks na Savassi",
                },
                {
                    "dia": "Dia 8 • 15/05 (Sáb)",
                    "atividade": "Passeio e almoço na Serra em Macacos",
                },
                {
                    "dia": "Dia 9 • 16/05 (Dom)",
                    "atividade": "Menu Degustação no Pacato + Mirante",
                },
                {
                    "dia": "Dia 10 • 17/05 (Seg)",
                    "atividade": "Jantar de despedida no Trintaeum",
                },
            ],
            "depositos": [],
                        "depositos": [],
            
            "checklist": [
                {"item": "Comprar passagem aérea BSB - CNF", "cat": "Reservas", "feito": False},
                {"item": "Reservar Airbnb no Lourdes", "cat": "Reservas", "feito": False},
                {"item": "Alugar carro compacto automático", "cat": "Reservas", "feito": False},
                {"item": "Reservar mesa Glouton, Pacato e Trintaeum", "cat": "Reservas", "feito": False},
                {"item": "Emitir ingressos gratuitos Inhotim (12/05)", "cat": "Reservas", "feito": False},
                {"item": "Separar CNH e cartão para caução do carro", "cat": "Malas", "feito": False},
                {"item": "Roupas esporte fino + jaqueta leve", "cat": "Malas", "feito": False},
                {"item": "Tênis confortável + carregador portátil", "cat": "Malas", "feito": False},
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
                else:
                    self.salvar_dados()
                self.lbl_status.configure(
                    text="🟢 Sincronizado", text_color="#10b981"
                )
        except Exception:
            self.lbl_status.configure(
                text="🔴 Sem internet", text_color="#ef4444"
            )
        if nuvem:
            self.dados["gastos"] = nuvem.get("gastos") or []
            self.dados["roteiro"] = nuvem.get("roteiro") or []
            self.dados["depositos"] = nuvem.get("depositos") or []
            self.dados["checklist"] = nuvem.get("checklist") or []
            
        self.atualizar_tudo()

    def salvar_dados(self):
        try:
            requests.put(ENDPOINT_NUVEM, json=self.dados, timeout=5)
            self.lbl_status.configure(
                text="🟢 Salvo na nuvem", text_color="#10b981"
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
        header = ctk.CTkFrame(self, height=70, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="✈️ Nossa Viagem • Belo Horizonte (08 a 18 de Maio)",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(side="left", padx=25, pady=18)
        self.lbl_status = ctk.CTkLabel(
            header,
            text="🟡 Conectando...",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.lbl_status.pack(side="left", padx=10)
        ctk.CTkButton(
            header,
            text="📊 Exportar Excel (.xlsx)",
            fg_color="#10b981",
            hover_color="#059669",
            font=ctk.CTkFont(weight="bold"),
            command=self.exportar_excel,
        ).pack(side="right", padx=25, pady=18)

        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.pack(fill="x", padx=20, pady=15)
        self.lbl_total_viagem = self.criar_card(
            cards, "Custo Total Planejado", "R$ 0,00", "#f59e0b"
        )
        self.lbl_total_guardado = self.criar_card(
            cards, "Já Guardado na Conta", "R$ 0,00", "#10b981"
        )
        self.lbl_falta_juntar = self.criar_card(
            cards, "Falta para a Meta", "R$ 0,00", "#3b82f6"
        )
        self.lbl_parte_dela = self.criar_card(
            cards, "Parte Dela (Meiado)", "R$ 0,00", "#a855f7"
        )

        self.tabview = ctk.CTkTabview(self, width=940, height=460)
        self.tabview.pack(padx=20, pady=5)
        self.tab_cofrinho = self.tabview.add("🏦 Cofrinho & Gráfico")
        self.tab_gastos = self.tabview.add("💰 Orçamento da Viagem")
        self.tab_roteiro = self.tabview.add("🗺️ Roteiro Dia a Dia")

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#18181b",
            foreground="white",
            fieldbackground="#18181b",
            rowheight=28,
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background="#27272a",
            foreground="white",
            font=("Segoe UI", 10, "bold"),
        )

        self.montar_aba_cofrinho()
        self.montar_aba_gastos()
        self.montar_aba_roteiro()
        self.tab_cofrinho = self.tabview.add("🏦 Cofrinho & Gráfico")
        self.tab_gastos = self.tabview.add("💰 Orçamento da Viagem")
        self.tab_roteiro = self.tabview.add("🗺️ Roteiro Dia a Dia")
        self.tab_check = self.tabview.add("✅ Checklist")

    def criar_card(self, parent, titulo, valor, cor):
        card = ctk.CTkFrame(parent, width=220, height=85, corner_radius=12)
        card.pack(side="left", expand=True, fill="both", padx=6)
        ctk.CTkLabel(
            card, text=titulo, font=ctk.CTkFont(size=12), text_color="#a1a1aa"
        ).pack(pady=(12, 2))
        lbl = ctk.CTkLabel(
            card,
            text=valor,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=cor,
        )
        lbl.pack(pady=(0, 12))
        return lbl

    def montar_aba_cofrinho(self):
        form = ctk.CTkFrame(self.tab_cofrinho, fg_color="transparent")
        form.pack(fill="x", pady=8)
        self.entry_dep_data = ctk.CTkEntry(form, width=140)
        self.entry_dep_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_dep_data.pack(side="left", padx=5)
        self.entry_dep_desc = ctk.CTkEntry(
            form, placeholder_text="Descrição (ex: Depósito Novembro)", width=350
        )
        self.entry_dep_desc.pack(side="left", padx=5)
        self.entry_dep_valor = ctk.CTkEntry(
            form, placeholder_text="Valor (R$)", width=130
        )
        self.entry_dep_valor.pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="+ Guardar Dinheiro",
            fg_color="#10b981",
            hover_color="#059669",
            width=140,
            command=self.adicionar_deposito,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.remover_deposito,
        ).pack(side="left", padx=5)

        content = ctk.CTkFrame(self.tab_cofrinho, fg_color="transparent")
        content.pack(fill="both", expand=True, pady=5)
        left = ctk.CTkFrame(content, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self.tree_depositos = ttk.Treeview(
            left,
            columns=("data", "desc", "valor", "acumulado"),
            show="headings",
            height=12,
        )
        for col, txt, w, anc in [
            ("data", "Data", 90, "center"),
            ("desc", "Descrição", 180, "w"),
            ("valor", "Depósito", 100, "e"),
            ("acumulado", "Saldo Total", 110, "e"),
        ]:
            self.tree_depositos.heading(col, text=txt)
            self.tree_depositos.column(col, width=w, anchor=anc)
        self.tree_depositos.pack(fill="both", expand=True)

        right = ctk.CTkFrame(content, width=410, corner_radius=12)
        right.pack(side="right", fill="both")
        right.pack_propagate(False)
        ctk.CTkLabel(
            right,
            text="📈 Evolução do Cofrinho rumo à Meta",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(10, 5))
        self.canvas_grafico = tk.Canvas(
            right, width=390, height=300, bg="#18181b", highlightthickness=0
        )
        self.canvas_grafico.pack(padx=10, pady=5)

    def montar_aba_gastos(self):
        form = ctk.CTkFrame(self.tab_gastos, fg_color="transparent")
        form.pack(fill="x", pady=10)
        self.entry_desc = ctk.CTkEntry(
            form, placeholder_text="Descrição do gasto...", width=350
        )
        self.entry_desc.pack(side="left", padx=5)
        self.entry_valor = ctk.CTkEntry(
            form, placeholder_text="Valor (R$)", width=130
        )
        self.entry_valor.pack(side="left", padx=5)
        self.combo_tipo = ctk.CTkComboBox(
            form, values=["Meiado (50/50)", "100% Ele", "100% Ela"], width=160
        )
        self.combo_tipo.pack(side="left", padx=5)
        ctk.CTkButton(
            form, text="+ Adicionar", width=110, command=self.adicionar_gasto
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.remover_gasto,
        ).pack(side="left", padx=5)

        self.tree_gastos = ttk.Treeview(
            self.tab_gastos,
            columns=("desc", "tipo", "total", "ele", "ela"),
            show="headings",
            height=12,
        )
        for col, txt, w, anc in [
            ("desc", "Item / Experiência", 340, "w"),
            ("tipo", "Divisão", 140, "center"),
            ("total", "Valor Total", 130, "e"),
            ("ele", "Parte Dele", 130, "e"),
            ("ela", "Parte Dela", 130, "e"),
        ]:
            self.tree_gastos.heading(col, text=txt)
            self.tree_gastos.column(col, width=w, anchor=anc)
        self.tree_gastos.pack(fill="both", expand=True, pady=5)

    def montar_aba_roteiro(self):
        form = ctk.CTkFrame(self.tab_roteiro, fg_color="transparent")
        form.pack(fill="x", pady=10)
        self.entry_dia = ctk.CTkEntry(
            form, placeholder_text="Dia (ex: Dia 11 • 18/05)", width=210
        )
        self.entry_dia.pack(side="left", padx=5)
        self.entry_ativ = ctk.CTkEntry(
            form, placeholder_text="Atividade ou Restaurante...", width=470
        )
        self.entry_ativ.pack(side="left", padx=5)
        ctk.CTkButton(
            form, text="+ Adicionar", width=110, command=self.adicionar_roteiro
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.remover_roteiro,
        ).pack(side="left", padx=5)

        self.tree_roteiro = ttk.Treeview(
            self.tab_roteiro,
            columns=("dia", "atividade"),
            show="headings",
            height=12,
        )
        self.tree_roteiro.heading("dia", text="Data / Dia")
        self.tree_roteiro.heading("atividade", text="Programação Planejada")
        self.tree_roteiro.column("dia", width=220)
        self.tree_roteiro.column("atividade", width=670)
        self.tree_roteiro.pack(fill="both", expand=True, pady=5)

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
            self.dados["depositos"].pop(self.tree_depositos.index(sel[0]))
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
            self.dados["gastos"].pop(self.tree_gastos.index(sel[0]))
            self.salvar_dados()

    def adicionar_roteiro(self):
        dia, ativ = (
            self.entry_dia.get().strip(),
            self.entry_ativ.get().strip(),
        )
        if dia and ativ:
            self.dados["roteiro"].append({"dia": dia, "atividade": ativ})
            self.entry_dia.delete(0, "end")
            self.entry_ativ.delete(0, "end")
            self.salvar_dados()

    def remover_roteiro(self):
        sel = self.tree_roteiro.selection()
        if sel:
            self.dados["roteiro"].pop(self.tree_roteiro.index(sel[0]))
            self.salvar_dados()

    def montar_aba_checklist(self):
        form = ctk.CTkFrame(self.tab_check, fg_color="transparent")
        form.pack(fill="x", pady=8)

        self.entry_chk_item = ctk.CTkEntry(
            form, placeholder_text="Novo item (ex: Levar protetor solar)...", width=420
        )
        self.entry_chk_item.pack(side="left", padx=5)

        self.combo_chk_cat = ctk.CTkComboBox(
            form, values=["Reservas", "Malas", "Documentos"], width=150
        )
        self.combo_chk_cat.pack(side="left", padx=5)

        ctk.CTkButton(
            form, text="+ Adicionar", width=110, command=self.adicionar_checklist
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            form,
            text="Excluir",
            width=80,
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.remover_checklist,
        ).pack(side="left", padx=5)

        prog_frame = ctk.CTkFrame(self.tab_check, fg_color="transparent")
        prog_frame.pack(fill="x", padx=5, pady=(0, 8))

        self.lbl_progresso = ctk.CTkLabel(
            prog_frame, text="Progresso: 0% (0/0)", font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_progresso.pack(side="left", padx=(0, 10))

        self.barra_progresso = ctk.CTkProgressBar(prog_frame, width=700)
        self.barra_progresso.set(0)
        self.barra_progresso.pack(side="left", fill="x", expand=True)

        self.tree_check = ttk.Treeview(
            self.tab_check, columns=("status", "cat", "item"), show="headings", height=11
        )
        self.tree_check.heading("status", text="Status (Duplo clique)")
        self.tree_check.heading("cat", text="Categoria")
        self.tree_check.heading("item", text="Tarefa / Item da Mala")
        self.tree_check.column("status", width=150, anchor="center")
        self.tree_check.column("cat", width=130, anchor="center")
        self.tree_check.column("item", width=610)
        self.tree_check.pack(fill="both", expand=True)
        self.tree_check.bind("<Double-1>", self.alternar_checklist)

    def adicionar_checklist(self):
        item = self.entry_chk_item.get().strip()
        if item:
            self.dados["checklist"].append(
                {"item": item, "cat": self.combo_chk_cat.get(), "feito": False}
            )
            self.entry_chk_item.delete(0, "end")
            self.salvar_dados()

    def alternar_checklist(self, event):
        sel = self.tree_check.selection()
        if sel:
            idx = self.tree_check.index(sel[0])
            self.dados["checklist"][idx]["feito"] = not self.dados["checklist"][idx]["feito"]
            self.salvar_dados()

    def remover_checklist(self):
        sel = self.tree_check.selection()
        if sel:
            self.dados["checklist"].pop(self.tree_check.index(sel[0]))
            self.salvar_dados()
        

    def formatar_moeda(self, v):
        return (
            f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    def desenhar_grafico(self, pontos, meta_total):
        c = self.canvas_grafico
        c.delete("all")
        w, h, pL, pR, pT, pB = 390, 300, 55, 20, 25, 40
        pW, pH = w - pL - pR, h - pT - pB
        c.create_line(pL, pT, pL, h - pB, fill="#3f3f46", width=2)
        c.create_line(pL, h - pB, w - pR, h - pB, fill="#3f3f46", width=2)
        max_v = max(meta_total, max([p[1] for p in pontos]) if pontos else 1000)
        if max_v == 0:
            max_v = 1000
        y_meta = (h - pB) - (meta_total / max_v) * pH
        c.create_line(
            pL, y_meta, w - pR, y_meta, fill="#f59e0b", dash=(4, 4), width=1.5
        )
        c.create_text(
            w - pR - 5,
            y_meta - 10,
            text=f"Meta: {self.formatar_moeda(meta_total)}",
            fill="#f59e0b",
            anchor="e",
            font=("Segoe UI", 8, "bold"),
        )
        c.create_text(
            pL - 8,
            h - pB,
            text="R$ 0",
            fill="#a1a1aa",
            anchor="e",
            font=("Segoe UI", 8),
        )

        if len(pontos) == 1:
            c.create_oval(
                pL - 4, h - pB - 4, pL + 4, h - pB + 4, fill="#10b981", outline=""
            )
            c.create_text(
                w / 2,
                h / 2,
                text="Começando do zero!\nRegistre o 1º depósito ao lado.",
                fill="#71717a",
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
                fill="#10b981",
                width=3,
            )
        for i, (x, y, lbl, val) in enumerate(coords):
            c.create_oval(x - 4, y - 4, x + 4, y + 4, fill="#10b981", outline="")
            if i == len(coords) - 1:
                c.create_text(
                    x,
                    y - 12,
                    text=self.formatar_moeda(val),
                    fill="#10b981",
                    font=("Segoe UI", 9, "bold"),
                )
            if i == 0 or i == len(coords) - 1:
                c.create_text(
                    x,
                    h - pB + 14,
                    text=lbl,
                    fill="#a1a1aa",
                    font=("Segoe UI", 8),
                )

    def atualizar_tudo(self):
        for tree in (
            self.tree_gastos,
            self.tree_depositos,
            self.tree_roteiro,
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
            self.tree_roteiro.insert(
                "", "end", values=(r["dia"], r["atividade"])
            )

        falta = max(0.0, tot_g - acum)
        self.lbl_total_viagem.configure(text=self.formatar_moeda(tot_g))
        self.lbl_total_guardado.configure(text=self.formatar_moeda(acum))
        self.lbl_falta_juntar.configure(
            text=self.formatar_moeda(falta),
            text_color="#10b981" if falta == 0 else "#3b82f6",
        )
        self.lbl_parte_dela.configure(text=self.formatar_moeda(tot_ela))
        self.desenhar_grafico(pts, tot_g)

        for item in self.tree_check.get_children():
            self.tree_check.delete(item)

        feitos = 0
        total_chk = len(self.dados.get("checklist", []))
        for c in self.dados.get("checklist", []):
            if c["feito"]:
                feitos += 1
            st = "✅ Concluído" if c["feito"] else "⬜ Pendente"
            self.tree_check.insert("", "end", values=(st, c["cat"], c["item"]))

        pct = (feitos / total_chk) if total_chk > 0 else 0
        self.barra_progresso.set(pct)
        self.lbl_progresso.configure(
            text=f"Progresso: {int(pct * 100)}% ({feitos}/{total_chk})"
        )


    def exportar_excel(self):
        wb = Workbook()
        h_fill = PatternFill(
            start_color="1E293B", end_color="1E293B", fill_type="solid"
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
        ws3.append(["Data / Dia", "Programação"])
        for c in range(1, 3):
            ws3.cell(1, c).fill = h_fill
            ws3.cell(1, c).font = h_font
        for r in self.dados["roteiro"]:
            ws3.append([r["dia"], r["atividade"]])

        for ws in (ws1, ws2, ws3):
            for col in ws.columns:
                ws.column_dimensions[col.column_letter].width = (
                    max(len(str(c.value or "")) for c in col) + 6
                )

        wb.save(ARQUIVO_EXCEL)
        messagebox.showinfo(
            "Sucesso!", f"Planilha '{ARQUIVO_EXCEL}' gerada com 3 abas!"
        )


if __name__ == "__main__":
    ViagemBHApp().mainloop()
