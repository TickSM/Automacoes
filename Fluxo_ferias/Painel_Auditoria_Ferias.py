import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
from datetime import datetime

# Garante a importação dos módulos da pasta local
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
if DIRETORIO_ATUAL not in sys.path:
    sys.path.insert(0, DIRETORIO_ATUAL)

from banco_ferias import (
    obter_meses_cadastrados,
    consultar_historico_consolidado,
    exportar_relatorio_excel
)

class PainelAuditoriaFerias(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Painel de Auditoria e Histórico de Férias - SEI")
        self.geometry("980x680")
        self.minsize(850, 550)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Cabeçalho Principal
        self.lbl_titulo = ctk.CTkLabel(
            self, 
            text="📊 Auditoria e Histórico de Férias (SEI)", 
            font=("Arial", 22, "bold")
        )
        self.lbl_titulo.pack(pady=(15, 2))

        self.lbl_subtitulo = ctk.CTkLabel(
            self, 
            text="Consolidação e conciliação automática entre Avisos emitidos e Recibos vinculados", 
            font=("Arial", 12), 
            text_color="gray70"
        )
        self.lbl_subtitulo.pack(pady=(0, 12))

        # ───────────── CARDS DE MÉTRICAS ─────────────
        self.frame_cards = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_cards.pack(fill="x", padx=20, pady=5)

        self.card_avisos = self._criar_card(self.frame_cards, "📑 Avisos Emitidos", "0", "#1f618d")
        self.card_recibos = self._criar_card(self.frame_cards, "🧾 Recibos Gerados", "0", "#2e86c1")
        self.card_vinculados = self._criar_card(self.frame_cards, "✅ Vinculados (Ok)", "0", "#229954")
        self.card_pendentes = self._criar_card(self.frame_cards, "⚠️ Recibos Pendentes", "0", "#d35400")

        # ───────────── FILTROS DE CONSULTA ─────────────
        self.frame_filtros = ctk.CTkFrame(self, fg_color="#2b2b2b", corner_radius=10)
        self.frame_filtros.pack(fill="x", padx=20, pady=(10, 8))

        # Filtro Mês/Ano
        ctk.CTkLabel(self.frame_filtros, text="Mês/Ano:", font=("Arial", 12, "bold")).pack(side="left", padx=(15, 5), pady=10)
        self.meses_disponiveis = ["TODOS"] + (obter_meses_cadastrados() or [])
        self.opt_mes = ctk.CTkOptionMenu(
            self.frame_filtros, 
            values=self.meses_disponiveis, 
            width=140,
            command=lambda _: self.atualizar_tabela()
        )
        self.opt_mes.set("TODOS")
        self.opt_mes.pack(side="left", padx=5, pady=10)

        # Filtro Nome
        ctk.CTkLabel(self.frame_filtros, text="Funcionário:", font=("Arial", 12, "bold")).pack(side="left", padx=(15, 5), pady=10)
        self.ent_busca = ctk.CTkEntry(self.frame_filtros, width=200, placeholder_text="Digite o nome...")
        self.ent_busca.pack(side="left", padx=5, pady=10)
        self.ent_busca.bind("<Return>", lambda _: self.atualizar_tabela())

        # Filtro Status
        ctk.CTkLabel(self.frame_filtros, text="Status:", font=("Arial", 12, "bold")).pack(side="left", padx=(15, 5), pady=10)
        self.opt_status = ctk.CTkOptionMenu(
            self.frame_filtros, 
            values=["TODOS", "CONCLUÍDO", "PENDENTE", "AVULSO"], 
            width=130,
            command=lambda _: self.atualizar_tabela()
        )
        self.opt_status.set("TODOS")
        self.opt_status.pack(side="left", padx=5, pady=10)

        # Botão Filtrar
        self.btn_filtrar = ctk.CTkButton(
            self.frame_filtros, 
            text="🔍 Filtrar", 
            width=80, 
            height=28, 
            command=self.atualizar_tabela
        )
        self.btn_filtrar.pack(side="left", padx=(10, 5), pady=10)

        # Botão Limpar Filtros
        self.btn_limpar = ctk.CTkButton(
            self.frame_filtros, 
            text="Limpar", 
            width=70, 
            height=28, 
            fg_color="#555555", 
            hover_color="#444444", 
            command=self.limpar_filtros
        )
        self.btn_limpar.pack(side="left", padx=5, pady=10)

        # ───────────── TABELA DE DADOS (TREEVIEW) ─────────────
        self.frame_tabela = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_tabela.pack(fill="both", expand=True, padx=20, pady=5)

        # Configura estilo do Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Auditoria.Treeview",
            background="#23272d",
            foreground="#f0f0f0",
            rowheight=26,
            fieldbackground="#23272d",
            font=("Arial", 10)
        )
        style.configure(
            "Auditoria.Treeview.Heading",
            background="#1a1c20",
            foreground="#ffffff",
            font=("Arial", 10, "bold")
        )
        style.map("Auditoria.Treeview", background=[('selected', '#1f618d')])

        colunas = ("funcionario", "unidade", "mes_ano", "processo_aviso", "processo_recibo", "status")
        self.tree = ttk.Treeview(
            self.frame_tabela, 
            columns=colunas, 
            show="headings", 
            style="Auditoria.Treeview"
        )

        self.tree.heading("funcionario", text="Funcionário")
        self.tree.heading("unidade", text="UF")
        self.tree.heading("mes_ano", text="Mês/Ano")
        self.tree.heading("processo_aviso", text="Processo do Aviso (SEI)")
        self.tree.heading("processo_recibo", text="Processo do Recibo (SEI)")
        self.tree.heading("status", text="Status")

        self.tree.column("funcionario", width=250, anchor="w")
        self.tree.column("unidade", width=60, anchor="center")
        self.tree.column("mes_ano", width=120, anchor="center")
        self.tree.column("processo_aviso", width=180, anchor="center")
        self.tree.column("processo_recibo", width=180, anchor="center")
        self.tree.column("status", width=170, anchor="center")

        # Scrollbars
        scrollbar_y = ttk.Scrollbar(self.frame_tabela, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(self.frame_tabela, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        # Cores para linhas conforme o status
        self.tree.tag_configure("concluido", foreground="#58d68d")
        self.tree.tag_configure("pendente", foreground="#f39c12")
        self.tree.tag_configure("avulso", foreground="#e74c3c")

        # ───────────── BARRA INFERIOR / AÇÕES ─────────────
        self.frame_rodape = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_rodape.pack(fill="x", padx=20, pady=(10, 15))

        self.lbl_total_linhas = ctk.CTkLabel(
            self.frame_rodape, 
            text="Exibindo 0 registros", 
            font=("Arial", 11), 
            text_color="gray70"
        )
        self.lbl_total_linhas.pack(side="left")

        self.btn_exportar = ctk.CTkButton(
            self.frame_rodape, 
            text="📊 Exportar Relatório Excel (.xlsx)", 
            font=("Arial", 12, "bold"),
            fg_color="#1e8449",
            hover_color="#196f3d",
            height=34,
            command=self.exportar_excel
        )
        self.btn_exportar.pack(side="right", padx=(10, 0))

        self.btn_atualizar = ctk.CTkButton(
            self.frame_rodape, 
            text="🔄 Atualizar", 
            width=90, 
            height=34, 
            command=self.atualizar_tudo
        )
        self.btn_atualizar.pack(side="right")

        # Carga inicial dos dados
        self.atualizar_tudo()

    def _criar_card(self, parent, titulo, valor_inicial, cor_destaque):
        card = ctk.CTkFrame(parent, fg_color="#2b2b2b", corner_radius=8, width=220, height=75)
        card.pack(side="left", fill="both", expand=True, padx=5)

        lbl_tit = ctk.CTkLabel(card, text=titulo, font=("Arial", 11, "bold"), text_color="gray70")
        lbl_tit.pack(pady=(8, 2))

        lbl_val = ctk.CTkLabel(card, text=valor_inicial, font=("Arial", 20, "bold"), text_color=cor_destaque)
        lbl_val.pack(pady=(0, 8))
        return lbl_val

    def limpar_filtros(self):
        self.opt_mes.set("TODOS")
        self.ent_busca.delete(0, "end")
        self.opt_status.set("TODOS")
        self.atualizar_tabela()

    def atualizar_tudo(self):
        # Atualiza a lista de meses disponíveis no dropdown
        meses = ["TODOS"] + (obter_meses_cadastrados() or [])
        self.opt_mes.configure(values=meses)
        self.atualizar_tabela()

    def atualizar_tabela(self):
        mes_selecionado = self.opt_mes.get()
        busca = self.ent_busca.get().strip()
        status_sel = self.opt_status.get()

        registros = consultar_historico_consolidado(mes_selecionado, busca, status_sel)

        # Limpa tabela
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Popula tabela
        total_avisos = set()
        total_recibos = 0
        total_vinculados = 0
        total_pendentes = 0

        for reg in registros:
            status = reg["status"]
            if "CONCLUÍDO" in status:
                tag = "concluido"
                total_vinculados += 1
            elif "PENDENTE" in status:
                tag = "pendente"
                total_pendentes += 1
            else:
                tag = "avulso"

            if reg["processo_aviso"] and reg["processo_aviso"] != "NÃO LOCALIZADO":
                total_avisos.add(reg["funcionario"])
            if reg["processo_recibo"] and reg["processo_recibo"] != "PENDENTE":
                total_recibos += 1

            self.tree.insert("", "end", values=(
                reg["funcionario"],
                reg["unidade"],
                reg["mes_ano"],
                reg["processo_aviso"],
                reg["processo_recibo"],
                reg["status"]
            ), tags=(tag,))

        # Atualiza métricas nos cards
        self.card_avisos.configure(text=str(len(total_avisos)))
        self.card_recibos.configure(text=str(total_recibos))
        self.card_vinculados.configure(text=str(total_vinculados))
        self.card_pendentes.configure(text=str(total_pendentes))

        self.lbl_total_linhas.configure(text=f"Exibindo {len(registros)} registro(s)")

    def exportar_excel(self):
        mes_selecionado = self.opt_mes.get()
        busca = self.ent_busca.get().strip()
        status_sel = self.opt_status.get()

        nome_padrao = f"Auditoria_Ferias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho_salvar = filedialog.asksaveasfilename(
            title="Salvar Relatório de Auditoria",
            initialfile=nome_padrao,
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")]
        )

        if not caminho_salvar:
            return

        sucesso, msg = exportar_relatorio_excel(caminho_salvar, mes_selecionado, busca, status_sel)
        if sucesso:
            messagebox.showinfo("Sucesso", msg)
        else:
            messagebox.showwarning("Atenção", msg)

if __name__ == "__main__":
    app = PainelAuditoriaFerias()
    app.mainloop()
