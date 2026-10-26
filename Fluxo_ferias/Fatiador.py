import customtkinter as ctk
import tkinter.filedialog as filedialog
import threading
import os
import re
import sys
from pypdf import PdfReader, PdfWriter
from extrator_inteligente import processar_pagina_pdf, verificar_ia_local

# Configuração de diretório para o PyInstaller
if getattr(sys, 'frozen', False):
    DIRETORIO_BASE = os.path.dirname(sys.executable)
else:
    DIRETORIO_BASE = os.path.dirname(os.path.abspath(__file__))

# Pasta onde os PDFs individuais serão salvos
CAMINHO_PASTA_SAIDA = os.path.join(DIRETORIO_BASE, "PDFs_Separados")

def executar_separacao(app_gui, caminho_pdf, unidade_padrao):
    try:
        if not os.path.exists(CAMINHO_PASTA_SAIDA):
            os.makedirs(CAMINHO_PASTA_SAIDA)

        app_gui.atualizar_status(f"Abrindo arquivo:\n{os.path.basename(caminho_pdf)}\n")
        
        # Checa status da IA Local
        if verificar_ia_local():
            app_gui.atualizar_status("🤖 IA Local (Ollama) conectada! Modo Inteligente ativo.\n")
        else:
            app_gui.atualizar_status("ℹ️ IA Local não detectada. Utilizando Extrator Heurístico Local (Regras).\n")
            
        reader = PdfReader(caminho_pdf)
        total_paginas = len(reader.pages)
        app_gui.atualizar_status(f"Total de páginas encontradas: {total_paginas}\n")

        for i in range(total_paginas):
            pagina = reader.pages[i]
            texto_cru = pagina.extract_text() or ""
            
            # Extração inteligente (tenta IA Local, senão usa regras resilientes)
            dados_extraidos, motor = processar_pagina_pdf(texto_cru, unidade_padrao)
            nome_cru = dados_extraidos.get("nome", "")
            unidade_cru = dados_extraidos.get("unidade", "")

            # =======================================================
            # TRATAMENTO E SALVAMENTO
            # =======================================================
            if nome_cru:
                nome_formatado = re.sub(r'\s+', '_', nome_cru.strip())
            else:
                nome_formatado = f"FUNCIONARIO_PAGINA_{i+1}"
            
            if unidade_cru:
                unidade_formatada = re.sub(r'\s+', '_', unidade_cru.strip().upper())
            else:
                unidade_formatada = unidade_padrao.strip().upper() if unidade_padrao else "UNIDADE_NAO_ENCONTRADA"
            
            # FILTRO ANTI-ERRO DO WINDOWS
            nome_formatado = re.sub(r'[\\/*?:"<>|]', '', nome_formatado)
            unidade_formatada = re.sub(r'[\\/*?:"<>|]', '', unidade_formatada)
            
            nome_arquivo_saida = f"{unidade_formatada}___{nome_formatado}.pdf"
            caminho_saida = os.path.join(CAMINHO_PASTA_SAIDA, nome_arquivo_saida)
            
            writer = PdfWriter()
            writer.add_page(pagina)
            
            with open(caminho_saida, "wb") as f_out:
                writer.write(f_out)
                
            if "FUNCIONARIO_PAGINA" in nome_formatado:
                app_gui.atualizar_status(f"⚠️ Salvo como {nome_arquivo_saida}")
            else:
                app_gui.atualizar_status(f"✔ [{motor}] Salvo: {nome_arquivo_saida}")
        
        app_gui.atualizar_status("\n===============================")
        app_gui.atualizar_status("Separação concluída com sucesso!")
        
    except Exception as e:
        app_gui.atualizar_status(f"\n[ERRO CRÍTICO]: {str(e)}")
    finally:
        app_gui.finalizar_execucao()


# ---------------------------------------------------------
# INTERFACE GRÁFICA DO SEPARADOR
# ---------------------------------------------------------
class InterfaceSeparador(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Fatiador de PDFs - RH")
        self.geometry("500x500")
        ctk.set_appearance_mode("dark")
        
        self.lbl_titulo = ctk.CTkLabel(self, text="Separador de Avisos/Recibos", font=("Arial", 18, "bold"))
        self.lbl_titulo.pack(pady=(20, 10))
        
        self.frame_inputs = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_inputs.pack(pady=10)
        
        # --- SELEÇÃO DO ARQUIVO PDF GIGANTE ---
        ctk.CTkLabel(self.frame_inputs, text="PDF Principal:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.caminho_pdf_var = ctk.StringVar()
        
        self.frame_pdf = ctk.CTkFrame(self.frame_inputs, fg_color="transparent")
        self.frame_pdf.grid(row=0, column=1, padx=0, pady=0, sticky="w")
        
        self.entry_pdf = ctk.CTkEntry(self.frame_pdf, width=140, textvariable=self.caminho_pdf_var, state="readonly")
        self.entry_pdf.pack(side="left", padx=(5, 5))
        
        self.btn_selecionar = ctk.CTkButton(self.frame_pdf, text="Buscar", width=50, command=self.selecionar_pdf)
        self.btn_selecionar.pack(side="left")
        
        # --- INPUT DA UNIDADE PADRÃO (Plano B caso não ache a UF no texto) ---
        ctk.CTkLabel(self.frame_inputs, text="Unidade Padrão (Opcional):").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        self.input_unidade = ctk.CTkEntry(self.frame_inputs, width=200, placeholder_text="Ex: DF (Usado se o robô não ler a UF)")
        self.input_unidade.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        # --- CAIXA DE STATUS ---
        self.caixa_logs = ctk.CTkTextbox(self, width=450, height=200)
        self.caixa_logs.pack(pady=10)
        self.caixa_logs.insert("0.0", "Selecione o PDF único com todos os recibos/avisos e clique em Iniciar.\n\n")
        
        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Separação", command=self.iniciar_thread)
        self.btn_iniciar.pack(pady=10)

    def selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o PDF com todos os arquivos",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if caminho:
            self.caminho_pdf_var.set(caminho)

    def iniciar_thread(self):
        caminho_pdf = self.caminho_pdf_var.get().strip()
        unidade_padrao = self.input_unidade.get().strip()
        
        if not caminho_pdf:
            self.atualizar_status("ERRO: Selecione o arquivo PDF primeiro!")
            return

        self.btn_iniciar.configure(state="disabled") 
        self.atualizar_status("Iniciando fatiamento do PDF...")
        
        # Roda a função em segundo plano para não travar a tela
        thread_separacao = threading.Thread(target=executar_separacao, args=(self, caminho_pdf, unidade_padrao))
        thread_separacao.start()

    def atualizar_status(self, mensagem):
        self.caixa_logs.insert("end", mensagem + "\n")
        self.caixa_logs.see("end")

    def finalizar_execucao(self):
        self.btn_iniciar.configure(state="normal", text="Executar Novamente")

if __name__ == "__main__":
    app = InterfaceSeparador()
    app.mainloop()