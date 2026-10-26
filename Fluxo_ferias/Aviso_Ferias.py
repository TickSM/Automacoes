import customtkinter as ctk
import threading
import os
import time
from collections import defaultdict
from playwright.sync_api import sync_playwright
from datetime import datetime
import pyautogui
import pyperclip
import pandas as pd
import re
import sys
from banco_ferias import registrar_aviso


# Descobre o caminho real onde o programa (ou .exe) está rodando
if getattr(sys, 'frozen', False):
    # Se for um executável compilado (PyInstaller)
    DIRETORIO_BASE = os.path.dirname(sys.executable)
else:
    # Se for o script .py rodando no VSCode/Terminal
    DIRETORIO_BASE = os.path.dirname(os.path.abspath(__file__))

# Cria os caminhos absolutos garantidos
CAMINHO_PASTA_PDFS = os.path.join(DIRETORIO_BASE, "PDFs_Avisos")

# 1. FUNÇÃO DE LEITURA
def processar_pasta_pdfs(pasta_pdfs):
    dados_agrupados = defaultdict(list)
    if not os.path.exists(CAMINHO_PASTA_PDFS):
        os.makedirs(CAMINHO_PASTA_PDFS)
        raise Exception(f"A pasta 'PDFs_Avisos' foi criada ao lado do programa. Coloque os PDFs lá dentro e tente novamente.")

    arquivos_pdf = [arq for arq in os.listdir(CAMINHO_PASTA_PDFS) if arq.lower().endswith('.pdf')]
    if not arquivos_pdf:
        raise Exception("Nenhum arquivo PDF encontrado.")

    for arquivo in arquivos_pdf:
        if "___" in arquivo[:-4]:
            partes = arquivo[:-4].split("___")
            unidade = partes[0].strip().upper()
            nome_formatado = partes[1].replace("_", " ").strip()
            dados_agrupados[unidade].append({
                'nome_interessado': nome_formatado,
                'nome_arquivo': arquivo
            })
    return dict(dados_agrupados)


# ---------------------------------------------------------
# 2. AUTOMAÇÃO NO SEI
# ---------------------------------------------------------
def executar_robo_sei(app_gui, mes_ano, usuario, senha):
    try:
        dados_agrupados = processar_pasta_pdfs(CAMINHO_PASTA_PDFS)
        app_gui.atualizar_status(f"Encontradas {len(dados_agrupados)} unidades para processar.\n")
        
        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=False, channel="chrome", args=['--start-maximized'])
            contexto = navegador.new_context(no_viewport=True)
            pagina = contexto.new_page()

            # Aceita pop-ups do SEI automaticamente
            pagina.on("dialog", lambda dialog: dialog.accept())

            app_gui.atualizar_status("Acessando a tela de Login...")
            pagina.goto("https://sei.assefaz.org.br/sip/login.php?sigla_orgao_sistema=ASSEFAZ&sigla_sistema=SEI") 
            
            # --- LOGIN ---
            pagina.locator("#txtUsuario").fill(usuario)
            pagina.locator("#pwdSenha").fill(senha)
            pagina.locator("#Acessar").click()

            # Espera carregar a tela inicial verificando se o menu lateral apareceu
            pagina.wait_for_selector("#divInfraSidebarMenu", timeout=15000)
            app_gui.atualizar_status("✔ Login realizado com sucesso!")

            # Cria uma lista vazia para guardar os dados do relatório
            relatorio_processos = []

            # --- LAÇO PRINCIPAL (CRIAR PROCESSOS) ---
            for unidade, lista_arquivos in dados_agrupados.items():
                # --- NOVO: TRY AQUI PARA PROTEGER A UNIDADE ---
                try:
                    app_gui.atualizar_status(f"\n➤ Criando processo da unidade: {unidade}")

                    menu_lateral = pagina.locator("#divInfraSidebarMenu")
                    
                    # Clica em Iniciar Processo no menu lateral
                    menu_lateral.get_by_role("link", name="Iniciar Processo").click()

                    try:
                        pagina.get_by_role("link", name="Exibir todos os tipos").click(timeout=2000)
                    except:
                        pass
                    
                    # Selecionando o processo
                    pagina.get_by_role("link", name="GNGDP - Aviso de férias").click() 
                    
                    # Preenche a especificação
                    pagina.get_by_role("textbox", name="Especificação:").fill(f"AVISO DE FÉRIAS - {unidade} - {mes_ano}")
                    
                    # Inserindo os Interessados
                    for item in lista_arquivos:
                        nome = item['nome_interessado']
                        app_gui.atualizar_status(f"  ↳ Add Interessado: {nome}")
                        
                        campo_interessado = pagina.get_by_role("textbox", name="Interessados:")
                        
                        campo_interessado.click()
                        campo_interessado.clear()
                        campo_interessado.fill(nome)
                        
                        campo_interessado.press("Space")
                        campo_interessado.press("Backspace")
                        
                        pagina.wait_for_timeout(1500) 
                        
                        campo_interessado.press("ArrowDown")
                        pagina.wait_for_timeout(500)
                        campo_interessado.press("Enter")
                        
                        pagina.wait_for_timeout(500)    

                    # Salva o processo                
                    pagina.locator("#divInfraBarraComandosSuperior").get_by_role("button", name="Salvar").click()
                    app_gui.atualizar_status(f"✔ Processo base criado!")
                                
                    # Aguarda o processo ser gerado
                    pagina.wait_for_selector("#ifrVisualizacao")

                    # ---------------------------------------------------------
                    # EXTRAÇÃO DO NÚMERO DO PROCESSO E DOS NOMES
                    # ---------------------------------------------------------
                    pagina.wait_for_selector("#ifrArvore")
                    frame_arvore = pagina.frame_locator("#ifrArvore")
                    
                    numero_processo = frame_arvore.get_by_role("link", name=re.compile(r"Assefaz\.\d")).first.inner_text()
                    app_gui.atualizar_status(f"  ↳ Processo gerado: {numero_processo}")
                    
                    # Pega a lista de nomes e junta todos separados por " / "
                    nomes_extraidos = [item['nome_interessado'] for item in lista_arquivos]
                    nomes_formatados = " / ".join(nomes_extraidos)
                    
                    relatorio_processos.append({
                        "Unidade": unidade,
                        "Processo": numero_processo,
                        "Interessados": nomes_formatados
                    })

                    # Salva no histórico local para conciliação automática com a etapa de recibos
                    try:
                        registrar_aviso(mes_ano, unidade, numero_processo, nomes_extraidos)
                        app_gui.atualizar_status("  ↳ Registrado no histórico local para conciliação automática!")
                    except Exception as e_banco:
                        app_gui.atualizar_status(f"  [Aviso Banco] {e_banco}")
                    # ---------------------------------------------------------

                    # entrando no iframe visualização
                    frame_vis = pagina.frame_locator("#ifrVisualizacao")

                    # ==========================================
                    # Inclusão do documento interno
                    # ==========================================
                    app_gui.atualizar_status(f"  ↳ Inserindo documento GNGDP - Orientações aviso de férias")
                    botao_incluir_interno = frame_vis.locator("img[title='Incluir Documento']")
                    botao_incluir_interno.wait_for(state="visible") 
                    botao_incluir_interno.click()
                    
                    try:
                        frame_vis.get_by_role("link", name="Exibir todos os tipos").click(timeout=2000)
                    except:
                        pass

                    pagina.wait_for_timeout(500)

                    frame_vis.get_by_role("link", name="GNGDP - Orientações aviso de férias", exact=True).click()              

                    frame_vis.get_by_role("textbox", name="Descrição:").fill(f"{mes_ano}")
                    frame_vis.get_by_role("textbox", name="Nome na Árvore:").fill(f"{mes_ano}")

                    frame_vis.locator("#divInfraBarraComandosSuperior").wait_for()            

                    with pagina.expect_popup() as nova_aba_info:
                        frame_vis.locator("#divInfraBarraComandosSuperior").locator("#btnSalvar").click()

                    aba_documento = nova_aba_info.value
                    aba_documento.wait_for_load_state()
                    aba_documento.close()

                    # ==========================================
                    # INCLUSÃO DOS PDFs
                    # ==========================================
                    pagina.wait_for_timeout(1000)

                    for item in lista_arquivos:
                        arquivo_pdf = item['nome_arquivo']
                        nome = item['nome_interessado']
                        caminho_relativo = os.path.join(CAMINHO_PASTA_PDFS, arquivo_pdf)
                        caminho_absoluto = os.path.abspath(caminho_relativo)
                        
                        app_gui.atualizar_status(f"  ↳ Anexando PDF: {arquivo_pdf}")
                        
                        botao_incluir = frame_vis.locator("img[title='Incluir Documento']")
                        botao_incluir.wait_for(state="visible") 
                        botao_incluir.click()
                        
                        frame_vis.get_by_role("link", name="Externo", exact=True).click()
                        frame_vis.get_by_label("Tipo do Documento:").select_option("305")

                        pagina.wait_for_timeout(1000)

                        frame_vis.locator("#divOptNato > .infraRadioDiv > .infraRadioLabel").click() 
                        pagina.wait_for_timeout(500) 

                        data_atual = datetime.now().strftime("%d/%m/%Y")
                        frame_vis.get_by_role("textbox", name="Data do Documento:").fill(data_atual)
                        frame_vis.get_by_role("textbox", name="Nome na Árvore:").fill(nome)      
                        
                        # ==========================================
                        # UPLOAD NATIVO E INVISÍVEL
                        # ==========================================
                        with pagina.expect_file_chooser() as fc_info:
                            frame_vis.locator("label#lblArquivo").click()
                        
                        janela_upload = fc_info.value
                        janela_upload.set_files(caminho_absoluto)
                        pagina.wait_for_timeout(2000)
                        # ==========================================
                        
                        frame_vis.locator("#divInfraBarraComandosSuperior").locator("#btnSalvar").click()
                        pagina.wait_for_timeout(1000)
                    
                    app_gui.atualizar_status(f"✔ Unidade {unidade} finalizada com sucesso!")

                    pagina.locator("#lnkControleProcessos").click()
                    pagina.wait_for_timeout(1500)
                    
                # --- EXCEPT PARA PROTEGER A UNIDADE ---
                except Exception as erro_unidade:
                    app_gui.atualizar_status(f"  [ERRO GRAVE] Falha ao processar unidade {unidade}. Pulando...")
                    
                    nomes_extraidos = [item['nome_interessado'] for item in lista_arquivos]
                    nomes_formatados = " / ".join(nomes_extraidos)
                    
                    relatorio_processos.append({
                        "Unidade": unidade,
                        "Processo": "ERRO DURANTE A CRIAÇÃO",
                        "Interessados": nomes_formatados
                    })
                    
                    # Força a voltar para a tela inicial do SEI para não quebrar a próxima unidade
                    try:
                        pagina.locator("#lnkControleProcessos").click(timeout=3000)
                    except:
                        pagina.goto("https://sei.assefaz.org.br/sip/login.php?sigla_orgao_sistema=ASSEFAZ&sigla_sistema=SEI")
                    continue
                # --------------------------------------------
            
            app_gui.atualizar_status("\nTodos os processos possíveis foram criados!")
            pagina.wait_for_timeout(1000)
            navegador.close()
            
    except Exception as e:
        app_gui.atualizar_status(f"\nERRO CRÍTICO NO ROBÔ: {str(e)}")
        
    # --- FINALLY PARA SALVAR A PLANILHA ACONTEÇA O QUE ACONTECER ---
    finally:
        if 'relatorio_processos' in locals() and relatorio_processos:
            try:
                df = pd.DataFrame(relatorio_processos)
                nome_excel = f"Relatorio_Avisos_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx"
                caminho_excel = os.path.join(DIRETORIO_BASE, nome_excel)
                
                with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, startrow=2, sheet_name="Processos")
                    planilha = writer.sheets["Processos"]
                    planilha["A1"] = f"Relatório de Avisos de Férias - Referência: {mes_ano}"
                    
                app_gui.atualizar_status(f"\n✔ Relatório Excel salvo com segurança: {nome_excel}")
            except Exception as e_excel:
                app_gui.atualizar_status(f"\n[ERRO] Falha ao salvar o Excel: {str(e_excel)}")
                
        app_gui.finalizar_execucao()


# ---------------------------------------------------------
# 3. INTERFACE GRÁFICA ATUALIZADA
# ---------------------------------------------------------
class InterfaceRobo(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Automação SEI - Aviso de Férias")
        self.geometry("450x550")
        ctk.set_appearance_mode("dark")
        
        self.lbl_titulo = ctk.CTkLabel(self, text="Gerador de Processos (SEI)", font=("Arial", 18, "bold"))
        self.lbl_titulo.pack(pady=(20, 10))
        
        self.frame_inputs = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_inputs.pack(pady=10)
        
        ctk.CTkLabel(self.frame_inputs, text="Usuário SEI:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.input_usuario = ctk.CTkEntry(self.frame_inputs, width=200)
        self.input_usuario.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_inputs, text="Senha:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.input_senha = ctk.CTkEntry(self.frame_inputs, width=200, show="*")
        self.input_senha.grid(row=1, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_inputs, text="Mês/Ano:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.input_mes_ano = ctk.CTkEntry(self.frame_inputs, width=200, placeholder_text="Ex: SETEMBRO/2026")
        self.input_mes_ano.grid(row=2, column=1, padx=5, pady=5)
        
        self.caixa_logs = ctk.CTkTextbox(self, width=400, height=180)
        self.caixa_logs.pack(pady=10)
        self.caixa_logs.insert("0.0", "Preencha os dados e clique em Iniciar.\n\n")
        
        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Automação", command=self.iniciar_thread)
        self.btn_iniciar.pack(pady=10)

    def iniciar_thread(self):
        usuario = self.input_usuario.get().strip()
        senha = self.input_senha.get().strip()
        mes_ano = self.input_mes_ano.get().strip().upper()
        
        if not usuario or not senha or not mes_ano:
            self.atualizar_status("ERRO: Preencha Usuário, Senha e Mês/Ano!")
            return

        self.btn_iniciar.configure(state="disabled") 
        self.atualizar_status("Preparando execução...")
        
        thread_robo = threading.Thread(target=executar_robo_sei, args=(self, mes_ano, usuario, senha))
        thread_robo.start()

    def atualizar_status(self, mensagem):
        self.caixa_logs.insert("end", mensagem + "\n")
        self.caixa_logs.see("end")

    def finalizar_execucao(self):
        self.btn_iniciar.configure(state="normal", text="Executar Novamente")


if __name__ == "__main__":
    app = InterfaceRobo()
    app.mainloop()