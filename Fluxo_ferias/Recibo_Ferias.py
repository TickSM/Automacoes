import customtkinter as ctk
import tkinter.filedialog as filedialog
import threading
import os
import re
import sys
from playwright.sync_api import sync_playwright
from datetime import datetime
import pandas as pd
import openpyxl
from banco_ferias import buscar_aviso_por_funcionario, registrar_recibo




if getattr(sys, 'frozen', False):
    DIRETORIO_BASE = os.path.dirname(sys.executable)
else:
    DIRETORIO_BASE = os.path.dirname(os.path.abspath(__file__))

CAMINHO_PASTA_PDFS = os.path.join(DIRETORIO_BASE, "PDFs_Recibos")

# ---------------------------------------------------------
# 1. FUNÇÕES DE LEITURA (PDF e EXCEL)
# ---------------------------------------------------------
def processar_pasta_pdfs(pasta_pdfs):
    """Lê os PDFs e cria uma lista de tarefas individual (1 PDF = 1 Processo)"""
    lista_arquivos = []
    if not os.path.exists(CAMINHO_PASTA_PDFS):
        os.makedirs(CAMINHO_PASTA_PDFS)
        raise Exception(f"A pasta 'PDFs_Recibos' foi criada. Coloque os PDFs lá dentro e tente novamente.")

    arquivos_pdf = [arq for arq in os.listdir(CAMINHO_PASTA_PDFS) if arq.lower().endswith('.pdf')]
    if not arquivos_pdf:
        raise Exception("Nenhum arquivo PDF encontrado.")

    for arquivo in arquivos_pdf:
        if "___" in arquivo[:-4]:
            partes = arquivo[:-4].split("___")
            unidade = partes[0].strip().upper()
            nome_formatado = partes[1].replace("_", " ").strip()
            lista_arquivos.append({
                'unidade': unidade,
                'nome_interessado': nome_formatado,
                'nome_arquivo': arquivo
            })
    return lista_arquivos

def carregar_planilha_avisos(caminho_excel):
    """Carrega a planilha gerada na etapa anterior como um dicionário"""
    try:
        # Pula as 2 primeiras linhas porque a linha 1 contém o título personalizado do código 1
        df = pd.read_excel(caminho_excel, skiprows=2) 
        return df.to_dict('records')
    except Exception as e:
        raise Exception(f"Erro ao ler a planilha: {str(e)}")

def encontrar_processo_aviso(unidade, nome_funcionario, dados_planilha):
    """Busca na planilha o processo correspondente cruzando apenas o Nome, ignorando a Unidade"""
    for linha in dados_planilha:
        # Os nomes na planilha foram salvos separados por " / "
        interessados_planilha = str(linha.get('Interessados', '')).upper()
        
        # Se achar o nome do funcionário em qualquer linha da planilha, retorna o processo
        if nome_funcionario.upper() in interessados_planilha:
            return str(linha.get('Processo', '')).strip()
            
    return None


# ---------------------------------------------------------
# 2. AUTOMAÇÃO NO SEI (RECIBOS)
# ---------------------------------------------------------
def executar_robo_sei(app_gui, mes_ano, usuario, senha, caminho_excel):
    try:
        # 1. Carrega os dados
        lista_tarefas_pdfs = processar_pasta_pdfs(CAMINHO_PASTA_PDFS)
        
        # SÓ TENTA CARREGAR A PLANILHA SE O CAMINHO FOI PREENCHIDO
        dados_planilha = []
        if caminho_excel:
            dados_planilha = carregar_planilha_avisos(caminho_excel)
            app_gui.atualizar_status("Planilha de vinculação detectada.")
        else:
            app_gui.atualizar_status("Nenhuma planilha anexada. Modo de criação SEM vinculação ativo.")
            
        app_gui.atualizar_status(f"Encontrados {len(lista_tarefas_pdfs)} recibos(PDFs) para processar.")
        
        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=False, channel="msedge", args=['--start-maximized'])
            contexto = navegador.new_context(no_viewport=True)
            pagina = contexto.new_page()

            pagina.on("dialog", lambda dialog: dialog.accept())

            app_gui.atualizar_status("Acessando a tela de Login...")
            pagina.goto("https://sei.assefaz.org.br/sip/login.php?sigla_orgao_sistema=ASSEFAZ&sigla_sistema=SEI") 
            
            pagina.locator("#txtUsuario").fill(usuario)
            pagina.locator("#pwdSenha").fill(senha)
            pagina.locator("#Acessar").click()

            pagina.wait_for_selector("#divInfraSidebarMenu", timeout=15000)
            app_gui.atualizar_status("✔ Login realizado com sucesso!")

            # Lista para armazenar os dados do relatório
            relatorio_processos = []


            # --- Variável para controlar o clique ---
            clicou_exibir_tipos = False
            clicou_exibir_documentos = False

            # --- LAÇO PRINCIPAL (UM PROCESSO POR FUNCIONÁRIO) ---
            for item in lista_tarefas_pdfs:
                unidade = item['unidade']
                nome = item['nome_interessado']
                arquivo_pdf = item['nome_arquivo']
                
                try:
                    app_gui.atualizar_status(f"\n➤ Criando processo para: {nome} ({unidade})")
                
                    processo_aviso_vinculado = None
                    status_do_vinculo = "NÃO SE APLICA (SEM VÍNCULO)"

                    # 1. Tenta localizar automaticamente no histórico local (banco SQLite)
                    try:
                        processo_aviso_vinculado = buscar_aviso_por_funcionario(nome, mes_ano)
                        if processo_aviso_vinculado:
                            app_gui.atualizar_status(f"  ↳ Aviso localizado no histórico local: {processo_aviso_vinculado}")
                    except Exception:
                        processo_aviso_vinculado = None

                    # 2. Se não achou no banco, tenta na planilha caso tenha sido informada
                    if not processo_aviso_vinculado and caminho_excel and dados_planilha:
                        processo_aviso_vinculado = encontrar_processo_aviso(unidade, nome, dados_planilha)
                        if processo_aviso_vinculado:
                            app_gui.atualizar_status(f"  ↳ Aviso localizado na planilha: {processo_aviso_vinculado}")

                    if processo_aviso_vinculado:
                        app_gui.atualizar_status(f"  ↳ Será vinculado ao Aviso Nº: {processo_aviso_vinculado}")
                    elif caminho_excel:
                        app_gui.atualizar_status(f"  [ERRO] Aviso não encontrado no histórico nem na planilha para {nome}. Pulando funcionário...")
                        relatorio_processos.append({
                            "Funcionário": nome,
                            "Processo do Recibo": "NÃO GERADO - AVISO NÃO ENCONTRADO",
                            "Processo do Aviso Vinculado": "NÃO ENCONTRADO"
                        })
                        continue
                    else:
                        app_gui.atualizar_status("  ↳ Gerando processo avulso (sem vinculação)...")

                    menu_lateral = pagina.locator("#divInfraSidebarMenu")
                    menu_lateral.get_by_role("link", name="Iniciar Processo").click()

                    # 1. Dá uma pequena pausa para o SEI carregar a tela central
                    pagina.wait_for_timeout(1500)

                    # 2. Só tenta clicar se for a primeira vez
                    if not clicou_exibir_tipos:
                        try:
                            pagina.get_by_role("link", name="Exibir todos os tipos").click(timeout=3000)
                            pagina.wait_for_timeout(1000) # Pausa rápida para a lista expandir
                        except Exception as e:
                            app_gui.atualizar_status(f"  [Info] Não foi necessário/possível clicar em 'Exibir todos os tipos'.")
                        finally:
                            # Independentemente de ter achado o botão ou não na primeira vez,
                            # ele marca como True para nunca mais perder tempo procurando nos próximos.
                            clicou_exibir_tipos = True
                    
                    # ATENÇÃO: Ajuste o nome do processo de recibo se for diferente
                    pagina.get_by_role("link", name="GNGDP - Recibo de férias").click() 
                    
                    # Especificação mais detalhada para facilitar a busca
                    pagina.get_by_role("textbox", name="Especificação:").fill(f"RECIBO DE FÉRIAS - {nome} - {mes_ano}")
                    
                    # Inserindo o Interessado (Apenas ele)
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
                    
                    pagina.locator("#divInfraBarraComandosSuperior").get_by_role("button", name="Salvar").click()

                    pagina.wait_for_timeout(3000)

                    # ---------------------------------------------------------
                    # VINCULAR O PROCESSO DE AVISO (CONDICIONAL)
                    # ---------------------------------------------------------
                    pagina.wait_for_selector("#ifrVisualizacao")
                    pagina.wait_for_selector("#ifrArvore")
                    
                    frame_vis = pagina.frame_locator("#ifrVisualizacao")
                    frame_arvore = pagina.frame_locator("#ifrArvore")
                    
                    pagina.wait_for_timeout(2000)
                    
                    if processo_aviso_vinculado:
                        botao_relacionar = frame_vis.get_by_role("link", name="Relacionamentos")
                        botao_relacionar.wait_for(state="visible") 
                        botao_relacionar.click()
                        
                        frame_vis.get_by_role("textbox", name="Processo Destino:").fill(processo_aviso_vinculado)
                        pagina.wait_for_timeout(1000)
                        frame_vis.get_by_role("button", name="Pesquisar").click() 
                        
                        status_do_vinculo = processo_aviso_vinculado 
                        
                        try:
                            frame_vis.get_by_role("button", name="Adicionar").click(timeout=3000)
                            pagina.wait_for_timeout(1000)
                        except Exception:
                            app_gui.atualizar_status(f"  [AVISO] Processo {processo_aviso_vinculado} não localizado no SEI. Seguindo sem vincular...")
                            status_do_vinculo = f"FALHOU - PROCESSO {processo_aviso_vinculado} NÃO ENCONTRADO NO SEI"

                        # === O CLIQUE DA ÁRVORE VEM PARA CÁ (DENTRO DO IF) ===
                        pagina.wait_for_timeout(1500)
                        frame_arvore.get_by_role("link", name=re.compile(r"Assefaz\.\d")).first.click()
                        pagina.wait_for_timeout(2000)
                        # =====================================================
                    
                    # Coleta do número do processo recém-criado
                    numero_processo_novo = frame_arvore.get_by_role("link", name=re.compile(r"Assefaz\.\d")).first.inner_text()
                    
                    relatorio_processos.append({
                        "Funcionário": nome,
                        "Processo do Recibo": numero_processo_novo,
                        "Processo do Aviso Vinculado": status_do_vinculo
                    })

                    # Salva no histórico local para rastreabilidade
                    try:
                        registrar_recibo(mes_ano, nome, numero_processo_novo, processo_aviso_vinculado, status_do_vinculo)
                    except Exception:
                        pass

                    # ==========================================
                    # INCLUSÃO DO DOCUMENTO INTERNO
                    # ==========================================

                    app_gui.atualizar_status(f"  ↳ Incluindo o documento Orientações recibo de férias")

                    botao_incluir = frame_vis.locator("img[title='Incluir Documento']")
                    botao_incluir.wait_for(state="visible") 
                    botao_incluir.click()

                    # --- Só tenta clicar se for a primeira vez ---
                    if not clicou_exibir_documentos:
                        try:
                            # Reduzimos o timeout para 3000 (3 segundos) para otimizar
                            frame_vis.get_by_role("link", name="Exibir todos os tipos").click(timeout=3000)
                            pagina.wait_for_timeout(1000) # Pausa rápida para a lista expandir
                        except Exception as e:
                            app_gui.atualizar_status(f"  [Info] Não foi necessário/possível clicar em 'Exibir todos os documentos'.")
                        finally:
                            # Marca como True para pular essa etapa nos próximos funcionários
                            clicou_exibir_documentos = True
                    # ---------------------------------------------------

                    frame_vis.get_by_role("link", name="GNGDP - Orientações recibo de férias", exact=True).click()

                    with pagina.expect_popup() as nova_aba_info:
                        frame_vis.locator("#divInfraBarraComandosSuperior").locator("#btnSalvar").click()

                    aba_documento = nova_aba_info.value
                    aba_documento.wait_for_load_state()
                    aba_documento.close()

                    # ==========================================
                    # INCLUSÃO DO PDF DO RECIBO
                    # ==========================================
                    caminho_relativo = os.path.join(CAMINHO_PASTA_PDFS, arquivo_pdf)
                    caminho_absoluto = os.path.abspath(caminho_relativo)
                    
                    app_gui.atualizar_status(f"  ↳ Anexando PDF: {arquivo_pdf}")
                    
                    botao_incluir = frame_vis.locator("img[title='Incluir Documento']")
                    botao_incluir.wait_for(state="visible") 
                    botao_incluir.click()
                    
                    frame_vis.get_by_role("link", name="Externo", exact=True).click()
                    frame_vis.get_by_label("Tipo do Documento:").select_option("305") # Verifique se 305 é o código correto para recibo no seu SEI

                    pagina.wait_for_timeout(1000)

                    frame_vis.locator("#divOptNato > .infraRadioDiv > .infraRadioLabel").click() 
                    pagina.wait_for_timeout(500) 

                    data_atual = datetime.now().strftime("%d/%m/%Y")
                    frame_vis.get_by_role("textbox", name="Data do Documento:").fill(data_atual)
                    frame_vis.get_by_role("textbox", name="Nome na Árvore:").fill(f"Recibo - {nome}")      
                    
                    # ==========================================
                    # UPLOAD NATIVO E INVISÍVEL
                    # ==========================================
                    
                    # 1. Avisa o Playwright para ficar esperando a janela de upload ser chamada
                    with pagina.expect_file_chooser() as fc_info:
                        # 2. Clica no botão do SEI que normalmente abriria o Windows Explorer
                        frame_vis.locator("label#lblArquivo").click()
                    
                    # 3. Pega o controle dessa janela "invisível" interceptada
                    janela_upload = fc_info.value
                    
                    # 4. Injeta o caminho absoluto do seu PDF diretamente nela
                    janela_upload.set_files(caminho_absoluto)
                    
                    # Dá um pequeno tempo para o SEI reconhecer o arquivo injetado
                    pagina.wait_for_timeout(2000)
                    
                    # ==========================================
                    
                    frame_vis.locator("#divInfraBarraComandosSuperior").locator("#btnSalvar").click()
                    pagina.wait_for_timeout(1000)
                    
                    app_gui.atualizar_status(f"✔ Recibo de {nome} finalizado!")

                    # Volta para controle de processos para iniciar o próximo funcionário
                    pagina.locator("#lnkControleProcessos").click()
                    pagina.wait_for_timeout(1500)
                    
                except Exception as erro_funcionario:
                    app_gui.atualizar_status(f"  [ERRO GRAVE] Falha ao processar {nome}. Pulando...")
                    
                    # Adiciona a falha no relatório
                    relatorio_processos.append({
                        "Funcionário": nome,
                        "Processo do Recibo": "ERRO DURANTE A CRIAÇÃO",
                        "Processo do Aviso Vinculado": "FALHOU"
                    })
                    
                    # Força a voltar para a tela inicial do SEI com tratamento de navegação
                    try:
                        # Avisa o Playwright que este clique vai gerar uma mudança de página
                        with pagina.expect_navigation(timeout=10000):
                            pagina.locator("#lnkControleProcessos").click()
                    except:
                        # Se não achar o botão ou travar, força pela URL limpa do controle de processos
                        try:
                            pagina.goto("https://sei.assefaz.org.br/sei/controlador.php?acao=procedimento_controlar&reset=1&infra_sistema=100000100&infra_unidade_atual=110000001")
                        except:
                            pass
                    continue # Pula para o próximo PDF da lista
                
            app_gui.atualizar_status("\nTodos os recibos possíveis foram processados!")
            pagina.wait_for_timeout(1000)
            navegador.close()
            
    except Exception as e:
        app_gui.atualizar_status(f"\nERRO CRÍTICO NO ROBÔ: {str(e)}")
        
    finally:
        # --- SALVAMENTO GARANTIDO DO EXCEL ---
        if 'relatorio_processos' in locals() and relatorio_processos:
            try:
                df_relatorio = pd.DataFrame(relatorio_processos)
                nome_excel_saida = f"Relatorio_Recibos_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx"
                caminho_excel_saida = os.path.join(DIRETORIO_BASE, nome_excel_saida)
                
                with pd.ExcelWriter(caminho_excel_saida, engine='openpyxl') as writer:
                    df_relatorio.to_excel(writer, index=False, startrow=2, sheet_name="Recibos Gerados")
                    planilha = writer.sheets["Recibos Gerados"]
                    planilha["A1"] = f"Relatório de Recibos de Férias - Referência: {mes_ano}"
                    
                app_gui.atualizar_status(f"\n✔ Relatório Excel salvo com segurança: {nome_excel_saida}")
            except Exception as e_excel:
                app_gui.atualizar_status(f"\n[ERRO] Falha ao salvar o Excel: {str(e_excel)}")
                
        app_gui.finalizar_execucao()


# ---------------------------------------------------------
# 3. INTERFACE GRÁFICA
# ---------------------------------------------------------
class InterfaceRobo(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Automação SEI - Recibo de Férias")
        self.geometry("500x600")
        ctk.set_appearance_mode("dark")
        
        self.lbl_titulo = ctk.CTkLabel(self, text="Gerador de Recibos (SEI)", font=("Arial", 18, "bold"))
        self.lbl_titulo.pack(pady=(20, 10))
        
        self.frame_inputs = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_inputs.pack(pady=10)
        
        ctk.CTkLabel(self.frame_inputs, text="Usuário SEI:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.input_usuario = ctk.CTkEntry(self.frame_inputs, width=200)
        self.input_usuario.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(self.frame_inputs, text="Senha:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.input_senha = ctk.CTkEntry(self.frame_inputs, width=200, show="*")
        self.input_senha.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(self.frame_inputs, text="Mês/Ano:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.input_mes_ano = ctk.CTkEntry(self.frame_inputs, width=200, placeholder_text="Ex: SETEMBRO/2026")
        self.input_mes_ano.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # --- SELEÇÃO DA PLANILHA ---
        ctk.CTkLabel(self.frame_inputs, text="Planilha Avisos:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.caminho_planilha_var = ctk.StringVar()
        
        self.frame_planilha = ctk.CTkFrame(self.frame_inputs, fg_color="transparent")
        self.frame_planilha.grid(row=3, column=1, padx=0, pady=0, sticky="w")
        
        self.entry_planilha = ctk.CTkEntry(self.frame_planilha, width=140, textvariable=self.caminho_planilha_var, state="readonly")
        self.entry_planilha.pack(side="left", padx=(5, 5))
        
        self.btn_selecionar = ctk.CTkButton(self.frame_planilha, text="Buscar", width=50, command=self.selecionar_planilha)
        self.btn_selecionar.pack(side="left")
        # ----------------------------------

        self.caixa_logs = ctk.CTkTextbox(self, width=450, height=180)
        self.caixa_logs.pack(pady=10)
        self.caixa_logs.insert("0.0", "Preencha os dados, anexe a planilha (Opcional) e clique em Iniciar.\nSe a planilha não for anexada, os recibos serão gerados sem vinculação.\n\n")
        
        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Automação", command=self.iniciar_thread)
        self.btn_iniciar.pack(pady=10)

    def selecionar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o Relatório de Avisos de Férias",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if caminho:
            self.caminho_planilha_var.set(caminho)

    def iniciar_thread(self):
        usuario = self.input_usuario.get().strip()
        senha = self.input_senha.get().strip()
        mes_ano = self.input_mes_ano.get().strip().upper()
        caminho_excel = self.caminho_planilha_var.get().strip()
        
        # REMOVIDA A OBRIGATORIEDADE DO caminho_excel
        if not usuario or not senha or not mes_ano:
            self.atualizar_status("ERRO: Preencha Usuário, Senha e Mês/Ano!")
            return

        self.btn_iniciar.configure(state="disabled") 
        self.atualizar_status("Preparando execução...")
        
        thread_robo = threading.Thread(target=executar_robo_sei, args=(self, mes_ano, usuario, senha, caminho_excel))
        thread_robo.start()

    def atualizar_status(self, mensagem):
        self.caixa_logs.insert("end", mensagem + "\n")
        self.caixa_logs.see("end")

    def finalizar_execucao(self):
        self.btn_iniciar.configure(state="normal", text="Executar Novamente")


if __name__ == "__main__":
    app = InterfaceRobo()
    app.mainloop()