import customtkinter as ctk
import threading
import os
import time
from playwright.sync_api import sync_playwright
from datetime import datetime
import pandas as pd
import re
import sys

# ---------------------------------------------------------
# 1. FUNÇÃO DE LEITURA DOS PDFs E CPFs
# ---------------------------------------------------------
def processar_pasta_pdfs(pasta_pdfs):
    if not pasta_pdfs or not os.path.exists(pasta_pdfs):
        raise Exception("A pasta selecionada é inválida ou não existe.")

    arquivos_pdf = [arq for arq in os.listdir(pasta_pdfs) if arq.lower().endswith('.pdf')]
    if not arquivos_pdf:
        raise Exception("Nenhum arquivo PDF encontrado na pasta selecionada.")

    lista_dados = []
    for arquivo in arquivos_pdf:
        nome_arquivo_bruto = arquivo[:-4].strip()
        lista_dados.append({
            'nome_bruto': nome_arquivo_bruto,
            'nome_arquivo': arquivo,
            'caminho_completo': os.path.join(pasta_pdfs, arquivo)
        })
    
    return lista_dados

def carregar_dicionario_cpfs(caminho_excel):
    """Lê a planilha de CPFs e formata para o padrão exigido pelo SEI (com pontos e traço)"""
    try:
        df = pd.read_excel(caminho_excel)
        colunas_upper = {str(c).strip().upper(): c for c in df.columns}
        
        if 'NOME' not in colunas_upper or 'CPF' not in colunas_upper:
            raise Exception("A planilha não possui as colunas 'NOME' e 'CPF'.")
            
        # Verifica se a nova coluna existe de forma flexível
        col_data = None
        for col in colunas_upper:
            if 'DATA DE INATIVA' in col or 'CANCELAMENTO' in col:
                col_data = colunas_upper[col]
                break
                
        dicionario_cpfs = {}
        for index, row in df.iterrows():
            nome_bruto = str(row[colunas_upper['NOME']]).strip().upper()
            nome_limpo = " ".join(nome_bruto.split())
            
            # Limpa o CPF para pegar apenas os números
            cpf_bruto = str(row[colunas_upper['CPF']])
            cpf_numeros = "".join(filter(str.isdigit, cpf_bruto))
            
            # Reconstrói a máscara do CPF (XXX.XXX.XXX-XX)
            if len(cpf_numeros) == 11:
                cpf_formatado = f"{cpf_numeros[:3]}.{cpf_numeros[3:6]}.{cpf_numeros[6:9]}-{cpf_numeros[9:]}"
            else:
                cpf_formatado = cpf_bruto
                
            # Extrai e formata a data, se a coluna existir e estiver preenchida
            data_inativacao = ""
            if col_data:
                valor_data = row[col_data]
                if pd.notna(valor_data):
                    # Se o Pandas leu como data, formata para DD/MM/AAAA
                    if isinstance(valor_data, pd.Timestamp) or isinstance(valor_data, datetime):
                        data_inativacao = valor_data.strftime('%d/%m/%Y')
                    else:
                        data_inativacao = str(valor_data).strip()
            
            if nome_limpo and nome_limpo != 'NAN':
                # Agora salvamos um dicionário com os dados da pessoa
                dicionario_cpfs[nome_limpo] = {
                    'cpf': cpf_formatado,
                    'data_inativacao': data_inativacao
                }
                
        return dicionario_cpfs
    except Exception as e:
        raise Exception(f"Erro ao ler a planilha de CPFs: {str(e)}")

# ---------------------------------------------------------
# 2. AUTOMAÇÃO NO SEI
# ---------------------------------------------------------
def executar_robo_sei(app_gui, usuario, senha, pasta_pdfs, planilha_cpfs, numero_bloco, opcao_inativo):
    try:
        lista_pdfs = processar_pasta_pdfs(pasta_pdfs)
        app_gui.atualizar_status(f"Encontrados {len(lista_pdfs)} PDFs para processar.\n")
        
        dicionario_cpfs = carregar_dicionario_cpfs(planilha_cpfs)
        app_gui.atualizar_status(f"Planilha de CPFs carregada com {len(dicionario_cpfs)} registros.")

        with sync_playwright() as p:
            navegador = p.chromium.launch(headless=False, channel="chrome", args=['--start-maximized'])
            contexto = navegador.new_context(no_viewport=True)
            pagina = contexto.new_page()

            pagina.on("dialog", lambda dialog: dialog.accept())

            app_gui.atualizar_status("Acessando a tela de Login...")
            pagina.goto("https://sei.assefaz.org.br/sip/login.php?sigla_orgao_sistema=ASSEFAZ&sigla_sistema=SEI") 
            
            # --- LOGIN ---
            pagina.locator("#txtUsuario").fill(usuario)
            pagina.locator("#pwdSenha").fill(senha)
            pagina.locator("#Acessar").click()

            # --- VALIDAÇÃO RÁPIDA DE LOGIN ---
            try:
                # Se o menu não aparecer em 8 segundos, assume que o login falhou
                pagina.wait_for_selector("#divInfraSidebarMenu", timeout=8000)
                app_gui.atualizar_status("✔ Login realizado com sucesso!")
            except:
                raise Exception("Não foi possível logar no SEI. Verifique seu Usuário e Senha.")

            # ==========================================
            # VERIFICAÇÃO E TROCA DE UNIDADE (ADB)
            # ==========================================
            app_gui.atualizar_status("Verificando unidade atual...")
            
            # Verifica se o link com o nome "ADB" exato está na tela
            botao_unidade_adb = pagina.get_by_role("link", name="ADB", exact=True)
            
            if not botao_unidade_adb.is_visible():
                app_gui.atualizar_status("  ↳ Unidade incorreta. Trocando para ADB...")
                
                # Usa o seletor ":visible" do próprio Playwright para pegar apenas o link que está aparecendo
                pagina.locator("#lnkInfraUnidade:visible").first.click()
                pagina.wait_for_timeout(2000) # Aguarda a tela de seleção carregar
                
                # Seleciona a unidade ADB na lista conforme o codegen
                pagina.get_by_role("row", name="ADB ADB Arquivo Digital de").locator("label").click()
                
                # Aguarda o redirecionamento e carregamento da nova unidade
                pagina.wait_for_timeout(3000)
                app_gui.atualizar_status("✔ Unidade alterada para ADB com sucesso!")
            else:
                app_gui.atualizar_status("✔ Usuário já está na unidade ADB!")

            relatorio_processos = []

            clicou_exibir_tipos = False
            clicou_exibir_documentos = False

            # --- LAÇO PRINCIPAL (1 PROCESSO POR PDF) ---
            for item in lista_pdfs:
                nome_bruto = item['nome_bruto']
                arquivo_pdf = item['nome_arquivo']
                caminho_absoluto = item['caminho_completo']
                
                # ==========================================
                # BUSCA NA PLANILHA E TRATAMENTO DE NOME
                # ==========================================
                # 1. Preserva o nome original (com o _) para buscar o CPF correto no dicionário
                nome_chave = " ".join(nome_bruto.upper().split())
                
                # 2. Cria a versão "limpa" (removendo o _) para digitar no SEI sem ficar estranho
                nome_limpo = " ".join(nome_bruto.upper().replace("_", " ").split())
                
                # 3. Busca os dados na planilha usando a CHAVE EXATA (com o _)
                dados_planilha = dicionario_cpfs.get(nome_chave, {})
                cpf_encontrado = dados_planilha.get('cpf', "")
                data_inativacao = dados_planilha.get('data_inativacao', "")
                
                # Variáveis de controle para o Relatório Excel
                numero_processo = "NÃO GERADO"
                etapa_atual = "PESQUISA PRÉVIA"

                try:
                    app_gui.atualizar_status(f"\n➤ Analisando: {nome_limpo}")

                    # ==========================================
                    # ETAPA NOVA: PESQUISA DETALHADA EM 2 PASSOS
                    # ==========================================
                    if cpf_encontrado:
                        app_gui.atualizar_status(f"  ↳ Abrindo pesquisa detalhada para o CPF: {cpf_encontrado}")
                        
                        # 1. Acessa a tela de pesquisa através da barra superior (compatível com múltiplas versões)
                        pagina.get_by_role("textbox", name="Pesquisar...").click()
                        pagina.get_by_role("textbox", name="Pesquisar...").press("Enter")
                        pagina.wait_for_timeout(2000)
                        
                        # 2. Preenche o CPF no campo de texto
                        pagina.get_by_role('textbox', name='Texto para Pesquisa:').fill(cpf_encontrado)
                        
                        # --- PRIMEIRA BUSCA ---
                        app_gui.atualizar_status("    ↳ Realizando 1ª busca (ADB Inativo)...")
                        
                        caixa_documento = pagina.get_by_label('Tipo do Documento:')
                        
                        # O robô procura a opção pelo texto de forma flexível e pega o "value" escondido dela
                        opcao_1 = caixa_documento.locator("option", has_text=re.compile(r"Acervo Arq de Dossiê de Beneficiário ADB - Inativo", re.IGNORECASE))
                        valor_opcao_1 = opcao_1.get_attribute("value")
                        
                        # Seleciona usando o valor que ele acabou de descobrir
                        caixa_documento.select_option(valor_opcao_1)
                        
                        # Clica no botão de pesquisar
                        pagina.get_by_role('button', name='Pesquisar').click()
                        pagina.wait_for_timeout(3500) 
                        
                        try:
                            pagina.get_by_text('Nenhum resultado encontrado.').wait_for(state="visible", timeout=3000)
                            achou_primeira = False
                        except:
                            achou_primeira = True
                            
                        processo_ja_existe = False
                        
                        if achou_primeira:
                            processo_ja_existe = True
                            app_gui.atualizar_status("    [!] Resultado localizado na 1ª busca!")
                        else:
                            # --- SEGUNDA BUSCA ---
                            app_gui.atualizar_status("    ↳ Realizando 2ª busca (ADB)...")
                            
                            # Faz o mesmo processo para a segunda opção
                            opcao_2 = caixa_documento.locator("option", has_text=re.compile(r"Acervo Arquivístico de Dossiê de Beneficiário ADB", re.IGNORECASE))
                            valor_opcao_2 = opcao_2.get_attribute("value")
                            
                            caixa_documento.select_option(valor_opcao_2)
                            
                            pagina.get_by_role('button', name='Pesquisar').click()
                            pagina.wait_for_timeout(3500)
                            
                            try:
                                pagina.get_by_text('Nenhum resultado encontrado.').wait_for(state="visible", timeout=3000)
                                achou_segunda = False
                            except:
                                achou_segunda = True
                                
                            if achou_segunda:
                                processo_ja_existe = True
                                app_gui.atualizar_status("    [!] Resultado localizado na 2ª busca!")
                                
                        # --- TOMA A DECISÃO FINAL ---
                        if processo_ja_existe:
                            app_gui.atualizar_status("  [!] Documento já existe no SEI. Pulando criação de processo.")
                            
                            # Adiciona no relatório
                            relatorio_processos.append({
                                "Nome da Pessoa": nome_limpo,
                                "CPF": cpf_encontrado,
                                "Processo": "NÃO GERADO - JÁ EXISTE",
                                "Arquivo PDF": arquivo_pdf,
                                "Status": "PULADO (DUPLICIDADE)",
                                "Marcador": "-",
                                "Atribuição": "-"
                            })
                            
                            # Segue direto para o próximo CPF do laço (reinicia o loop já clicando em Pesquisa de novo)
                            continue
                        else:
                            app_gui.atualizar_status("  ↳ Nenhum documento localizado nas buscas. Preparando criação...")
                            # Apenas segue o fluxo para baixo, direto para a criação do processo, sem voltar para o Controle
                    else:
                        app_gui.atualizar_status(f"  [!] Sem CPF na planilha para realizar a busca. Indo direto para criação...")

                    # ==========================================
                    # CONTINUAÇÃO NORMAL: CRIANDO O PROCESSO
                    # ==========================================
                    etapa_atual = "CRIANDO PROCESSO BASE"
                    app_gui.atualizar_status(f"  ↳ Criando processo base...")
                    
                    menu_lateral = pagina.locator("#divInfraSidebarMenu")
                    botao_iniciar = menu_lateral.get_by_role("link", name="Iniciar Processo")
                    
                    # Abordagem de tentativa: tenta clicar rápido. Se algo estiver na frente (menu oculto), ele falha e vai para o "except" abrir o menu.
                    try:
                        botao_iniciar.first.click(timeout=2000)
                    except:
                        # Se caiu aqui, é porque o menu está escondido. Vamos abri-lo!
                        botao_menu = pagina.get_by_role("link", name="Exibir/Ocultar Menu do Sistema")
                        botao_menu.first.click(force=True)
                        
                        # Aguarda a animação do menu abrir
                        pagina.wait_for_timeout(1500)
                        
                        # Agora clica com segurança
                        botao_iniciar.first.click(timeout=5000)
                    
                    pagina.wait_for_timeout(1500)

                    if not clicou_exibir_tipos:
                        try:
                            pagina.get_by_role("link", name="Exibir todos os tipos").click(timeout=3000)
                            pagina.wait_for_timeout(1000) 
                        except Exception:
                            pass
                        finally:
                            clicou_exibir_tipos = True
                    
                    pagina.get_by_role("link", name="GRC - Acervo Arquivístico de Dossiê de Beneficiário").click(timeout=10000) 
                    
                    pagina.wait_for_timeout(2000)
                    
                    # Preenche a Especificação com o Nome Limpo
                    pagina.get_by_role("textbox", name="Especificação:").fill(nome_limpo)
                    
                    # Inserindo os Interessados Fixos
                    interessados_fixos = [
                        'Arquivo Digital de beneficiário - ADB', 
                        'Coordenação de documentação e informação'
                    ]
                    
                    campo_interessado = pagina.get_by_role("textbox", name="Interessados:")
                    
                    for interessado in interessados_fixos:
                        app_gui.atualizar_status(f"  ↳ Add Interessado: {interessado}")
                        campo_interessado.click()
                        campo_interessado.clear()
                        campo_interessado.fill(interessado)
                        campo_interessado.press("Space")
                        campo_interessado.press("Backspace")
                        pagina.wait_for_timeout(1500) 
                        campo_interessado.press("ArrowDown")
                        pagina.wait_for_timeout(500)
                        campo_interessado.press("Enter")
                        pagina.wait_for_timeout(500)    

                    # Salva o processo                
                    pagina.locator("#divInfraBarraComandosSuperior").get_by_role("button", name="Salvar").click()
                    pagina.wait_for_timeout(5000)
                                
                    pagina.wait_for_selector("#ifrVisualizacao", timeout=15000)
                    pagina.wait_for_selector("#ifrArvore", timeout=10000)
                    frame_arvore = pagina.frame_locator("#ifrArvore")
                    numero_processo = frame_arvore.get_by_role("link", name=re.compile(r"Assefaz\.\d")).first.inner_text()
                    app_gui.atualizar_status(f"  ↳ Processo gerado: {numero_processo}")
                    
                    frame_vis = pagina.frame_locator("#ifrVisualizacao")

                    # ==========================================
                    # Inclusão do documento interno (FORMULÁRIO)
                    # ==========================================
                    etapa_atual = "PREENCHENDO FORMULÁRIO INTERNO"
                    nome_doc_interno = "Acervo Arq de Dossiê de Beneficiário ADB - Inativo (Formulário)"
                    app_gui.atualizar_status(f"  ↳ Inserindo formulário: {nome_doc_interno}")
                    
                    botao_incluir_interno = frame_vis.locator("img[title='Incluir Documento']")
                    botao_incluir_interno.wait_for(state="visible", timeout=10000) 
                    botao_incluir_interno.click()
                    
                    pagina.wait_for_timeout(2000)
                    
                    if not clicou_exibir_documentos:
                        try:
                            botao_exibir = frame_vis.get_by_role("link", name="Exibir todos os tipos")
                            botao_exibir.wait_for(state="visible", timeout=4000)
                            botao_exibir.click()
                            pagina.wait_for_timeout(1500)
                        except:
                            pass
                        finally:
                            clicou_exibir_documentos = True
                    
                    frame_vis.get_by_role("link", name="Acervo Arq de Dossiê de").click()              
                    
                    app_gui.atualizar_status(f"  ↳ Preenchendo campos do formulário (00, Nome, CPF)")
                    
                    frame_vis.get_by_role("textbox", name="MÓDULO:").wait_for(state="visible", timeout=10000)
                    
                    frame_vis.get_by_role("textbox", name="MÓDULO:").click()
                    frame_vis.get_by_role("textbox", name="MÓDULO:").fill("00")
                    
                    frame_vis.get_by_role("textbox", name="LADO:").click()
                    frame_vis.get_by_role("textbox", name="LADO:").fill("00")
                    
                    frame_vis.get_by_role("textbox", name="PRATELEIRA:").click()
                    frame_vis.get_by_role("textbox", name="PRATELEIRA:").fill("00")
                    
                    frame_vis.get_by_role("textbox", name="CAIXA :").click()
                    frame_vis.get_by_role("textbox", name="CAIXA :").fill("00")

                    if data_inativacao and data_inativacao.upper() != 'NAN':
                        try:
                            frame_vis.get_by_role("textbox", name="DATA DA INATIVAÇÃO/CANCELAMENTO:").click()
                            frame_vis.get_by_role("textbox", name="DATA DA INATIVAÇÃO/CANCELAMENTO:").fill(data_inativacao)
                        except Exception as e:
                            app_gui.atualizar_status(f"  [!] Não foi possível preencher a Data de Inativação.")
                    
                    frame_vis.get_by_role("textbox", name="NOME DO BENEFICIÁRIO:").click()
                    frame_vis.get_by_role("textbox", name="NOME DO BENEFICIÁRIO:").fill(nome_limpo)
                    
                    if cpf_encontrado:
                        frame_vis.get_by_role("textbox", name="CPF:").click()
                        frame_vis.get_by_role("textbox", name="CPF:").fill(cpf_encontrado)
                    
                    frame_vis.locator("#divInfraBarraComandosInferior").get_by_role("button", name="Salvar").click()
                    pagina.wait_for_timeout(3000)

                    # ==========================================
                    # UPLOAD DO PDF CORRESPONDENTE
                    # ==========================================
                    etapa_atual = "UPLOAD DO PDF EXTERNO"
                    pagina.wait_for_timeout(1000)
                    app_gui.atualizar_status(f"  ↳ Anexando PDF: {arquivo_pdf}")
                    
                    botao_incluir = frame_vis.locator("img[title='Incluir Documento']")
                    botao_incluir.wait_for(state="visible") 
                    botao_incluir.click()
                    
                    frame_vis.get_by_role("link", name="Externo", exact=True).click()
                    frame_vis.get_by_label("Tipo do Documento:").select_option("305")
                    
                    pagina.wait_for_timeout(2000)
                    
                    frame_vis.locator("#divOptDigitalizado > .infraRadioDiv > .infraRadioLabel").click()
                    pagina.wait_for_timeout(500) 

                    frame_vis.locator("#selTipoConferencia").select_option(index=4)
                    pagina.wait_for_timeout(500)

                    data_atual = datetime.now().strftime("%d/%m/%Y")
                    frame_vis.get_by_role("textbox", name="Data do Documento:").fill(data_atual)
                    frame_vis.get_by_role("textbox", name="Nome na Árvore:").fill(nome_limpo)      
                    
                    with pagina.expect_file_chooser() as fc_info:
                        frame_vis.locator("label#lblArquivo").click()
                    
                    janela_upload = fc_info.value
                    janela_upload.set_files(caminho_absoluto)
                    
                    # Aguarda dinamicamente até que o arquivo seja processado pelo sistema
                    # Coloquei um timeout de 30 segundos (30000 ms) como margem de segurança para arquivos grandes
                    frame_vis.get_by_text("Lista de Anexos (1 registro):").wait_for(state="visible", timeout=30000)
                    
                    frame_vis.locator("#divInfraBarraComandosSuperior").locator("#btnSalvar").click()
                    
                    # Mantemos uma pequena pausa após o clique de salvar apenas para o SEI processar a requisição de fechamento/atualização
                    pagina.wait_for_timeout(2000)
                    
                    app_gui.atualizar_status(f"✔ Processo {numero_processo} finalizado!")
                    
                    # SUCESSO TOTAL - ADICIONA NO RELATÓRIO
                    relatorio_processos.append({
                        "Nome da Pessoa": nome_limpo,
                        "CPF": cpf_encontrado,
                        "Processo": numero_processo,
                        "Arquivo PDF": arquivo_pdf,
                        "Status": "SUCESSO",
                        "Marcador": "-",
                        "Atribuição": "-"
                    })
                    
                except Exception as erro_unidade:
                    app_gui.atualizar_status(f"  [ERRO] Falha na etapa: {etapa_atual}")
                    app_gui.atualizar_status(f"  Motivo: {str(erro_unidade)}")
                    
                    # CASO DÊ ERRO, SALVA O STATUS EXATO ONDE O ROBÔ PAROU
                    relatorio_processos.append({
                        "Nome da Pessoa": nome_limpo,
                        "CPF": cpf_encontrado,
                        "Processo": numero_processo,
                        "Arquivo PDF": arquivo_pdf,
                        "Status": f"FALHA: {etapa_atual}",
                        "Marcador": "-",
                        "Atribuição": "-"
                    })
                    
                    # TENTA VOLTAR PARA A TELA INICIAL PARA O PRÓXIMO FUNCIONÁRIO NÃO TRAVAR
                    try:
                        with pagina.expect_navigation(timeout=10000):
                            pagina.locator("#lnkControleProcessos").click()
                    except:
                        try:
                            pagina.goto("https://sei.assefaz.org.br/sei/controlador.php?acao=procedimento_controlar&reset=1&infra_sistema=100000100&infra_unidade_atual=110000001")
                        except:
                            pass
                    continue
            
            app_gui.atualizar_status("\nTodos os processos possíveis foram criados!")

            # ==========================================
            # NOVO BLOCO: ADICIONAR MARCADORES
            # ==========================================
            processos_sucesso = [p for p in relatorio_processos if p['Status'] == 'SUCESSO']
            
            if processos_sucesso:
                app_gui.atualizar_status("\nIniciando a marcação dos processos criados...")
                try:
                    pagina.get_by_role("link", name="Controle de Processos", description="Controle de Processos").click()
                    pagina.wait_for_timeout(3000)
                    
                    # LISTA DE VERIFICAÇÃO: Guarda apenas os que ele realmente conseguiu clicar
                    processos_selecionados = []
                    
                    # 2. Selecionar os processos criados com segurança
                    for proc in processos_sucesso:
                        num_processo = proc['Processo']
                        try:
                            # Encontra a linha específica do processo e espera ela estar visível
                            linha = pagina.get_by_role("row", name=num_processo)
                            linha.wait_for(state="visible", timeout=4000)
                            
                            # Encontra a checkbox, força a rolagem da tela até ela e clica
                            checkbox = linha.locator("label")
                            checkbox.scroll_into_view_if_needed()
                            checkbox.click(timeout=2000)
                            pagina.wait_for_timeout(400) # Pausa rápida para o SEI computar o clique
                            
                            # Se chegou aqui sem dar erro, adiciona na lista de selecionados
                            processos_selecionados.append(proc)
                        except Exception:
                            app_gui.atualizar_status(f"  [!] Não foi possível selecionar o processo {num_processo} na tela.")
                            proc['Marcador'] = "FALHA NA SELEÇÃO"
                    
                    # Só prossegue para salvar o marcador se conseguiu selecionar pelo menos um
                    if processos_selecionados:
                        pagina.wait_for_timeout(1000)
                        
                        # 3. Clicar no botão de Adicionar Marcador
                        pagina.get_by_role("link", name="Adicionar Marcador").click()
                        pagina.wait_for_timeout(2000)

                        app_gui.atualizar_status("  ↳ Abrindo a lista de marcadores...")
                        pagina.locator(".dd-pointer").click()
                        pagina.wait_for_timeout(1500)
                        
                        primeiro_nome = usuario.split('.')[0].upper()
                        texto_marcador = rf"DOSSIÊS INATIVOS {primeiro_nome}"
                        marcador_locator = pagina.locator("a").filter(has_text=re.compile(texto_marcador, re.IGNORECASE))
                        
                        if marcador_locator.count() > 0:
                            marcador_locator.first.click()
                            pagina.wait_for_timeout(1000)
                            
                            texto_complementar = f"{numero_bloco} - {opcao_inativo}"
                            pagina.get_by_role("textbox", name="Texto:").fill(texto_complementar)
                            pagina.wait_for_timeout(500)
                            
                            pagina.get_by_role("button", name="Salvar").click()
                            pagina.wait_for_timeout(2000)
                            
                            app_gui.atualizar_status(f"✔ Marcador 'DOSSIÊS INATIVOS {primeiro_nome}' adicionado e salvo!")
                            
                            # Atualiza o relatório informando sucesso APENAS para os que foram selecionados na tela
                            for proc in processos_selecionados:
                                proc['Marcador'] = "ADICIONADO"
                        else:
                            app_gui.atualizar_status(f"  [!] Marcador não localizado para o usuário: {primeiro_nome}")
                            for proc in processos_selecionados:
                                proc['Marcador'] = "NÃO LOCALIZADO"
                            
                except Exception as erro_marcador:
                    app_gui.atualizar_status(f"  [ERRO] Falha na rotina de marcadores: {str(erro_marcador)}")
                    for proc in processos_sucesso:
                        # Se não for "FALHA NA SELEÇÃO", marca o erro genérico
                        if proc.get('Marcador') not in ["FALHA NA SELEÇÃO"]:
                            proc['Marcador'] = "FALHA NO SISTEMA"

            # ==========================================
            # NOVO BLOCO: ATRIBUIÇÃO DE PROCESSOS
            # ==========================================
            if processos_sucesso:
                app_gui.atualizar_status("\nIniciando a atribuição dos processos...")
                try:
                    pagina.get_by_role("link", name="Controle de Processos", description="Controle de Processos").click()
                    pagina.wait_for_timeout(3000)
                    
                    # LISTA DE VERIFICAÇÃO: Guarda apenas os que ele realmente conseguiu clicar
                    processos_selecionados_atribuicao = []
                    
                    for proc in processos_sucesso:
                        num_processo = proc['Processo']
                        try:
                            # Mesma lógica de rolagem da tela e verificação
                            linha = pagina.get_by_role("row", name=num_processo)
                            linha.wait_for(state="visible", timeout=4000)
                            
                            checkbox = linha.locator("label")
                            checkbox.scroll_into_view_if_needed()
                            checkbox.click(timeout=2000)
                            pagina.wait_for_timeout(400)
                            
                            processos_selecionados_atribuicao.append(proc)
                        except Exception:
                            app_gui.atualizar_status(f"  [!] Não foi possível selecionar o processo {num_processo} para atribuição.")
                            proc['Atribuição'] = "FALHA NA SELEÇÃO"
                    
                    # Só prossegue se selecionou algo
                    if processos_selecionados_atribuicao:
                        pagina.wait_for_timeout(1000)
                        
                        pagina.get_by_role("link", name="Atribuição de Processos").click()
                        pagina.wait_for_timeout(2000)
                        
                        app_gui.atualizar_status(f"  ↳ Atribuindo para o usuário logado: {usuario}")
                        
                        try:
                            pagina.get_by_label("Atribuir para:").wait_for(state="visible", timeout=5000)
                            caixa_atribuicao = pagina.get_by_label("Atribuir para:")
                            
                            caixa_atribuicao.click()
                            pagina.wait_for_timeout(1500)
                            
                            opcoes = caixa_atribuicao.locator("option").all()
                            valor_encontrado = None
                            
                            for op in opcoes:
                                texto_opcao = op.inner_text().lower()
                                if usuario.lower() in texto_opcao:
                                    valor_encontrado = op.get_attribute("value")
                                    break
                                    
                            if valor_encontrado:
                                caixa_atribuicao.select_option(valor_encontrado)
                                pagina.wait_for_timeout(1000)
                                
                                pagina.get_by_role("button", name="Salvar").click()
                                pagina.wait_for_timeout(2000)
                                
                                app_gui.atualizar_status("✔ Processos atribuídos com sucesso!")
                                
                                # Atualiza APENAS os que ele conseguiu selecionar
                                for proc in processos_selecionados_atribuicao:
                                    proc['Atribuição'] = "ATRIBUÍDO"
                            else:
                                raise Exception("Usuário não encontrado no laço de opções.")
                                
                        except Exception as e:
                            app_gui.atualizar_status(f"  [!] O usuário '{usuario}' não foi encontrado na lista de atribuição.")
                            try:
                                pagina.get_by_role("button", name="Cancelar").click(timeout=2000)
                            except:
                                pass
                                
                            for proc in processos_selecionados_atribuicao:
                                proc['Atribuição'] = "NÃO LOCALIZADO NA LISTA"
                        
                except Exception as erro_atribuicao:
                    app_gui.atualizar_status(f"  [ERRO] Falha na rotina de atribuição: {str(erro_atribuicao)}")
                    
                    for proc in processos_sucesso:
                        if proc.get('Atribuição') not in ["FALHA NA SELEÇÃO"]:
                            proc['Atribuição'] = "FALHA NO SISTEMA"         

            pagina.wait_for_timeout(1000)
            navegador.close()
            
    except Exception as e:
        app_gui.atualizar_status(f"\nERRO CRÍTICO NO ROBÔ: {str(e)}")
        
    finally:
        if 'relatorio_processos' in locals() and relatorio_processos:
            try:
                if getattr(sys, 'frozen', False):
                    dir_base = os.path.dirname(sys.executable)
                else:
                    dir_base = os.path.dirname(os.path.abspath(__file__))

                df = pd.DataFrame(relatorio_processos)
                nome_excel = f"Relatorio_Processos_{usuario}_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx"
                caminho_excel = os.path.join(dir_base, nome_excel)
                
                with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, startrow=2, sheet_name="Processos")
                    planilha = writer.sheets["Processos"]
                    planilha["A1"] = f"Relatório de Criação de Processos - Acervo de Beneficiário"
                    planilha["A2"] = f"Relatório gerado por: {usuario}"
                    
                app_gui.atualizar_status(f"\n✔ Relatório Excel salvo com segurança em:\n{caminho_excel}")
            except Exception as e_excel:
                app_gui.atualizar_status(f"\n[ERRO] Falha ao salvar o Excel: {str(e_excel)}")
                
        app_gui.finalizar_execucao()


# ---------------------------------------------------------
# 3. INTERFACE GRÁFICA
# ---------------------------------------------------------
class InterfaceRobo(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Automação SEI - Acervo de Beneficiário")
        self.geometry("500x650")
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

        # --- PASTA DOS PDFs ---
        ctk.CTkLabel(self.frame_inputs, text="Pasta dos PDFs:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.input_pasta = ctk.CTkEntry(self.frame_inputs, width=200, state="disabled")
        self.input_pasta.grid(row=2, column=1, padx=5, pady=5)
        self.btn_pasta = ctk.CTkButton(self.frame_inputs, text="📁", width=30, command=self.selecionar_pasta)
        self.btn_pasta.grid(row=2, column=2, padx=5, pady=5)

        # --- PLANILHA DE CPFs ---
        ctk.CTkLabel(self.frame_inputs, text="Planilha de CPFs:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.input_planilha_cpf = ctk.CTkEntry(self.frame_inputs, width=200, state="disabled")
        self.input_planilha_cpf.grid(row=3, column=1, padx=5, pady=5)
        self.btn_planilha = ctk.CTkButton(self.frame_inputs, text="📊", width=30, command=self.selecionar_planilha_cpf)
        self.btn_planilha.grid(row=3, column=2, padx=5, pady=5)
        
        self.caixa_logs = ctk.CTkTextbox(self, width=420, height=180)
        self.caixa_logs.pack(pady=10)
        self.caixa_logs.insert("0.0", "Preencha os dados, selecione a pasta e a planilha de CPFs, e clique em Iniciar.\n\n")
        
        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Automação", command=self.iniciar_thread)
        self.btn_iniciar.pack(pady=10)

        # --- NÚMERO DO BLOCO ---
        ctk.CTkLabel(self.frame_inputs, text="Número do Bloco:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.input_bloco = ctk.CTkEntry(self.frame_inputs, width=200)
        self.input_bloco.grid(row=4, column=1, padx=5, pady=5)
        
        # --- OPÇÕES DE INATIVO (Radio Buttons) ---
        ctk.CTkLabel(self.frame_inputs, text="Status (Inativo):").grid(row=5, column=0, padx=5, pady=5, sticky="ne")
        
        # Variável que guarda qual opção está marcada (já com valor padrão)
        self.var_inativo = ctk.StringVar(value="Inativo fora do sistema") 
        
        self.frame_radios = ctk.CTkFrame(self.frame_inputs, fg_color="transparent")
        self.frame_radios.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        
        self.rb_fora = ctk.CTkRadioButton(self.frame_radios, text="Inativo fora do sistema", variable=self.var_inativo, value="Inativo fora do sistema")
        self.rb_fora.pack(anchor="w", pady=2)
        
        self.rb_sistema = ctk.CTkRadioButton(self.frame_radios, text="Inativo do sistema", variable=self.var_inativo, value="Inativo do sistema")
        self.rb_sistema.pack(anchor="w", pady=2)

    def selecionar_pasta(self):
        pasta_selecionada = ctk.filedialog.askdirectory(title="Selecione a pasta onde estão os arquivos PDF")
        if pasta_selecionada:
            self.input_pasta.configure(state="normal")
            self.input_pasta.delete(0, "end")
            self.input_pasta.insert(0, pasta_selecionada)
            self.input_pasta.configure(state="disabled")

    def selecionar_planilha_cpf(self):
        arquivo_selecionado = ctk.filedialog.askopenfilename(
            title="Selecione a Planilha de CPFs",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if arquivo_selecionado:
            self.input_planilha_cpf.configure(state="normal")
            self.input_planilha_cpf.delete(0, "end")
            self.input_planilha_cpf.insert(0, arquivo_selecionado)
            self.input_planilha_cpf.configure(state="disabled")

    def iniciar_thread(self):
        usuario = self.input_usuario.get().strip()
        senha = self.input_senha.get().strip()
        pasta_pdfs = self.input_pasta.get().strip()
        planilha_cpfs = self.input_planilha_cpf.get().strip()
        
        # Resgata os novos valores
        numero_bloco = self.input_bloco.get().strip()
        opcao_inativo = self.var_inativo.get()
        
        # Verifica se o bloco não está em branco também
        if not usuario or not senha or not pasta_pdfs or not planilha_cpfs or not numero_bloco:
            self.atualizar_status("ERRO: Preencha todos os campos, incluindo o Número do Bloco!")
            return
            
        self.btn_iniciar.configure(state="disabled") 
        self.atualizar_status("Preparando execução...")
        
        # Passa as novas variáveis (numero_bloco, opcao_inativo) para o robô
        thread_robo = threading.Thread(target=executar_robo_sei, args=(self, usuario, senha, pasta_pdfs, planilha_cpfs, numero_bloco, opcao_inativo))
        thread_robo.start()

    def atualizar_status(self, mensagem):
        # Usa o comando .after() para mandar a mensagem para a thread principal da interface
        self.after(0, self._inserir_texto, mensagem)

    def _inserir_texto(self, mensagem):
        self.caixa_logs.insert("end", mensagem + "\n")
        self.caixa_logs.see("end")

    def finalizar_execucao(self):
        # Retorna o botão ao estado normal usando a thread principal
        self.after(0, lambda: self.btn_iniciar.configure(state="normal", text="Executar Novamente"))

if __name__ == "__main__":
    app = InterfaceRobo()
    app.mainloop()