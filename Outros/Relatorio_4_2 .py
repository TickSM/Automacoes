import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import json
import re
import threading
import time


# ───────────── FUNÇÕES DE APOIO ─────────────
def limpar_nome_arquivo(nome):
    return re.sub(r'[\\/*?:"<>|]', "-", nome)

driver = None


def criar_janela_status(root):
    status_win = tk.Toplevel(root)
    status_win.title("Status da Automação")
    status_win.geometry("350x150")
    status_win.attributes("-topmost", True)

    def confirmar_interrupcao():
        if messagebox.askyesno("Interromper", "Deseja realmente interromper?", parent=status_win):
            try:
                if driver:
                    # Tenta fechar de forma amigável primeiro
                    driver.quit()
            except:
                pass
            finally:
                # Mata o processo do Python e interrompe tudo
                os._exit(0)

    status_win.protocol("WM_DELETE_WINDOW", confirmar_interrupcao)

    label_msg = ttk.Label(status_win, text="Iniciando...", padding=10, wraplength=300, font=("Arial", 10, "bold"))
    label_msg.pack()

    label_detalhe = ttk.Label(status_win, text="Preparando ambiente...", padding=5)
    label_detalhe.pack()

    progress = ttk.Progressbar(status_win, mode='indeterminate')
    progress.pack(fill='x', padx=20, pady=10)
    progress.start()

    return status_win, label_msg, label_detalhe


# ───────────── INTERFACE DE CONFIGURAÇÃO ─────────────
def coletar_dados(root):
    dados = {}
    CONFIG_FILE = "config_sei.json"

    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config_anterior = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        config_anterior = {}

    usuario_var = tk.StringVar(value=config_anterior.get('usuario', ''))
    senha_var = tk.StringVar(value='')
    unidade_var = tk.StringVar(value=config_anterior.get('unidade_geradora', ''))
    tipo_proc_var = tk.StringVar(value=config_anterior.get('tipo_processo', ''))
    tipo_doc_var = tk.StringVar(value=config_anterior.get('tipo_documento', ''))
    form_var = tk.StringVar(value=config_anterior.get('nome_formulario', ''))
    campos_str_var = tk.StringVar(value=", ".join(config_anterior.get('lista_campos_excel', [])) or "TODOS")
    data_inicio_var = tk.StringVar(value=config_anterior.get('data_inicio', ''))
    data_final_var = tk.StringVar(value=config_anterior.get('data_final', ''))
    tramitacao_var = tk.BooleanVar(value=config_anterior.get('tramitacao', False))
    caminho_var = tk.StringVar(
        value=config_anterior.get('caminho_salvamento', os.path.join(os.path.expanduser("~"), "Desktop")))

    ok_clicado = tk.BooleanVar(value=False)
    dialogo = tk.Toplevel(root)
    dialogo.title("Configurar Automação - Versão 4.1")
    dialogo.grab_set()

    frame = ttk.Frame(dialogo, padding="15")
    frame.pack(expand=True, fill="both")

    def add_field(label, var, is_pass=False):
        row = frame.grid_size()[1]
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=var, width=40, show="*" if is_pass else "").grid(row=row, column=1, pady=3,
                                                                                       padx=5, sticky="w")

    add_field("Usuário SEI:", usuario_var)
    add_field("Senha SEI:", senha_var, True)
    add_field("Unidade Geradora (Opcional):", unidade_var)
    add_field("Tipo de Processo:", tipo_proc_var)
    add_field("Tipo do Documento:", tipo_doc_var)
    add_field("Nome do Formulário:", form_var)
    add_field("Campos (ou 'TODOS'):", campos_str_var)
    add_field("Data Início (DD/MM/AAAA):", data_inicio_var)
    add_field("Data Final (DD/MM/AAAA):", data_final_var)

    # Layout do Checkbox e Salvamento
    ttk.Checkbutton(frame, text="Pesquisar com Tramitação na Unidade", variable=tramitacao_var).grid(row=10, column=0,
                                                                                                     columnspan=2,
                                                                                                     sticky="w", pady=5)

    ttk.Label(frame, text="Salvar em:").grid(row=11, column=0, sticky="w", pady=3)
    caminho_frame = ttk.Frame(frame)
    caminho_frame.grid(row=11, column=1, sticky="w", pady=3)
    ttk.Entry(caminho_frame, textvariable=caminho_var, width=33).pack(side="left", padx=(5, 2))

    def selecionar_pasta():
        pasta = filedialog.askdirectory()
        if pasta: caminho_var.set(pasta)

    ttk.Button(caminho_frame, text="...", width=3, command=selecionar_pasta).pack(side="left")

    def on_ok():
        # Dicionário de campos obrigatórios para validação
        obrigatorios = {
            "Usuário SEI": usuario_var.get(),
            "Senha SEI": senha_var.get(),
            "Tipo de Processo": tipo_proc_var.get(),
            "Tipo do Documento": tipo_doc_var.get(),
            "Nome do Formulário": form_var.get(),
            "Campos": campos_str_var.get(),
            "Data Início": data_inicio_var.get(),
            "Data Final": data_final_var.get()
        }

        # Verifica se algum campo obrigatório está vazio ou apenas com espaços
        vazios = [campo for campo, valor in obrigatorios.items() if not valor.strip()]

        if vazios:
            messagebox.showwarning("Campos Obrigatórios",
                                   f"Os seguintes campos devem ser preenchidos:\n\n- " + "\n- ".join(vazios),
                                   parent=dialogo)
            return

        # Se passar na validação, atualiza o dicionário e fecha a janela
        dados.update({
            'usuario': usuario_var.get(), 'senha': senha_var.get(), 'unidade_geradora': unidade_var.get(),
            'tipo_processo': tipo_proc_var.get(), 'tipo_documento': tipo_doc_var.get(),
            'nome_formulario': form_var.get(), 'data_inicio': data_inicio_var.get(),
            'data_final': data_final_var.get(), 'tramitacao': tramitacao_var.get(),
            'caminho_salvamento': caminho_var.get(),
            'lista_campos_excel': [c.strip() for c in campos_str_var.get().split(',') if c.strip()]
        })

        # Salva as configurações para a próxima vez (sem a senha)
        config_save = dados.copy()
        del config_save['senha']
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_save, f, indent=4, ensure_ascii=False)

        ok_clicado.set(True)
        dialogo.destroy()

    ttk.Button(frame, text="Iniciar Automação Oculta", command=on_ok).grid(row=13, column=0, columnspan=2, pady=15)
    root.wait_window(dialogo)
    if not ok_clicado.get(): os._exit(0)
    return dados


# ───────────── LÓGICA CORE (SELENIUM) ─────────────
def executar_automacao(dados, lb_msg, lb_detalhe, win_status):
    global driver
    try:
        lb_msg.config(text="Iniciando Navegador...")

        chrome_options = Options()
        #chrome_options.add_argument("--headless")  # Modo oculto ativo para o executável
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-gpu")

        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 15)

        lb_msg.config(text="Realizando Login no SEI...")
        driver.get("https://sei.assefaz.org.br/sip/login.php?sigla_orgao_sistema=ASSEFAZ&sigla_sistema=SEI")

        driver.find_element(By.ID, "txtUsuario").send_keys(dados['usuario'])
        driver.find_element(By.ID, "pwdSenha").send_keys(dados['senha'])
        driver.find_element(By.ID, "Acessar").click()

        driver.find_element(By.ID, "txtPesquisaRapida").send_keys(Keys.RETURN)

        lb_msg.config(text="Aplicando Filtros de Pesquisa...")
        if dados['unidade_geradora']:
            campo_u = driver.find_element(By.ID, "txtUnidade")
            campo_u.send_keys(dados['unidade_geradora'])
            time.sleep(1)
            campo_u.send_keys(Keys.ARROW_DOWN, Keys.ENTER)

        Select(driver.find_element(By.ID, "selTipoProcedimentoPesquisa")).select_by_visible_text(dados['tipo_processo'])
        Select(driver.find_element(By.ID, "selSeriePesquisa")).select_by_visible_text(dados['tipo_documento'])
        driver.find_element(By.ID, "txtDataInicio").send_keys(dados['data_inicio'])
        driver.find_element(By.ID, "txtDataFim").send_keys(dados['data_final'])

        # ID Corrigido para chkSinTramitacao
        if dados.get('tramitacao'):
            try:
                chk_tramitacao = driver.find_element(By.ID, "chkSinTramitacao")
                if not chk_tramitacao.is_selected():
                    driver.execute_script("arguments[0].click();", chk_tramitacao)
            except:
                pass

        driver.find_element(By.ID, 'sbmPesquisar').click()

        processos_detalhes = []
        contador = 0
        extrair_tudo = "TODOS" in [c.upper() for c in dados['lista_campos_excel']]

        while True:
            tabela = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "pesquisaResultado")))
            linhas = tabela.find_elements(By.XPATH, './/tr[@class="pesquisaTituloRegistro"]')
            aba_original = driver.current_window_handle

            for linha in linhas:
                contador += 1

                link = None
                num_proc = "-"
                eh_processo_valido = False

                # --- CORREÇÃO 1: GARANTIR A EXTRAÇÃO DO NUP ---
                links_da_linha = linha.find_elements(By.TAG_NAME, "a")

                for l in links_da_linha:
                    href = l.get_attribute("href") or ""
                    texto_link = l.get_attribute("textContent").strip()

                    # Exige que o link contenha o padrão do processo (ex: Assefaz.017150/2026-08)
                    match_nup = re.search(r'[\w\.]*\d{4,6}/\d{4}-\d{2}', texto_link)

                    if match_nup and "procedimento_trabalhar" in href:
                        link = l
                        num_proc = match_nup.group(0)  # Salva estritamente o texto do NUP
                        eh_processo_valido = True
                        break

                if not eh_processo_valido:
                    try:
                        link_generico = linha.find_element(By.XPATH, './/a[contains(@class, "protocoloNormal")]')
                        texto_gen = link_generico.get_attribute("textContent").strip()
                        num_proc = texto_gen if texto_gen else f"Item {contador} (Não identificado)"
                    except Exception:
                        num_proc = f"Item {contador} (Sem link)"

                try:
                    meta_tr = linha.find_element(By.XPATH, "./following-sibling::tr[2]")
                    usuario_nome = meta_tr.find_element(By.XPATH,
                                                        ".//b[contains(text(), 'Usuário:')]/following-sibling::a").get_attribute(
                        "title")
                    data_inc = meta_tr.find_element(By.XPATH, ".//td[b[contains(text(), 'Inclusão:')]]").get_attribute(
                        "textContent").split(":")[-1].strip().replace('"', '')
                except Exception:
                    usuario_nome = "Não extraído"
                    data_inc = "Não extraído"

                if not eh_processo_valido:
                    linha_erro = [num_proc, usuario_nome, data_inc]
                    for _ in dados['lista_campos_excel']:
                        linha_erro.append("NÃO FOI POSSÍVEL IDENTIFICAR COMO PROCESSO")
                    processos_detalhes.append(linha_erro)
                    continue

                lb_msg.config(text=f"Processando: {num_proc}")

                driver.execute_script("window.open(arguments[0].href, '_blank');", link)
                driver.switch_to.window(driver.window_handles[-1])

                try:
                    wait.until(EC.frame_to_be_available_and_switch_to_it("ifrArvore"))

                    while True:
                        pastas = driver.find_elements(By.XPATH, "//img[starts-with(@title, 'Abrir Pasta')]")
                        if not pastas: break
                        for p in pastas:
                            try:
                                driver.execute_script("arguments[0].click();", p)
                                try:
                                    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located(
                                        (By.XPATH, "//span[contains(@id, 'spanAGUARDE')]")))
                                except:
                                    pass
                            except:
                                pass
                        time.sleep(0.3)

                    xpath_f = f"//a[contains(@class, 'infraArvoreNo') and not(contains(@title, 'Documento Cancelado')) and .//span[contains(text(), '{dados['nome_formulario']}')]]"
                    links_f = driver.find_elements(By.XPATH, xpath_f)

                    if links_f:
                        ids_vistos = []
                        f_count = 0
                        for l_form in links_f:
                            id_l = l_form.get_attribute("id")
                            if id_l in ids_vistos: continue
                            ids_vistos.append(id_l)

                            dados_f = {}
                            f_count += 1
                            lb_detalhe.config(text=f"Doc {f_count} de {len(links_f)} no processo {num_proc}")

                            try:
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", l_form)
                                driver.execute_script("arguments[0].click();", l_form)

                                driver.switch_to.default_content()
                                wait.until(EC.frame_to_be_available_and_switch_to_it("ifrVisualizacao"))
                                wait.until(EC.frame_to_be_available_and_switch_to_it("ifrArvoreHtml"))

                                time.sleep(0.5)

                                elemento_conteudo = wait.until(EC.presence_of_element_located((By.ID, "conteudo")))
                                texto_bruto = elemento_conteudo.get_attribute("innerText")

                                # --- CORREÇÃO 2: MANTER LINHAS EM BRANCO E MELHORAR IDENTIFICAÇÃO DE RÓTULOS ---
                                # Removido o 'if l.strip()' para não destruir a estrutura visual das observações
                                linhas_txt = [l.strip() for l in texto_bruto.splitlines()]

                                if extrair_tudo and (
                                        not dados['lista_campos_excel'] or dados['lista_campos_excel'] == ["TODOS"]):
                                    campos_detectados = []
                                    for l in linhas_txt:
                                        # Detecta campos dinamicamente apenas se for curto e não estiver tudo em maiúsculo
                                        if l.endswith(":") and len(l) <= 45 and not l.isupper():
                                            campos_detectados.append(l)
                                    if campos_detectados:
                                        dados['lista_campos_excel'] = campos_detectados

                                for i, txt in enumerate(linhas_txt):
                                    if txt in dados['lista_campos_excel']:
                                        valores = []
                                        j = i + 1
                                        while j < len(linhas_txt):
                                            ln = linhas_txt[j]

                                            # Para a extração APENAS se bater com um campo conhecido...
                                            if ln in dados['lista_campos_excel']:
                                                break

                                            # ...OU se tiver formato exato de rótulo (curto, não-maiúsculo, com dois pontos).
                                            # Isso impede de parar em observações como "SETEMBRO:" ou "PAGTº:"
                                            if ln.endswith(":") and len(
                                                    ln) <= 45 and not ln.isupper() and not ln.startswith(
                                                    ("1 -", "2 -", "3 -", "4 -")):
                                                break

                                            valores.append(ln)
                                            j += 1

                                        dados_f[txt] = "\n".join(valores).strip() if valores else "-"

                                linha_final = [num_proc, usuario_nome, data_inc]
                                for campo in dados['lista_campos_excel']:
                                    linha_final.append(dados_f.get(campo, "-"))
                                processos_detalhes.append(linha_final)

                            except Exception as e:
                                print(f"Erro no processo {num_proc}: {e}")
                                linha_erro = [num_proc, usuario_nome, data_inc]
                                for _ in dados['lista_campos_excel']:
                                    linha_erro.append(f"ERRO NA EXTRAÇÃO ({type(e).__name__})")
                                processos_detalhes.append(linha_erro)
                                driver.switch_to.default_content()
                                try:
                                    wait.until(EC.frame_to_be_available_and_switch_to_it("ifrArvore"))
                                except:
                                    pass
                                continue

                            driver.switch_to.default_content()
                            wait.until(EC.frame_to_be_available_and_switch_to_it("ifrArvore"))
                    else:
                        processos_detalhes.append(
                            [num_proc, usuario_nome, data_inc] + ["NÃO ENCONTRADO"] * len(dados['lista_campos_excel']))

                finally:
                    driver.close()
                    driver.switch_to.window(aba_original)

            try:
                btn_p = driver.find_element(By.LINK_TEXT, "Próxima")
                driver.execute_script("arguments[0].click();", btn_p);
                time.sleep(1.5)
            except:
                break

        lb_msg.config(text="Gerando relatório em excel...")
        # Geração do Excel
        wb = Workbook()
        ws = wb.active
        ws.merge_cells('A1:G1');
        ws['A1'] = "Relatório de Extração SEI"
        ws['A1'].font = Font(bold=True, size=14);
        ws['A1'].alignment = Alignment(horizontal="center")
        ws['A3'], ws['B3'] = "Tipo de Processo:", dados['tipo_processo']
        ws['A4'], ws['B4'] = "Tipo de Documento:", dados['tipo_documento']
        ws['A5'], ws['B5'] = "Formulário:", dados['nome_formulario']
        ws['A6'], ws['B6'] = "Período:", f"{dados['data_inicio']} - {dados['data_final']}"
        for r in range(3, 7): ws[f'A{r}'].font = Font(bold=True)

        headers = ["NUP SEI", "USUÁRIO DO RESULTADO DA PESQUISA", "DATA INCLUSÃO"] + dados['lista_campos_excel']
        for c_idx, h_text in enumerate(headers, 1):
            cell = ws.cell(row=8, column=c_idx, value=h_text)
            cell.font = Font(bold=True)

        for r_idx, r_data in enumerate(processos_detalhes, 9):
            for c_idx, val in enumerate(r_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = 40

        nome_f = limpar_nome_arquivo(dados['nome_formulario'])
        caminho_final = os.path.join(dados['caminho_salvamento'],
                                     f"Relatorio_{nome_f}_{datetime.now().strftime('%H-%M')}.xlsx")
        wb.save(caminho_final)
        messagebox.showinfo("Sucesso", f"Relatório gerado!\n{caminho_final}")


    except Exception as e:

        import traceback

        erro_completo = traceback.format_exc()

        print(erro_completo)

        messagebox.showerror(

            "Erro",

            f"{type(e).__name__}\n\n{erro_completo}"

        )
    finally:
        if driver: driver.quit()
        win_status.destroy()


if __name__ == "__main__":
    root = tk.Tk();
    root.withdraw()
    config = coletar_dados(root)
    win_status, lb_msg, lb_detalhe = criar_janela_status(root)
    threading.Thread(target=executar_automacao, args=(config, lb_msg, lb_detalhe, win_status)).start()
    root.mainloop()