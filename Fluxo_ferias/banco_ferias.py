import os
import sqlite3
from datetime import datetime

DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
CAMINHO_BANCO = os.path.join(DIRETORIO_ATUAL, "historico_ferias.db")

def inicializar_banco():
    """Garante a criação das tabelas necessárias para persistência das férias."""
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        
        # Tabela de Avisos de Férias
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS avisos_ferias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                mes_ano TEXT NOT NULL,
                unidade TEXT NOT NULL,
                numero_processo TEXT NOT NULL,
                interessados TEXT NOT NULL,
                data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Tabela de Recibos de Férias
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS recibos_ferias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                mes_ano TEXT NOT NULL,
                nome_funcionario TEXT NOT NULL,
                numero_processo_recibo TEXT NOT NULL,
                processo_aviso_vinculado TEXT,
                status_vinculo TEXT,
                data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

def registrar_aviso(mes_ano, unidade, numero_processo, lista_nomes):
    """Registra um processo de aviso de férias e os funcionários vinculados."""
    inicializar_banco()
    interessados_str = " / ".join(lista_nomes)
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO avisos_ferias (mes_ano, unidade, numero_processo, interessados)
            VALUES (?, ?, ?, ?)
        """, (mes_ano.upper(), unidade.upper(), numero_processo, interessados_str.upper()))
        conn.commit()

def registrar_recibo(mes_ano, nome_funcionario, numero_processo_recibo, processo_aviso_vinculado, status_vinculo):
    """Registra um processo de recibo de férias e o status de sua vinculação."""
    inicializar_banco()
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO recibos_ferias (mes_ano, nome_funcionario, numero_processo_recibo, processo_aviso_vinculado, status_vinculo)
            VALUES (?, ?, ?, ?, ?)
        """, (mes_ano.upper(), nome_funcionario.upper(), numero_processo_recibo, processo_aviso_vinculado, status_vinculo))
        conn.commit()

def buscar_aviso_por_funcionario(nome_funcionario, mes_ano=None):
    """
    Localiza o número do processo de aviso correspondente pelo nome do funcionário.
    Elimina a necessidade do usuário leigo ter que procurar planilhas antigas.
    """
    inicializar_banco()
    nome_alvo = nome_funcionario.strip().upper()
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        
        if mes_ano:
            cursor.execute("""
                SELECT numero_processo, interessados FROM avisos_ferias
                WHERE mes_ano = ?
                ORDER BY id DESC
            """, (mes_ano.upper(),))
        else:
            cursor.execute("""
                SELECT numero_processo, interessados FROM avisos_ferias
                ORDER BY id DESC
            """)
            
        linhas = cursor.fetchall()
        for num_processo, interessados in linhas:
            if nome_alvo in interessados:
                return num_processo
    return None

def auditar_conciliacao(mes_ano):
    """
    Auditor inteligente: Cruza quem recebeu aviso vs quem já recebeu recibo.
    Retorna estatísticas e lista de pendências.
    """
    inicializar_banco()
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        
        # Coleta todos os funcionários de avisos no mês
        cursor.execute("SELECT interessados FROM avisos_ferias WHERE mes_ano = ?", (mes_ano.upper(),))
        avisos = cursor.fetchall()
        funcionarios_com_aviso = set()
        for (interessados,) in avisos:
            for nome in interessados.split(" / "):
                if nome.strip():
                    funcionarios_com_aviso.add(nome.strip().upper())
                    
        # Coleta recibos já gerados
        cursor.execute("SELECT nome_funcionario FROM recibos_ferias WHERE mes_ano = ?", (mes_ano.upper(),))
        recibos = cursor.fetchall()
        funcionarios_com_recibo = {r[0].strip().upper() for r in recibos}
        
        pendentes_recibo = list(funcionarios_com_aviso - funcionarios_com_recibo)
        
        return {
            "total_avisos": len(funcionarios_com_aviso),
            "total_recibos": len(funcionarios_com_recibo),
            "pendentes": pendentes_recibo
        }

def obter_meses_cadastrados():
    """Retorna a lista de meses/anos com registros no banco de dados."""
    inicializar_banco()
    meses = set()
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT mes_ano FROM avisos_ferias")
        for (m,) in cursor.fetchall():
            if m:
                meses.add(m.strip().upper())
        cursor.execute("SELECT DISTINCT mes_ano FROM recibos_ferias")
        for (m,) in cursor.fetchall():
            if m:
                meses.add(m.strip().upper())
    return sorted(list(meses), reverse=True)

def consultar_historico_consolidado(mes_ano=None, busca_nome=None, filtro_status=None):
    """
    Cruza avisos e recibos gerando um relatório consolidado completo por funcionário.
    """
    inicializar_banco()
    with sqlite3.connect(CAMINHO_BANCO) as conn:
        cursor = conn.cursor()
        
        # 1. Carrega todos os avisos
        query_avisos = "SELECT mes_ano, unidade, numero_processo, interessados, data_criacao FROM avisos_ferias"
        params_avisos = []
        if mes_ano and mes_ano != "TODOS":
            query_avisos += " WHERE mes_ano = ?"
            params_avisos.append(mes_ano.upper())
        cursor.execute(query_avisos, params_avisos)
        linhas_avisos = cursor.fetchall()
        
        # 2. Carrega todos os recibos
        query_recibos = "SELECT mes_ano, nome_funcionario, numero_processo_recibo, processo_aviso_vinculado, status_vinculo, data_criacao FROM recibos_ferias"
        params_recibos = []
        if mes_ano and mes_ano != "TODOS":
            query_recibos += " WHERE mes_ano = ?"
            params_recibos.append(mes_ano.upper())
        cursor.execute(query_recibos, params_recibos)
        linhas_recibos = cursor.fetchall()

    # Mapeamento de recibos por (nome, mes_ano) e por nome
    mapa_recibos = {}
    recibos_usados = set()
    for r in linhas_recibos:
        r_mes, r_nome, r_proc, r_aviso_vinc, r_status, r_data = r
        chave = (r_nome.strip().upper(), r_mes.strip().upper())
        mapa_recibos[chave] = {
            "processo_recibo": r_proc,
            "aviso_vinculado": r_aviso_vinc,
            "status_vinculo": r_status,
            "data_recibo": r_data[:10] if r_data else "-"
        }

    registros_consolidados = []

    # 3. Itera avisos explodindo por interessado
    for a_mes, a_und, a_proc, a_interessados, a_data in linhas_avisos:
        data_aviso_formatada = a_data[:10] if a_data else "-"
        for nome_func in a_interessados.split(" / "):
            nome_limpo = nome_func.strip().upper()
            if not nome_limpo:
                continue

            chave = (nome_limpo, a_mes.strip().upper())
            recibo_info = mapa_recibos.get(chave)
            
            if recibo_info:
                recibos_usados.add(chave)
                proc_recibo = recibo_info["processo_recibo"]
                status = "CONCLUÍDO (VINCULADO)"
                data_recibo = recibo_info["data_recibo"]
            else:
                proc_recibo = "PENDENTE"
                status = "RECIBO PENDENTE"
                data_recibo = "-"

            registros_consolidados.append({
                "funcionario": nome_limpo,
                "unidade": a_und,
                "mes_ano": a_mes,
                "processo_aviso": a_proc,
                "processo_recibo": proc_recibo,
                "status": status,
                "data_aviso": data_aviso_formatada,
                "data_recibo": data_recibo
            })

    # 4. Adiciona recibos avulsos (que não tiveram aviso cadastrado)
    for r in linhas_recibos:
        r_mes, r_nome, r_proc, r_aviso_vinc, r_status, r_data = r
        chave = (r_nome.strip().upper(), r_mes.strip().upper())
        if chave not in recibos_usados:
            registros_consolidados.append({
                "funcionario": r_nome.strip().upper(),
                "unidade": "AVULSO",
                "mes_ano": r_mes.strip().upper(),
                "processo_aviso": r_aviso_vinc or "NÃO LOCALIZADO",
                "processo_recibo": r_proc,
                "status": "RECIBO AVULSO",
                "data_aviso": "-",
                "data_recibo": r_data[:10] if r_data else "-"
            })

    # 5. Aplica filtros de pesquisa e status
    resultado_filtrado = []
    for reg in registros_consolidados:
        if busca_nome:
            if busca_nome.strip().upper() not in reg["funcionario"]:
                continue
        if filtro_status and filtro_status != "TODOS":
            if filtro_status.upper() not in reg["status"].upper():
                continue
        resultado_filtrado.append(reg)

    return sorted(resultado_filtrado, key=lambda x: (x["mes_ano"], x["unidade"], x["funcionario"]))

def exportar_relatorio_excel(caminho_saida, mes_ano=None, busca_nome=None, filtro_status=None):
    """
    Gera uma planilha Excel estilizada e profissional pronta para envio ou arquivo do RH.
    """
    import pandas as pd
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    dados = consultar_historico_consolidado(mes_ano, busca_nome, filtro_status)
    if not dados:
        return False, "Nenhum registro encontrado para exportar."

    df = pd.DataFrame([{
        "Funcionário": d["funcionario"],
        "Unidade (UF)": d["unidade"],
        "Mês/Ano": d["mes_ano"],
        "Processo de Aviso (SEI)": d["processo_aviso"],
        "Data Aviso": d["data_aviso"],
        "Processo de Recibo (SEI)": d["processo_recibo"],
        "Data Recibo": d["data_recibo"],
        "Status da Conciliação": d["status"]
    } for d in dados])

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=3, sheet_name="Auditoria Férias")
        ws = writer.sheets["Auditoria Férias"]

        # Título do Relatório
        ws.merge_cells("A1:H1")
        ws["A1"] = "RELATÓRIO CONSOLIDADO DE AUDITORIA - FLUXO DE FÉRIAS (SEI)"
        ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["A1"].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

        # Subtítulo com Data de Emissão e Filtro
        ws.merge_cells("A2:H2")
        ws["A2"] = f"Referência: {mes_ano or 'Todos os Meses'} | Emitido em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}"
        ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="555555")
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        # Estilo do Cabeçalho da Tabela
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        for col_idx in range(1, 9):
            cell = ws.cell(row=4, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Bordas e Destaque de Status
        borda_fina = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        fill_concluido = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        font_concluido = Font(color="276A3C", bold=True)

        fill_pendente = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        font_pendente = Font(color="B25900", bold=True)

        fill_avulso = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        font_avulso = Font(color="9C0006", bold=True)

        for row in range(5, 5 + len(df)):
            for col in range(1, 9):
                c = ws.cell(row=row, column=col)
                c.border = borda_fina
                if col in (2, 3, 5, 7):
                    c.alignment = Alignment(horizontal="center")
                
                # Destaca a coluna de Status
                if col == 8:
                    val = str(c.value)
                    if "CONCLUÍDO" in val:
                        c.fill = fill_concluido
                        c.font = font_concluido
                    elif "PENDENTE" in val:
                        c.fill = fill_pendente
                        c.font = font_pendente
                    else:
                        c.fill = fill_avulso
                        c.font = font_avulso

        # Ajuste de largura das colunas
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    return True, f"Relatório exportado com sucesso para:\n{caminho_saida}"
