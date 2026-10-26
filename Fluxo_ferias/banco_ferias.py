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
