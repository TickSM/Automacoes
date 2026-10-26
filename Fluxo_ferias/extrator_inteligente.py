import os
import re
import json
import requests
from pypdf import PdfReader

# Configurações do Ollama Local
OLLAMA_URL = os.getenv("OLLAMA_HOST", "http://localhost:11434/api/generate")
MODELO_PADRAO = "llama3.2:3b"

def verificar_ia_local():
    """Verifica se o serviço do Ollama está ativo na máquina ou rede interna."""
    try:
        resposta = requests.get("http://localhost:11434", timeout=1.5)
        return resposta.status_code == 200
    except Exception:
        return False

def extrair_dados_com_ia_local(texto_pagina, tipo_esperado="AUTO"):
    """
    Envia o texto da página para a LLM Open Source local (Ollama)
    para extração estrita em formato JSON.
    """
    prompt = f"""Você é um extrator de dados de documentos de RH (Avisos e Recibos de Férias).
Analise o texto abaixo e responda APENAS com um objeto JSON válido, sem comentários ou texto adicional.

Campos obrigatórios no JSON:
- "tipo_documento": "AVISO" ou "RECIBO"
- "nome_funcionario": Nome completo em letras maiúsculas
- "unidade_uf": Sigla do estado (ex: RJ, SP, BA, DF) ou "" se não identificada

Texto do documento:
\"\"\"
{texto_pagina}
\"\"\"
"""
    try:
        payload = {
            "model": MODELO_PADRAO,
            "prompt": prompt,
            "stream": False,
            "format": "json"
        }
        res = requests.post(OLLAMA_URL, json=payload, timeout=12)
        if res.status_code == 200:
            resultado = res.json()
            resposta_texto = resultado.get("response", "{}")
            dados = json.loads(resposta_texto)
            return {
                "tipo": dados.get("tipo_documento", "DESCONHECIDO").upper(),
                "nome": dados.get("nome_funcionario", "").strip().upper(),
                "unidade": dados.get("unidade_uf", "").strip().upper()
            }
    except Exception:
        pass
    return None

def extrair_dados_por_regras(texto_pagina, unidade_padrao=""):
    """
    Fallback seguro: Algoritmo heurístico local baseado em regex e palavras-chave.
    Garante que a aplicação nunca pare, mesmo sem a IA instalada.
    """
    texto_limpo = texto_pagina.replace('\n', ' ').replace('\r', ' ')
    
    nome_cru = ""
    unidade_cru = ""
    tipo_doc = "DESCONHECIDO"

    # 1. É um Recibo de Férias?
    if "EMPREGADO" in texto_limpo.upper():
        tipo_doc = "RECIBO"
        match_nome = re.search(r'Empregado[^\:]*:\s*([^|]+)', texto_limpo, re.IGNORECASE)
        if match_nome:
            nome_cru = match_nome.group(1).strip()
        unidade_cru = ""

    # 2. É um Aviso de Férias?
    elif "SR(A)" in texto_limpo.upper() or "AVISO" in texto_limpo.upper():
        tipo_doc = "AVISO"
        match_nome = re.search(r'([A-ZÀ-Ú][A-ZÀ-Ú \.]{8,})(?=[\d\W]*CTPS)', texto_limpo)
        if match_nome:
            nome_cru = match_nome.group(1).strip()
        
        match_depto = re.search(r'DEPTO[\W_]*([A-ZÀ-Úa-zà-ú \-]+)', texto_limpo, re.IGNORECASE)
        if match_depto:
            unidade_cru = match_depto.group(1).strip()
            unidade_cru = re.sub(r'(?i)\s*Comunicamos.*', '', unidade_cru).strip()

    # Tratamento final
    if not unidade_cru and unidade_padrao:
        unidade_cru = unidade_padrao

    return {
        "tipo": tipo_doc,
        "nome": nome_cru.strip().upper(),
        "unidade": unidade_cru.strip().upper()
    }

def processar_pagina_pdf(texto_pagina, unidade_padrao="", usa_ia=True):
    """
    Função principal que orquestra a extração:
    Tenta IA local se disponível, caso contrário recorre ao extrator determinístico.
    """
    if usa_ia and verificar_ia_local():
        dados_ia = extrair_dados_com_ia_local(texto_pagina)
        if dados_ia and dados_ia.get("nome"):
            if not dados_ia.get("unidade") and unidade_padrao:
                dados_ia["unidade"] = unidade_padrao.upper()
            return dados_ia, "IA_LOCAL (Ollama)"

    dados_regras = extrair_dados_por_regras(texto_pagina, unidade_padrao)
    return dados_regras, "EXTRATOR_LOCAL (Regras)"
