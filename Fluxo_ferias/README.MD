# 🏖️ Automações - Fluxo de Férias (RH)

Esta pasta abriga os projetos criados para otimizar o fluxo de férias, reduzindo significativamente o trabalho manual no tratamento de PDFs e na criação de processos no sistema SEI.

## 📜 Scripts e Módulos

*   **`extrator_inteligente.py`**: Módulo inteligente de extração que atua com **LLM Open Source Local (Ollama)**. Se o serviço do Ollama estiver ativo (ex: modelo `llama3.2:3b`), ele realiza a extração semântica dos dados da página em JSON sem depender de formatações rígidas. Se o serviço não for detectado, aciona automaticamente o extrator heurístico local de alta tolerância.
*   **`banco_ferias.py`**: Banco de dados local em SQLite (`historico_ferias.db`) que memoriza automaticamente os avisos gerados no SEI. Elimina a necessidade de o usuário guardar e selecionar planilhas antigas na etapa de recibos, além de permitir auditoria e conciliação prévia.
*   **`Fatiador.py`**: Utilitário com interface gráfica que separa o PDF compilado em arquivos individuais, agora integrado ao `extrator_inteligente` para identificação de nome e unidade.
*   **`Aviso_Ferias.py`**: Automação em `Playwright` que cria os processos de aviso no SEI por Unidade e salva os registros tanto em Excel quanto no banco local de conciliação.
*   **`Recibo_Ferias.py`**: Automação em `Playwright` que cria os processos de recibos individuais e faz a vinculação automática com o aviso correspondente, consultando o banco local automaticamente (com fallback para planilha se fornecida).
*   **`Painel_Auditoria_Ferias.py`**: Interface gráfica moderna de **Auditoria e Relatórios** para o usuário e gestão de RH. Exibe cards de métricas (total de avisos, recibos e pendências), tabela interativa com busca em tempo real por funcionário, filtros por Mês/Ano e Status, além de botão para **exportação de relatórios em Excel (.xlsx) com layout profissional** e estilização visual.

## 🚀 Como Utilizar

1. **Modo Tradicional (100% Offline e sem dependências extras):** Basta executar os scripts normalmente (`Fatiador.py` -> `Aviso_Ferias.py` -> `Recibo_Ferias.py`). O sistema utiliza o extrator de regras e o banco de dados SQLite local.
2. **Modo com IA Open Source Local (Ollama):** Para habilitar a inteligência semântica:
   * Instale o [Ollama](https://ollama.com).
   * No terminal, execute uma única vez: `ollama run llama3.2:3b`
   * Pronto! O `Fatiador.py` detectará a IA local automaticamente e usará a LLM local para a extração dos PDFs com privacidade total.