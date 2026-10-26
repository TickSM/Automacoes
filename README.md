# 🤖 Automações - Portfólio de Projetos

Bem-vindo ao meu repositório central de automações! Aqui você encontrará uma coleção de scripts e projetos desenvolvidos em Python focados em otimizar rotinas administrativas, reduzir o tempo de execução de tarefas repetitivas e aumentar a precisão na extração e inserção de dados em sistemas web.

A maioria dos projetos utiliza a biblioteca **Selenium** ou **Playwright** para a automação web e **OpenPyXL** / **Pandas** para a manipulação e geração de relatórios estruturados em Excel.

## 📂 Estrutura do Repositório

O repositório está organizado nas seguintes pastas, cada uma focada em uma frente de automação específica:

*   [**506**](./506): Scripts focados no projeto de digitalização (Projeto 506), incluindo a criação automatizada de processos no sistema SEI e a extração e consolidação de dados de acervos arquivísticos.
*   [**Fluxo_ferias**](./Fluxo_ferias): Automações para o setor de RH com suporte a **IA Open Source Local (Ollama)** para parsing inteligente de PDFs e banco de dados SQLite para **conciliação e vinculação automática** entre avisos e recibos no SEI sem necessidade de planilhas manuais.
*   [**Outros**](./Outros): Contém ferramentas gerais para geração de relatórios de desempenho e relatórios de extração de dados do SEI (Versão 4.2).
*   [**Regionais**](./Regionais): Script para acompanhamento e extração de dados da digitalização na Gerência Regional (GR), com suporte a **execução em lote de múltiplas UFs em sequência** (RJ, SP, BA e personalizadas).

## 🛠️ Tecnologias Utilizadas

*   **Linguagem:** Python
*   **Inteligência Artificial (LLM Local):** Ollama (Llama 3.2 / Qwen 2.5) com fallback resiliente para extração determinística
*   **Banco de Dados:** SQLite (persistência de processos e conciliação)
*   **Interface Gráfica (GUI):** Tkinter e CustomTkinter
*   **Automação Web:** Selenium WebDriver e Playwright
*   **Manipulação de Dados e Planilhas:** Pandas e OpenPyXL
*   **Manipulação de PDF:** PyPDF

---
*Este é um repositório pessoal utilizado para documentar e organizar os meus códigos de automação.*