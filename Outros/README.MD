# 🔧 Automações - Relatórios Gerais

Pasta destinada a scripts de extração genérica de dados do sistema SEI, muito úteis para auditoria e controle de desempenho da equipe.

## 📜 Scripts

*   **`Relatorio_desempenho.py`**: Um robô (`Selenium`) que acessa a pesquisa do SEI, permite filtrar por Unidade ou por um Colaborador específico (Usuário Gerador) e por um período de datas. Ele navega pelas páginas de resultados extraindo o Título, NUP, Data de Inclusão e o Usuário Gerador, compilando tudo em uma planilha de Excel (via `Pandas`) formatada e pronta para apresentação.
*   **`Relatorio_4_2.py`**: Automação focada na extração de dados *dentro* do formulário dos processos. O usuário configura (via uma interface em `Tkinter`) o tipo de processo, o formulário específico e quais campos textuais (rótulos) ele deseja extrair. O robô abre processo por processo da pesquisa, lê o documento interno e estrutura os dados encontrados em um arquivo `.xlsx` (via `OpenPyXL`).

## 🚀 Como Executar

Execute o arquivo Python desejado e preencha os dados na interface gráfica que será aberta. O `Relatorio_4_2.py` utiliza um arquivo `config_sei.json` para lembrar as suas últimas configurações, facilitando o reuso.