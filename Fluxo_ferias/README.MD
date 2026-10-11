# 🏖️ Automações - Fluxo de Férias (RH)

Esta pasta abriga os projetos criados para otimizar o fluxo de férias, reduzindo significativamente o trabalho manual no tratamento de PDFs e na criação de processos no sistema SEI.

## 📜 Scripts

*   **`Fatiador.py`**: Um utilitário construído com `PyPDF` que recebe um único arquivo PDF gigante (contendo vários recibos ou avisos de férias) e o separa em PDFs individuais. O código utiliza Expressões Regulares (`regex`) para ler o conteúdo da página e nomear automaticamente o arquivo fatiado com a Unidade (UF) e o Nome do funcionário.
*   **`Aviso_Ferias.py`**: Baseado em `Playwright`, este script lê os PDFs individuais gerados pelo fatiador, agrupa-os por Unidade, e cria automaticamente o processo base de "Aviso de Férias" no SEI para cada unidade. Ele anexa o documento de orientações, faz o *upload* nativo do PDF de cada funcionário e retorna uma planilha com o número do processo gerado.
*   **`Recibo_Ferias.py`**: Automação (também em `Playwright`) que cria os recibos de férias. O diferencial é que ele lê a planilha gerada pelo `Aviso_Ferias.py` e, para cada funcionário, não apenas cria o recibo, mas faz a **vinculação (relacionamento)** automática dentro do SEI com o processo de Aviso correspondente.

## 🚀 Como Executar

Os scripts possuem interface gráfica (`CustomTkinter`) e devem ser executados na sequência lógica do fluxo de RH (Fatiar -> Avisos -> Recibos). Ao abrir cada programa, basta preencher os dados solicitados (credenciais, Mês/Ano e os caminhos dos arquivos).