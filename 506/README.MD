# 📁 Automações - Projeto 506 (Digitalização)

Esta pasta contém os scripts desenvolvidos para apoiar as rotinas da etapa do Projeto de Digitalização 506. O foco principal é a integração com o sistema SEI para a criação em lote de processos e a extração consolidada de dados dos acervos.

## 📜 Scripts

*   **`ADB_Criar_Processo.py`**: Automatiza a criação de processos no SEI para os Dossiês de Beneficiários (ADB). O script lê uma pasta contendo PDFs e uma planilha de CPFs, realiza duas etapas de pesquisa no SEI para evitar duplicidade, cria o processo base, preenche o formulário interno com os dados (como Módulo, Lado, Prateleira, etc.), anexa o PDF externo correspondente, e finaliza adicionando os marcadores e atribuindo o processo ao usuário logado. Tudo isso gerando um relatório em Excel com o status de cada tentativa.
*   **`Digitalizacao_506.py`**: Robô que acessa o SEI, realiza uma pesquisa avançada baseada em filtros de unidade e data, e extrai dados (NUP, Usuário, Data de Inclusão, e Status) de todos os processos de "Acervo Arquivístico de Dossiê de Beneficiário ADB" encontrados. Os dados são salvos periodicamente (sistema de *checkpoint*) em uma planilha de acompanhamento (Excel).

## 🚀 Como Executar

Ambos os scripts possuem uma interface gráfica amigável desenvolvida com `CustomTkinter`. Ao rodar o script (ex: `python ADB_Criar_Processo.py`), uma janela solicitará as credenciais do SEI e os caminhos (pastas ou planilhas) necessários para iniciar a automação.