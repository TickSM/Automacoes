# 📍 Automações - Regionais (GR)

Nesta pasta ficam os scripts desenvolvidos para auxiliar as Gerências Regionais (GR) nas suas rotinas específicas de digitalização.

## 📜 Scripts

*   **`Acompanhamento_GR.py`**: Similar ao script do projeto 506, este robô (`Selenium`) acessa o SEI, faz o login e, caso não esteja na unidade correta, altera para a unidade geradora especificada. Ele pesquisa processos, verifica as observações (identificando se pertencem à "2ª ETAPA DO PROJETO DE DIGITALIZAÇÃO") e extrai os dados consolidando-os diretamente em uma planilha padrão de acompanhamento (`AFD_processos_GR.xlsx`). O robô possui salvamento parcial a cada página lida, garantindo que os dados não sejam perdidos em caso de queda de conexão.

## 🚀 Como Executar

Inicie o script para abrir a interface gráfica. Se a planilha de acompanhamento padrão não for encontrada nos caminhos mapeados (discos de rede `J:` ou `I:`), o sistema solicitará que você localize o arquivo manualmente antes de iniciar a extração.