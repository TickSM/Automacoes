# 📍 Automações - Regionais (GR)

Nesta pasta ficam os scripts desenvolvidos para auxiliar as Gerências Regionais (GR) nas suas rotinas específicas de digitalização.

## 📜 Scripts

*   **`Acompanhamento_GR.py`**: Robô em `Selenium` que acessa o SEI, faz login e extrai dados dos processos de digitalização da Gerência Regional. O script possui suporte a **execução em lote de múltiplas UFs em sequência** (com seleção facilitada por checkboxes para RJ, SP, BA e inclusão de novas UFs). O robô verifica as observações (identificando se pertencem à "2ª ETAPA DO PROJETO DE DIGITALIZAÇÃO") e consolida tudo na planilha padrão de acompanhamento (`AFD_processos_GR.xlsx`), com salvamento parcial a cada página e relatório final consolidado.

## 🚀 Como Executar

Inicie o script para abrir a interface gráfica (`CustomTkinter`).
1. Informe suas credenciais de acesso ao SEI.
2. Selecione as Regionais desejadas (as principais RJ, SP e BA já vêm pré-marcadas). Você também pode adicionar novas UFs no campo dedicado.
3. Clique em **Iniciar Automação**. O robô fará todo o processo em fila única automaticamente.
4. Se a planilha de acompanhamento padrão não for encontrada nos caminhos mapeados (discos de rede `J:` ou `I:`), o sistema solicitará que você localize o arquivo manualmente antes de iniciar a extração.