# Dashboard de Performance de Vendas v5.0

Uma aplicação desktop completa desenvolvida em Python para monitoramento, visualização e reporte da performance da equipe de vendas, baseada em dados de planilhas Excel.

## ✨ Funcionalidades Principais

* **Visualização em Tempo Real:** Monitora arquivos Excel (`Planilha_Pontuacao_Atualizada.xlsx`) e atualiza a interface automaticamente (`watchdog`).
* **Múltiplas Telas:** Interface organizada com navegação lateral, incluindo:
    * **Líder Atual:** Exibe o vendedor com maior pontuação (com tratamento para empates).
    * **Ranking Top 5:** Classificação geral dos melhores vendedores.
    * **Jornada da Performance:** Gráfico de linhas mostrando a evolução da pontuação dos Top 5 ao longo do tempo (incluindo perdas de pontos).
    * **Desempenho por Equipe:** Rankings individuais para cada equipe de coordenador.
    * **Mural de Conquistas:** Destaca feitos notáveis (Liderança, Maior Ganho de Pontos, Maior Escalada no Ranking).
    * **Crônica da Liderança:** Histórico persistente (SQLite) das mudanças de liderança e pontuações registradas manualmente.
    * **Relógio:** Exibe a hora atual e uma saudação.
* **Entrada de Dados Interativa:** Tela "Pontuar" para registrar novas pontuações baseadas em ações predefinidas (com checkboxes), associadas a clientes (`Tabela_Cliente.xlsx`), com autocomplete para vendedores. Atualiza a planilha de origem.
* **Persistência de Dados:** Utiliza um banco de dados SQLite (`dashboard_history.db`) para salvar o histórico de pontuações (para o gráfico) e os eventos da Crônica, garantindo que os dados não sejam perdidos ao fechar o programa.
* **Relatórios Automáticos por E-mail:**
    * Agenda o envio de e-mails para as 15:00 (Seg-Sex) (`schedule`).
    * Tira screenshots das telas relevantes (Líder, Ranking, Equipes).
    * Envia automaticamente via **Outlook** (`pywin32`) para a gerência e coordenadores, com anexos e assinatura digital.
    * Inclui botão para envio manual de teste.
* **Modo Apresentação:** Esconde a barra lateral e alterna automaticamente entre as telas principais e as telas de cada coordenador em intervalos definidos.
* **Interface Moderna:** GUI desenvolvida com Tkinter, utilizando temas de cores customizáveis e animações sutis.
* **Configuração Centralizada:** Fácil personalização de caminhos, e-mails, cores, fontes e intervalos através do dicionário `CONFIG`.

## 🛠️ Tecnologias Utilizadas

* **Python 3**
* **Tkinter:** Para a interface gráfica.
* **Pandas:** Para leitura e processamento inicial das planilhas Excel.
* **Openpyxl:** Para escrita na planilha Excel (`Planilha_Pontuacao_Atualizada.xlsx`).
* **SQLite3:** Para armazenamento persistente do histórico.
* **Watchdog:** Para monitoramento de arquivos em tempo real.
* **Schedule:** Para agendamento do envio de e-mails.
* **pywin32:** Para automação do Microsoft Outlook.
* **Pillow (PIL):** Para manipulação de imagens (ícones e screenshots).
* **re:** Para limpeza de strings (nomes de colunas).

## 🚀 Como Usar

1.  Clone o repositório.
2.  Instale as dependências: `pip install pandas openpyxl watchdog schedule pywin32 Pillow`
3.  Certifique-se que os arquivos `Planilha_Pontuacao_Atualizada.xlsx` e `Tabela_Cliente.xlsx` estão na mesma pasta do script.
4.  Crie uma pasta chamada `icons` e adicione os ícones necessários.
5.  **Configure os e-mails** no dicionário `CONFIG` dentro do script.
6.  Execute o script: `python seu_script_nome.py`
7.  Certifique-se que o Microsoft Outlook está aberto e configurado para o envio de e-mails.

# Sales Performance Dashboard v5.0

A comprehensive desktop application developed in Python for monitoring, visualizing, and reporting sales team performance based on data from Excel spreadsheets.

## ✨ Key Features

* **Real-time Visualization:** Monitors Excel files (`Planilha_Pontuacao_Atualizada.xlsx`) and updates the interface automatically (`watchdog`).
* **Multiple Screens:** Organized interface with sidebar navigation, including:
    * **Current Leader:** Displays the top-scoring salesperson (handles ties).
    * **Top 5 Ranking:** Overall leaderboard of the best performers.
    * **Performance Journey:** Line chart showing the score evolution of the Top 5 over time (including score deductions).
    * **Team Performance:** Individual rankings for each coordinator's team.
    * **Achievements Board:** Highlights notable feats (Leadership, Biggest Score Gain, Fastest Rank Climber).
    * **Leadership Chronicle:** Persistent history (SQLite) of leadership changes and manually logged scores.
    * **Clock:** Displays the current time and a greeting.
* **Interactive Data Entry:** "Score" screen to register new scores based on predefined actions (using checkboxes), linked to clients (`Tabela_Cliente.xlsx`), with salesperson autocomplete. Updates the source spreadsheet.
* **Data Persistence:** Uses an SQLite database (`dashboard_history.db`) to save score history (for the chart) and Chronicle events, ensuring no data is lost when the program closes.
* **Automated Email Reports:**
    * Schedules email dispatch for 3:00 PM (Mon-Fri) (`schedule`).
    * Takes screenshots of relevant screens (Leader, Ranking, Teams).
    * Automatically sends emails via **Outlook** (`pywin32`) to management and coordinators, with attachments and digital signature.
    * Includes a button for manual test sending.
* **Presentation Mode:** Hides the sidebar and automatically cycles through the main screens and individual coordinator screens at defined intervals.
* **Modern Interface:** GUI developed with Tkinter, featuring customizable color themes and subtle animations.
* **Centralized Configuration:** Easy customization of paths, emails, colors, fonts, and intervals via the `CONFIG` dictionary.

## 🛠️ Technologies Used

* **Python 3**
* **Tkinter:** For the graphical user interface.
* **Pandas:** For reading and initial processing of Excel spreadsheets.
* **Openpyxl:** For writing to the Excel spreadsheet (`Planilha_Pontuacao_Atualizada.xlsx`).
* **SQLite3:** For persistent storage of historical data.
* **Watchdog:** For real-time file monitoring.
* **Schedule:** For scheduling email dispatches.
* **pywin32:** For Microsoft Outlook automation.
* **Pillow (PIL):** For image manipulation (icons and screenshots).
* **re:** For string cleaning (column names).

## 🚀 How to Use

1.  Clone the repository.
2.  Install dependencies: `pip install pandas openpyxl watchdog schedule pywin32 Pillow`
3.  Ensure `Planilha_Pontuacao_Atualizada.xlsx` and `Tabela_Cliente.xlsx` files are in the same directory as the script.
4.  Create a folder named `icons` and add the necessary icon files.
5.  **Configure the email addresses** in the `CONFIG` dictionary within the script.
6.  Run the script: `python your_script_name.py`
7.  Ensure Microsoft Outlook is open and configured for sending emails.
