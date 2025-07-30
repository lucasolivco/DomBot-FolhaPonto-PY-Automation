DomBot - Folha de Ponto

Automação para geração de folhas de ponto no sistema Domínio Folha com interface gráfica amigável.

VISÃO GERAL:
O DomBot é uma solução de automação que simplifica o processo de geração de folhas de ponto no sistema Domínio Folha.

FUNCIONALIDADES:
- Interface gráfica moderna e intuitiva
- Processamento em segundo plano
- Sistema de logs detalhado
- Barra de progresso visual
- Controle de execução (iniciar/parar)
- Gerenciamento de erros robusto
- Exportação de PDFs automática
- Suporte a múltiplas empresas

PRÉ-REQUISITOS:
- Windows 10/11
- Python 3.9+
- Sistema Domínio Folha instalado
- Microsoft Excel

INSTALAÇÃO:
1. Clone o repositório

2. Crie e ative um ambiente virtual (recomendado):
python -m venv venv
venv\Scripts\activate

3. Instale as dependências:
pip install -r requirements.txt

COMO USAR:
1. Prepare seu arquivo Excel com as colunas:
   - Nº: Número da empresa no Domínio
   - EMPRESA: Nome da empresa (para logs)
   - data inicio: Data inicial (DD/MM/AAAA)
   - data final: Data final (DD/MM/AAAA)
   - nome pdf: Nome do arquivo PDF (sem extensão)

2. Execute o programa:
python DomBot-FolhaPonto.py

3. Na interface:
   - Selecione o arquivo Excel
   - Defina a linha inicial (opcional)
   - Clique em "Iniciar"

ESTRUTURA DE ARQUIVOS:
DomBot-FolhaPonto/
├── assets/                     # Ícones e recursos visuais
├── logs/                       # Logs de execução automáticos
├── DomBot-FolhaPonto.py        # Script principal
├── requirements.txt            # Dependências
└── README.md                   # Documentação

FLUXO DE TRABALHO:
1. Conecta ao Domínio Folha
2. Para cada linha do Excel:
   - Troca para a empresa especificada
   - Acessa o módulo de relatórios
   - Configura o período da folha de ponto
   - Gera o relatório
   - Salva o PDF no diretório padrão
3. Registra sucessos e erros em logs
4. Atualiza interface em tempo real

SOLUÇÃO DE PROBLEMAS:
- Janelas não encontradas: Verifique os títulos das janelas do Domínio
- Erros de tempo: Aumente os tempos de espera
- Problemas de foco: Certifique-se que o Domínio está visível

REQUIREMENTS.TXT:
customtkinter==5.2.0
pandas==2.0.3
pywinauto==0.6.8
pywin32==306
pillow==10.1.0
