# Autoamacao_Dou
# 📄 Atualizador Automático de Prazos (DOU → Excel)

Um script em Python desenvolvido para automatizar a leitura de arquivos em PDF do Diário Oficial da União (DOU) e atualizar automaticamente planilhas de controle de prazos regulatórios.

O projeto resolve o problema da extração manual de dados, lendo múltiplos PDFs em paralelo, identificando os números de expedientes e suas respectivas datas de publicação, e inserindo essas informações diretamente nas abas de controle de equipe (ex: Márcia, Patrícia, etc.) no Excel.

## 🚀 Principais Funcionalidades

- **Extração Inteligente:** Utiliza Expressões Regulares (Regex) para identificar datas no padrão do DOU e números de expediente.
- **Processamento Paralelo:** Implementação de `ThreadPoolExecutor` para leitura simultânea de múltiplos PDFs, reduzindo drasticamente o tempo de execução.
- **Segurança dos Dados:**
  - Cria backups automáticos do arquivo Excel original antes de qualquer modificação.
  - Verifica se a célula já possui uma data preenchida e altera **apenas** células vazias.
- **Modo de Simulação (Dry-Run):** Permite rodar o script para verificar quais dados seriam alterados, sem de fato modificar nenhum arquivo.
- **Logging Estruturado:** Sistema de logs detalhado para facilitar o rastreamento de operações e erros.

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**
- **[pdfplumber](https://github.com/jsvine/pdfplumber):** Para extração precisa de texto dos arquivos PDF do Diário Oficial.
- **[openpyxl](https://openpyxl.readthedocs.io/):** Para leitura, manipulação e salvamento das planilhas Excel sem perder a formatação original.
- **Bibliotecas Nativas:** `concurrent.futures`, `dataclasses`, `argparse`, `logging`, `re`, `pathlib`.

## 📦 Como Instalar

1. Clone este repositório:
   ```bash
   git clone [https://github.com/seu-usuario/nome-do-repositorio.git](https://github.com/seu-usuario/nome-do-repositorio.git)
   cd nome-do-repositorio

Crie um ambiente virtual (recomendado):
  python -m venv venv
  source venv/bin/activate  # No Linux/Mac
  venv\Scripts\activate     # No Windows

  pip install -r requirements.txt

garanta que a estrutura do diretório esteja assim:
  📂 seu_projeto/
 ┣ 📜 atualizador_dou_v2.py
 ┣ 📜 requirements.txt
 ┣ 📂 arquivos_dou/                   # Coloque os PDFs baixados do DOU aqui
 ┗ 📊 (COPIA DE SEGURANÇA) CONTROLE DOS PRAZOS DE ANÁLISE_CCTAB_2025.xlsx  # Arquivo alvo

 Para rodar o script no modo padrão (atualizando o Excel e criando o backup):
 python atualizador_dou_v2.py
