# Case de Análise de Dados Financeiros

Este projeto realiza a extração, consolidação, análise de retornos e validação de dados históricos de ativos financeiros a partir de uma planilha Excel e da API do Yahoo Finance.

## 🛠️ Instalação

1.  **Clone o repositório:**
    ```bash
    git clone <url-do-repositorio>
    cd case
    ```

2.  **Crie um ambiente virtual (recomendado):**
    ```bash
    python -m venv .venv
    ```

3.  **Ative o ambiente virtual:**
    *   **Windows:** `.venv\Scripts\activate`
    *   **Linux/Mac:** `source .venv/bin/activate`

4.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

## 🏃 Como Rodar

Basta executar o script principal:

```bash
python main.py
```

Após a execução, os seguintes arquivos serão gerados/atualizados:
*   `consolidated_data.xlsx`: Contendo as abas `consolidated_data`, `top_5` e `erros`.
*   `cumulative_returns.png`: Gráfico de evolução dos retornos.

## 📦 Bibliotecas Utilizadas

*   `pandas`: Manipulação e análise de dados.
*   `openpyxl`: Integração e edição de arquivos Excel.
*   `matplotlib`: Geração de gráficos.
*   `yfinance`: Extração de dados financeiros em tempo real.