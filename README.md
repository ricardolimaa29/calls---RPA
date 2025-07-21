# calls---RPA

Automatiza o login no sistema SENAI, acessa a página de “Turmas” e extrai dados de presença de uma planilha Google Sheets.

---

## 📋 Pré-requisitos

- **Python 3.8+**
- Navegador **Google Chrome**
- **Chromedriver** compatível com sua versão do Chrome e disponível no PATH
- Conta no **Google Cloud** com APIs habilitadas:
  - Google Sheets API
  - Google Drive API
- Bibliotecas Python: `selenium`, `gspread`, `google-auth`, `oauth2client`
- (Opcional) `webdriver_manager` para gerenciamento automático do Chromedriver

---

## 🧩 Instalação

1. Clone seu repositório:
   ```bash
   git clone https://seu-repo.git
   cd seu-repo
2. Crie e ative o ambiente virtual:
    python -m venv venv
    source venv/bin/activate   # Windows: venv\Scripts\activate
   
3. Instale as dependências:
   pip install selenium gspread google-auth oauth2client python-dotenv webdriver_manager

4. Garanta que o Chromedriver esteja disponível, de preferência via webdriver_manager.
   
