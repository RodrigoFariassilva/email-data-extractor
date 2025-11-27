
# Email-to-Excel Automation

## 📖 Descrição
Este projeto automatiza a extração de dados estruturados de e-mails recebidos no **Microsoft Outlook** e exporta-os para uma planilha Excel.  
Ele foi desenvolvido para simplificar tarefas repetitivas, como copiar manualmente **CNPJs** e **valores financeiros** de fundos e carteiras, garantindo agilidade e redução de erros.

---

## ✅ Funcionalidades
- Conexão com Outlook via `win32com.client`
- Leitura dos e-mails mais recentes
- Extração de **CNPJs** e **valores monetários** usando expressões regulares
- Exportação dos dados para um arquivo Excel provisório (`dados_extraidos.xlsx`)
- Estrutura pronta para integração com planilha real (mapeamento de células)

---

## 🛠 Tecnologias
- **Python 3.x**
- Bibliotecas:
  - `pywin32` (integração com Outlook)
  - `pandas` (manipulação de dados)
  - `openpyxl` (exportação para Excel)
  - `re` (expressões regulares)

---

## 🚀 Como instalar e rodar
1. Clone este repositório:
   ```bash
   git clone https://github.com/RodrigoFariassilva/email-data-extractor.git
   cd email-to-excel
