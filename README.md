# Email Reports

Email Reports é um projeto desenvolvido para automatizar o envio de relatórios de vendas por email. Este projeto foi criado como parte do meu aprendizado em Python e suas bibliotecas, sendo meu segundo projeto prático usando Python.

## Descrição

O Email Reports permite importar uma base de dados em Excel, calcular métricas de vendas por loja, e enviar essas métricas por email de forma automática. As métricas incluem faturamento, quantidade de produtos vendidos e ticket médio por produto em cada loja. Importante lembrar que o código presente nesse repositório só valerá para a base de dados na pasta "example_database".

## Funcionalidades

- Seleção de arquivo de base de dados em Excel.
- Cálculo do faturamento por loja.
- Cálculo da quantidade de produtos vendidos por loja.
- Cálculo do ticket médio por produto em cada loja.
- Envio automático de relatórios por email em formato HTML.

## Tecnologias Utilizadas

- Python 3
- Biblioteca `tkinter` para interface gráfica de seleção de arquivo.
- Biblioteca `pandas` para manipulação de dados.
- Biblioteca `smtplib` para envio de emails.
- Biblioteca `email` para construção de mensagens de email em HTML.

## Como Usar

1. **Clone o repositório**:
   ```sh
   git clone <URL_DO_REPOSITÓRIO>
   cd <NOME_DO_REPOSITÓRIO>
   ```

2. **Instale as dependências**:
   - Certifique-se de ter o Python 3 instalado.
   - Instale as bibliotecas necessárias:
     ```sh
     pip install pandas openpyxl
     ```

3. **Configure o script**:
   - Edite o script `email_reports.py` para incluir seu email e senha de app na seção de configuração do email:
     ```python
     from_address = "seuemail@gmail.com"  # Seu Email
     to_address = "email@gmail.com"  # Destinatário
     login = "seuemail@gmail.com"  # Seu email
     password = "sua-senha"  # Sua senha de app
     ```

4. **Execute o script**:
   ```sh
   python email_reports.py
   ```

5. **Selecione a base de dados**:
   - Uma janela pop-up permitirá que você selecione o arquivo Excel contendo os dados de vendas.

6. **Verifique o envio do email**:
   - O script calculará as métricas e enviará o relatório por email.

## Código

```python
from tkinter.filedialog import askopenfile
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def format_currency(x):
    return f'R${x:,.2f}'

database = askopenfile().name
tabela_vendas = pd.read_excel(database)
pd.set_option('display.max_columns', None)

faturamento_loja = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
html_faturamento_loja = faturamento_loja.to_html(formatters={'Valor Final': format_currency})

Qproduto_loja = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
html_Qproduto_loja = Qproduto_loja.to_html()

ticket_med = (faturamento_loja['Valor Final'] / Qproduto_loja['Quantidade']).to_frame()
ticket_med = ticket_med.rename(columns={0: 'Ticket Médio'})
html_ticket_med = ticket_med.to_html(formatters={'Ticket Médio': format_currency})

from_address = "seuemail@gmail.com"
to_address = "email@gmail.com"
subject = "Relatório de Vendas"

body = f'''
<html>
<head></head>
<body>
<p>Prezados,</p>
<p>Segue o relatório de vendas por cada loja.</p>
<p><b>Faturamento:</b></p>
{html_faturamento_loja}
<p><b>Quantidade Vendida:</b></p>
{html_Qproduto_loja}
<p><b>Ticket médio dos produtos de cada loja:</b></p>
{html_ticket_med}
<p>Qualquer dúvida estou à disposição.</p>
<p>Att.,<br>Junior</p>
</body>
</html>
'''

msg = MIMEMultipart()
msg['From'] = from_address
msg['To'] = to_address
msg['Subject'] = subject
msg.attach(MIMEText(body, 'html'))

smtp_server = "smtp.gmail.com"
smtp_port = 587
login = "seuemail@gmail.com"
password = "sua-senha"

try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(login, password)
    server.send_message(msg)
    server.quit()
    print("Email enviado com sucesso!")
except Exception as e:
    print(f"Falha ao enviar email: {e}")
```

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e pull requests para melhorias ou correções.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## Contato

Se você tiver dúvidas ou sugestões, sinta-se à vontade para me contatar.

## Conecte-se comigo

[![Gmail](https://img.shields.io/badge/Gmail-333333?style=for-the-badge&logo=gmail&logoColor=red)](mailto:juniorbmelo12@gmail.com)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/alexsandro-junior-576719297/)
[![Instagram](https://img.shields.io/badge/-Instagram-%23E4405F?style=for-the-badge&logo=instagram&logoColor=white)](https://www.instagram.com/juniorbm.wn/)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/junioom)
