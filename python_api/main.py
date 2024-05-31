from tkinter.filedialog import askopenfile
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Função de formatação
def format_currency(x):
    return f'R${x:,.2f}'

# Abrir a base de dados
database = askopenfile().name

# Importar a base de dados
tabela_vendas = pd.read_excel(database)

# Configuração para exibir todas as colunas da base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
faturamento_loja = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
html_faturamento_loja = faturamento_loja.to_html(formatters={'Valor Final': format_currency})

# Quantidade de produtos vendidos por loja
Qproduto_loja = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
html_Qproduto_loja = Qproduto_loja.to_html()

# Ticket médio por produto em cada loja
ticket_med = (faturamento_loja['Valor Final'] / Qproduto_loja['Quantidade']).to_frame()
ticket_med = ticket_med.rename(columns={0: 'Ticket Médio'})
html_ticket_med = ticket_med.to_html(formatters={'Ticket Médio': format_currency})

# Configurar o email
from_address = "seuemail@gmail.com" # Seu Email
to_address = "email@gmail.com" # Destinatário
subject = "Relatório de Vendas"

# Corpo do email em HTML
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

# Montar a mensagem
msg = MIMEMultipart()
msg['From'] = from_address
msg['To'] = to_address
msg['Subject'] = subject
msg.attach(MIMEText(body, 'html'))  # Definindo o conteúdo como HTML

# Configurar servidor SMTP (exemplo com Gmail)
smtp_server = "smtp.gmail.com"
smtp_port = 587
login = "seuemail@gmail.com"  # Seu email
password = "sua-senha"  # Sua senha de app

# Enviar o email
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # Upgrade para conexão segura
    server.login(login, password)
    server.send_message(msg)
    server.quit()
    print("Email enviado com sucesso!")
except Exception as e:
    print(f"Falha ao enviar email: {e}")
