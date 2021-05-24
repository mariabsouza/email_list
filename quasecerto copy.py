import win32com.client as win32
import pandas as pd

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

email_list = pd.read_excel('D:\Projeto Zurich/email_list.xlsx')
emails = email_list['EMAIL']

# iterate through the records
for i in range(len(emails)):
  
    # for every record get the name and the email addresses
    email = emails[i]


# configurar as informações do seu e-mail
email.To = emails
email.Subject = "Informações sobre o seu sinistro!"
email.HTMLBody = """
<p>Prezado(a) segurado(a),</p>
<p>Foram realizadas diversas tentativas de contato sem sucesso para agendar a realização da visita da assistência técnica Electrolux, por esse motivo estamos cancelando o seu sinistro.</p>
<p>Para solicitar a reabertura, você pode entrar em contato pelos nossos canais de atendimento:</p>
<p>4020 4848 (capitais e regiões metropolitanas)</p>
<p>0800 285 4141 (demais localidades)</p>
<p>Ou através dos nossos canais digitais em https://www.zurich.com.br/pt-br/atendimento</p>

<p>Atenciosamente,</p>
<p>Zurich Seguros</p>
"""

email.Send()
print("Email Enviado")
