
#1 - Carregar planilha com os destinatarios (Nome, Email, Assunto)
#2 - Carregar HTML com o corpo do email a ser enviado
#3 - Colocar o nome e assunto no html
#4 - Enviar email para o destinatario da planilha, com o corpo em html e com anexo

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd

x = pd.read_excel(r"destinatarios.xls")

nomes = []
emails = []
assuntos = []
datas = []

for registro in range(len(x)):
    nomes.append(x['nome'] [registro])
    emails.append(x['e-mail'] [registro])
    assuntos.append(x['assunto']  [registro])
    datas.append(x['data']  [registro].strftime('%d/%m/%Y'))

email_from = 'seu@email'
email_password = 'suasenha'
email_smtp_server = 'seusmtp'

for i in range(len(nomes)):
    print(emails[i])
    destination = [emails[i]]
    subject = 'Email referente - ' + assuntos[i] 


    text = """<html>
                    <h3>Olá {0} </h3>
                    </br>
                    <p>Precisamos do seu parecer da {1} realizada no dia {2}, para isso preparamos um modelo de parecer que esta anexo, nele voce encontra algumas instruções de como preenche-lo.</p>
                    <p>Em caso de duvidas, estou a disposição.<p>
                    </br>
                    <p>Atenciosamente</p>
                    <p>Python</p><p>Email Sender<p/>
              </html>"""
    text = text.format(nomes[i], assuntos[i], datas[i])

    print(text)
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['Subject'] = subject
    msg_text =  MIMEText(text, 'html')
    msg.attach(msg_text)

    filename = 'arquivo.doc'
    attachment = open(filename,'rb')


    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    attachment.close()

    try:
        smtp = smtplib.SMTP(email_smtp_server, 587)
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(email_from, email_password)
        smtp.sendmail(email_from, ','.join(destination), msg.as_string())
        smtp.quit()
        print('Email enviado')

    except Exception as err:
        print(f'Falha ao enviar email: {err}')

#print(msg)