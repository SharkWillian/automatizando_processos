#------------------------ Importações ------------------------#

import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
import smtplib,math
import numpy as np
import matplotlib.pyplot as plt
import io
import base64

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

import pymysql
import pandas as pd
pd.options.mode.copy_on_write = True
pymysql.install_as_MySQLdb()
from sqlalchemy import create_engine
from datetime import date, timedelta,datetime


#------------------------ Acessando arquivos e credenciais ------------------------#

from_email = '<loggin e-mail>'
from_psw = '<senha e-mail>'

destinatarios = [
    '<destinatário1@...>',
    '<destinatário1@...>',
    '<destinatário1@...>']

#destinatarios = ['<seu e-mail>'] # Testes

host = "<host e-mail>"
port = "<port e-mail>"

yesterday = date.today() - timedelta(days=1)
arquivo = f"<caminho para o arquivo>{yesterday.strftime('%Y-%m-%d')}.xlsx"
pool_size = 50
sqlalchemy_uri = '<uri do mysql>'
engine = create_engine(sqlalchemy_uri, pool_size=50, max_overflow=0)


#------------------------ Criando funções ------------------------#

#----------Corpo do Email---------#
def email_corpo():
    hora = int(datetime.now().strftime('%H'))
    corpo = f"""<html>
    <body>
        <div>
        <p>
            <span>
            Prezados, {'bom dia' if hora < 12 else 'boa tarde' if hora < 19 else 'boa noite'}.<br>
            </span>
        </p>
        <div>
            <p>
            <span> Segue anexo o report diário.<br> 
                <br>
                   Dados referentes ao mês de <mês> até {yesterday.strftime('%Y-%m-%d')}. <br> <br>
                <br>Att.
            </span>
            </p>
            <p>
            <span>&nbsp;
            <img width=439 height=126 src="<assinatura>" >
            </span>
            </p>
        </div>
        <p>
            <span style='font-size:12.0pt'>&nbsp;
            </span>
        </p>
        </div>
    </body>
    </html>
    """
    return corpo

#---------- Carregando os dados ---------#
def gera_xlsx():

    QUERY = '''
    <Query para formar a tabela>'''

    df = pd.read_sql_query(QUERY,engine)
    
    <Tratamento dos dados com Pandas>

    css_alt_rows = 'background-color: #D9D9D9; color: black;'
    css_indexes = 'background-color: #0F243E; color: white;'
    
    df.style.apply(lambda col: np.where(col.index % 2, css_alt_rows, None)).applymap_index(lambda _: css_indexes, axis=0).applymap_index(lambda _: css_indexes, axis=1).to_excel(arquivo, sheet_name='<nome da sheet>',index=False)


#---------- Enviando o e-mail ---------#
def enviar_email(destinatarios_email,assunto_email,corpo_email):

    gera_xlsx()
    
    server = smtplib.SMTP(host,port)
    server.starttls()
    server.login(from_email,from_psw)

    email_msg = MIMEMultipart()
    email_msg['From'] = from_email
    email_msg['To'] =  "; ".join(destinatarios_email)
    email_msg['Subject'] = assunto_email
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(arquivo, 'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 
                    f"attachment; filename={arquivo.split('/')[-1]}")
    email_msg.attach(part)
    
    email_msg.attach(MIMEText(corpo_email,'html'))

    server.sendmail(from_email,destinatarios_email,email_msg.as_string())
    server.quit()


if __name__ == '__main__':
    enviar_email(destinatarios,f"<Assunto do e-mail>",email_corpo())
    print('Tá pronto o SorvErtinho!')
