import streamlit as st
import pandas as pd
import xlsxwriter
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def enviar_email(email, anexo):

    hora_atual = datetime.now().time()
    if hora_atual >= datetime.strptime('05:00:00', '%H:%M:%S').time() and hora_atual < datetime.strptime('12:00:00', '%H:%M:%S').time():
        msg_hr = 'Bom dia!'
    elif hora_atual >= datetime.strptime('12:00:00', '%H:%M:%S').time() and hora_atual < datetime.strptime('18:00:00', '%H:%M:%S').time():
        msg_hr = 'Boa tarde!'
    else:
        msg_hr = 'Boa noite!'

    # Configurações do servidor SMTP do Gmail
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    # Construindo o email
    msg = MIMEMultipart()
    msg['From'] = 'equipe.desenvolvimento.cip@gmail.com'
    msg['To'] = email
    msg['Subject'] = 'Base Dados Filtrada'

    texto_corpo = f"""
                {msg_hr}
                Segue sua base filtrada via sistema de busca interno
                Atenciosamente,<br>Equipe de Desenvolvimento
                """
    # Corpo do email
    msg.attach(MIMEText(texto_corpo, 'plain'))

    # Adicionando o anexo
    with open(anexo, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {anexo}')
        msg.attach(part)

    # Conectando-se ao servidor SMTP do Gmail e enviando o email
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login('equipe.desenvolvimento.cip@gmail.com', 'Equipe_Desenvolvimento')
        server.sendmail('equipe.desenvolvimento.cip@gmail.com', email, msg.as_string())

def get_excel_bytes(df):
    with pd.ExcelWriter("tabela_filtrada.xlsx", engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    with open("tabela_filtrada.xlsx", "rb") as file:
        data = file.read()
    return data

dic = {'CPF':str, 'NOME':str, 'NUM_ACORDO':str,'VALOR_NEGOCIADO':str}
df = pd.read_csv(r'dataframe.csv',sep=';', dtype=dic)

st.set_page_config(layout='wide', page_title='CIP')

col_1, col_2, col_3 = st.columns(3)
with col_2:
    titulo = st.title('C.I.P')

with st.sidebar:
    cpfcnpj = st.text_input(
        label="CPF CLIENTE"
    )

if cpfcnpj != '':
    tabela_filtrada = df[df['CPF'] == cpfcnpj]
    st.dataframe(tabela_filtrada,use_container_width=True,hide_index=True)
    if not df.empty:
        file_bytes = get_excel_bytes(tabela_filtrada)
        st.download_button(label='Baixar tabela filtrada', data=file_bytes, file_name='tabela_filtrada.xlsx', mime='application/octet-stream')
        enviar_email_1 = st.button('Enviar por email?')
        if enviar_email_1:
            email = st.text_input('Digite seu email')
            enviar_email_2 = st.button('Enviar')
            if enviar_email_2 and email != '':
                enviar_email(email,'tabela_filtrada.xlsx')
            else:
                st.info('Você deve inserir um endereço de email!')
    else:
        st.info('A busca não retornou resultados!')
else:
    st.info('Você deve procurar um CPF ou CNPJ')
