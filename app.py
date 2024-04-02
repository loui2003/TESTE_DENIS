import streamlit as st
import pandas as pd
import xlsxwriter
import win32com.client as win32
from datetime import datetime

def send_email(email,anexo):

    hora_atual = datetime.now().time()
    if hora_atual >= datetime.strptime('05:00:00', '%H:%M:%S').time() and hora_atual < datetime.strptime('12:00:00', '%H:%M:%S').time():
        msg_hr = 'Bom dia!'
    elif hora_atual >= datetime.strptime('12:00:00', '%H:%M:%S').time() and hora_atual < datetime.strptime('18:00:00', '%H:%M:%S').time():
        msg_hr = 'Boa tarde!'
    else:
        msg_hr = 'Boa noite!'

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'Tabela de Dados filtrada'
    mail.Attachments.Add(anexo)
    mail.Body = f"""
                <p>{msg_hr}</p>
                <p>Segue sua base filtrada via sistema de busca interno</p>
                <p>Atenciosamente,<br>Equipe de Desenvolvimento</p>
                """
    mail.To = email
    mail.Send()

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
        enviar_email = st.button('Enviar por email')
        if enviar_email:
            email = st.text_input('Digite seu email')
            send_email(email,'tabela_filtrada.xlsx')
    else:
        st.info('A busca não retornou resultados!')
else:
    st.info('Você deve procurar um CPF ou CNPJ')