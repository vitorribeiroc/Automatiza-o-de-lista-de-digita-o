#!/usr/bin/env python
# coding: utf-8

# In[2]:


# Importar bibliotecas e módulos necessários.
import win32com.client as win32
import numpy as np
import pandas as pd
from datetime import date

#0 - Criar lista e dicionário com {nome: email} dos digitadores.
digitadores = {1: ['Vítor', 'emaildovitor@gmail.com'], 2: ['Eduarda', 'emaildaduda@gmail.com'], 3:['Vinícius', 'emaildovini@gmail.com']}
lista_digitadores = ['Vítor', 'Eduarda', 'Vinícius']

# PARTE 1 - IMPORTANDO E LIMPANDO OS DADOS
#1 - Importar planilha.
lista_digitacao = pd.read_excel('C:/Users/danim/Downloads/protocolos_2021_10_23_15_07_05.xls')

#2 - Limpar dados da planilha. (como reduzir quantidade de código?)
#2.1 - Slice só com os dados necessários.
listalimpa = lista_digitacao[['Análise do Alvará', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5']]
#2.2 - Renomeando colunas com nome escroto.
listalimpa.rename(columns={'Análise do Alvará': 'Protocolo', 'Unnamed: 3':'CPF/CNPJ','Unnamed: 4':'Razão Social', 'Unnamed: 5':'Data' }, inplace=True)
#2.3 - Deletando linha desnecessária.
listalimpa.drop(0, inplace = True)

# PARTE 2 - CÁLCULOS E DISTRIBUIÇÃO
#1 - Contar nº total de processos.
totalprocessos = listalimpa['Protocolo'].count()
print(totalprocessos)
#1.1 - Dividir quantidade de processos pelo número de digitadores.
proxdig = (listalimpa['Protocolo'].count()/len(digitadores))
print(proxdig)
#2 - Atribuir processos aos digitadores.
#2.1 - Criar coluna 'Digitadores' no DF.
listalimpa['Digitador'] = ''
#2.2 - Preencher coluna 'Digitadores' com os nomes, efetivamente atribuindo processo ao digitador.
lista_proxdig = []
for i in range(len(listalimpa)):
    for n in range(len(lista_digitadores)):
        if len(lista_proxdig) < len(listalimpa):
            lista_proxdig.append(lista_digitadores[n])        
listalimpa['Digitador'] = lista_proxdig
#2.3 - Separando os DF's por digitador e limpando-os (como fazer com menos linhas de código?).
df_vitorbase = listalimpa.loc[(listalimpa['Digitador'] == 'Vítor')]
df_vitor = df_vitorbase[['Protocolo','CPF/CNPJ','Razão Social']]
df_vinibase = listalimpa.loc[(listalimpa['Digitador'] == 'Vinícius')]
df_vini = df_vinibase[['Protocolo','CPF/CNPJ','Razão Social']]
df_dudsbase = listalimpa.loc[(listalimpa['Digitador'] == 'Eduarda')]
df_duds = df_dudsbase[['Protocolo','CPF/CNPJ','Razão Social']]
#2.4 - Contabilizando os processos por digitador:
procvitor = len(df_vitor)
procvini = len(df_vini)
procduds = len(df_duds)

# PARTE 3 - EMAILS.
#1 - Pegar dia de hoje para colocar no email.
hoje = date.today().strftime('%d/%m/%Y')
print(hoje)
#2 - Envio dos emails.
#2.1 - Enviar email com a lista completa para controle ao email do setor. 
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'destinatario@conta.com.br'
email.Subject = 'Controle de processos para digitar - Auto Teste'
email.HTMLBody = f'''
<p>Bom dia, senhores, </p>

<p>segue a lista de digitação distribuída hoje, {hoje}. </p>
<p>VÍTOR: </p>
<p>{df_vitor.to_html()} </p>
<p> Total de {procvitor}  processos. </p>

<p>VINÍCIUS:</p>
<p>{df_vini.to_html()} </p>
<p> Total de {procvini} processos. </p>

<p>EDUARDA</p>
<p>{df_duds.to_html()} </p>
<p> Total de {procduds} processos. </p>

Foram distribuídos, no total, {totalprocessos} processos em {hoje}. </p>

<p> Atenciosamente, </p>
<p> Coordenação de Cadastro Mobiliário </p>


'''
email.Send()
print('O email de controle foi enviado.')

