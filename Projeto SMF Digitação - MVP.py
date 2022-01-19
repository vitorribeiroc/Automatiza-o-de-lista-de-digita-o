#!/usr/bin/env python
# coding: utf-8

# In[1]:


# PARTE 0 - IMPORTAÇÕES E ATRIBUIÇÕES INICIAIS


# In[2]:


# Importar bibliotecas e módulos necessários.
import win32com.client as win32
import numpy as np
import pandas as pd
from datetime import date


# In[3]:


#0 - Criar lista e dicionário com {nome: email} dos digitadores.
digitadores = {1: ['Vítor', 'vitorribeiro.fp@gmail.com'], 2: ['Eduarda', 'eduardasut.fp@gmail.com'], 3:['Vinícius', 'viniciusguioto.fp@gmail.com']}
lista_digitadores = ['Vítor', 'Eduarda', 'Vinícius']


# In[4]:


# PARTE 1 - IMPORTANDO E LIMPANDO OS DADOS


# In[5]:


#1 - Importar planilha.
lista_digitacao = pd.read_excel('C:/Users/danim/Downloads/protocolos_2021_10_23_15_07_05.xls')
# Buscar automatização para o procedimento.


# In[6]:


#2 - Limpar dados da planilha. (como reduzir quantidade de código?)


# In[7]:


#2.1 - Slice só com os dados necessários.
listalimpa = lista_digitacao[['Análise do Alvará', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5']]


# In[8]:


#2.2 - Renomeando colunas com nome escroto.
listalimpa.rename(columns={'Análise do Alvará': 'Protocolo', 'Unnamed: 3':'CPF/CNPJ','Unnamed: 4':'Razão Social', 'Unnamed: 5':'Data' }, inplace=True)


# In[9]:


#2.3 - Deletando linha desnecessária.
listalimpa.drop(0, inplace = True)


# In[10]:


# PARTE 2 - CÁLCULOS E DISTRIBUIÇÃO


# In[11]:


#3 - Contar nº total de processos.
totalprocessos = listalimpa['Protocolo'].count()
print(totalprocessos)


# In[12]:


#3.1 - Dividir quantidade de processos pelo número de digitadores.
proxdig = (listalimpa['Protocolo'].count()/len(digitadores))
print(proxdig)


# In[13]:


#4 - Atribuir processos aos digitadores.       


# In[14]:


#4.1 - Criar coluna 'Digitadores' no DF.
listalimpa['Digitador'] = ''


# In[15]:


#4.2 - Preencher coluna 'Digitadores' com os nomes, efetivamente atribuindo processo ao digitador.
lista_proxdig = []
for i in range(len(listalimpa)):
    for n in range(len(lista_digitadores)):
        if len(lista_proxdig) < len(listalimpa):
            lista_proxdig.append(lista_digitadores[n])
            
listalimpa['Digitador'] = lista_proxdig
           


# In[16]:


#4.3 - Separando os DF's por digitador e limpando-os (como fazer com menos linhas de código?).
df_vitorbase = listalimpa.loc[(listalimpa['Digitador'] == 'Vítor')]
df_vitor = df_vitorbase[['Protocolo','CPF/CNPJ','Razão Social']]

df_vinibase = listalimpa.loc[(listalimpa['Digitador'] == 'Vinícius')]
df_vini = df_vinibase[['Protocolo','CPF/CNPJ','Razão Social']]

df_dudsbase = listalimpa.loc[(listalimpa['Digitador'] == 'Eduarda')]
df_duds = df_dudsbase[['Protocolo','CPF/CNPJ','Razão Social']]


# In[17]:


#4.4 - Contabilizando os processos por digitador:
procvitor = len(df_vitor)

procvini = len(df_vini)

procduds = len(df_duds)


# In[18]:


# PARTE 3 - EMAILS.


# In[19]:


#4.5 - Pegar dia de hoje para colocar no email.
hoje = date.today().strftime('%d/%m/%Y')
print(hoje)


# In[20]:


#5 - Envio dos emails.


# In[21]:


#5.1 - Enviar email com a lista completa para controle ao email do setor. 
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'alvara@fazenda.niteroi.rj.gov.br'
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


# In[22]:


#5.2 - Enviar email ao digitador 1.
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'vitorribeiro.fp@gmail.com'
email.Subject = 'Lista de processos para digitar - Auto Teste'
email.HTMLBody = f'''
<p>Bom dia, Sr.Vítor 1, </p>

<p>segue a lista de digitação de hoje, {hoje}. </p>

<p>{df_vitor.to_html()} </p>

<p> Atenciosamente, </p>
<p> Coordenação de Cadastro Mobiliário </p>


'''
email.Send()

print('O email foi enviado para Vítor 1.')


# In[23]:


#5.3 - Enviar email ao digitador 2.
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'vitorribeiroc@hotmail.com'
email.Subject = 'Lista de processos para digitar - Auto Teste'
email.HTMLBody = f'''
<p>Bom dia, Sr.Vítor 2, </p>

<p>segue a lista de digitação de hoje, {hoje}. </p>

<p>{df_vini.to_html()} </p>

<p> Atenciosamente, </p>
<p> Coordenação de Cadastro Mobiliário </p>


'''
email.Send()

print('O email foi enviado para Vítor 2.')


# In[24]:


#5.4 - Enviar email ao digitador 3
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.To = 'eduardasut.fp@gmail.com'
email.Subject = 'Lista de processos para digitar - Auto Teste'
email.HTMLBody = f'''
<p>Bom dia, Sra.Eduarda Maria, </p>

<p>segue a lista de digitação de hoje, {hoje}. </p>

<p>EDUARDA</p>
<p>{df_duds.to_html()}</p>

<p> Atenciosamente, </p>
<p> Coordenação de Cadastro Mobiliário </p>

'''
email.Send()

print('O email foi enviado para Eduarda.')

