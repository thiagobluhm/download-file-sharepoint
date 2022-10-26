#!/usr/bin/env python
# coding: utf-8

# # Download arquivo Sharepoint 

# In[77]:


from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 

import requests
import pandas as pd
import io

import os
# Aqui colocar a pasta onde fica o script 
os.chdir(r'C:\Users\ThiagoBluhm\OneDrive - Grupo Portfolio\Documentos\ESTUDO\sharepoint')


# ### Set de acordo com o usuario

# In[82]:


sharepoint_base_url = 'https://grupoportfoliocombr.sharepoint.com/sites/bi/'
sharepoint_user = 'thiago.bluhm@portfoliotech.com.br'
#essa senha é gerada pela conta do usuario e deve ser a senha de APP
sharepoint_password = 'lbnshzkfmhzylrry' 
####################################################################
pasta_no_sharepoint = '/sites/bi/Pasta/Consulting/Laredo'


# In[103]:


def inicio(sharepoint_base_url, sharepoint_user, sharepoint_password, pasta_no_sharepoint):
    try:
        auth = AuthenticationContext(sharepoint_base_url) 
        auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
        ctx = ClientContext(sharepoint_base_url, auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        # Confirmando onde estamos no SHAREPOINT
        print('Connected to SharePoint: ',web.properties['Title'])
        print('Iniciar download...')
        propriedades = carregar_pastas(ctx, pasta_no_sharepoint)        
        
    except:
        print('Problemas com a autenticacao.')
        
    return propriedades

# Buscando pastas de acordo com a URL relativa passada acima.
def carregar_pastas(ctx, pasta_no_sharepoint):
    try:
        pasta = ctx.web.get_folder_by_server_relative_url(pasta_no_sharepoint)
        propriedades = []  
        sub_pasta = pasta.files   
        ctx.load(sub_pasta)  
        ctx.execute_query()  
        for conteudo in sub_pasta:
            propriedades.append(conteudo.properties) 

        #gravando arquivo na pasta especificada
        #especifiquei propriedades[0]['Name'] pq so existe um arquivo na pasta do sharepoint
        #fazer um loop se houver mais de um arquivo, ou especificar o nome do arquivo a ser trazido
        grava_arquivo(propriedades[0]['ServerRelativeUrl'], propriedades[0]['Name'])
        
    except:
        print('Algum problema com a leitura das pastas. Confira a autenticacao.')
    
    return propriedades

def grava_arquivo(caminho_relativo_url, nome_arquivo):
    # aqui colocar a pasta dentro da pasta onde fica o script que recebera o download + nome do arquivo
    diretorio = checa_dir(r'./arquivo/')
    pasta_para_download = diretorio + nome_arquivo
    try:
        with open(pasta_para_download, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_path(caminho_relativo_url)
            file.download(local_file).execute_query()

        print("[Ok] Arquivo foi baixado na pasta: {0}".format(nome_arquivo))
    
    except:
        print('Algum problema com o download!')

    return 0

def checa_dir(diretorio):
    if os.path.isdir(diretorio) is False:
        os.mkdir(diretorio)
        dir_ = diretorio
        print(f'Diretorio criado: {diretorio}')
    else:
        print(f'Diretorio já {diretorio} existe.')
        dir_ = diretorio
    
    return dir_


# ### Chamando funcao inicio

# In[104]:


props_lista = inicio(sharepoint_base_url, sharepoint_user, sharepoint_password, pasta_no_sharepoint) 


# In[96]:


# Constatando outros dados que vem nas propriedades do CTX
props_lista


# ### EOF 
