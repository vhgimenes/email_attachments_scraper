"""
Author: Victor Gimenes
Date: 08/11/2022
Módulo responsável por armazenar as funções de download de aquivos em anexos de e-mails.
"""
# Ignoras warnings desnecessários
import warnings
warnings.filterwarnings("ignore")

# Importando libs
from datetime import datetime as dt
import workdays as wd
import pandas as pd
import win32com.client
import glob
import os
import time 
import sys

# Importando módulo com meu logger customizado
from bz_holidays import scrape_anbima_holidays as bz 

# Criando as variáveis globais
global TITLE, PATH, FOLDER, ATTACHMENT, FILE_NAME, FILE_EXT
# Instanciando as variáveis relativas a rotina
TITLE = 'enter the title here!' # Título do E-mail 
PATH = 'enter the path here!' # Pasta onde será salvo  
FOLDER = 'enter the folder here' # Pasta dentro da Caixa de Entrada reservada para armazenar os e-mails de interesse  
ATTACHMENT = 7 # Qual a posição do anexo dentro dos anexos do e-mail (tentativa e erro)
FILE_NAME = 'enter the desire folder name here!' # Nome final do arquivo
FILE_EXT = 'enter the extension of the file here!' # Extensão do arquivo (exemplo: xlsx)

def get_last_refresh_date():
    """Função criada para retornar a última data de download do arquivo"""
    file_type = r'\*xlsx'
    files = glob.glob(PATH + file_type)
    max_file = max(files, key=os.path.getctime)
    date = pd.to_datetime(max_file[-13:-5],format='%Y%m%d').date()
    return date 

def download_btg_vols(init_date,final_date,holidays):    # sourcery skip: remove-redundant-pass
    """
    Função resposável por realizar o download das planilhas de preços dos e-mails do BTG. 
    
    Existe a necessidade de criar umas pasta 'BACK_BTG' entro da 'Caixa de Entrada' e manter o Outlook
    aberto para que a rotina possa ser executada.
    
    - Código referência:
      https://towardsdatascience.com/automatic-download-email-attachment-with-python-4aa59bc66c25
    """
    # Instanciando as variáveis estáticas
    print("Iniciando a rotina de download do arquivo.\n")

    # Criando contador de tempo de processamento
    processing_time = time.process_time()    

    # Criando o range de datas para atualização
    weekmask = "Mon Tue Wed Thu Fri"
    mydates = pd.bdate_range(start=init_date, end=final_date,
                             holidays=holidays, freq='C',
                             weekmask = weekmask).tolist()

    # Criando a conexão com outlook email
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # Filtrando a pasta de interesse (Levamos em cosideração que a pasta está dentro da caixa de entrada - 6)
    print("Acessando a pasta dentro da caixa de entrada.\n")
    inbox = outlook.GetDefaultFolder(6).folders(FOLDER)
    messages = inbox.Items
    message = messages.GetFirst()

    try:
        # While para percorrer por todos os emails da caixa BTG_BACK
        while message != None:
            # current_sender = str(message.Sender).lower()
            current_subject = str(message.Subject)
            current_date = message.senton.date()
            str_current_date  = current_date.strftime('%Y%m%d')
            # Caso o título do e-mail seja igual ao setado anteriormente acessaremos esse email
            if  (TITLE == current_subject) and (current_date in mydates): # and re.search(SENDER,current_sender) != None
                print(f'Extraindo o arquvo {current_subject} referente ao dia {current_date}')
                # Acessando os attachments do email de interesse 
                attachments = message.Attachments
                attachment = attachments.Item(2)
                # Fitrando nome do arquivo e extenção
                file_name = f'{FILE_NAME}_{str_current_date}.{FILE_EXT}'
                # Salvando o arquivo no path indicado no início do script
                file_path = os.path.join(PATH, file_name)
                # Checando se o arquivo já existe na rede
                check = os.path.exists(file_path)
                if check==False:
                    attachment.SaveASFile(os.path.join(PATH, file_name))
                    print('Arquivo salvo na pasta com sucesso!\n')
                else:
                    print('Arquivo já foi salvo na pasta anteriormente.\n')
            # Acessando a próxima mensagem
            message = messages.GetNext()
        # Finalizando a rotina e mostrando o tempo de execução
        print('Rotina finalizada com sucesso!')
        print(f'Tempo de execução: {processing_time}\n')
    except Exception as e:
        print('Erro na rotina, stopando o processo, checar manualmente.')
        raise e
    
def main():
    # Criando conexão com o canal de avisos do Teams 
    teams_conn = pm.get_connector_mesa_teams()
    # Inciando a rotina
    try:
        # Data de referência
        now = dt.now() # Extraindo horário
        today = now.date() # Extraindo data de rodagem
        holidays = bz.holidays() # Extraindo os feriados da Anbima
        yesterday = wd.workday(today,-1,holidays)
        final_date = yesterday if now.hour < 22 else today # data final de análise dependera do horário
        # Ultima data com dados da tabela
        last_refresh = get_last_refresh_date()
        # Check para ver se já foi atualizada
        if final_date > last_refresh:
            # Criando as datas para busca das marcações
            init_date = wd.workday(last_refresh,1,holidays)
            #! Descomentar para recalcular o histórico
            # init_date = dt(2022,12,1).date()
            # Extraido as marcações para a data desejada   
            download_btg_vols(init_date,final_date,holidays)
            # Checando se o último arquivo foi atualizado, caso não retornaremos um erro
            new_path_file = os.path.join(PATH, f"{FILE_NAME}_{init_date.strftime('%Y%m%d')}.{FILE_EXT}")
            check_new_file = os.path.exists(new_path_file)
            if check_new_file == True:
                print("Download do arquivo realizado com sucesso!")
            else:
                print("Erro no download do arquivo, checar manualmente.")
        else:
            print("Download já foi feito hoje!")
    except Exception as e:
        print("Erro no download do arquivo, checar manualmente.")

if __name__ == '__main__':
    main()
