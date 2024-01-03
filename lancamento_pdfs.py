import configparser
import json
import time
import zipfile
import os
import shutil
import sys
import pyodbc
from datetime import datetime, timedelta,date
from os import listdir, path, remove, unlink, walk
from shutil import rmtree, move
from subprocess import CREATE_NEW_CONSOLE, Popen, call
from threading import *
from lançamento_dassa import converterXmlServicoToJson
from lancador_equatorial import *


import pandas as pd
import pyautogui
import pyperclip as pc
import win32api

now = datetime.now()
today = date.today()
current_time = now.strftime("%H:%M:%S")
current_day = now.strftime("%d/%m/%Y")

if getattr(sys, 'frozen', False):
    application_path = path.dirname(sys.executable)
elif __file__:
    application_path = path.dirname(__file__)

cfg = configparser.ConfigParser()
cfg.read(application_path+'\\lancamento_servicos_procfit.ini')

EXE_PROCFIT = cfg.get('CONFIG', 'exe_procfit')
USUARIO_PROCFIT = cfg.get('CONFIG', 'usuario_procfit')
SENHA_PROCFIT = cfg.get('CONFIG', 'senha_procfit')
DIR_XMLS = cfg.get('CONFIG', 'dir_xml')
DIR_XMLS_PROCESSADAS = cfg.get('CONFIG', 'dir_processadas')
DIR_XML_N_PROCESSADAS = cfg.get('CONFIG', 'dir_nao_resolvida')
DIR_PDF = cfg.get('CONFIG','dir_pdf')
DIR_PDFS_N_PROCESSADOS = cfg.get('CONFIG','dir_pdf_n_resolvida')
DIR_PDFS_PROCESSADOS = cfg.get('CONFIG','dir_pdf_resolvida')
DIR_PDF_CLARO = cfg.get('CONFIG','dir_pdf_claro')

MENU_ERP_X = cfg.getint('MENU_ERP', 'X')
MENU_ERP_Y = cfg.getint('MENU_ERP', 'Y')

if(True == True):
    qtd_sistemas = 1
    for z in range(qtd_sistemas):
        time.sleep(1)
        pyautogui.hotkey('super', 'd')
        time.sleep(2)

        pyautogui.hotkey('winleft', 'r')
        pyautogui.write(EXE_PROCFIT)
        pyautogui.press('enter')
        time.sleep(10)

        # LOGIN USUARIO
        pyautogui.press('tab')
        pyautogui.press('delete', presses=253)
        pyautogui.write(USUARIO_PROCFIT)
        pyautogui.press('tab')
        pyautogui.press('delete', presses=253)
        pyautogui.write(SENHA_PROCFIT)
        pyautogui.press('tab')
        pyautogui.press('enter')
        time.sleep(10)

        pyautogui.moveTo(MENU_ERP_X, MENU_ERP_Y)
        pyautogui.mouseDown()
        pyautogui.mouseUp()
        time.sleep(5)

        pyautogui.press('right', presses = 12)
        pyautogui.press('enter')
        time.sleep(5)
        
        #leitor de pdf
        fileExt = r".pdf"
        arquivos = [path.join(DIR_PDF, nome) for nome in listdir(
            DIR_PDF) if nome.endswith(fileExt) or nome.endswith(fileExt.upper())]
        for arq in arquivos:
            try:
                dados = extract_text_from_pdf_equatorial(arq)
            except:
                try:
                    dados = extract_text_from_pdf_energisa(arq)
                except:
                    dados = extract_text_from_pdf_neoenergia(arq)
                    
            print(dados)  
            dados = list(dados)     
            file_name = arq.split("[")[-1][0:15]
            print(file_name[14])
            if file_name[14] == ']':                                    #se o 13 charactere do nome for ] dá ele pega os numeros.
                dados[1] = file_name[0:14]
                
                
            print(arq, len(arquivos))
            print(dados)
                     
            CLASS_FIN = cfg.get(dados[3], 'class_fin')
            SERIE= cfg.get(dados[3],'serie')
            NF_ESPECIE = cfg.get(dados[3],'nf_especie')
            NF_MODELO = cfg.get(dados[3],'nf_modelo')
            
                        
            SERVER = cfg.get('CONFIG','SERVER')
            DATABASE = cfg.get('CONFIG','DATABASE')
            USERNAME = cfg.get('CONFIG','USERNAME')
            PASSWORD = cfg.get('CONFIG','PASSWORD')
            DRIVER = cfg.get('CONFIG','DRIVER')

            CNPJ_DESTINATARIO = (f'{dados[1][0:2]}.{dados[1][2:5]}.{dados[1][5:8]}/{dados[1][8:12]}-{dados[1][12:15]}')
            CNPJ_EMITENTE = (f'{dados[3][0:2]}.{dados[3][2:5]}.{dados[3][5:8]}/{dados[3][8:12]}-{dados[3][12:15]}')
            NF_NUMERO = dados[2]
            NF_SERIE = SERIE
            
            if NF_NUMERO[0]== '0':
                NF_NUMERO = NF_NUMERO[1:len(NF_NUMERO)]

            sql = f'''
                SELECT B.INSCRICAO_FEDERAL, C.INSCRICAO_FEDERAL, A.NF_NUMERO, A.NF_ESPECIE, NF_SERIE 
                FROM NF_COMPRA A 
                JOIN ENTIDADES B ON A.ENTIDADE = B.ENTIDADE 
                JOIN ENTIDADES C ON A.EMPRESA  = C.ENTIDADE
                WHERE B.INSCRICAO_FEDERAL      = '{CNPJ_EMITENTE}'
                AND C.INSCRICAO_FEDERAL        = '{CNPJ_DESTINATARIO}'
                AND A.NF_NUMERO                = '{NF_NUMERO}'
                AND A.NF_ESPECIE               = '{NF_ESPECIE}'
                AND A.NF_SERIE                 = '{NF_SERIE}'
            '''
            
            try:
                print("Iniciando conexão com o DB")
                connectionString = f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
                conn = pyodbc.connect(connectionString)
                print("Conexão estabelecida")
                cursor = conn.cursor()

                print("Executando consulta SQL")
                cursor.execute(sql)
                records = cursor.fetchall()
                print("Exibindo resultado")
                print("")
                print("")
                print(sql)
                print("")
                print("")
                print(records)
                print("")
                print("")

                totalRegistros = len(records)
                if(totalRegistros > 0):
                    print("Nota já lançada")
                    src_path = os.path.join(arq)
                    dst_path = os.path.join(DIR_PDFS_PROCESSADOS)
                    shutil.move(src_path, dst_path)
                    print(f"{today} {current_time} Arquivo {arq} já processado.\n")
                    with open(r'C:/SEVEN/teste joao/logs.txt', "a") as arquivo:
                        arquivo.write(f"{today} {current_time} Arquivo {arq} já processado'.\n")
                    continue    
                else:
                    print("Nota NÃO lançada") 
                    print("")
                    print("")
                    print("Fechando conexão com o DB")
                    conn.close()
                    print("Conexão fechada")

                    pyautogui.hotkey('ctrl', 'i')
                    time.sleep(5)
                    pyautogui.press('tab', presses= 8)
                    pyautogui.press('enter')
                    time.sleep(5)
                    pyautogui.moveTo(980, 644)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(1)
                    pyautogui.moveTo(541, 111)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    #na ordem, 0 data, 1 cnpj do tomador, 2 numero da nota fiscal, 3 cnpj do prestador e 4 codigo de verificação, 5 valor de serviço
                    pyautogui.write((f'{dados[1][0:2]}.{dados[1][2:5]}.{dados[1][5:8]}/{dados[1][8:12]}-{dados[1][12:15]}'))
                    time.sleep(5)
                    pyautogui.press('enter')
                    time.sleep(10)
                    pyautogui.press('down')
                    pyautogui.press('enter')
                    time.sleep(1)
                    pyautogui.press('tab', presses= 2)
                    pyautogui.press('enter')
                    time.sleep(5)
                    pyautogui.press('right', presses=2)
                    pyautogui.write((f'{dados[3][0:2]}.{dados[3][2:5]}.{dados[3][5:8]}/{dados[3][8:12]}-{dados[3][12:15]}'))
                    time.sleep(3)
                    pyautogui.press('enter')
                    time.sleep(10)
                    pyautogui.press('down')
                    pyautogui.press('enter')
                    time.sleep(3)
                    pyautogui.press('esc')
                    time.sleep(5)
                    pyautogui.hotkey('shift', 'tab')
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(5)
                    pyautogui.press('tab', presses= 5)
                    pyautogui.press('enter')
                    if NF_ESPECIE=='NFS':
                        pyautogui.press('down', presses=2)
                    if NF_ESPECIE =='NEE':
                        pyautogui.press('down', presses=8)
                    if NF_ESPECIE =='NF3':   
                        pyautogui.press('down', presses= 6)
                        
                    
                    pyautogui.press('enter')
                    pyautogui.press('down')
                    pyautogui.press('enter')
                    pyautogui.write(dados[2])
                    pyautogui.press('tab', presses=2)
                    pyautogui.press('enter')

                    if SERIE == 'E':
                        pyautogui.press('down', presses = 7)
                    elif SERIE == 'U':
                        pyautogui.press('down', presses = 2)
                    elif SERIE == 'B':
                        pyautogui.press('down', presses = 7)
                    else:
                        pyautogui.press('down')
                        pyautogui.press('end') 
                    
                    pyautogui.press('enter')
                    pyautogui.press('tab', presses=2)
                    
                    if NF_MODELO == 66:
                        pyautogui.press('down')
                        pyautogui.press('enter')
                    
                    pyautogui.press('tab')
                    pyautogui.write(dados[-1])
                    pyautogui.press('tab')
                    pyautogui.write(current_day.replace("/",""))
                    pyautogui.press('tab')
                    pyautogui.press('space')
                    pyautogui.press('tab')
                    pyautogui.write(dados[11])
                    time.sleep(2)

                    pyautogui.press('f2')
                    pyautogui.moveTo(421, 164)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(2)
                    pyautogui.press('tab', presses=38)
                    time.sleep(5)
                    pyautogui.write(dados[5].replace('.',''))
                    pyautogui.press('tab')
                    time.sleep(3)

                    pyautogui.press('f6')
                    pyautogui.press('tab')
                    pyautogui.write(dados[0])   
                    pyautogui.press('tab')
                    time.sleep(3)
                    pyautogui.write('100')
                    time.sleep(3)
                    pyautogui.press('tab')
                    time.sleep(3)
                    pyautogui.write(dados[5].replace('.',''))
                    pyautogui.press('tab', presses=9)
                    pyautogui.write(CLASS_FIN)
                    pyautogui.press('tab')
                    pyautogui.press('enter')
                    time.sleep(3)

                    pyautogui.press('f12')
                    pyautogui.write('01') 
                    pyautogui.press('tab', presses= 2)
                    pyautogui.write('3') 
                    pyautogui.press('tab', presses= 2)
                    pyautogui.write('12') 
                    pyautogui.press('tab', presses= 2)
                    time.sleep(2)
                    pyautogui.write(dados[6])
                    pyautogui.press('tab')
                    time.sleep(2)
                    pyautogui.write(dados[7])
                    pyautogui.press('tab')
                    time.sleep(2)
                    pyautogui.write(dados[4])
                    pyautogui.press('tab')
                    time.sleep(2)
                    pyautogui.write(dados[7])
                    pyautogui.press('tab', presses= 3)
                    time.sleep(2)
                    pyautogui.write(dados[8])
                    pyautogui.press('tab')
                    time.sleep(2)
                    pyautogui.write(dados[9])
                    pyautogui.press('tab')
                    time.sleep(2)
                    pyautogui.write(dados[10])
                    time.sleep(1)
                        
                    pyautogui.hotkey('ctrl', 'g')
                    time.sleep(15)

                    pyautogui.moveTo(994, 76)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(5)         
                    pyautogui.press('enter')
                    time.sleep(5)  
                    
                    connectionString = f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
                    conn = pyodbc.connect(connectionString)
                    cursor = conn.cursor()
                    cursor.execute(sql)
                    records = cursor.fetchall()
                    time.sleep(2)
                    
                    if totalRegistros == 0:
                        print("Nota com problema")
                        src_path = os.path.join(arq)
                        dst_path = os.path.join(DIR_PDFS_N_PROCESSADOS)
                        shutil.move(src_path, dst_path)
                        print(f"Arquivo '{arq}' movido para '{DIR_PDFS_N_PROCESSADOS}'.")
                        with open(r'C:/SEVEN/teste joao/logs.txt', "a") as arquivo:
                            arquivo.write(f"'{today}, {current_time} ' Arquivo '{arq}' não está no banco de dados\n")
                        pyautogui.hotkey('alt', 'f4')
                        pyautogui.hotkey('alt', 'f4')
                        handle = Popen(r"C:\SEVEN\teste joao\lancamento_servicos_procfit.exe", creationflags=CREATE_NEW_CONSOLE)
                        exit()
                    else:              
                        src_path = os.path.join(arq)
                        dst_path = os.path.join(DIR_PDFS_PROCESSADOS)
                        shutil.move(src_path, dst_path)
                        print(f"Arquivo '{arq}' movido para '{DIR_PDFS_PROCESSADOS}'.")
                        with open(r'C:/SEVEN/teste joao/logs.txt', "a") as arquivo:
                            arquivo.write(f"{today} {current_time} Arquivo {arq} foi cadastrado corretamente!\n")
                            
            except Exception as e:
                print(e)  
                              
        fileExt = r".pdf"
        arquivos = [path.join(DIR_PDF_CLARO, nome) for nome in listdir(
            DIR_PDF_CLARO) if nome.endswith(fileExt) or nome.endswith(fileExt.upper())]
        for arq in arquivos:
            dados = extract_text_from_pdf_claro(arq)
            print(arq, len(arquivos))
            

            CLASS_FIN = cfg.get(dados[3], 'class_fin')
            SERIE= cfg.get(dados[3],'serie')
            NF_ESPECIE = cfg.get(dados[3],'nf_especie')
            NF_MODELO = cfg.get(dados[3],'nf_modelo')
                     
            SERVER = '192.168.51.9'
            DATABASE = 'PBS_PERMANENTE_DADOS'
            USERNAME = 'SEVEN.CONTABIL'
            PASSWORD = 'S3v3N.C0nt@biL@2023!*@'
            DRIVER = 'SQL Server'

            CNPJ_DESTINATARIO = (f'{dados[1][0:2]}.{dados[1][2:5]}.{dados[1][5:8]}/{dados[1][8:12]}-{dados[1][12:15]}')
            CNPJ_EMITENTE = (f'{dados[3][0:2]}.{dados[3][2:5]}.{dados[3][5:8]}/{dados[3][8:12]}-{dados[3][12:15]}')
            NF_NUMERO = dados[2]
            NF_SERIE = SERIE

            sql = f'''
                SELECT B.INSCRICAO_FEDERAL, C.INSCRICAO_FEDERAL, A.NF_NUMERO, A.NF_ESPECIE, NF_SERIE 
                FROM NF_COMPRA A 
                JOIN ENTIDADES B ON A.ENTIDADE = B.ENTIDADE 
                JOIN ENTIDADES C ON A.EMPRESA  = C.ENTIDADE
                WHERE B.INSCRICAO_FEDERAL      = '{CNPJ_EMITENTE}'
                AND C.INSCRICAO_FEDERAL        = '{CNPJ_DESTINATARIO}'
                AND A.NF_NUMERO                = '{NF_NUMERO}'
                AND A.NF_ESPECIE               = '{NF_ESPECIE}'
                AND A.NF_SERIE                 = '{NF_SERIE}'
            '''

            try:
                print("Iniciando conexão com o DB")
                connectionString = f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
                conn = pyodbc.connect(connectionString)
                cursor = conn.cursor()

                cursor.execute(sql)
                records = cursor.fetchall()
                print(sql)
                print("")
                print("")

                totalRegistros = len(records)
                if(totalRegistros > 0):
                    print("Nota já lançada")
                    src_path = os.path.join(arq)
                    dst_path = os.path.join(DIR_PDFS_PROCESSADOS)
                    shutil.move(src_path, dst_path)
                    print(f"{today} {current_time} Arquivo {arq} já processado.\n")
                    with open('logs.txt', "a") as arquivo:
                        arquivo.write(f"{today} {current_time} Arquivo {arq} já processado'.\n")
                    continue    
                else:
                    print("Nota NÃO lançada") 
                    print("")
                    print("")
                    print("Fechando conexão com o DB")
                    conn.close()
                    print("Conexão fechada")

                    pyautogui.hotkey('ctrl', 'i')
                    time.sleep(5)
                    pyautogui.press('tab', presses= 8)
                    pyautogui.press('enter')
                    time.sleep(5)
                    pyautogui.moveTo(980, 644)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(1)
                    pyautogui.moveTo(541, 111)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    #na ordem, 0 data, 1 cnpj do tomador, 2 numero da nota fiscal, 3 cnpj do prestador e 4 codigo de verificação, 5 valor de serviço
                    pyautogui.write((f'{dados[1][0:2]}.{dados[1][2:5]}.{dados[1][5:8]}/{dados[1][8:12]}-{dados[1][12:15]}'))
                    time.sleep(5)
                    pyautogui.press('enter')
                    time.sleep(3)
                    pyautogui.press('down')
                    pyautogui.press('enter')
                    time.sleep(1)
                    pyautogui.press('tab')
                    pyautogui.write('996394')
                    pyautogui.press('tab')
                    time.sleep(1)
                    pyautogui.press('esc')
                    time.sleep(5)
                    pyautogui.hotkey('shift', 'tab')
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(5)
                    pyautogui.press('tab', presses= 5)
                    pyautogui.press('enter')
                    time.sleep(1)
                    if NF_ESPECIE=='NFS':
                        pyautogui.press('down', presses=2)
                    if NF_ESPECIE =='NEE':
                        pyautogui.press('down', presses=8)
                    if NF_ESPECIE =='NF3':   
                        pyautogui.press('down', presses= 6)
                        
                    pyautogui.press('enter')
                    pyautogui.press('down')
                    pyautogui.press('enter')
                    time.sleep(1)
                    pyautogui.press('tab')
                    pyautogui.write(dados[2])
                    pyautogui.press('tab', presses=2)
                    pyautogui.press('enter')
                    time.sleep(1)

                    if SERIE == 'E':
                        pyautogui.press('down', presses = 7)
                    elif SERIE == 'U':
                        pyautogui.press('down', presses = 2)
                    elif SERIE =='B':
                        pyautogui.press('down', presses = 13)
                    else:
                        pyautogui.press('down')
                        pyautogui.press('end')
                        
                    pyautogui.press('enter')
                    pyautogui.press('tab', presses=2)
                    time.sleep(1)
                        
                    if NF_MODELO == 22:
                        pyautogui.press('down', presses= 4)
                        pyautogui.press('enter')
                        
                    pyautogui.press('tab')
                    pyautogui.write(dados[-1])
                    time.sleep(1)
                    pyautogui.press('tab')
                    pyautogui.write(current_day.replace("/",""))
                    pyautogui.press('tab')
                    pyautogui.press('space')
                    pyautogui.press('tab')
                    pyautogui.write(dados[-2])

                    pyautogui.press('f2')
                    pyautogui.moveTo(538, 161)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(2)
                    pyautogui.press('tab', presses=38)
                    time.sleep(2)
                    pyautogui.write(dados[4].replace('.',''))
                    pyautogui.press('tab')

                    pyautogui.press('f6')
                    pyautogui.press('tab')
                    pyautogui.write(dados[0])   
                    pyautogui.press('tab')
                    pyautogui.write('100')
                    pyautogui.press('tab', presses=10)
                    pyautogui.write(CLASS_FIN)
                    pyautogui.press('tab')
                        
                    pyautogui.hotkey('ctrl', 'g')
                    time.sleep(15)
                                
                    pyautogui.moveTo(994, 76)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(5)         
                    pyautogui.press('enter')
                    time.sleep(5)  
                    
                    conn = pyodbc.connect(connectionString)
                    cursor.execute(sql)
                    records = cursor.fetchall()
                    totalRegistros = len(records)
                    time.sleep(2)
                    
                    if totalRegistros == 0:
                        print("Nota com problema")
                        src_path = os.path.join(arq)
                        dst_path = os.path.join(DIR_PDFS_N_PROCESSADOS)
                        shutil.move(src_path, dst_path)
                        print(f"Arquivo '{arq}' movido para '{DIR_PDFS_N_PROCESSADOS}'.")
                        with open('logs.txt', "a") as arquivo:
                            arquivo.write(f"'{today}, {current_time} ' Arquivo '{arq}' não está no banco de dados")
                        pyautogui.press('esc')
                        pyautogui.hotkey('alt', 'f4')
                        handle = Popen("C:\SEVEN\teste joao\lancamento_servicos_procfit.exe", creationflags=CREATE_NEW_CONSOLE)
                        print('falta ver o [WinError 5] Acesso negado')
                        exit()
                    if totalRegistros > 0:              
                        src_path = os.path.join(arq)
                        dst_path = os.path.join(DIR_PDFS_PROCESSADOS)
                        shutil.move(src_path, dst_path)
                        with open('logs.txt', "a") as arquivo:
                            arquivo.write(f"{today} {current_time} Arquivo {arq} foi cadastrado corretamente!\n")
                    
                        pyautogui.moveTo(994, 76)
                        pyautogui.mouseDown()
                        pyautogui.mouseUp()
                        time.sleep(2)
                        
                        pyautogui.press('right', presses=7)
                        pyautogui.press('enter')
                        time.sleep(2)
                        
                        pyautogui.hotkey('ctrl', 'i')
                        pyautogui.press('tab')
                        pyautogui.write('13')
                        pyautogui.press('tab', presses= 5)
                        pyautogui.press('enter')
                        time.sleep(2)
                        pyautogui.press('right', presses=2)
                        pyautogui.write((f'{dados[3][0:2]}.{dados[3][2:5]}.{dados[3][5:8]}/{dados[3][8:12]}-{dados[3][12:15]}'))
                        time.sleep(3)
                        pyautogui.press('enter')
                        time.sleep(5)
                        pyautogui.press('down')
                        pyautogui.press('enter')
                        time.sleep(3)
                        pyautogui.press('esc')
                        pyautogui.press('tab')
                        pyautogui.write(dados[6])
                        
                        pyautogui.press('f1')
                        pyautogui.press('tab', presses= 9)
                        pyautogui.write(CLASS_FIN)
                        pyautogui.press('tab')
                        
                        pyautogui.press('f2')
                        time.sleep(2)
                        pyautogui.write(dados[0])
                        pyautogui.press('tab')
                        pyautogui.write(dados[2])
                        pyautogui.press('tab', presses= 3)
                        pyautogui.write('1')
                        pyautogui.press('tab')
                        time.sleep(1)
                        
                        pyautogui.hotkey('ctrl', 'g')
                        time.sleep(15)         
                                         
            except Exception as e:
                print(e)        
                
        
    pyautogui.moveTo(1005, 706)
    pyautogui.mouseDown()
    pyautogui.mouseUp()
    pyautogui.press('right')
    pyautogui.press('enter')
    print('script rodado completamente')
sys.exit()  