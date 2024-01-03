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
from lancador_equatorial import extract_text_from_pdf_equatorial

#alowlaoslaoslalslasol

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

# FECHAR_X = cfg.getint('FECHAR', 'X')
# FECHAR_Y = cfg.getint('FECHAR', 'Y')

pyautogui.PAUSE = 0.5

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
        
        fileExt = r".xml"
        arquivos = [path.join(DIR_XMLS, nome) for nome in listdir(
            DIR_XMLS) if nome.endswith(fileExt)]
        for arq in arquivos:
            dados = converterXmlServicoToJson(arq)
            print(arq, len(arquivos))
            try:
                SERVICO = cfg.get(dados[3], 'servico')
                OPERACAO = cfg.get(dados[3], 'operacao')
                CLASS_FIN = cfg.get(dados[3], 'class_fin')
                F6 = cfg.get(dados[3], 'f6')
                SERIE= cfg.get(dados[3],'serie')
                NF_ESPECIE = cfg.get(dados[3],'nf_especie')
            except:
                SERVICO = 43
                OPERACAO = 152
                CLASS_FIN = []
                F6 = 'False'
                SERIE = 0
                NF_ESPECIE = 'NFS'
                NF_SERIE = 0
            try:
                RETENCAO = cfg.get(dados[3], 'retencao')
            except:
                RETENCAO = 'false'
                          
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
                    dst_path = os.path.join(DIR_XMLS_PROCESSADAS)
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
                    time.sleep(3)
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
                    time.sleep(5)
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
                    pyautogui.press('down', presses=2)
                    pyautogui.press('enter')
                    pyautogui.press('tab')
                    pyautogui.write(dados[2])
                    pyautogui.press('tab', presses=2)
                    pyautogui.press('enter')
                    pyautogui.press('down')
                    if SERIE == 'E':
                        pyautogui.press('down', presses = 6)
                    #falta colocar os outros steps
                    else:
                        pyautogui.press('end')
                        
                    pyautogui.press('enter')
                    pyautogui.press('tab', presses=3)
                    pyautogui.write(dados[0])
                    pyautogui.press('tab')
                    pyautogui.write(dados[0])
                    pyautogui.press('tab')
                    pyautogui.press('space')
                    pyautogui.press('tab')
                    pyautogui.press('f1')
                    time.sleep(5)
                    
                    pyautogui.press('tab', presses=28)
                    time.sleep(5)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(5)
                    pyautogui.press('tab')
                    time.sleep(3)
                    
                    
                    pyautogui.press('f3')
                    time.sleep(3)
                    pyautogui.write(str(SERVICO))
                    pyautogui.press('tab')
                    pyautogui.write(str(OPERACAO))
                    pyautogui.press('tab')
                    if RETENCAO == 'true':
                        pyautogui.write('S')
                    pyautogui.press('tab')
                    if RETENCAO == 'true' and dados[5].replace(".",",") > 499.99:
                        pyautogui.write('S')
                    pyautogui.press('tab', presses=3)
                    pyautogui.write('1')
                    pyautogui.press('tab', presses=7)
                    pyautogui.press('backspace')
                    pyautogui.write(dados[5].replace(".",","))
                    pyautogui.press('tab', presses=3)
                    pyautogui.write('1009') 
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('up')
                    pyautogui.moveTo(571, 520)
                    pyautogui.mouseDown()
                    pyautogui.mouseUp()
                    time.sleep(2)
                    
                    pyautogui.press('f4')
                    pyautogui.press('tab', presses=11)
                    pyautogui.press('del', presses=4)
                    pyautogui.write(dados[5].replace(".",","))
                    pyautogui.press('tab')
                    pyautogui.write(dados[5].replace(".",","))
                    time.sleep(2)
                    
                    pyautogui.press('f5')
                    time.sleep(3)
                    pyautogui.press('tab')
                    pyautogui.press('tab', presses=2)
                    pyautogui.press('space', presses=2)
                    time.sleep(2)
                    time.sleep(5)
                    
                    if F6 == 'true':
                        pyautogui.press('f6')
                        pyautogui.press('tab')
                        if dados[3] == '02535864000133':
                            pyautogui.write(f'20{current_day[3:5]}{current_day[6:11]}')
                        else:
                            pyautogui.write(f'{current_day[0:2]}{current_day[3:5]}{current_day[6:11]}')
                        pyautogui.press('tab')
                        pyautogui.write('100')
                        pyautogui.press('tab', presses=10)
                        pyautogui.write(CLASS_FIN)
                        pyautogui.press('tab')
                        
                    
                    pyautogui.press('f11')
                    pyautogui.write(dados[4])
                    time.sleep(2)
                    
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
                        dst_path = os.path.join(DIR_XML_N_PROCESSADAS)
                        shutil.move(src_path, dst_path)
                        print(f"Arquivo '{arq}' movido para '{DIR_XML_N_PROCESSADAS}'.")
                        with open(r'C:/SEVEN/teste joao/logs.txt', "a") as arquivo:
                            arquivo.write(f"'{today}, {current_time} ' Arquivo '{arq}' não está no banco de dados")
                        pyautogui.press('esc')
                        pyautogui.hotkey('alt', 'f4')
                        handle = Popen(r"C:\SEVEN\teste joao\lancamento_servicos_procfit.exe", creationflags=CREATE_NEW_CONSOLE)
                        exit()
                    else:              
                        src_path = os.path.join(arq)
                        dst_path = os.path.join(DIR_XMLS_PROCESSADAS)
                        shutil.move(src_path, dst_path)
                        print(f"Arquivo '{arq}' movido para '{DIR_XMLS_PROCESSADAS}'.")
                        with open(r'C:/SEVEN/teste joao/logs.txt', "a") as arquivo:
                            arquivo.write(f"{today} {current_time} Arquivo {arq} foi cadastrado corretamente!\n")
                                                  
            except Exception as e:
                print(e)   
                                       
    pyautogui.moveTo(1005, 706)
    pyautogui.mouseDown()
    pyautogui.mouseUp()
    pyautogui.press('right')
    pyautogui.press('enter')
    print('script rodado completamente')
sys.exit()

