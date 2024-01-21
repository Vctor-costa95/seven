import os
import csv  
import pdfplumber
import pandas as pd

def leitor(dir):
    csv_file_path = dir + '/pdf_analise.csv'
    colunas = ('chave','cnpj_prestador','cnpj_tomador','valor_servico','numero_nota','data','email_tomador','email_prestador','empresa_tomador','empresa_prestador','codigo_tributacao','IRRF,CP,CSLL')            
    with open(csv_file_path,'w',newline='') as f:
        csv_writer = csv.writer(f)
        csv_writer.writerow(colunas)
    try:   
        fileExt = r".pdf"
        files = os.listdir(dir)
        pdf_files = [arq for arq in files if arq.lower().endswith('.pdf')]
        for pdf_file in pdf_files:
            with pdfplumber.open(f"{dir}/{pdf_file}") as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text()
                    a = text
                    lines = a.split('\n')
                    cda=False
                    tds=False
                    vds=False
                    nfs=False
                    ne=False
                    em=0
                    cdt=False
                    counter=0
                    p_irff = False
                    for line in lines:
                        if p_irff == True:
                            irrf = line.split(' ')[0]
                            p_irff = False
                        if cdt==True:
                            counter=counter+1
                            if counter==1:
                                codigo_tributacao=line.split(' ')[0].split('-')[0]
                                desc1=line.split(' ')[1]
                                
                            if counter==2:
                                desc2=line.split(' ')[0]
                                desc=desc1+desc2
                                cdt=False
                            cdt=cdt+1
                            
                        if ne==True:
                            if em == 0:
                                email_prestador=line.split(' ')[1]
                                empresa_tomador = line.split('-')[0]
                                em =1
                            email_tomador=line.split(' ')[1]
                            empresa_prestador = line.split(' ')[0]
                            ne=False
                        if nfs == True:
                            numero_nota= line.split(' ')[0]
                            data = line.split(' ')[1]
                            nfs=False
                        if vds == True:
                            valor_servico=line.split(' ')[0]
                            vds=False
                        if tds== True:
                            cnpj_tomador = line.split(' ')[0]
                            tds = False
                        if cda == True:
                            chave = line.replace(' ','')[0:44]
                            cda=False
                        if 'ChavedeAcessodaNFS' in line:
                            cda=True
                        
                        if 'PrestadordoServiço' in line:
                            cnpj_prestador = line.split(' ')[1]
                        if 'TOMADORDOSERVIÇO' in line:
                            tds=True
                            
                        if 'ValordoServiço' in line:
                            vds=True
                            
                        if 'NúmerodaNFS' in line:
                            nfs=True
                            
                        if 'Nome/NomeEmpresarial' in line:
                            ne=True
                            
                        if 'CódigodeTributaçãoNacional' in line:
                            cdt=True
                            
                        if 'IRRF,CP,CSLL' in line:
                            p_irff=True
                    with open(csv_file_path, mode='a', encoding= 'utf8',newline='') as resultado_pesquisa:
                        writer = csv.writer(resultado_pesquisa)
                        linhas = (chave,cnpj_prestador,cnpj_tomador,valor_servico,numero_nota,data,email_tomador,email_prestador,empresa_tomador,empresa_prestador,codigo_tributacao,irrf)
                        writer.writerow(linhas)
    except:
        print('não foi possível ler todos os dados')
            
