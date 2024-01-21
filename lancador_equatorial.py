import sys
from os import path

import pandas as pd
import pdfplumber
import PyPDF2

if getattr(sys, 'frozen', False):
    application_path = path.dirname(sys.executable)
elif __file__:
    application_path = path.dirname(__file__)


def extract_text_from_pdf_equatorial(arq):
        with open(arq, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                lines = text.split('\n')
                    
                data_leitura_anterior = ""
                data_leitura_atual = ""
                numero_dias = ""
                data_leitura_proxima = ""
                conta_contrato = ""
                total_pagar = ""
                vencimento = ""
                emissao = ""
                cnpj_prestador = ""
                numero_nota = ""
                leitura_atual = ""
                leitura_anterior = ""
                total_kwh = ""
                conta_mes = ""
                consumo_total = ""
                chave = ""
                
                chave_de_acesso = False
                cc = False
                tp = False
                venc = False
                cl = False
                
                cont = 0
                cont_leituras = 0
                
                for line in lines:
                    cont += 1
                    
                    if chave_de_acesso == True:
                        chave = line
                        chave_de_acesso = False
                    
                    if cc:
                        cc = False
                        conta_contrato = line.strip()
                    elif tp:
                        tp = False
                        total_pagar = line.replace("Vencimento", "")
                        total_pagar = total_pagar.replace("R$", "")
                        
                        venc = True
                    elif venc:
                        venc = False
                        vencimento = line.replace("Conta Contrato", "")
                        
                    if(cont == 3):
                        cnpj_prestador = line.split(" ")
                        cnpj_prestador = cnpj_prestador[1]
                        
                    if(cl):
                        cont_leituras += 1
                        
                        if (cont_leituras == 3):
                            cl = False
                            linha_leituras = line.split(" ")
                            t = len(linha_leituras) - 1
                            leitura_anterior = linha_leituras[t-4]
                            leitura_atual = linha_leituras[t-3]
                            total_kwh = linha_leituras[t-1]

                    if "Leituras" in line:
                        datas = line.split(" ")
                        data_leitura_anterior = datas[1]
                        data_leitura_atual = datas[2]
                        numero_dias = datas[3]
                        data_leitura_proxima = datas[4]
                        
                    elif "Conta Contrato" in line:
                        cc = True
                    elif "Total a Pagar" in line:
                        tp = True       
                    elif "NOTA FISCAL Nº" in line:
                        conta_mes = line[0:7]
                        numero_nota = line.replace("NOTA FISCAL Nº", "")
                        numero_nota = numero_nota.split(" ")[1]
                        
                        
                    elif "Medidor Grandeza" in line:
                        cont_leituras += 1
                        cl = True
                    elif "DATA DE EMISSÃO:" in line:
                        emissao = line.replace("DATA DE EMISSÃO:", "").strip()
                        
                    elif "Consumo (kWh)" in line:
                        consumo_total = line.split(" ")
                        consumo_total = consumo_total[len(consumo_total) - 1]
                        
                    elif 'chave de acesso:' in line:
                        chave_de_acesso = True
                        
                vencimento = vencimento.replace('/','')
                cnpj_prestador = cnpj_prestador.replace('/','').replace('.','').replace('-','')
                total_pagar = total_pagar.replace('.','').replace(' ','')
                data_leitura_anterior = data_leitura_anterior.replace('/','')
                data_leitura_atual = data_leitura_atual.replace('/','')
                total_kwh = total_kwh.replace('.','')
                conta_mes = conta_mes.replace('/','')
                emissao = emissao.replace('/','')
                chave = chave.replace(' ','')
                
                lista = vencimento,numero_nota,cnpj_prestador,conta_contrato,total_pagar,data_leitura_anterior,data_leitura_atual,total_kwh,consumo_total,conta_mes,chave,emissao
                list_lista = list(lista)
                df = pd.read_excel("C:/SEVEN/teste joao/relacao_nome_matricula.xlsx")
                corresp = df.loc[df['MATRÍCULA']==int(lista[3])]['CNPJ']
                corresp = corresp.to_string()
                corresp = corresp.split(sep=' ')
                list_lista.insert(1,corresp[4].replace(".","").replace("/","").replace("-",""))
            return list_lista
        
    #problemas acima, tem 1 modelo q ta dando erro no leitura anterior. investigar  
    #na ordem, 0 data, 1 cnpj do tomador, 2 numero da nota fiscal, 3 cnpj do prestador e 4 codigo de verificação, 5 valor de serviço,6 leitura anterior, 7 leitura atual, 8 totale m kwh , 9 consumo acumulado, 10 conta mes, 11 chave, 12data de emissao
def extract_text_from_pdf_neoenergia(arquivo):
   with open(arquivo, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)
        c_a = False
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            lines = text.split('\n')
            
            for line in lines:
                if 'COMPANHIA ENERGÉTICA DE PERNAMBUCO AV.JOÃO DE BARROS' in line or 'INSCRIÇÃO ESTADUAL ' in line:
                    
                    try:
                        linea = line.split(",")
                        cnpj_prestador = linea[4].split(' ')[5]
                    except:
                        if 'INSCRIÇÃO ESTADUAL ' in line:
                            linea = line.split(" ")
                            cnpj_prestador = linea[0]
                        
                    
                if 'NOME DO CLIENTE: DROGATIM DROGARIAS LTDA CNPJ:' in line or 'NOME DO CLIENTE: COMERCIAL DRUGSTORE LTDA CNPJ:' in line:
                    cnpj_cliente = line.split(" ")
                    cnpj_cliente=cnpj_cliente[8]
                    
                if 'TOTAL A PAGAR R$' in line:
                    linha = line.split(' ')
                    data = linha
                    total_pagar = data[-2]
                    conta_mes = data[0]
                if 'DATAS DE LEITURAS' in line:
                    linha = line.split(' ')
                    leitura_anterior = linha[8]
                    leitura_atual = linha[13]
                    prox_leitura = linha[25]
                    n_nota = linha[30]
                    emissao = linha[39]
                    
                if 'HORÁRIOS ANTERIOR ATUAL MEDIDOR kWh' in line:
                    total_kwh = line.split(' ')[-6]
                    
                if 'TRIBUTO  BASE DE' in line:
                    consumo_acumulado = line.replace(' ','').split('ICMS')
                    c_a = True
                    
                if 'consulta chave de acesso:' in line:
                    chave = line.replace(' ','').split(':')[-4][0:44]
                    
                if 'TOTAL A PAGAR R$' in line:
                    codigo_cliente = line.split(' ')
                if c_a == True:
                    lst = (consumo_acumulado[1][0:8].split(',')[0],",",consumo_acumulado[1][0:8].split(',')[1][0:2])
                    consumo_acumulado_limpo = " ".join(lst).replace(' ','')
                    consumo_acumulado_limpo
                    c_a == False
                    
        try:
            cnpj_cliente[1]
        except:
            cnpj_cliente='cnpj_cliente'
                    
        return data[9].replace('/',''),cnpj_cliente.replace('/','').replace('.','').replace('-',''),n_nota,cnpj_prestador.replace('/','').replace('.','').replace('-',''),codigo_cliente[5],total_pagar.replace('.',''),leitura_anterior.replace('/',''),leitura_atual.replace('/',''),total_kwh.replace('.',''),consumo_acumulado_limpo.replace('.',''),conta_mes.replace('/',''),chave,emissao.replace('/','')
                
def extract_text_from_pdf_energisa(arquivo):
   with pdfplumber.open(arquivo) as pdf:
        text = ""
        cont_cc=False
        v_c = False
        carregarVencimento = False
        l_a = False
        lx = False
        t_n = False
        counter = 0  
        acres="" 
        tem_chave = False
        cnt = 0
        ct=0
        cod_cliente = 0
        conta = 0
        chave = ""
        ven = False
        cnpj_cliente=0
        for page in pdf.pages:
            text += page.extract_text()
            a = text
            lines = a.split(sep='\n')
            
            for line in lines:
                if cont_cc== True:
                    x=x+1
                    if x==1:
                        if cod_cliente== 0:
                            cod_cliente = line.split(' ')[-1]
                            cont_cc=False
                            
                if t_n == True:
                    ct= ct+1
                    if ct==3:
                        cod_cliente = line.split(' ')[-1]
                        t_n=False
                    
                if v_c:
                    valor = line.split(' ')[-1].replace('.','')
                    try:
                        int(vencimento)
                    except:
                        vencimento = line.split(' ')[1]
                    v_c = False

                if tem_chave:
                    if cnt ==1:
                        chave = chave1+line
                        tem_chave = False
                    chave1 = line.replace(' ','')
                    cnt = cnt+1                    
                    
                if ven == True:
                    vencimento = line.split(' ')[3]
                    ven = False
                    
                if carregarVencimento == True:
                    vencimento = line.split(' ')[1]
                    carregarVencimento = False
                    
                if chave[0:3] == 'htt':
                    conta = conta +1
                    if conta == 2:
                        chave = line.replace(' ','').replace('-','')
                    
                if l_a == True:
                    linha2 = line.split(' ')   
                    l_a=False
                    leitura_anterior = linha2[0].replace('/','')  
                    leitura_atual = linha2[1].replace('/','')
                    
                if '00:00:00' in line and l_a == False:
                    linha2 = line.split('00:00:00')
                    acres="1"   
                    leitura_anterior = linha2[0].replace('/','')  
                    leitura_atual = linha2[1].replace('/','')

  
                if 'Autorização:' in line:
                    ven = True     
                    
                if 'VENCIMENTO' in line:
                    carregarVencimento = True
                                      
                if 'CPF/CNPJ/RANI:' in line:
                    cnpj_cliente = line.split(':')[-1].replace(' ','').replace('/','').replace('-','').replace('.','')
                    v_c = True       
                    
                if 'NOTA FISCAL N' in line:
                    print(line)   
                    try:
                        n_nota = line.split(' ')[3].replace('.','')
                        if (bool(int(n_nota))) == False:
                            pass
                    except:
                        n_nota = line.split(' ')[5].replace('.','')
                    
                    try:
                        int(n_nota)
                    except:
                        n_nota = line.split(' ')[7].replace('.','')
                        
                    cod_cliente = line.split(' ')[1]
                    if cod_cliente == 'PAULO':
                        cod_cliente = line.split(' ')[3]
                        
                if 'Insc.Est.' in line:
                    cnpj_prestador = line.replace(' ','').split('Insc.Est.')[0][4:22].replace('/','').replace('-','').replace('.','') 
                    
                if f'DROGATIM DROGARIAS LTDA' in line or 'COMERCIAL DRUGSTORE LTDA' in line:
                    l_a=True
                    cont_cc = True
                    x=0
                    
                if 'Consumo em kWh' in line:
                    linha = line.split(' ')
                    total_kwh = line.split(' ')[4].replace('.','')
                    consumo_acumulado = line.split(' ')[6].replace('.','')
                    
                if 'TOTAL:' in line:
                    valor = line.split(' ')[1].replace('.','')
                  
                if 'DATA EMISSÂO' in line or 'DATA DE EMISSÃO:' in line:
                    emissao = line.split(':')[-1]
                    
                if 'Chave de Acesso' in line or 'chave de acesso' in line:
                    tem_chave = True
                     
                if 'COMERCIAL/COMERCIAL' in line:
                    t_n=True
                    
                # if 'INTERMARES' in line or 'TORRE' in line or 'MANGABEIRA' in line:
                #     if(cod_cliente == ''):
                #         cod_cliente = line.split(' ')[1]
                #         print(cod_cliente)
                   
            vencimento = vencimento.replace('/','')
            
            lista = vencimento,cnpj_cliente,n_nota,cnpj_prestador,cod_cliente.replace('-',''),valor,leitura_anterior,leitura_atual,total_kwh,consumo_acumulado,leitura_atual[2:12],chave[0:44],emissao.replace('/','')
            df = pd.read_excel(application_path + "\\relacao_nome_matricula.xlsx")            
            corresp = df.loc[df['MATRÍCULA']==(lista[4])]['CNPJ']
            corresp = corresp.to_string()
            if(corresp != 'Series([], )'):
                corresp = corresp.split(sep=' ')[4]
                lista = list(lista)
                lista[1]=corresp.replace('.','').replace('-','').replace('/','')
                lista[4]=lista[4].replace('/','')
            
                    
        return lista



def extract_text_from_pdf_claro(arquivo):
    with pdfplumber.open(arquivo) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()
            a = text
            lines = a.split(sep='\n')
            
            data_vencimento = ''
            
            vs = False
            cc = False
            tp = False
            venc = False
            cl = False
            
            cont = 0
            cont_leituras = 0
            
            for line in lines:
                cont += 1
                
                if vs:
                    linha = line.split(' ')
                    valor_servico = linha[-1]
                    vs = False
                
                if cc:
                    cc = False
                    conta_contrato = line.strip()
                elif tp:
                    tp = False
                    total_pagar = line.replace("Vencimento", "")
                    total_pagar = total_pagar.replace("R$", "")
                    
                    venc = True
                elif venc:
                    venc = False
                    vencimento = line.replace("Conta Contrato", "")
                    
                if(cont == 3):
                    data_vencimento = line.split(" ")
                    data_vencimento = data_vencimento[-1]
                    
                if(cont == 2):
                    valor_fatura = line.split(" ")
                    valor_fatura = valor_fatura[0]
                    
                if 'Número:' in line:
                    linha = line.split(" ")
                    nf_numero = linha[-3]
                    data_emissao = linha[-1]
            
                elif 'I.E.:' in line:
                    linha = line.split(' ')
                    cnpj_drug = linha[3]
                    
                elif 'CNPJ:' in line:
                    linha = line.split(' ')
                    cnpj_claro = linha[1]
                    
                elif 'BANDA LARGA ICMS' in line:
                    vs = True
                    
                elif 'C CE NP P:' in line:
                    linha = line.replace(' ','').split(':')
                    cnpj_certo = linha[2]
                    
                    
                
                           
        data_vencimento = data_vencimento.replace('/','')
        data_emissao = data_emissao.replace('/','')
        cnpj_drug = cnpj_drug.replace('/','').replace('.','').replace('-','')
        cnpj_claro = cnpj_claro.replace('/','').replace('.','').replace('-','')
        outros_valores = str(outros_valores).replace('.',',')
               
    return data_vencimento,cnpj_drug,nf_numero,'40432544015250',valor_servico,data_emissao        
# 0 vencimento, 1 cnpj da drugstore, 2 numero da nota fiscal, 3 cnpj claro, 4 valor do servico, 5 data de emissao, 6 valor resto     
                
                
#print(extract_text_from_pdf_energisa("D:\\temp\\[06198619003740]LOJA60ENERGISA.pdf"))

def extract_text_from_pdf_VFS(arquivo):
        with pdfplumber.open(arquivo) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text()
                a = text
                lines = a.split(sep='\n')
                if lines[1] == 'DESCRITIVO DE LOCAÇÃO':
                    desc_lotacao = True
                cnpj_vfs = 0
                data_vencimento = ''
                
                cont = 0
                cont_leituras = 0
                
                for line in lines:
                    cont += 1
                    if desc_lotacao == True:
                        if 'Data emissão' in line:
                            data_emissao = line.split(' ')[2]
                            d_e = line.split(' ')[2][3:10]
                            d_e_mes=d_e[0:2]
                            d_e_ano = d_e[2:8]
                            data_vencimento = f'01/0{int(d_e_mes)+1}{d_e_ano}'
                            
                        if 'CPF/CNPJ:' in line:
                            if cnpj_vfs == 0:
                                cnpj_vfs = line.replace(' ','').split(':')[1]
                            cnpj_drug = line.replace(' ','').split(':')[1][0:18]
                            
                        if 'VALOR R$' in line:
                            valor_servico = line.split(' ')[3]
                            
                        if 'N° :' in line:
                            nf_numero = line.split(' ')[2]
                            
                data_vencimento = data_vencimento.replace('/','')        
                cnpj_drug = cnpj_drug.replace('.','').replace('/','').replace('-','') 
                cnpj_vfs =  cnpj_vfs.replace('.','').replace('/','').replace('-','')  
                data_emissao = data_emissao.replace('/','') 
                    
        return data_vencimento,cnpj_drug,nf_numero,cnpj_vfs,valor_servico,data_emissao," "