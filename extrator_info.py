def ExtratordeInfo2(file):
    import os
    import csv  
    import pdfplumber
    import pandas as pd
    
    with open('extração_de_dados.csv','w') as fp:
        fp.write('cnpj_prestador,razao_social,uf_serv,endereco,n_nota,serie,data,VALOR_SERVIÇO ,DESCONTO_INCONDICIONAL,VALOR_SERVIÇO,DEDUÇÕES,VALOR_CONTÁBIL,BASE_CÁLCULO,ALIQUOTA_ISS,ISS,PIS,COFINS,CSLL,INSS\n')
    fileExt = r".pdf"
    files = os.listdir(file)
    pdf_files = [arq for arq in files if arq.lower().endswith('.pdf')]
    for pdf_file in pdf_files:
        try:
            with pdfplumber.open(f'{file}/{pdf_file}') as pdf:
                text = ""
                cnpj_prestador = 0
                naprox = False
                csll_prox = False
                ded_prox = False
                end = 0
                rs=0
                uf=False
                serie = '-'
                CRF = '-'
                for page in pdf.pages:
                    text += page.extract_text()
                    a = text
                    lines = a.split('\n')
                    for line in lines: 
                        if 'NOTA FISCAL DE SERVIÇOS ELETRÔNICA - NFSe' == lines[0]:
                            if uf==True:
                                uf_serv=line.split('-')[1]
                                uf=False
                            if ded_prox == True:
                                VALOR_SERVIÇO = line.split(' ')[0].replace(',','.')
                                DEDUÇÕES = line.split(' ')[1].replace(',','.')
                                DESCONTO_INCONDICIONAL = line.split(' ')[2].replace(',','.')
                                BASE_CÁLCULO = line.split(' ')[3].replace(',','.')
                                
                                
                                ALIQUOTA_ISS = line.split(' ')[4].replace(',','.')
                                ISS = line.split(' ')[5].replace(',','.')
                                ded_prox = False
                            if csll_prox == True:
                                INSS = line.split(' ')[0].replace(',','.')
                                IR = line.split(' ')[1].replace(',','.')
                                CSLL = line.split(' ')[2].replace(',','.')
                                COFINS = line.split(' ')[3].replace(',','.')
                                PIS = line.split(' ')[4].replace(',','.')
                                csll_prox = False
                            if naprox == True:
                                n_nota = line.split(' ')[0]
                                naprox=False
                            if 'CPF/CNPJ:' in line:
                                cnpj_tomador = line.split(' ')[-1]
                                if cnpj_prestador == 0:
                                    cnpj_prestador = line.split(' ')[-1]
                            if 'Emitido em' in line:
                                data = line.split(' ')[2] 
                            if 'Exigível Tributacao Normal' in line:
                                naprox=True
                            if 'CSLL (R$)' in line:
                                csll_prox = True
                                
                            if 'DEDUÇÕES' in line:
                                ded_prox = True
                            
                            if 'Endereço:' in line:
                                end = end+1
                                if end ==2:
                                    endereco=line.split(':')[1].replace(',','')
                                    uf=True
                                    
                            if 'Razão Social:' in line:
                                rs=rs+1
                                if rs==1:
                                    razao_social=line.split(':')[1]
                    dados = cnpj_prestador,razao_social,uf_serv,endereco,n_nota,serie,data,VALOR_SERVIÇO ,DESCONTO_INCONDICIONAL,VALOR_SERVIÇO,DEDUÇÕES,VALOR_SERVIÇO,BASE_CÁLCULO,ALIQUOTA_ISS,ISS,PIS,COFINS,CSLL,INSS
                    dados=str(dados)
                    with open('extração_de_dados.csv','a') as fp:
                        fp.write(f'{dados}\n')
                    print(f'lido {pdf_file} com sucesso')
        except:
            try:
                with pdfplumber.open(f'{file}/{pdf_file}') as pdf:
                    for page in pdf.pages:
                        text += page.extract_text()
                        a = text
                        cnpj = 0
                        razao = 0
                        endereco=0
                        uf = 0
                        serie = 0
                        situacao = 0
                        descontos = 0
                        prox = False
                        prox1 = 0
                        rpt = False
                        lines = a.split('\n')
                        for line in lines: 
                            #print(line)
                            if lines[0][0:4] == 'N° :':
                                if rpt == True:
                                    COFINS = line.split("R$")[1]
                                    CSLL =line.split("R$")[2]
                                    INSS =line.split("R$")[3]
                                    IRPJ =line.split("R$")[4]
                                    PIS =line.split("R$")[5]
                                    rpt = False
                                if prox == True:
                                    if prox1 == 0:
                                        deducoes = line.split("R$")[1]
                                        Base_calculo =line.split("R$")[2]
                                        aliquota =line.split("R$")[3]
                                        ISS =line.split("R$")[4]
                                        prox1 = prox+1
                                    prox=False
                                n_nota = lines[0].split(' ')[2].strip()
                                if 'CPF/CNPJ:' in line:
                                    if cnpj == 0:
                                        cnpj = line.split(':')[1].strip()
                                if 'Nome/Razão Social' in line:
                                    if razao == 0:
                                        razao = line.split(':')[1].strip()
                                if 'Endereço:' in line:
                                    if endereco == 0:
                                        endereco = line.split(':')[1].split(',')[0].strip().replace(':','')
                                if 'UF:' in line:
                                    if uf ==0:
                                        uf = line.split(' ')[3]
                                        municipio = line.split(' ')[1]
                                if 'Data emissão:' in line:
                                    data = line.split(' ')[2]
                                if 'VALOR R$ ' in line:
                                    VALOR_SERVIÇO = line.split(' ')[3]
                                if 'Valor Total das Deduções' in line:
                                    prox = True
                                if 'COFINS CSLL INSS IRPJ PIS' in line:
                                    rpt = True  
                        dados = cnpj,razao,uf,municipio,endereco,n_nota,serie,data,serie,VALOR_SERVIÇO ,deducoes,VALOR_SERVIÇO,Base_calculo,aliquota,ISS,PIS,COFINS,CSLL,IRPJ,INSS
                        dados=str(dados)
                        with open('extração_de_dados.csv','a') as fp:
                            fp.write(f'{dados}\n')
                        print(f"lido {pdf_file} com sucesso")
            except:       
                try:
                    with pdfplumber.open(f'{file}/{pdf_file}') as pdf:
                        text = ""
                        prox = False
                        razao = 0
                        outra = False
                        uf = 0
                        notaprox = False
                        n_nota = 0
                        situacao = 0
                        prox_val = False
                        proxiss = False
                        INSS = 0
                        aplicada=False
                        for page in pdf.pages:
                            text += page.extract_text()
                            a = text
                            lines = a.split('\n')
                            for line in lines: 
                                #print(line)
                                if lines[1][0:10] == 'DANFSev1.0':
                                    if aplicada == True:
                                        ISS = line.split(' ')[0]
                                        aliquota = line.split(' ')[1]
                                        aplicada = False
                                    if proxiss == True:
                                        IRRF=line.split(' ')[0]
                                        CP=line.split(' ')[0]
                                        CSLL=line.split(' ')[0]
                                        PIS=line.split(' ')[1]
                                        COFINS=line.split(' ')[1]
                                        proxiss=False
                                    if prox_val == True:
                                        VALOR_SERVIÇO= line.split(' ')[0]
                                        descontos=line.split(' ')[1]
                                        deducoes=line.split(' ')[2]
                                        Base_calculo=line.split(' ')[3]
                                        prox_val=False
                                    if notaprox == True:
                                        if n_nota==0:
                                            n_nota = line.split(' ')[0]
                                            data = line.split(' ')[1]
                                            notaprox = False
                                    if outra == True:
                                        if uf == 0:
                                            uf = line.split('-')[1].split(' ')[0]
                                            endereco = line.split(' ')[0]
                                            municipio = line.split(' ')[1].split('-')[0]
                                        outra = False
                                    if prox == True:
                                        if razao == 0:
                                            razao = line.split(' ')[0]
                                        prox = False
                                        line.split
                                    if 'PrestadordoServiço' in line:
                                        cnpj = line.split(' ')[1]
                                    if 'Nome/NomeEmpresarial E-mail' in line:
                                        prox = True
                                    if 'Endereço Município CEP' in line:
                                        outra = True
                                    if 'NúmerodaNFS-e' in line:
                                        notaprox = True
                                    if 'autenticidadedestaNFS-' in line:
                                        serie = line.split(' ')[1]
                                    if 'TotalDeduções/Reduções CálculodoBM' in line:
                                        prox_val = True
                                    if 'IRRF,CP,CSLL-Retidos' in line:
                                        proxiss=True   
                                    if 'AlíquotaAplicada' in line:
                                        aplicada = True
                            dados = cnpj,razao,uf,endereco,n_nota,serie,data,VALOR_SERVIÇO ,descontos,VALOR_SERVIÇO,deducoes,VALOR_SERVIÇO,Base_calculo,aliquota,ISS,PIS,COFINS,CSLL,INSS
                            dados=str(dados)
                            with open('extração_de_dados.csv','a') as fp:
                                fp.write(f'{dados}\n')
                            print(f'lido {pdf_file} com sucesso')
  
                except:
                    print(f'Não foi possível ler o arquivo {pdf_file}')