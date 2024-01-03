import pandas as pd
from bs4 import BeautifulSoup
import json

def converterXmlServicoToJson(arquivoXml):
    with open(arquivoXml, 'r', encoding="utf8") as f:
        data = f.read()
    
    if data[1:5] == '?xml':

        Bs_data = BeautifulSoup(data, "xml")

        prefixo = "ns2:"

        # GERAL
        try:
            numero = Bs_data.find(prefixo+'Numero').text
        except:
            prefixo = ""
            numero = Bs_data.find(prefixo+'Numero').text

        data_emissao = Bs_data.find(prefixo+'DataEmissao').text
        data_emissao = data_emissao[8:10] +  \
            data_emissao[5:7] +  data_emissao[0:4]
        
        valor_liquido = Bs_data.find(prefixo+'ValorLiquidoNfse').text
        vl_liquido = str(valor_liquido).replace(",", "")
        vl_liquido = vl_liquido.replace(".", ",")
        
        cod_ver = Bs_data.find(prefixo + 'CodigoVerificacao').text
        val_ser = Bs_data.find(prefixo + 'ValorServicos').text

        valor_iss = Bs_data.find(prefixo+'ValorIss').text
        vl_iss = str(valor_iss).replace(",", "")
        vl_iss = vl_iss.replace(".", ",")

        percentual_aliquota = Bs_data.find(prefixo+'Aliquota').text
        pc_aliquota = str(percentual_aliquota).replace(",", "")
        pc_aliquota = pc_aliquota.replace(".", ",")

        valor_base_calculo = Bs_data.find(prefixo+'BaseCalculo').text
        vl_base_calculo = str(valor_base_calculo).replace(",", "")
        vl_base_calculo = vl_base_calculo.replace(".", ",")

        inf_declaracao_servico = Bs_data.find(
            prefixo+'InfDeclaracaoPrestacaoServico')
        competencia = inf_declaracao_servico.find(prefixo+'Competencia').text
        competencia = competencia[5:7]+competencia[8:10]+competencia[0:4]
        cod_situacao = 1 if Bs_data.find(
            prefixo+'InfPedidoCancelamento') is None else 2
        situacao = 'ATIVA' if cod_situacao == 1 else 'CANCELADA'

        # VALORES
        servico = Bs_data.find(prefixo+'Servico')
        iss_retido = servico.find(prefixo+'IssRetido').text

        valores = Bs_data.find(prefixo+'Valores')

        try:
            valor_deducoes = valores.find(prefixo+'ValorDeducoes').text
            vl_deducoes = str(valor_deducoes).replace(",", "")
            vl_deducoes = vl_deducoes.replace(".", ",")
        except:
            vl_deducoes = "0,00"

        try:
            valor_pis = valores.find(prefixo+'ValorPis').text
            vl_pis = str(valor_pis).replace(",", "")
            vl_pis = vl_pis.replace(".", ",")
        except:
            vl_pis = "0,00"

        try:
            valor_cofins = valores.find(prefixo+'ValorCofins').text
            vl_cofins = str(valor_cofins).replace(",", "")
            vl_cofins = vl_cofins.replace(".", ",")
        except:
            vl_cofins = "0,00"

        try:
            valor_inss = valores.find(prefixo+'ValorInss').text
            vl_inss = str(valor_inss).replace(",", "")
            vl_inss = vl_inss.replace(".", ",")
        except:
            vl_inss = "0,00"

        try:
            valor_ir = valores.find(prefixo+'ValorIr').text
            vl_ir = str(valor_ir).replace(",", "")
            vl_ir = vl_ir.replace(".", ",")
        except:
            vl_ir = "0,00"

        try:
            valor_csll = valores.find(prefixo+'ValorCsll').text
            vl_csll = str(valor_csll).replace(",", "")
            vl_csll = vl_csll.replace(".", ",")
        except:
            vl_csll = "0,00"

        # TOMADOR
        tomador_servico = Bs_data.find(prefixo+'TomadorServico')
        inf_tomador = tomador_servico.find(prefixo+'IdentificacaoTomador')
        try:
            cnpj_cnpj_tomador = inf_tomador.find(prefixo+'Cpf').text if inf_tomador.find(
                prefixo+'Cpf') is not None else inf_tomador.find(prefixo+'Cnpj').text
        except:
            cnpj_cnpj_tomador = ''
        rz_tomador = tomador_servico.find(prefixo+'RazaoSocial').text
        endereco_tomador_servico = tomador_servico.find(prefixo+'Endereco')
        uf_tomador = endereco_tomador_servico.find(prefixo+'Uf').text
        codigo_municipio_tomador = endereco_tomador_servico.find(
            prefixo+'CodigoMunicipio').text

        # PRESTADOR
        prestador_servico = Bs_data.find(prefixo+'PrestadorServico')
        rz_prestador = prestador_servico.find(prefixo+'RazaoSocial').text
        inf_prestador = Bs_data.find(prefixo+'Prestador')
        cnpj_cnpj_prestador = inf_prestador.find(prefixo+'Cpf').text if inf_prestador.find(
            prefixo+'Cpf') is not None else inf_prestador.find(prefixo+'Cnpj').text
        endereco_prestador_servico = prestador_servico.find(prefixo+'Endereco')
        codigo_municipio_prestador = endereco_prestador_servico.find(
            prefixo+'CodigoMunicipio').text
        if(codigo_municipio_prestador == "0"):
            codigo_municipio_prestador = "2704302"
            uf_prestador = "AL"
        else:
            uf_prestador = endereco_prestador_servico.find(prefixo+'Uf').text
            
        return data_emissao, cnpj_cnpj_tomador, numero, cnpj_cnpj_prestador, cod_ver, val_ser
    
    elif data[0:5] == '<NFe ':
        Bs_data = BeautifulSoup(data, "xml")
        data_emissao = Bs_data.find('DataEmissaoNFe').text
        data_emissao = data_emissao[8:10] +  \
            data_emissao[5:7] +  data_emissao[0:4]
            
        # TOMADOR
        tomador_servico = Bs_data.find('CPFCNPJTomador')
        cnpj_cnpj_tomador = tomador_servico.find('CPFCNPJTomador').text if tomador_servico.find(
        'CPF') is not None else tomador_servico.find('CNPJ').text
        try:
            numero = Bs_data.find('Numero').text
        except:
            prefixo = ""
            numero = Bs_data.find('NumeroNFe').text   
            if 'Numero' in Bs_data:
                numero = Bs_data.find('Numero').text 
  
        # PRESTADOR 
        prestador_servico = Bs_data.find('CPFCNPJPrestador')
        cnpj_cnpj_prestador = prestador_servico.find('Cpf').text if prestador_servico.find(
        'CPF') is not None else prestador_servico.find('CNPJ').text
        
        cod_ver = Bs_data.find('CodigoVerificacao').text
        
        val_ser = Bs_data.find('ValorServicos').text
        
        return data_emissao, cnpj_cnpj_tomador, numero, cnpj_cnpj_prestador, cod_ver, val_ser
        #na ordem, 0 data, 1 cnpj do tomador, 2 numero da nota fiscal, 3 cnpj do prestador e 4 codigo de verificação, 5 valor de serviço 
        
    elif data[0:5] == '<retu':
        Bs_data = BeautifulSoup(data, "xml")
        data_emissao = Bs_data.find('DataEmissao').text
        data_emissao = data_emissao[8:10] +  \
            data_emissao[5:7] +  data_emissao[0:4]
            
        # TOMADOR
        tomador_servico = Bs_data.find('IdentificacaoTomador')
        cnpj_cnpj_tomador = tomador_servico.find('CpfCnpj').text if tomador_servico.find(
        'CpfCnpj') is not None else tomador_servico.find('CNPJ').text
        cnpj_cnpj_tomador = cnpj_cnpj_tomador.strip()
        
        numero = Bs_data.find('IdentificacaoRps').text
        numero = numero.split(sep=' ')
        numero = numero[1]
        
        # PRESTADOR     
        prestador_servico = Bs_data.find('IdentificacaoPrestador')
        cnpj_cnpj_prestador = prestador_servico.find('Cpf').text if prestador_servico.find(
        'Cpf') is not None else prestador_servico.find('Cnpj').text
        cnpj_cnpj_prestador = cnpj_cnpj_prestador.strip()
        
        cod_ver = Bs_data.find('CodigoVerificacao').text
        
        val_ser = Bs_data.find('ValorServicos').text
        
        return data_emissao, cnpj_cnpj_tomador, numero, cnpj_cnpj_prestador, cod_ver, val_ser
    
    elif data[0:5] == '<ns3:':
        Bs_data = BeautifulSoup(data, "xml")

        prefixo = "s4:"
        
        # GERAL
        try:
            numero = Bs_data.find(prefixo+'Numero').text
        except:
            prefixo = ""
            numero = Bs_data.find(prefixo+'Numero').text
            
        data_emissao = Bs_data.find(prefixo+'DataEmissao').text
        data_emissao = data_emissao[8:10] +  \
            data_emissao[5:7] +  data_emissao[0:4]

        valor_liquido = Bs_data.find(prefixo+'ValorLiquidoNfse').text
        vl_liquido = str(valor_liquido).replace(",", "")
        vl_liquido = vl_liquido.replace(".", ",")
        
        cod_ver = Bs_data.find(prefixo + 'CodigoVerificacao').text
        val_ser = Bs_data.find(prefixo + 'ValorServicos').text

        valor_iss = Bs_data.find(prefixo+'ValorIss').text
        vl_iss = str(valor_iss).replace(",", "")
        vl_iss = vl_iss.replace(".", ",")
        
        percentual_aliquota = Bs_data.find(prefixo+'Aliquota').text
        pc_aliquota = str(percentual_aliquota).replace(",", "")
        pc_aliquota = pc_aliquota.replace(".", ",")

        valor_base_calculo = Bs_data.find(prefixo+'BaseCalculo').text
        vl_base_calculo = str(valor_base_calculo).replace(",", "")
        vl_base_calculo = vl_base_calculo.replace(".", ",")

        # VALORES
        servico = Bs_data.find(prefixo+'Servico')
        iss_retido = servico.find(prefixo+'IssRetido').text

        valores = Bs_data.find(prefixo+'Valores')

        try:
            valor_deducoes = valores.find(prefixo+'ValorDeducoes').text
            vl_deducoes = str(valor_deducoes).replace(",", "")
            vl_deducoes = vl_deducoes.replace(".", ",")
        except:
            vl_deducoes = "0,00"

        try:
            valor_pis = valores.find(prefixo+'ValorPis').text
            vl_pis = str(valor_pis).replace(",", "")
            vl_pis = vl_pis.replace(".", ",")
        except:
            vl_pis = "0,00"

        try:
            valor_cofins = valores.find(prefixo+'ValorCofins').text
            vl_cofins = str(valor_cofins).replace(",", "")
            vl_cofins = vl_cofins.replace(".", ",")
        except:
            vl_cofins = "0,00"

        try:
            valor_inss = valores.find(prefixo+'ValorInss').text
            vl_inss = str(valor_inss).replace(",", "")
            vl_inss = vl_inss.replace(".", ",")
        except:
            vl_inss = "0,00"

        try:
            valor_ir = valores.find(prefixo+'ValorIr').text
            vl_ir = str(valor_ir).replace(",", "")
            vl_ir = vl_ir.replace(".", ",")
        except:
            vl_ir = "0,00"

        try:
            valor_csll = valores.find(prefixo+'ValorCsll').text
            vl_csll = str(valor_csll).replace(",", "")
            vl_csll = vl_csll.replace(".", ",")
        except:
            vl_csll = "0,00"

        # TOMADOR
        tomador_servico = Bs_data.find(prefixo+'TomadorServico')
        inf_tomador = tomador_servico.find(prefixo+'IdentificacaoTomador')
        try:
            cnpj_cnpj_tomador = inf_tomador.find(prefixo+'Cpf').text if inf_tomador.find(
                prefixo+'Cpf') is not None else inf_tomador.find(prefixo+'Cnpj').text
        except:
            cnpj_cnpj_tomador = ''
        rz_tomador = tomador_servico.find(prefixo+'RazaoSocial').text
        endereco_tomador_servico = tomador_servico.find(prefixo+'Endereco')
        uf_tomador = endereco_tomador_servico.find(prefixo+'Uf').text
        codigo_municipio_tomador = endereco_tomador_servico.find(
            prefixo+'CodigoMunicipio').text

        # PRESTADOR
        prestador_servico = Bs_data.find(prefixo+'PrestadorServico')
        rz_prestador = prestador_servico.find(prefixo+'RazaoSocial').text
        cnpj_cnpj_prestador = prestador_servico.find(prefixo+'Cpf').text if prestador_servico.find(
            prefixo+'Cpf') is not None else prestador_servico.find(prefixo+'Cnpj').text
        endereco_prestador_servico = prestador_servico.find(prefixo+'Endereco')
        codigo_municipio_prestador = endereco_prestador_servico.find(
            prefixo+'CodigoMunicipio').text
        if(codigo_municipio_prestador == "0"):
            codigo_municipio_prestador = "2704302"
            uf_prestador = "AL"
        else:
            uf_prestador = endereco_prestador_servico.find(prefixo+'Uf').text
            
        return data_emissao, cnpj_cnpj_tomador, numero, cnpj_cnpj_prestador, cod_ver, val_ser
    
    elif data[0:5] == '<Comp':
        Bs_data = BeautifulSoup(data, "xml")
        
        
        valor_pis = 0
        valor_cofins = 0
        valor_csll = 0
        try:
            valor_ir = Bs_data.find('ValorIr').text
        except:
            valor_ir = 0
        valor_pis = Bs_data.find('ValorPis').text
        valor_cofins = Bs_data.find('ValorCofins').text
        valor_csll = Bs_data.find('ValorCsll').text
        
        total_pcc = float(valor_pis)+float(valor_cofins)+float(valor_csll)
        
        data_emissao = Bs_data.find('DataEmissao').text
        data_emissao = data_emissao[8:10] +  \
            data_emissao[5:7] +  data_emissao[0:4]
            
        # TOMADOR
        tomador_servico = Bs_data.find('IdentificacaoTomador')
        cnpj_cnpj_tomador = tomador_servico.find('Cpf').text if tomador_servico.find(
        'Cpf') is not None else tomador_servico.find('Cnpj').text
        numero = Bs_data.find('Numero').text
        
        
        
        # PRESTADOR     
        prestador_servico = Bs_data.find('IdentificacaoPrestador')
        cnpj_cnpj_prestador = prestador_servico.find('Cpf').text if prestador_servico.find(
        'Cpf') is not None else prestador_servico.find('Cnpj').text
        cnpj_cnpj_prestador = cnpj_cnpj_prestador.strip()
        
        cod_ver = Bs_data.find('CodigoVerificacao').text
        
        val_ser = Bs_data.find('ValorServicos').text
        
        return data_emissao, cnpj_cnpj_tomador, numero, cnpj_cnpj_prestador, cod_ver, val_ser,valor_ir,total_pcc