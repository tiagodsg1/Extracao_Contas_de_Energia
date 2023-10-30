import openpyxl
import PyPDF2
import os 
import re

def modelo1(arquivo, sheet, index):
    print('modelo1')
    pdfs = open('PDF/' + arquivo, 'rb')
    leitor_pdf = PyPDF2.PdfReader(pdfs)
    pagina = leitor_pdf.pages[0]
    texto = pagina.extract_text()
    linhas = texto.split('\n')
    for x, linha in enumerate(linhas):
        if 'CNPJ' in linha:
            cnpj = linhas[x-2]
            cnpj = cnpj.replace('-', '', 1)
            cnpj = cnpj.replace('\n', '')

        if 'Pis/Cofins (R$)' in linha:
            infos = linha
            infos = infos.split(' ')
            numero_instalacao = infos[2]
            numero_instalacao = numero_instalacao.replace('\n', '')

        if 'Nota Fiscal/Conta de Energia Elétrica Nº' in linha:
            infos = linha
            padrao = r'Nº (\d+\.\d+\.\d+)'
            nota_fiscal = re.search(padrao, infos)
            nota_fiscal = nota_fiscal.group(0)
            nota_fiscal = nota_fiscal.replace('Nº ', '')
            nota_fiscal = nota_fiscal.replace('\n', '')
            

        if 'Emissão:' in linha:
            infos = linha
            padrao = r'Emissão: (\d+\/\d+\/\d+)'
            data_emissao = re.search(padrao, infos)
            data_emissao = data_emissao.group(0)
            data_emissao = data_emissao.replace('Emissão:', '')
            data_emissao = data_emissao.replace(' ', '')
            data_emissao = data_emissao.replace('\n', '')
            

        if 'R$' in linha:
            if x == 4:
                infos = linha
                padrao = r'(\d+\/\d+\/\d+)'
                data_vencimento = re.search(padrao, infos)
                data_vencimento = data_vencimento.group(0)
                data_vencimento = data_vencimento.replace('\n', '')

        if 'Código de Classificação do Item Total:' in linha:
            infos = linha
            infos = infos.split(' ')
            icms = infos[13]
            icms = icms.replace('\n', '').replace(' ', '')
    
    sheet[f'A{index}'] = cnpj
    sheet[f'B{index}'] = numero_instalacao
    sheet[f'C{index}'] = nota_fiscal
    sheet[f'D{index}'] = data_emissao
    sheet[f'E{index}'] = data_vencimento
    sheet[f'F{index}'] = icms

def modelo2(arquivo, sheet, index):
    print('modelo2')
    pdfs = open('PDF/' + arquivo, 'rb')
    leitor_pdf = PyPDF2.PdfReader(pdfs)
    pagina = leitor_pdf.pages[0]
    texto = pagina.extract_text()
    linhas = texto.split('\n')
    for x, linha in enumerate(linhas):
        if 'CNPJ' in linha:
            if x >= 1 and x <=5:
                infos = linha
                infos = infos.replace('-', '', 1).replace(' ', '')
                padrao = r'(\d+\.\d+\.\d+\/\d+\-\d+)'
                cnpj = re.search(padrao, infos)
                cnpj = cnpj.group(0)

        if 'PIX!' in linha:
            infos = linha
            padrao = r'PIX! (\d+\/\d+\-\d+)'
            numero_instalacao = re.search(padrao, infos)
            numero_instalacao = numero_instalacao.group(0)
            numero_instalacao = numero_instalacao.replace('PIX! ', '')

        if 'Nota Fiscal/Conta de Energia Elétrica Nº' in linha:
            infos = linha
            padrao = r'(\d+\.\d+\.\d+\ \d+)'
            nota_fiscal = re.search(padrao, infos)
            nota_fiscal = nota_fiscal.group(0)
            nota_fiscal = nota_fiscal.replace(' ', '')

        if ' Emissão:' in linha:
            infos = linha
            padrao = r'Emissão: (\d+\/\d+\/\d+)'
            data_emissao = re.search(padrao, infos)
            data_emissao = data_emissao.group(0)
            data_emissao = data_emissao.replace('Emissão:', '').replace(' ', '')

        if 'R$' in linha:
            if x >= 5 and x <=9:
                infos = linha
                padrao = r'(\d+\/\d+\/\d+)'
                data_vencimento = re.search(padrao, infos)
                data_vencimento = data_vencimento.group(0)
                data_vencimento = data_vencimento.replace('\n', '')

        if 'Consumo em kWh' in linha:
            infos = linha
            infos = infos.split(' ')
            icms = infos[12]
        
    sheet[f'A{index}'] = cnpj
    sheet[f'B{index}'] = numero_instalacao
    sheet[f'C{index}'] = nota_fiscal
    sheet[f'D{index}'] = data_emissao
    sheet[f'E{index}'] = data_vencimento
    sheet[f'F{index}'] = icms

def modelo3(arquivo, sheet, index):
    print('modelo3')
    pdfs = open('PDF/' + arquivo, 'rb')
    leitor_pdf = PyPDF2.PdfReader(pdfs)
    pagina = leitor_pdf.pages[0]
    texto = pagina.extract_text()
    linhas = texto.split('\n')
    for x, linha in enumerate(linhas):
        if 'CPF/CNPJ/RANI:' in linha:
            infos = linha
            padrao = r'(\d+\.\d+\.\d+\/\d+\-\d+)'
            cnpj = re.search(padrao, infos)
            if cnpj:
                cnpj = cnpj.group(0)
            else:
                cnpj = 'Não encontrado'

        if '18 ' in linha:
            infos = linha
            padrao = r'(\d+\/\d+\-\d+)'
            numero_instalacao = re.search(padrao, infos)
            numero_instalacao = numero_instalacao.group(0)

        if 'NOTA FISCAL N°' in linha:
            infos = linha
            padrao = r'(\d+\.\d+\.\d+)'
            nota_fiscal = re.search(padrao, infos)
            nota_fiscal = nota_fiscal.group(0)

        if 'EMISSÂO/APRESENTAÇÂO' in linha:
            infos = linha
            padrao = r'(\d+\/\d+\/\d+)'
            data_emissao = re.search(padrao, infos)
            data_emissao = data_emissao.group(0)

        if 'R$' in linha:
            if x >=0 and x <=8:
                infos = linha
                padrao = r'(\d+\/\d+\/\d+)'
                data_vencimento = re.search(padrao, infos)
                data_vencimento = data_vencimento.group(0)

        if 'PIS/PASEP' in linha:
            infos = linha
            padrao = r'(\d+\,\d+)'
            icms = re.search(padrao, infos)
            icms = icms.group(0)

    sheet[f'A{index}'] = cnpj
    sheet[f'B{index}'] = numero_instalacao
    sheet[f'C{index}'] = nota_fiscal
    sheet[f'D{index}'] = data_emissao
    sheet[f'E{index}'] = data_vencimento
    sheet[f'F{index}'] = icms

def extracao_energisa():
    print(''' 
        ###############################################
        #                                             #
        # Extrator de dados de notas fiscais Energisa #
        #                                             #
        ###############################################''')
    
    pasta = ('PDF/')
    arquivos = os.listdir(pasta)
    contador = len(arquivos)

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['A1'] = 'CNPJ'
    sheet['B1'] = 'Numero da instalação'
    sheet['C1'] = 'Numero da nota fiscal'
    sheet['D1'] = 'Data de emissão'
    sheet['E1'] = 'Data de vencimento'
    sheet['F1'] = 'Icms'
    index = 2

    for i in range(contador):
        arquivo = arquivos[i]
        int_arquivo = int(arquivo[0:4])

        if int_arquivo >=2018 and int_arquivo <= 2020:
            modelo1(arquivo, sheet, index)

        if int_arquivo == 2021:
            modelo2(arquivo, sheet, index)

        if int_arquivo >= 2022:
            modelo3(arquivo, sheet, index)
            
        
        index += 1
        wb.save('Extracao_Energisa.xlsx')

    print('\nExtracao concluida com sucesso!')
                


            



extracao_energisa()