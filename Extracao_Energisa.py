import openpyxl
import PyPDF2
import os 
import re

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
        pdfs = open('PDF/' + arquivos[i], 'rb')
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
            
                
            index += 1
                


            



extracao_energisa()