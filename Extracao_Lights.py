import openpyxl
import PyPDF2
import os 
import re
def extracao_lights():

    print(''' 
        ##############################################
        #                                            #
        #  Extrator de dados de notas fiscais Lights #
        #                                            #
        ##############################################''')
    
    arquivo = ('PDF/')
    lista = os.listdir(arquivo)
    contador = len(lista)

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
        pdf = open('PDF/' + lista[i], 'rb')
        leitor_pdf = PyPDF2.PdfReader(pdf)
        pagina = leitor_pdf.pages[0]
        texto = pagina.extract_text()
        linhas = texto.split('\n')
        numero_linhas = len(linhas)
        for x, linha in enumerate(linhas):
            if numero_linhas >= 76:
                if 'CNPJ' in linha:
                    infos = linha
                    infos = infos.split(' ')
                    cnpj = infos[1]
                    cod_instalacao = infos[3]
                if 'Nota Fiscal' in linha:
                    nota_incompleta = linha
                    padrao_nota = r'no. (\d+)'
                    nota_fiscal = re.search(padrao_nota, nota_incompleta)
                    nota_fiscal = nota_fiscal.group(1)
                    padrao_data = r'(\d{2}/\d{2}/\d{4})'
                    data_emissao = re.search(padrao_data, texto)
                    data_emissao = data_emissao.group(0)

                if 'R$' in linha:
                    if x > 20 and x < 30:
                        infos = linha
                        infos = infos.split(' ')
                        data_vencimento = infos[1]
                if 'ICMS' in linha:
                    infos = (linhas[x+2])
                    icms = infos.replace('COFINS', '')
                    sheet[f'A{index}'] = cnpj
                    sheet[f'B{index}'] = cod_instalacao
                    sheet[f'C{index}'] = nota_fiscal
                    sheet[f'D{index}'] = data_emissao
                    sheet[f'E{index}'] = data_vencimento
                    sheet[f'F{index}'] = icms
                    index += 1
                    wb.save('Extracao_Lights.xlsx') 

            if numero_linhas <= 75:
                if 'CNPJ' in linha:
                    cnpj = linha
                    cnpj = cnpj.replace('CNPJ ', '')

                if 'Conta Contrato' in linha:
                    cod_instalacao = linha
                    cod_instalacao = cod_instalacao[-10:]

                if 'NOTA FISCAL' in linha:
                    nota_dataemi = linha
                    padrao_nota = r'Nº (\d+) -'
                    nota_fiscal = re.search(padrao_nota, nota_dataemi)
                    nota_fiscal = nota_fiscal.group(1)

                    padrao_data = r'DATA DE EMISSÃO: (\d{2}/\d{2}/\d{4})'
                    data_emissao = re.search(padrao_data, nota_dataemi)
                    data_emissao = data_emissao.group(1)

                if 'R$' in linha:
                    if x < 30:
                        valor = linha
                        padrao_data = r'(\d{2}/\d{2}/\d{4})'
                        data_vencimento = re.search(padrao_data, valor)
                        if data_vencimento:
                            data_vencimento = data_vencimento.group(0)
                        else:
                            data_vencimento = 'XX/XX/XXXX'
                    

                if 'Energia Elétrica kWh' in linha:
                    info_gerais = linha
                    info_gerais = info_gerais.split(' ')
                    icms = info_gerais[6]

                    sheet[f'A{index}'] = cnpj
                    sheet[f'B{index}'] = cod_instalacao
                    sheet[f'C{index}'] = nota_fiscal
                    sheet[f'D{index}'] = data_emissao
                    sheet[f'E{index}'] = data_vencimento
                    sheet[f'F{index}'] = icms
                    index += 1
                    wb.save('Extracao_Lights.xlsx')
    print('Extração concluida com sucesso!')
extracao_lights()
