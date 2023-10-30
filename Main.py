from src.Extracao_Energisa import extracao_energisa
from src.Extracao_Lights import extracao_lights

def main():
    print('''
        ############################################
        #                                          #
        #    Extrator de dados de notas fiscais    #
        #    Digite 1 para Lights                  #
        #    Digite 2 para Energisa                #
        ############################################''')
    escolha = int(input('Digite sua escolha: '))

    if escolha == 1:
        extracao_lights()
        
    elif escolha == 2:
        extracao_energisa()

main()
    