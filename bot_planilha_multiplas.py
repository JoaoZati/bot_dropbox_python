from macro_file import executar_macro

import os

#path = r"C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC-MB_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos\PRJ_2212_MVP_Piloto_AHK.xlsb"

pasta = r"C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC-MB_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos"

caminhos = [os.path.join(pasta, nome) for nome in os.listdir(pasta)]
arquivos = [arq for arq in caminhos if os.path.isfile(arq)]
xlsb = [arq for arq in arquivos if arq.lower().endswith(".xlsb")]

list_excluir_xlsb = []
Atualizar_link = "Atualizar_link"

for caminho in xlsb:
    if caminho[111:114] == 'PRJ':
        list_excluir_xlsb.append(caminho)
        executar_macro(caminho, Atualizar_link)

#executar_macro(pf, Atualizar_link)
