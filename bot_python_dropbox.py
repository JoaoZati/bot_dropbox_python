from macro_file import executar_macro, open_personal, tear_down_personal

import os

# Caminho para executar o Bot
pasta = r"C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC-MB_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos"

# Codigo para gerar lista de todas os projeto terminados em xlsb
caminhos = [os.path.join(pasta, nome) for nome in os.listdir(pasta)]
arquivos = [arq for arq in caminhos if os.path.isfile(arq)]
xlsb = [arq for arq in arquivos if arq.lower().endswith(".xlsb")]

# Utilizar nome da macro
Atualizar_link = "Atualizar_link"

# Abrir planilha com macro
pensonal_path = r'\arquivos_excell\PRJ_2212_MVP_Piloto_AHK.xlsb'
wb_personal = open_personal(pensonal_path)

# Looping para executar macro
for caminho in xlsb:
    if caminho[111:114] == 'PRJ':
        executar_macro(caminho, Atualizar_link)

tear_down_personal(wb_personal)
