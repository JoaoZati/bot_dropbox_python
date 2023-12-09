from macro_file import executar_macro, open_personal, tear_down_personal

import os

# Caminho para executar o Bot
pasta = r"C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos"

# Codigo para gerar lista de todas os projeto terminados em xlsb
caminhos = [os.path.join(pasta, nome) for nome in os.listdir(pasta)]
arquivos = [arq for arq in caminhos if os.path.isfile(arq)]
xlsb = [arq for arq in arquivos if arq.lower().endswith(".xlsb")]

# Utilizar nome da macro
Atualizar_link = "Atualizar_link"

# Abrir planilha com macro
# pensonal_path = os.path.dirname(__file__) + r'\arquivos_excell\PRJ_2212_MVP_Piloto_AHK.xlsb'
pensonal_path = r"C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos\00_script_não_apagar.XLSB"
wb_personal = open_personal(pensonal_path)

# Looping para executar macro
for caminho in xlsb:
    nome_do_arquivo = caminho.split("\\")[-1]
    if nome_do_arquivo[0:3] == 'PRJ':
        executar_macro(caminho, Atualizar_link)

tear_down_personal(wb_personal)
