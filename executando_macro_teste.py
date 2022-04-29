# -*- coding: utf-8 -*-
"""
Apenas comentário: Codigo usado na macro!!!

Sub UpdateLink()
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources, Type:=xlExcelLinks
End Sub

Sub Atualizar_link()
'
' Atualizar_link Macro
'

'
    ActiveWorkbook.UpdateLink Name:= _
        "C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC-MB_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos\Dados_Horas-Cargo-Projeto-Macroprocesso_2021_Bi.xlsb" _
        , Type:=xlExcelLinks
End Sub
"""

import xlwings as xw

excel_file_path = r'C:\Users\CESAR\Documents\GitHub\bot_dropbox_python\teste_macro.xlsb'

wb = xw.Book(excel_file_path, update_links = False) #update_links travava ele ao abrir
app = wb.app

macro_vba = app.macro("'PERSONAL.XLSB'!Macro2") 
macro_vba()

path_xlsb = 'C:\\Users\\CESAR\\Dropbox (Mitsidi Projetos)\\05_TEC-MB_Dir+Ger+Anl\\TEC_Plan.-Técnico\\TEC_Gerenciamento-Projetos\\PRJ_2155.21144_Giner_CRB_Iguatemi-II.xlsb'
wb = xw.Book(path_xlsb, update_links = False)

app = wb.app
macro_vba = app.macro("'PERSONAL.XLSB'!Macro2") 