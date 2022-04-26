# -*- coding: utf-8 -*-
"""
Apenas comentário: Codigo usado na macro!!!

Sub Macro1()
'
' Macro1 Macro
' Execular 1+1
'

'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-2]C+R[-1]C"
    Range("A4").Select
End Sub

Sub UpdateLink()
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources, Type:=xlExcelLinks
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