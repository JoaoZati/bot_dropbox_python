# -*- coding: utf-8 -*-
"""
Python Bot para atualizaçao do drive
"""
from openpyxl import load_workbook
from pyxlsb import open_workbook
import xlwings as xw


path_xlsb = 'C:\\Users\\CESAR\\Dropbox (Mitsidi Projetos)\\05_TEC-MB_Dir+Ger+Anl\\TEC_Plan.-Técnico\\TEC_Gerenciamento-Projetos\\PRJ_2155.21144_Giner_CRB_Iguatemi-II.xlsb'
path_xlsx = 'C:\\Users\\CESAR\\Dropbox (Mitsidi Projetos)\\05_TEC-MB_Dir+Ger+Anl\\TEC_Plan.-Técnico\\TEC_Gerenciamento-Projetos\\Cronograma-Mapa tipologias.xlsx'

wb_xlsx = load_workbook(path_xlsx)
wb_xlsx.close()

wb_xlsb = open_workbook(path_xlsb)

app = xw.app()
book_xlsb = xw.Book(path_xlsb)
book_xlsb.close()
app.kill()

# comentario