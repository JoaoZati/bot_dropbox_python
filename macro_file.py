import xlwings as xw


def executar_macro(excel_path, name_macro):
    wb = xw.Book(excel_path, update_links = False)
    app = wb.app

    macro_vba = app.macro(f"'PERSONAL.XLSB'!{name_macro}") 
    macro_vba()
    
    wb.save()
    wb.close()


if __name__ == '__main__':
    excel_file_path = r'C:\Users\CESAR\Documents\GitHub\bot_dropbox_python\teste_macro.xlsb'
    macro_name = "Macro1"
    executar_macro(excel_file_path, macro_name)
