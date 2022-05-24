import xlwings as xw


def executar_macro(excel_path, name_macro):
    wb = xw.Book(excel_path, update_links=False)
    app = wb.app

    macro_vba = app.macro(f"'00_script_não_apagar.XLSB'!{name_macro}")
    macro_vba()

    wb.save()
    wb.close()


def open_personal(excel_path):
    wb_personal = xw.Book(excel_path, update_links=False)
    return wb_personal


def tear_down_personal(wb_personal):
    wb_personal.close()


if __name__ == '__main__':
    excel_file_path = r'C:\Users\CESAR\Documents\GitHub\bot_dropbox_python\teste_macro.xlsb'
    macro_name = "Macro1"
    executar_macro(excel_file_path, macro_name)
