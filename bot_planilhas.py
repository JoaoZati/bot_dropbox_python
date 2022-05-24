from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.metrics import dp
import os
from kivy.core.window import Window
# Window.size = (900, 850)
import time
from datetime import timedelta
from macro_file import executar_macro, open_personal, tear_down_personal


class MainApp(MDApp):
    def build(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "DeepPurple"
        self.theme_cls.accent_palette = "Lime"
        return Builder.load_file('bot_planilhas.kv')

    def get_docs(self):
        print('Enter')

        # Caminho para executar o Bot
        self.pasta = r"C:\Users\Mitsidi\Dropbox (Mitsidi Projetos)\05_TEC-MB_Dir+Ger+Anl\TEC_Plan.-Técnico\TEC_Gerenciamento-Projetos"

        # Codigo para gerar lista de todas os projeto terminados em xlsb
        caminhos = [os.path.join(self.pasta, nome)
                    for nome in os.listdir(self.pasta)]
        arquivos = [arq for arq in caminhos if os.path.isfile(arq)]
        xlsb = [arq for arq in arquivos if arq.lower().endswith(".xlsb")]
        nome = [nome[len(self.pasta)+1:]
                for nome in xlsb if nome[len(self.pasta)+1:len(self.pasta)+4] == 'PRJ']
        self.add_datatable(nome)
        self.files = []

    def add_datatable(self, nome):
        file = tuple([[file, " "] for file in nome])
        self.data_tables = MDDataTable(
            size_hint=(0.9, 0.8),
            check=True,
            use_pagination=True,
            rows_num=20,
            pagination_menu_height='240dp',
            pagination_menu_pos="auto",
            column_data=[
                ("Nome", dp(130)),
                (" ", dp(2))
            ],
            row_data=file
        )
        self.root.ids.data_layout.add_widget(self.data_tables)

        self.data_tables.bind(on_check_press=self.checked)

    def checked(self, instance_table, current_row):

        if current_row[0] in self.files:
            self.files.remove(current_row[0])
        else:
            print(current_row)
            self.files.append(current_row[0])

    def update(self):
        data = []
        pensonal_path = self.pasta + r'\00_script_não_apagar.xlsb'
        wb_personal = open_personal(pensonal_path)
        Atualizar_link = "Atualizar_link"

        for file in self.files:
            start = time.time()
            caminho = self.pasta + '\\' + file
            executar_macro(caminho, Atualizar_link)
            end = time.time()
            tempo = end - start
            
            print(end, start, tempo)

            data.append([file, "{:.3f}".format(tempo)])
            datas = tuple(data)

            self.data_tables = MDDataTable(
                size_hint=(0.9, 0.8),
                use_pagination=True,
                rows_num=20,
                pagination_menu_height='240dp',
                pagination_menu_pos="auto",
                column_data=[
                    ("Nome", dp(100)),
                    ("Tempo", dp(60))
                ],
                row_data=datas
            )
            self.root.ids.data_layout.add_widget(self.data_tables)

        tear_down_personal(wb_personal)


MainApp().run()
