import sys

sys.path.append(r"C:\rpa\Python")
from Classes.GoogleSheets.GoogleSheets.GoogleSheets import GoogleSheets



class GeraTabelasSheetsIsencao:
    def __init__(self):
        pass


    def gera_tabela_sheets_isencao():
        token_path = r"C:\rpa\Python\Analisa Itens Reforma Tributaria\Nova Analise\token.json"
        id_planilha = "1LD087-ZwcgPK1lkNZ0lJCP-qhaIOnubhl3dnEprJ3AM"
        range_dados = "Anexos!A2:F"

        sheets = GoogleSheets(diretorio_json=token_path, id_planilha=id_planilha, range_dados=range_dados)
        tabela = sheets.solicita_tabela()

        return tabela