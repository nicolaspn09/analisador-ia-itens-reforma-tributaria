from geraSQLProdutos import GeraSQLProdutos
from geraTabelaExcel import GeraTabelaExcel


def main():
    # Movimenta os arquivos para iniciar com uma planilha limpa
    GeraTabelaExcel.movimenta_arquivos()

    # Gera o SQL dos produtos
    sql_produtos = GeraSQLProdutos.gera_sql_produtos()

    # Gera a tabela Excel com os dados obtidos
    GeraTabelaExcel.gera_tabela_excel(sql_produtos)