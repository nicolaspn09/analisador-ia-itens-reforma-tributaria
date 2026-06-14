import sys

sys.path.append(r"C:\rpa\Python")
from Classes.Hangouts.Hangouts.Hangouts import Hangouts


def envia_mensagem_hangouts(mensagem):
    url = f"https://chat.googleapis.com/v1/spaces/AAQA_nTpXnI/messages?key='SECRET_REMOVED_BY_AI'&token=A3uqSJ-wwTTNlSuMWEbk5tFkWCUI4jOaRaRiQjIf9Tw"

    hangouts = Hangouts(url=url, mensagem=mensagem)
    hangouts.retorna_google_chat()


def main():
    envia_mensagem_hangouts(f"*Início do processo de análise de isenção - Reforma Tributária*")

    from GeraDadosIniciais import main as gera_dados_iniciais
    from analiseDadosIsencaoGPT import lista_itens_excel as analisa_dados_isencao_gpt
    # from analisaDadosIsencao import lista_itens_excel as analisa_dados_isencao     ##### EM DESUSO
    from enviaEmailRelatorioFinalizado import main as envia_email_relatorio_finalizado

    envia_mensagem_hangouts(f"Início da geração dos dados iniciais")
    
    try:
        gera_dados_iniciais()
    except Exception as e:
        envia_mensagem_hangouts(f"Erro ao importar GeraDadosIniciais: {e}")
        return
    
    envia_mensagem_hangouts(f"Fim da geração dos dados iniciais")
    envia_mensagem_hangouts(f"Início da análise de isenção")

    try:
        analisa_dados_isencao_gpt()
    except Exception as e:
        envia_mensagem_hangouts(f"Erro ao importar analisaDadosIsencao: {e}")
        return
    
    envia_mensagem_hangouts(f"Fim da análise de isenção")
    envia_mensagem_hangouts(f"Início do envio do email com o relatório finalizado")
    
    try:
        envia_email_relatorio_finalizado()
    except Exception as e:
        envia_mensagem_hangouts(f"Erro ao importar enviaEmailRelatorioFinalizado: {e}")
        return
    
    envia_mensagem_hangouts(f"Fim do envio do email com o relatório finalizado") 
    envia_mensagem_hangouts(f"*Fim do processo de análise de isenção - Reforma Tributária*")



main()