import sys

sys.path.append(r"C:\rpa\Python")
from Classes.ZimbraMailer.ZimbraMailer.Zimbra import ZimbraMailer


def envia_email_relatorio_finalizado(arquivo_anexo, destinatarios, assunto, corpo_email):
    """
    Envia um email com o relatório finalizado como anexo.
    """
    zimbra = ZimbraMailer()
    zimbra.envia_email(
        anexos=arquivo_anexo,
        assunto_email=assunto,
        mensagem_email=corpo_email,
        destinatarios_email=destinatarios
    )


def main():
    arquivo_anexo = []
    arquivo_anexo.append(r"C:\rpa\Python\Analisa Itens Reforma Tributaria\Nova Analise\ANALISE.xlsx")

    destinatarios = []
    destinatarios.append("nicolas.nasario@COMPANY_NAME.com.br")
    destinatarios.append("israel.martins@COMPANY_NAME.com.br")
    destinatarios.append("fiscal@COMPANY_NAME.com.br")
    destinatarios.append("marcella.barros@COMPANY_NAME.com.br")
    destinatarios.append("debora.arceno@COMPANY_NAME.com.br")

    assunto = "Análise de Isenção - Reforma Tributária"
    corpo_email = f"""
    Olá!<br><br>

    Segue em anexo o relatório finalizado da análise de isenção referente à Reforma Tributária.<br><br>

    Atenciosamente,<br>
    Equipe de RPA COMPANY_NAME
    """

    envia_email_relatorio_finalizado(arquivo_anexo=arquivo_anexo, destinatarios=destinatarios, assunto=assunto, corpo_email=corpo_email)
    print("Email enviado com sucesso!")


# Exemplo de uso
if __name__ == "__main__":
    main()