import subprocess
import sys

def iniciar_analise_background():
    """
    Inicia a análise em background e retorna imediatamente
    """
    caminho = r"C:\rpa\Python\Analisa Itens Reforma Tributaria\Nova Analise\main.pyw"
    
    processo = subprocess.Popen([sys.executable, caminho])
    print(f"Análise iniciada em background (PID: {processo.pid})")
    
    return processo.pid

# Uso
pid = iniciar_analise_background()
print(f"Análise rodando em background! PID: {pid}")
print("Código principal finalizado, mas análise continua rodando...")