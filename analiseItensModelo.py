from sentence_transformers import SentenceTransformer, util
import unidecode

# Modelo pré-treinado que transforma frases em vetores
modelo = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

def normalizar(texto):
    return unidecode.unidecode(texto.upper()).strip()

# Exemplo
entrada = "1652 - GLICOSE - 0"
principio_ativo_extraido = normalizar(entrada.split(" - ")[1])

# Lista de princípios ativos
lista_principios = [
    "CLORIDRATO DE METFORMINA",
    "HIDRÓXIDO DE ALUMÍNIO",
    "IBUPROFENO",
    "GLICOSE"
]

# Normaliza e codifica em vetores
principio_vetor = modelo.encode(principio_ativo_extraido, convert_to_tensor=True)
lista_vetores = modelo.encode([normalizar(p) for p in lista_principios], convert_to_tensor=True)

# Calcula similaridades
similaridades = util.cos_sim(principio_vetor, lista_vetores)[0]

# Pega o melhor match
indice_max = similaridades.argmax().item()
print(f'Melhor match: {lista_principios[indice_max]} (similaridade: {similaridades[indice_max]:.2f})')