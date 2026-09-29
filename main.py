from dotenv import load_dotenv
from openpyxl import Workbook
import os 
import json
import requests
import math

# Carrega as configurações e credenciais armazenadas no arquivo .env
load_dotenv()

url_base = os.getenv("URL_BASE")
id_parceiro = int(os.getenv("ID_PARCEIRO"))
cnpj = os.getenv("CNPJ")
token = os.getenv("TOKEN")
cmun = os.getenv("CMUN")
regime = os.getenv("REGIME")
crt = int(os.getenv("CRT"))
# Monta a URL do endpoint de Revisão Fiscal V3
url_revisao = f"{url_base}/revisao/{id_parceiro}/{cnpj}/{token}"

# Lê os produtos do arquivo JSON de entrada
produtos = []

with open("produtos.json", "r", encoding="utf-8") as arquivo:
    for linha in arquivo:
        produto = json.loads(linha)
        produtos.append(produto)
# Calcula a quantidade de produtos e de lotes necessários
total_produtos = len(produtos)
total_lotes = math.ceil(total_produtos / 300)
print(f"""Importando arquivo
{total_produtos} produtos encontrados""")
print("Iniciando processamento...")

# Armazena separadamente as respostas bem-sucedidas e os lotes com erro
lotes_certos = []
lotes_errados = []
# TESTE DE ERRO
#lotes_errados.append({
#    "numero_lote": 99,
#    "produtos": produtos[:2],
#    "motivo": "Erro de teste"
#})

# Divide automaticamente os produtos em lotes de no máximo 300 itens
for numero_lote, inicio in enumerate(range(0, total_produtos, 300), start= 1):
    lote = produtos[inicio : inicio + 300]    
    lote_api = []
    print(f"Processando lote {numero_lote}/{total_lotes} - {len(lote)} produtos")
     # Converte os produtos do arquivo para o formato esperado pela API
    for produto in lote:
        tamanho_ean = len(produto["codEan"])
        # A API aceita EAN com no máximo 14 caracteres.
        # Quando ultrapassa esse limite, utiliza o código interno como EAN.
        if tamanho_ean > 14:
            produto["codEan"] = str(produto["idproduto"])
        produto_api = { "codinterno": str(produto["idproduto"]),
                    "ean": produto["codEan"], 
                    "descricao": produto["descricao"],
                    "cMun": cmun,
                    "regime": regime,
                    "crt": crt,
                    "csosn": "000",
                    "ncm": "00000000",
                    "cfop": "0000",
                    "pICMS": 0,
                    "cst_pis_saida": "00",
                    "cst_cofins_saida": "00" }
        lote_api.append(produto_api)
    # Envia o lote para o endpoint de Revisão Fiscal
    try:
            resposta = requests.post(url_revisao, json=lote_api)
            if (resposta.status_code == 200) or (resposta.status_code == 201):
                print(f"Carregamento do lote {numero_lote} concluído - Status: {resposta.status_code}")
                lotes_certos.append(resposta.json())
            else:
                erro_lote = {
                        "numero_lote": numero_lote,
                        "produtos": lote,
                        "motivo": resposta.json()}
                print(f"Falha ao carregar lote {numero_lote} - Status: {resposta.status_code}")
                lotes_errados.append(erro_lote)
                print(resposta.json())
    except Exception as erro:
            erro_lote = {
                "numero_lote": numero_lote,
                "produtos": lote,
                "motivo": str(erro)
            }
            lotes_errados.append(erro_lote)
            print(f"Falha de conexão no lote {numero_lote}: {erro}")
# Consolida os produtos retornados pelos lotes processados com sucesso
itens = []
for lote_certo in lotes_certos:
    itens.extend(lote_certo.get("itens", []))
# Obtém os códigos internos dos produtos retornados pela API
retornados = []
for item in itens:
    retornados.append(int(item["produtos"]["produto_cliente"]["cod_interno_cliente"]))
print(f"Total de itens retornados pela API: {len(retornados)}")
# Identifica os produtos enviados que não foram retornados pela API
nao_encontrados = []
for produto in produtos:
    if produto["idproduto"] not in retornados:
        nao_encontrados.append(produto)
print(f"Produtos não encontrados pela API: {len(nao_encontrados)}")

# Achata recursivamente os objetos JSON retornados pela API.
# Exemplo:
# tributos.icms.cst passa a ser uma coluna do Excel.
def achatar_json(dados, caminho = "", resultado = None):
    if resultado is None:
        resultado = {}
    for chave, valor in dados.items():
        if caminho:
            novo_caminho = caminho + "." + chave
        else:
            novo_caminho = chave
        if isinstance(valor, dict):
            achatar_json(valor, novo_caminho, resultado)
        else:
            resultado[novo_caminho] = valor
    return resultado

# Aplica o achatamento a todos os produtos retornados
itens_achatados = []
for item in itens:
    itens_achatados.append(achatar_json(item))

# Monta dinamicamente as colunas a partir dos campos retornados pela API
colunas = []

for item in itens_achatados:
    for chave in item.keys():
        if chave not in colunas:
            colunas.append(chave)

# Gera o Excel principal com os dados retornados pela API
workbook = Workbook()
planilha = workbook.active
planilha.append(colunas)

for item in itens_achatados:
    linha = []
    for coluna in colunas:
        linha.append(item.get(coluna, ""))
    planilha.append(linha)
workbook.save("output/revisao_fiscal_v3.xlsx")

# Transforma os lotes com erro em uma lista de produtos individuais
produtos_com_erro = []

for lote_com_erro in lotes_errados:
    for produto in lote_com_erro["produtos"]:
        produto_erro = {
   "numero_lote": lote_com_erro["numero_lote"],
   "idproduto": str(produto["idproduto"]),
   "ean": produto["codEan"],
   "descricao": produto["descricao"],
   "motivo_erro": lote_com_erro["motivo"]
   }
        produtos_com_erro.append(produto_erro)
# Gera o Excel de erros somente quando houver produtos com falha
if produtos_com_erro:
    workbook_erros = Workbook()
    planilha_erros = workbook_erros.active

    colunas_com_erros = ["numero_lote", "idproduto", "ean", "descricao", "motivo_erro"]
    planilha_erros.append(colunas_com_erros)

    for produto in produtos_com_erro:
        linhas_com_erro = []
        for coluna in colunas_com_erros:
            linhas_com_erro.append(produto.get(coluna, ""))
        planilha_erros.append(linhas_com_erro)    
    workbook_erros.save("output/revisao_fiscal_v3_erros.xlsx")

# Exibe o resumo geral da execução
print("\n--- Resumo final ---")
print(f"Produtos enviados: {total_produtos}")
print(f"Total de lotes: {total_lotes}")
print(f"Lotes processados com sucesso: {len(lotes_certos)}")
print(f"Lotes processados com erro: {len(lotes_errados)}")
print(f"Produtos encontrados pela API: {len(retornados)}")
print(f"Produtos não encontrados pela API: {len(nao_encontrados)}")