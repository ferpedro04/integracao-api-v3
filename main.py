from dotenv import load_dotenv
from openpyxl import Workbook
import os 
import json
import requests
import math

load_dotenv()

url_base = os.getenv("URL_BASE")
id_parceiro = int(os.getenv("ID_PARCEIRO"))
cnpj = os.getenv("CNPJ")
token = os.getenv("TOKEN")
cmun = os.getenv("CMUN")
regime = os.getenv("REGIME")
crt = int(os.getenv("CRT"))
url_revisao = f"{url_base}/revisao/{id_parceiro}/{cnpj}/{token}"


produtos = []

with open("produtos.json", "r", encoding="utf-8") as arquivo:
    for linha in arquivo:
        produto = json.loads(linha)
        produtos.append(produto)
total_produtos = len(produtos)
total_lotes = math.ceil(total_produtos / 300)
print(f"""Importando arquivo
{total_produtos} produtos encontrados""")
print("Iniciando processamento...")

lotes_certos = []
lotes_errados = []
#lotes_errados.append({
#    "numero_lote": 99,
#    "produtos": produtos[:2],
#    "motivo": "Erro de teste"
#})

for numero_lote, inicio in enumerate(range(0, total_produtos, 300), start= 1):
    lote = produtos[inicio : inicio + 300]    
    lote_api = []
    print(f"Processando lote {numero_lote}/{total_lotes} - {len(lote)} produtos")
    for produto in lote:
        tamanho_ean = len(produto["codEan"])
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

itens = []
for lote_certo in lotes_certos:
    itens.extend(lote_certo.get("itens", []))

retornados = []
for item in itens:
    retornados.append(int(item["produtos"]["produto_cliente"]["cod_interno_cliente"]))
print(f"Total de itens retornados pela API: {len(retornados)}")

nao_encontrados = []
for produto in produtos:
    if produto["idproduto"] not in retornados:
        nao_encontrados.append(produto)
print(f"Produtos não encontrados pela API: {len(nao_encontrados)}")

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

itens_achatados = []
for item in itens:
    itens_achatados.append(achatar_json(item))

colunas = []

for item in itens_achatados:
    for chave in item.keys():
        if chave not in colunas:
            colunas.append(chave)

workbook = Workbook()
planilha = workbook.active
planilha.append(colunas)

for item in itens_achatados:
    linha = []
    for coluna in colunas:
        linha.append(item.get(coluna, ""))
    planilha.append(linha)
workbook.save("output/revisao_fiscal_v3.xlsx")

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

print("\n--- Resumo final ---")
print(f"Produtos enviados: {total_produtos}")
print(f"Total de lotes: {total_lotes}")
print(f"Lotes processados com sucesso: {len(lotes_certos)}")
print(f"Lotes processados com erro: {len(lotes_errados)}")
print(f"Produtos encontrados pela API: {len(retornados)}")
print(f"Produtos não encontrados pela API: {len(nao_encontrados)}")