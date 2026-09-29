from dotenv import load_dotenv
from openpyxl import Workbook
import os 
import json
import requests
import math
import re 

#Listas do projeto
produtos = []
lotes_certos = []
lotes_errados = []
itens = []
retornados = []
nao_encontrados = []
itens_achatados = []
colunas = []
produtos_com_erro = []
erros_entrada = []

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

# Recupera campos de texto de linhas com JSON inválido para registrar no relatório de erros
def extrair_campo(texto, campo):
    padrao = rf'"{campo}"\s*:\s*"([^"]*)"'
    resultado = re.search(padrao, texto)

    if resultado:
        return resultado.group(1)

    return ""

# Recupera o idproduto numérico de linhas com JSON inválido para registrar no relatório de erros
def extrair_idproduto(texto):
    padrao = r'"idproduto"\s*:\s*(\d+)'
    resultado = re.search(padrao, texto)

    if resultado:
        return resultado.group(1)

    return ""

# Lê e valida os produtos do arquivo de entrada
# Produtos com campos obrigatórios ausentes são separados para o relatório de erros
with open("produtos.json", "r", encoding="utf-8") as arquivo:
    for numero_linha, linha in enumerate(arquivo, start =1):
        try:
            produto = json.loads(linha)

            if "idproduto" not in produto:
                erro_entrada = {
                    "numero_linha": numero_linha,
                    "numero_lote": "",
                    "idproduto": "",
                    "ean": produto.get("codEan", ""),
                    "descricao": produto.get("descricao", ""),
                    "conteudo": linha.strip(),
                    "motivo_erro": "Campo idproduto ausente"
                }
                erros_entrada.append(erro_entrada)
                continue

            if "codEan" not in produto:
                erro_entrada = {
                "numero_linha": numero_linha,
                "numero_lote": "",
                "idproduto": str(produto.get("idproduto", "")),
                "ean": "",
                "descricao": produto.get("descricao", ""),
                "conteudo": linha.strip(),
                "motivo_erro": "Campo codEan ausente"
            }
                erros_entrada.append(erro_entrada)
                continue

            if "descricao" not in produto:
                erro_entrada = {
                "numero_linha": numero_linha,
                "numero_lote": "",
                "idproduto": str(produto.get("idproduto", "")),
                "ean": produto.get("codEan", ""),
                "descricao": "",
                "conteudo": linha.strip(),
                "motivo_erro": "Campo descricao ausente"
            }
                erros_entrada.append(erro_entrada)
                continue
            
            produtos.append(produto)
            
        # Em linhas com JSON inválido, tenta recuperar os campos disponíveis
        # e registra todos os problemas encontrados sem interromper a execução
        except json.JSONDecodeError as erro:
            motivos = []

            idproduto_extraido = extrair_idproduto(linha)
            ean_extraido = extrair_campo(linha, "codEan")
            descricao_extraida = extrair_campo(linha, "descricao")

            if idproduto_extraido == "":
                motivos.append("Campo idproduto ausente ou inválido")

            if ean_extraido == "":
                motivos.append("Campo codEan ausente, vazio ou inválido")

            if descricao_extraida == "":
                motivos.append("Campo descricao ausente, vazio ou inválido")

            motivos.append(f"JSON inválido: {erro}")
            motivo_completo = " | ".join(motivos)
            
            erro_entrada = {
                "numero_linha": numero_linha,
                "ean": ean_extraido,
                "descricao": descricao_extraida,
                "conteudo": linha.strip(),
                "motivo_erro": motivo_completo,
                "idproduto": idproduto_extraido,
            }
            erros_entrada.append(erro_entrada)

# Calcula a quantidade de produtos e de lotes necessários
total_produtos = len(produtos)
total_lotes = math.ceil(total_produtos / 300)
print(f"""Importando arquivo
{total_produtos} produtos encontrados""")
print("Iniciando processamento...")

# TESTE DE ERRO
# lotes_errados.append({
#     "numero_lote": 99,
#     "produtos": produtos[:2],
#     "motivo": "Erro de teste"
# })

# Divide automaticamente os produtos em lotes de no máximo 300 itens
for numero_lote, inicio in enumerate(range(0, total_produtos, 300), start= 1):
    lote = produtos[inicio : inicio + 300]    
    lote_api = []
    print(f"Processando lote {numero_lote}/{total_lotes} - {len(lote)} produtos")
    
     # Converte os produtos do arquivo para o formato esperado pela API
    for produto in lote:
        tamanho_ean = len(produto["codEan"])

        # Valida o EAN antes do envio.
        # Quando estiver vazio, não numérico ou possuir mais de 14 caracteres,
        # utiliza o código interno do produto como EAN.
        if (tamanho_ean > 14) or (produto["codEan"] == "") or not (produto["codEan"].isdigit()):
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
                try:
                    motivo_erro = json.dumps(resposta.json(), ensure_ascii=False)
                except ValueError:
                    motivo_erro = resposta.text

                erro_lote = {
                        "numero_lote": numero_lote,
                        "produtos": lote,
                        "motivo": motivo_erro}
                print(f"Falha ao carregar lote {numero_lote} - Status: {resposta.status_code}")
                lotes_errados.append(erro_lote)
                print(motivo_erro)
    except Exception as erro:
            erro_lote = {
                "numero_lote": numero_lote,
                "produtos": lote,
                "motivo": str(erro)
            }
            lotes_errados.append(erro_lote)
            print(f"Falha de conexão no lote {numero_lote}: {erro}")

# Consolida os produtos retornados pelos lotes processados com sucesso
for lote_certo in lotes_certos:
    itens.extend(lote_certo.get("itens", []))

# Obtém os códigos internos dos produtos retornados pela API
for item in itens:
    retornados.append(int(item["produtos"]["produto_cliente"]["cod_interno_cliente"]))
print(f"Total de itens retornados pela API: {len(retornados)}")

# Identifica os produtos enviados que não foram retornados pela API
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
for item in itens:
    itens_achatados.append(achatar_json(item))

# Monta dinamicamente as colunas a partir dos campos retornados pela API
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

# Consolida erros de entrada e de processamento da API em um único relatório
relatorio_de_erros = produtos_com_erro + erros_entrada
if relatorio_de_erros:
    workbook_erros = Workbook()
    planilha_erros = workbook_erros.active

    colunas_com_erros = ["numero_linha","numero_lote", "idproduto", "ean", "descricao", "conteudo", "motivo_erro"]
    planilha_erros.append(colunas_com_erros)

    for produto in relatorio_de_erros:
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