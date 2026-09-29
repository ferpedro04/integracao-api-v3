# Integração API V3 - Revisão Fiscal

Projeto desenvolvido em Python para consumir o endpoint de Revisão Fiscal da API V3 da Avant Fiscal.

A aplicação lê produtos a partir de um arquivo JSON, divide automaticamente os produtos em lotes de até 300 itens e envia cada lote para a API.

Os dados retornados são consolidados e exportados para um arquivo Excel, utilizando achatamento genérico dos campos JSON. Caso ocorram erros durante o processamento, os produtos dos lotes com falha são registrados separadamente em um arquivo de erros.

## Requisitos

- Python 3
- Bibliotecas disponíveis em `requirements.txt`

## Instalação

Instale as dependências do projeto:

pip install -r requirements.txt

## Configuração

Crie um arquivo `.env` na raiz do projeto e configure as seguintes variáveis:

URL_BASE= URL usada do portal V3
ID_PARCEIRO= Seu ID parceiro
CNPJ= Seu CNPJ
TOKEN= Seu Token
CMUN= Código do seu múnicipio
REGIME= simples, lucro presumido ou real
CRT= 1(Simples Nacional), 2(Simples Nacional com Excesso de Sublimite) ou 3(Regime Normal)

## Como executar

Com as dependências instaladas e o arquivo `.env` configurado, execute:

python main.py

O processamento dos lotes e o status das requisições serão exibidos no terminal durante a execução.

## Arquivo de entrada

O projeto espera um arquivo chamado `produtos.json` na raiz da aplicação.

Cada linha do arquivo deve representar um produto em formato JSON, contendo os seguintes campos:

```json
{
  "idproduto": 123456,
  "codEan": "7891234567890",
  "descricao": "Produto de exemplo"
}
```

O arquivo pode conter vários produtos, com um objeto JSON por linha.

## Arquivos de saída

Os arquivos gerados pela aplicação são armazenados na pasta `output/`.

### Revisão Fiscal

output/revisao_fiscal_v3.xlsx

Contém os produtos retornados pela API. Os dados são achatados automaticamente, transformando propriedades aninhadas do JSON em colunas do Excel.

### Erros

output/revisao_fiscal_v3_erros.xlsx

Gerado somente quando ocorrerem erros no processamento. Contém os produtos dos lotes com falha, o número do lote e o motivo do erro.