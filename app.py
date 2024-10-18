import openpyxl

book = openpyxl.load_workbook("./Vendas.xlsx")
sheet_vendas = book["Outubro"]

# Criar a lista com os produtos
vendas = []

for linha in sheet_vendas.iter_rows(min_row=2, values_only=True):
    produto = linha[0]
    quantidade = linha[1]
    preco = linha[2]
    vendas.append({"produto": produto, "quantidade": quantidade, "preco": preco})


# Calcular o total de vendas de um produto
def calcular_total_vendas(vendas):
    total_vendas = 0
    for venda in vendas:
        total_vendas += venda["quantidade"] * venda["preco"]
    return total_vendas


# Calcular lucro total (considerando custo de produção ficticio)
def calcular_lucro_total(vendas, custo_producao=0.6):
    lucro_total = 0
    for venda in vendas:
        lucro_por_item = venda["preco"] - venda["preco"] * custo_producao
        lucro_total += venda["quantidade"] * lucro_por_item
    return lucro_total


# Calcular media de vendas por produto
def calcular_media_vendas(vendas):
    return calcular_total_vendas(vendas) / len(vendas)


# Produto mais vendido
def produto_mais_vendido(vendas):
    mais_vendido = vendas[0]
    for venda in vendas:
        if venda["quantidade"] > mais_vendido["quantidade"]:
            mais_vendido = venda
    return mais_vendido["produto"]


# Calcular desemprenho da meta
def desempenho_meta(vendas, meta = 25800.00):
    meta_alcance = ((calcular_total_vendas(vendas)*100) / meta)
    return meta_alcance


def gerar_relatorio(vendas):
    print(
        f"""
"Relatório de vendas:"
{("-" * 50)}
Total de vendas:
R$ {calcular_total_vendas(vendas):.2f}

Lucro total:
R$ {calcular_lucro_total(vendas):.2f}

Média de vendas por produto:
R$ {calcular_media_vendas(vendas):.2f}

Produto mais vendido:
{produto_mais_vendido(vendas)}.

Desemprenho de vendas em relação a meta:
{desempenho_meta(vendas):.2f}% da meta atingido.

{("-" * 50)}
Observações:
O produto mais vendido foi responsável por uma significativa parcela das vendas.
O desempenho {"superou" if (desempenho_meta(vendas)> 100) else "não superou"} a meta, indicando um período {"positivo" if (desempenho_meta(vendas)> 100) else "negativo"}.
"""
)


gerar_relatorio(vendas)
