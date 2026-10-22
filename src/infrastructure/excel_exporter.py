

from openpyxl import Workbook

from domain.text import normalizar_texto

MESES = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3,
    "ABRIL": 4, "MAIO": 5, "JUNHO": 6, "JULHO": 7,
    "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10,
    "NOVEMBRO": 11, "DEZEMBRO": 12,
}

def converter_numero_br(valor):
    if valor is None or valor == "":
        return None
    return float(valor.replace(".", "").replace(",", "."))


def gerar_excel(registros, vd_escolhido, destino):
    vd_escolhido_norm = normalizar_texto(vd_escolhido)

    # Correspondência por texto normalizado. Permite escolher a denominação exibida.
    lancamentos_filtrados = [
        r for r in registros
        if normalizar_texto(r["denominacao"]) == vd_escolhido_norm
    ]

    if not lancamentos_filtrados:
        raise ValueError("Nenhum lançamento encontrado para a rubrica escolhida.")

    # Ordena cronologicamente; linhas sem mês identificado ficam no final.
    ordem_meses = {nome.title(): numero for nome, numero in MESES.items()}
    lancamentos_filtrados.sort(key=lambda r: (
        r["ano"] is None,
        r["ano"] or 9999,
        ordem_meses.get(r["mes"], 99),
        r["pagina"],
    ))

    wb = Workbook()
    ws = wb.active
    ws.title = "Extração"

    cabecalhos = [
        "Ano", "Mês", "Rubrica", "VD", "Qtde", "Página do PDF",
    ]
    ws.append(cabecalhos)

    for lancamento in lancamentos_filtrados:
        ws.append([
            lancamento["ano"],
            lancamento["mes"],
            lancamento["denominacao"],
            f"{lancamento["codigo"]} {lancamento["natureza"]}",
            converter_numero_br(lancamento["qtde"]),
            lancamento["pagina"],
        ])

    wb.save(destino)
    return len(lancamentos_filtrados)
