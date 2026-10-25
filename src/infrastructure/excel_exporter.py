from collections import defaultdict

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from constants import MESES
from domain.numbers import converter_numero_br
from domain.text import normalizar_texto


def filtrar_por_vd_escolhido(registros, vd_escolhido):
    vd_escolhido_norm = normalizar_texto(vd_escolhido)

    lancamentos_filtrados = [
        r
        for r in registros
        if normalizar_texto(r["denominacao"]) == vd_escolhido_norm
        if r["natureza"] == "5"
    ]

    ordem_meses = {nome.title(): numero for nome, numero in MESES.items()}
    lancamentos_filtrados.sort(
        key=lambda r: (
            r["ano"] is None,
            r["ano"] or 9999,
            ordem_meses.get(r["mes"], 99),
            r["pagina"],
        )
    )

    return lancamentos_filtrados


def agrupar_por_mes(registros):
    agrupados = defaultdict(float)

    for lancamento in registros:
        ano = str(lancamento["ano"]).strip()
        mes = str(lancamento["mes"]).strip().upper()
        chave = (ano, mes)

        if not lancamento["qtde"]:
            agrupados[chave] += 0
        else:
            agrupados[chave] += converter_numero_br(lancamento["qtde"])  # type: ignore

    return agrupados


def gerar_excel(registros, nome_funcionario: str, vds: list[str], destino):
    vd_hora_aula = vds[0]
    vd_coordenacao = vds[1]
    vd_cp_oe = vds[2]
    vd_hae_outros = vds[3]
    registros_por_rubrica = {
        vd_hora_aula: agrupar_por_mes(
            filtrar_por_vd_escolhido(registros, vds[0])
        ),
        vd_coordenacao: agrupar_por_mes(
            filtrar_por_vd_escolhido(registros, vd_coordenacao)
        ),
        vd_cp_oe: agrupar_por_mes(filtrar_por_vd_escolhido(registros, vd_cp_oe)),
        vd_hae_outros: agrupar_por_mes(
            filtrar_por_vd_escolhido(registros, vd_hae_outros)
        ),
    }

    hora_aula = registros_por_rubrica[vd_hora_aula]
    coordenacao = registros_por_rubrica[vd_coordenacao]
    cp_oe = registros_por_rubrica[vd_cp_oe]
    hae_outros = registros_por_rubrica[vd_hae_outros]

    periodos = set(hora_aula) | set(coordenacao) | set(cp_oe) | set(hae_outros)

    ordem_meses = {
        "JANEIRO": 1,
        "FEVEREIRO": 2,
        "MARÇO": 3,
        "MARCO": 3,
        "ABRIL": 4,
        "MAIO": 5,
        "JUNHO": 6,
        "JULHO": 7,
        "AGOSTO": 8,
        "SETEMBRO": 9,
        "OUTUBRO": 10,
        "NOVEMBRO": 11,
        "DEZEMBRO": 12,
    }

    periodos = sorted(
        periodos,
        key=lambda periodo: (
            int(periodo[0]),
            ordem_meses.get(str(periodo[1]).upper(), 99),
        ),
    )

    wb = Workbook()
    ws = wb.active
    assert ws is not None

    ws.title = "Carga Horária"

    cinza = "D0D0D0"
    preto = "000000"

    preenchimento_cabecalho = PatternFill(
        fill_type="solid",
        fgColor=cinza,
    )

    borda_fina = Side(
        style="thin",
        color=preto,
    )

    borda = Border(
        left=borda_fina,
        right=borda_fina,
        top=borda_fina,
        bottom=borda_fina,
    )

    alinhamento_central = Alignment(
        horizontal="center",
        vertical="center",
    )

    fonte_cabecalho = Font(
        name="Arial",
        size=10,
        bold=True,
        color=preto,
    )

    fonte_dados = Font(
        name="Arial",
        size=10,
        color=preto,
    )

    ws.merge_cells("B2:C2")
    ws.merge_cells("D2:G2")

    ws["B2"] = "NOME"
    ws["D2"] = nome_funcionario.upper()

    for linha in ws.iter_rows(
        min_row=2,
        max_row=2,
        min_col=2,
        max_col=7,
    ):
        for celula in linha:
            celula.fill = preenchimento_cabecalho
            celula.border = borda
            celula.alignment = alinhamento_central
            celula.font = fonte_cabecalho

    ws.merge_cells("B4:B5")
    ws.merge_cells("C4:C5")
    ws.merge_cells("D4:G4")

    ws["B4"] = "ANO"
    ws["C4"] = "MÊS"
    ws["D4"] = "RUBRICAS"

    ws["D5"] = "HORA AULA SEMANAL"
    ws["E5"] = "COORDENAÇÃO"
    ws["F5"] = "CP/OE"
    ws["G5"] = "HAE OUTROS"

    for linha in ws.iter_rows(
        min_row=4,
        max_row=5,
        min_col=2,
        max_col=7,
    ):
        for celula in linha:
            celula.fill = preenchimento_cabecalho
            celula.border = borda
            celula.alignment = alinhamento_central
            celula.font = fonte_cabecalho

    for numero_linha, (ano, mes) in enumerate(periodos, start=6):
        valores = [
            int(ano),
            str(mes).upper(),
            hora_aula.get((ano, mes), 0),
            coordenacao.get((ano, mes), 0),
            cp_oe.get((ano, mes), 0),
            hae_outros.get((ano, mes), 0),
        ]

        for coluna, valor in enumerate(valores, start=2):
            celula = ws.cell(
                row=numero_linha,
                column=coluna,
                value=valor,
            )

            celula.font = fonte_dados
            celula.border = borda
            celula.alignment = alinhamento_central

            if coluna >= 4:
                celula.number_format = "#,##0.00"

    larguras = {
        "A": 3,
        "B": 9,
        "C": 14,
        "D": 22,
        "E": 22,
        "F": 22,
        "G": 22,
        "H": 3,
    }

    for coluna, largura in larguras.items():
        ws.column_dimensions[coluna].width = largura

    ws.row_dimensions[2].height = 20
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 20

    preenchimento_zero = PatternFill(fill_type="solid", fgColor="D9D9D9")
    fonte_zero = Font(name="Calibri", size=10, color="808080")

    ultima_linha = 5 + len(periodos)

    ws.conditional_formatting.add(
        f"D6:G{ultima_linha}",
        CellIsRule(
            operator="equal", formula=["0"], fill=preenchimento_zero, font=fonte_zero
        ),
    )  # type: ignore

    for linha in range(6, 6 + len(periodos)):
        ws.row_dimensions[linha].height = 20

    ws.sheet_view.showGridLines = False
    assert ws.sheet_properties.pageSetUpPr is not None

    wb.save(destino)

    return len(periodos)
