import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import pdfplumber
from openpyxl import Workbook

MESES = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3,
    "ABRIL": 4, "MAIO": 5, "JUNHO": 6, "JULHO": 7,
    "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10,
    "NOVEMBRO": 11, "DEZEMBRO": 12,
}

# Posições horizontais aproximadas da coluna Qtde na ficha financeira.
# O script identifica a coluna pela posição no PDF, não pela ordem dos números.
QTDE_X_MIN = 118
QTDE_X_MAX = 166


def normalizar(texto):
    """Normaliza espaços e caixa para comparar denominações."""
    return re.sub(r"\s+", " ", texto or "").strip().upper()


def mes_ano_da_linha(texto):
    """Retorna (mês, ano) quando a linha contém, por exemplo, JANEIRO/2025."""
    texto = normalizar(texto)
    match = re.search(
        r"\b(" + "|".join(MESES.keys()) + r")\s*/\s*(20\d{2})\b",
        texto,
    )
    if not match:
        return None
    return match.group(1).title(), int(match.group(2))


def extrair_linhas(pdf_path):
    """
    Extrai as linhas usando coordenadas dos elementos do PDF.
    Isso permite distinguir Qtde de Valor, Percentual e outros números.
    """
    registros = []
    rubricas = set()
    mes_atual = None
    secao_atual = ""

    with pdfplumber.open(pdf_path) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            palavras = pagina.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            )

            # Agrupa palavras que estão na mesma linha visual.
            linhas = {}
            for palavra in palavras:
                chave_y = round(palavra["top"] / 3) * 3
                linhas.setdefault(chave_y, []).append(palavra)

            for y in sorted(linhas):
                palavras_linha = sorted(linhas[y], key=lambda p: p["x0"])
                texto_linha = " ".join(p["text"] for p in palavras_linha)
                texto_norm = normalizar(texto_linha)

                # A ficha tem cabeçalhos como "FOLHA NORMAL JANEIRO/2025".
                candidato_mes = mes_ano_da_linha(texto_linha)
                if candidato_mes and (
                    "FOLHA" in texto_norm
                    or "JANEIRO/" in texto_norm
                    or "FEVEREIRO/" in texto_norm
                    or "MARÇO/" in texto_norm
                    or "MARCO/" in texto_norm
                    or "ABRIL/" in texto_norm
                    or "MAIO/" in texto_norm
                    or "JUNHO/" in texto_norm
                    or "JULHO/" in texto_norm
                    or "AGOSTO/" in texto_norm
                    or "SETEMBRO/" in texto_norm
                    or "OUTUBRO/" in texto_norm
                    or "NOVEMBRO/" in texto_norm
                    or "DEZEMBRO/" in texto_norm
                ):
                    mes_atual = candidato_mes

                # Guarda o tipo de folha/seção para ajudar a identificar duplicidades.
                if any(k in texto_norm for k in (
                    "FOLHA NORMAL", "FERIAS COLETIVAS", "FÉRIAS COLETIVAS",
                    "ADIANTAMENTO", "13.SALARIO", "13. SALARIO",
                    "BR RESULTADOS",
                )):
                    secao_atual = texto_linha.strip()

                # Rubricas começam com código de 6 dígitos e têm a denominação à direita.
                if not re.match(r"^\s*\d{6}\b", texto_linha):
                    continue

                # Denominação está na parte direita da tabela; ignora campos antes dela.
                palavras_denominacao = [p for p in palavras_linha if p["x0"] >= 300]
                denominacao = " ".join(p["text"] for p in palavras_denominacao).strip()
                if not denominacao:
                    continue

                rubricas.add(denominacao)

                # Natureza fica depois da rubrica, e tem sempre valor 5 ou 8
                natureza = re.search(r"\b(5|8)\b", texto_linha)
                natureza = natureza.group(1) if natureza else None

                # Identifica a quantidade usando a posição horizontal da coluna Qtde.
                # Se a célula estiver vazia, Qtde fica como None.
                qtde_palavras = [
                    p for p in palavras_linha
                    if QTDE_X_MIN <= p["x0"] <= QTDE_X_MAX
                    and re.fullmatch(r"-?\d[\d.]*,\d{2}", p["text"])
                ]
                qtde = qtde_palavras[0]["text"] if qtde_palavras else None

                codigo = re.match(r"^\s*(\d{6})", texto_linha).group(1)

                registros.append({
                    "ano": mes_atual[1] if mes_atual else None,
                    "mes": mes_atual[0] if mes_atual else None,
                    "codigo": codigo,
                    "denominacao": denominacao,
                    "natureza": natureza,
                    "qtde": qtde,
                    "secao": secao_atual,
                    "pagina": numero_pagina,
                    "linha_original": texto_linha.strip(),
                })

    return registros, sorted(rubricas, key=str.casefold)


def converter_numero_br(valor):
    if valor is None or valor == "":
        return None
    return float(valor.replace(".", "").replace(",", "."))


def gerar_excel(registros, vd_escolhido, destino):
    vd_escolhido_norm = normalizar(vd_escolhido)

    # Correspondência por texto normalizado. Permite escolher a denominação exibida.
    lancamentos_filtrados = [
        r for r in registros
        if normalizar(r["denominacao"]) == vd_escolhido_norm
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


def main():
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Selecione a ficha financeira em PDF",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
    if not pdf_path:
        return

    try:
        registros, rubricas = extrair_linhas(pdf_path)
    except Exception as exc:
        messagebox.showerror(
            "Erro ao ler PDF",
            f"Não foi possível ler o arquivo.\n\nDetalhes: {exc}\n\n"
            "Confira se o PDF contém texto selecionável (não apenas imagens).",
        )
        return

    if not rubricas:
        messagebox.showerror(
            "Nenhuma rubrica encontrada",
            "Não foram identificadas rubricas no PDF.",
        )
        return

    # Escolha por número para manter compatibilidade com diferentes versões do Python.
    lista = "\n".join(f"{i + 1}. {nome}" for i, nome in enumerate(rubricas))
    escolha = simpledialog.askinteger(
        "Escolher rubrica",
        "Digite o número da rubrica que deseja extrair:\n\n" + lista,
        minvalue=1,
        maxvalue=len(rubricas),
    )
    if escolha is None:
        return

    rubrica = rubricas[escolha - 1]
    destino = filedialog.asksaveasfilename(
        title="Salvar planilha Excel",
        defaultextension=".xlsx",
        initialfile=f"extracao_{re.sub(r'[^A-Za-z0-9]+', '_', rubrica).strip('_')}.xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")],
    )
    if not destino:
        return

    try:
        quantidade = gerar_excel(registros, rubrica, destino)
        messagebox.showinfo(
            "Extração concluída",
            f"Rubrica: {rubrica}\n"
            f"Lançamentos encontrados: {quantidade}\n\n"
            f"Planilha salva em:\n{destino}",
        )
    except Exception as exc:
        messagebox.showerror("Erro ao gerar planilha", str(exc))


if __name__ == "__main__":
    main()
