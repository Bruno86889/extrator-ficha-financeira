import os
import re

import pdfplumber

from domain.dates import extrair_mes_ano_da_linha
from domain.models import Lancamento
from domain.text import normalizar_texto

QTDE_X_MIN = 118
QTDE_X_MAX = 166


MESES = {
    "JANEIRO": "Janeiro",
    "FEVEREIRO": "Fevereiro",
    "MARCO": "Março",
    "ABRIL": "Abril",
    "MAIO": "Maio",
    "JUNHO": "Junho",
    "JULHO": "Julho",
    "AGOSTO": "Agosto",
    "SETEMBRO": "Setembro",
    "OUTUBRO": "Outubro",
    "NOVEMBRO": "Novembro",
    "DEZEMBRO": "Dezembro",
}


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
                texto_norm = normalizar_texto(texto_linha)

                # A ficha tem cabeçalhos como "FOLHA NORMAL JANEIRO/2025".
                candidato_mes = extrair_mes_ano_da_linha(texto_linha)
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

