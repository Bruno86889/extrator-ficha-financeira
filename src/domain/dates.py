import re

from .text import normalizar_texto

MESES = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3,
    "ABRIL": 4, "MAIO": 5, "JUNHO": 6, "JULHO": 7,
    "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10,
    "NOVEMBRO": 11, "DEZEMBRO": 12,
}

def extrair_mes_ano_da_linha(texto: str)-> tuple[str,str]:
    normalizado = normalizar_texto(texto)

    match = re.search(
            r"\b(" + "|".join(MESES.keys()) + r")\s*/\s*(20\d{2})\b",
            normalizado,
        )
    if not match:
        return None
    return match.group(1).title(), int(match.group(2))



