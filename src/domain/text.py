import re


def normalizar_texto(texto):
     """Normaliza espaços e caixa para comparar denominações."""
     return re.sub(r"\s+", " ", texto or "").strip().upper()
