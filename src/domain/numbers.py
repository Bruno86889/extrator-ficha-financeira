def converter_numero_br(valor):
    if valor is None or valor == "":
        return None
    return float(valor.replace(".", "").replace(",", "."))
