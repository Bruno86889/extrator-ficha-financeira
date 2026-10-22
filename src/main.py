import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

from infrastructure.excel_exporter import gerar_excel
from infrastructure.pdf_reader import extrair_linhas


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
    except Exception as exc:  # noqa: BLE001
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
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Erro ao gerar planilha", str(exc))


if __name__ == "__main__":
    main()
