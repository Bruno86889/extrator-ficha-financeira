import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from infrastructure.excel_exporter import gerar_excel
from infrastructure.pdf_reader import extrair_linhas

# Paleta visual
COLORS = {
    "background": "#F4F6FA",
    "surface": "#FFFFFF",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "text": "#172033",
    "muted": "#667085",
    "border": "#D9E0EA",
    "success": "#16A34A",
    "danger": "#DC2626",
    "soft_blue": "#EFF6FF",
    "soft_green": "#ECFDF3",
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Extrator de Folha de Pagamento")
        self.geometry("900x720")
        self.minsize(760, 650)
        self.configure(bg=COLORS["background"])

        # Estado da interface
        self.pdf_path = tk.StringVar()
        self.rubrica_var_01 = tk.StringVar()

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "App.TCombobox",
            padding=9,
            fieldbackground=COLORS["surface"],
            background=COLORS["surface"],
            foreground=COLORS["text"],
            arrowcolor=COLORS["muted"],
            bordercolor=COLORS["border"],
        )
        style.map(
            "App.TCombobox",
            fieldbackground=[("readonly", COLORS["surface"])],
            selectbackground=[("readonly", COLORS["surface"])],
            selectforeground=[("readonly", COLORS["text"])],
        )

        style.configure(
            "Primary.TButton",
            padding=(16, 11),
            background=COLORS["primary"],
            foreground="#FFFFFF",
            borderwidth=0,
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["primary_hover"]), ("disabled", "#AAB8D0")],
            foreground=[("disabled", "#FFFFFF")],
        )

        style.configure(
            "Secondary.TButton",
            padding=(12, 9),
            background=COLORS["surface"],
            foreground=COLORS["text"],
            borderwidth=1,
            font=("Segoe UI", 9),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#F2F5FA")],
        )

    def _build_ui(self):
        # Container principal
        root = tk.Frame(self, bg=COLORS["background"])
        root.pack(fill="both", expand=True, padx=30, pady=26)

        # Cabeçalho
        header = tk.Frame(root, bg=COLORS["background"])
        header.pack(fill="x", pady=(0, 22))

        icon_box = tk.Frame(header, bg=COLORS["primary"], width=52, height=52)
        icon_box.pack(side="left", padx=(0, 14))
        icon_box.pack_propagate(False)
        tk.Label(
            icon_box, text="", bg=COLORS["primary"], fg="#FFFFFF",
            font=("Segoe UI", 13, "bold"),
        ).pack(expand=True)

        title_box = tk.Frame(header, bg=COLORS["background"])
        title_box.pack(side="left", fill="x", expand=True)

        tk.Label(
            title_box, text="Extrator de Folha de Pagamento",
            bg=COLORS["background"], fg=COLORS["text"],
            font=("Segoe UI", 20, "bold"), anchor="w",
        ).pack(anchor="w")


        # Cartão: arquivo de origem
        file_card = self._card(root)
        self._section_title(file_card, "1. Arquivo de origem")
        file_row = tk.Frame(file_card, bg=COLORS["surface"])
        file_row.pack(fill="x", pady=(12, 0))

        self.file_icon = tk.Label(
            file_row, text="PDF", bg="#FEF2F2", fg=COLORS["danger"],
            font=("Segoe UI", 10, "bold"), width=5, pady=10,
        )
        self.file_icon.pack(side="left", padx=(0, 12))

        file_info = tk.Frame(file_row, bg=COLORS["surface"])
        file_info.pack(side="left", fill="x", expand=True)

        self.file_name_label = tk.Label(
            file_info, text="Nenhum arquivo selecionado",
            bg=COLORS["surface"], fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"), anchor="w",
        )
        self.file_name_label.pack(fill="x")

        self.file_path_label = tk.Label(
            file_info, text="Selecione a ficha financeira em formato PDF",
            bg=COLORS["surface"], fg=COLORS["muted"],
            font=("Segoe UI", 9), anchor="w",
        )
        self.file_path_label.pack(fill="x", pady=(4, 0))

        ttk.Button(
            file_row, text="Selecionar PDF", style="Secondary.TButton",
            command=self.select_pdf,
        ).pack(side="right", padx=(12, 0))

        # Cartão: rubrica
        rubric_card = self._card(root)
        self._section_title(rubric_card, "2. Rubrica para extrair")

        self.rubrica_combo = ttk.Combobox(
            rubric_card,
            textvariable=self.rubrica_var_01,
            values=[],
            state="readonly",
            style="App.TCombobox",
            font=("Segoe UI", 10),
        )

        self.rubrica_combo.pack(fill="x", pady=(6, 0))
        self.rubrica_combo.set("Selecione primeiro um arquivo PDF")

        # Rodapé com status e exportação
        footer = tk.Frame(root, bg=COLORS["background"])
        footer.pack(fill="x", pady=(2, 0))


        self.export_button = ttk.Button(
            footer, text="Exportar para Excel", style="Primary.TButton",
            command=self.export_excel,
        )
        self.export_button.pack(side="right")

    def _card(self, parent):
        card = tk.Frame(
            parent, bg=COLORS["surface"], padx=18, pady=16,
            highlightbackground=COLORS["border"], highlightthickness=1,
        )
        card.pack(fill="x", pady=(0, 14))
        return card

    def _section_title(self, parent, text):
        tk.Label(
            parent, text=text, bg=COLORS["surface"], fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"), anchor="w",
        ).pack(fill="x")

    def select_pdf(self):
        """Abre o seletor de PDF. Integre aqui a leitura e análise do arquivo."""
        selected = filedialog.askopenfilename(
            title="Selecione a ficha financeira",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )
        if not selected:
            return

        _, rubricas = extrair_linhas(selected)

        self.rubrica_combo["values"] = rubricas

        if rubricas:
            self.rubrica_combo.current(0)

        self.pdf_path.set(selected)
        self.file_name_label.config(text=Path(selected).name)
        self.file_path_label.config(text=selected)

    def export_excel(self):
        """Valida a seleção e pede o destino. Integre aqui a geração do Excel."""
        if not self.pdf_path.get():
            messagebox.showwarning("Arquivo necessário", "Selecione um arquivo PDF primeiro.")
            return

        rubrica = self.rubrica_var_01.get().strip()
        if not rubrica or rubrica == "Selecione primeiro um arquivo PDF":
            messagebox.showwarning("Rubrica necessária", "Selecione uma rubrica.")
            return

        destino = filedialog.asksaveasfilename(
            title="Salvar planilha Excel",
            defaultextension=".xlsx",
            initialfile="extracao_folha.xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if not destino:
            return

        try:
            print(self.pdf_path.get())
            registros, _ = extrair_linhas(self.pdf_path.get())
            quantidade = gerar_excel(registros, rubrica, destino)
            messagebox.showinfo(
                    "Extração concluída",
                    f"Rubrica: {rubrica}\n"
                    f"Lançamentos encontrados: {quantidade}\n\n"
                    f"Planilha salva em:\n{destino}",
                )

            self.pdf_path.set("")
            self.rubrica_combo.set("")
            self.file_name_label.config(text="")
            self.file_path_label.config(text="")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erro ao gerar planilha", str(exc))
        
