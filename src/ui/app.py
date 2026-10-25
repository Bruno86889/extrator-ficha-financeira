import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import constants
from infrastructure.excel_exporter import gerar_excel
from infrastructure.pdf_reader import extrair_linhas

from .styles import COLORS
from .widgets.header import AppHeader


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Extrator de Folha de Pagamento")
        self.configure(bg=COLORS["background"])

        # Estado da interface
        self.pdf_path = tk.StringVar()
        self.vds_selecionados = tk.StringVar()
        self.vds_selecionados = []
        self.nome_empregado = tk.StringVar()

        self.vd_hora_aula_var = tk.StringVar()
        self.vd_coordenacao_var = tk.StringVar()
        self.vd_cp_oe_var = tk.StringVar()
        self.vd_hae_outros_var = tk.StringVar()

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

        AppHeader(root, COLORS).pack(fill="x")

        # Cartão: arquivo de origem
        file_card = self._card(root)
        self._section_title(file_card, "Arquivo de origem")
        file_row = tk.Frame(file_card, bg=COLORS["surface"])
        file_row.pack(fill="x", pady=(12, 0))

        self.file_icon = tk.Label(
            file_row,
            text="PDF",
            bg="#FEF2F2",
            fg=COLORS["danger"],
            font=("Segoe UI", 10, "bold"),
            width=5,
            pady=10,
        )
        self.file_icon.pack(side="left", padx=(0, 12))

        file_info = tk.Frame(file_row, bg=COLORS["surface"])
        file_info.pack(side="left", fill="x", expand=True)

        self.file_name_label = tk.Label(
            file_info,
            text="Nenhum arquivo selecionado",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        )
        self.file_name_label.pack(fill="x")

        self.file_path_label = tk.Label(
            file_info,
            text="Selecione a ficha financeira em formato PDF",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            anchor="w",
        )
        self.file_path_label.pack(fill="x", pady=(4, 0))

        self.select_pdf_button = ttk.Button(
            file_row,
            text="Selecionar PDF",
            style="Secondary.TButton",
            command=self.select_pdf,
        )
        self.select_pdf_button.pack(side="right", padx=(12, 0))

        employee_card = self._card(root)
        self._section_title(employee_card, "Nome do Funcionário")

        self.input_empregado = ttk.Entry(
            employee_card,
            font=("Segoe UI", 10),
            background=COLORS["surface"],
            foreground=COLORS["text"],
            textvariable=self.nome_empregado,
        )

        self.input_empregado.pack(fill="x", pady=(12, 0))

        # Card de selação de VD de hora aula
        hora_aula_card = self._card(root)
        self._section_title(hora_aula_card, "Hora Aula")

        self.vd_combo_hora_aula = ttk.Combobox(
            hora_aula_card,
            textvariable=self.vd_hora_aula_var,
            state="readonly",
            style="App.TCombobox",
            font=("Segoe UI", 10),
        )

        self.vd_combo_hora_aula.pack(fill="x", pady=(6, 0))
        self.vd_combo_hora_aula.set(constants.VD_HORA_AULA)

        # Card de selação de VD de coordenação
        coordenacao_card = self._card(root)
        self._section_title(coordenacao_card, "HAE Coordenação")

        self.vd_combo_coordenacao = ttk.Combobox(
            coordenacao_card,
            textvariable=self.vd_coordenacao_var,
            state="readonly",
            style="App.TCombobox",
            font=("Segoe UI", 10),
        )

        self.vd_combo_coordenacao.pack(fill="x", pady=(6, 0))
        self.vd_combo_coordenacao.set(constants.VD_COORDENACAO)

        # Card de selação de VD de CP ou OE
        cp_or_card = self._card(root)
        self._section_title(cp_or_card, "CP/OE")

        self.vd_combo_cp_oe = ttk.Combobox(
            cp_or_card,
            textvariable=self.vd_cp_oe_var,
            state="readonly",
            style="App.TCombobox",
            font=("Segoe UI", 10),
        )

        self.vd_combo_cp_oe.pack(fill="x", pady=(6, 0))
        self.vd_combo_cp_oe.set(constants.VD_CP_OE)

        # Card de selação de HAE Outros
        cp_or_card = self._card(root)
        self._section_title(cp_or_card, "HAE Outros")

        self.vd_combo_hae_outros = ttk.Combobox(
            cp_or_card,
            textvariable=self.vd_hae_outros_var,
            state="readonly",
            style="App.TCombobox",
            font=("Segoe UI", 10),
        )

        self.vd_combo_hae_outros.pack(fill="x", pady=(6, 0))
        self.vd_combo_hae_outros.set(constants.VD_HAE_OUTROS)

        # Rodapé com status e exportação
        footer = tk.Frame(root, bg=COLORS["background"])
        footer.pack(fill="x", pady=(2, 0))

        self.export_button = ttk.Button(
            footer,
            text="Exportar para Excel",
            style="Primary.TButton",
            command=self.export_excel,
        )
        self.export_button.pack(side="right")
        
        self.clear_button = ttk.Button(
            footer,
            text="Limpar",
            style="Secondary.TButton",
            command=self._clear,
        )
        self.clear_button.pack(side="right")

    def _card(self, parent):
        card = tk.Frame(
            parent,
            bg=COLORS["surface"],
            padx=18,
            pady=16,
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        card.pack(fill="x", pady=(0, 14))
        return card

    def _section_title(self, parent, text):
        tk.Label(
            parent,
            text=text,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        ).pack(fill="x")

    def select_pdf(self):
        """Seleciona o PDF e inicia a leitura em segundo plano."""
        selected = filedialog.askopenfilename(
            title="Selecione a ficha financeira",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )

        if not selected:
            return

        # Atualiza a interface antes de iniciar o processamento.
        self.select_pdf_button.config(state="disabled")
        self.export_button.config(state="disabled")

        self.file_name_label.config(text=Path(selected).name)
        self.file_path_label.config(text="Processando PDF, aguarde...")

        # A leitura será executada fora da thread da interface.
        thread = threading.Thread(
            target=self._process_pdf,
            args=(selected,),
            daemon=True,
        )
        thread.start()

    def _process_pdf(self, selected):
        """Executa o processamento sem bloquear a interface."""
        try:
            registros, rubricas = extrair_linhas(selected)

            # A atualização dos widgets ocorre na thread principal.
            self.after(
                0,
                lambda: self._on_pdf_processed(selected, registros, rubricas, None),
            )

        except Exception as exc:  # noqa: BLE001
            self.after(
                0,
                lambda error=exc: self._on_pdf_processed(selected, None, None, error),
            )

    def _on_pdf_processed(self, selected, registros, rubricas, error):
        """Finaliza o processamento e atualiza a interface."""
        self.select_pdf_button.config(state="normal")
        self.export_button.config(state="normal")

        if error is not None:
            self.file_path_label.config(text="Não foi possível processar o PDF.")
            messagebox.showerror(
                "Erro ao ler PDF",
                str(error),
            )
            return

        self.pdf_path.set(selected)
        self.registros = registros

        self.vd_combo_hora_aula["values"] = rubricas
        self.vd_combo_coordenacao["values"] = rubricas
        self.vd_combo_cp_oe["values"] = rubricas
        self.vd_combo_hae_outros["values"] = rubricas

        self.file_name_label.config(text=Path(selected).name)
        self.file_path_label.config(text="PDF processado com sucesso.")

    def export_excel(self):
        caminho_arquivo = self.pdf_path.get()

        if not caminho_arquivo:
            messagebox.showwarning(
                "Arquivo necessário", "Selecione um arquivo PDF primeiro."
            )
            return

        rubrica = self.vd_hora_aula_var.get().strip()

        if not rubrica or rubrica == "Selecione primeiro um arquivo PDF":
            messagebox.showwarning("Rubrica necessária", "Selecione uma rubrica.")
            return

        nome_arquivo = os.path.basename(caminho_arquivo).split(".")[0]
        nome_arquivo_resuldado = nome_arquivo + "_extraido"

        destino = filedialog.asksaveasfilename(
            title="Salvar planilha Excel",
            defaultextension=".xlsx",
            initialfile=nome_arquivo_resuldado,
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if not destino:
            return

        try:
            vds = [
                self.vd_hora_aula_var.get(),
                self.vd_coordenacao_var.get(),
                self.vd_cp_oe_var.get(),
                self.vd_hae_outros_var.get(),
            ]
            gerar_excel(self.registros, self.input_empregado.get(), vds, destino)
            messagebox.showinfo(
                "Extração concluída",
                f"Planilha salva em:\n{destino}",
            )

            self._clear()

        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erro ao gerar planilha", str(exc))

    def _clear(self):
        self.pdf_path.set("")
        self.vd_hora_aula_var.set(constants.VD_HORA_AULA)
        self.vd_coordenacao_var.set(constants.VD_COORDENACAO)
        self.vd_cp_oe_var.set(constants.VD_CP_OE)
        self.vd_hae_outros_var.set(constants.VD_HAE_OUTROS)
        self.file_name_label.config(text="Nenhum arquivo selecionado")
        self.file_path_label.config(text="Selecione a ficha financeira em formato PDF")
        self.nome_empregado.set("")
