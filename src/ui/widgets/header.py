import tkinter as tk


class AppHeader(tk.Frame):
 def __init__(self, master, colors):
        super().__init__(
            master,
            bg=colors["surface"],
            padx=18,
            pady=16,
            highlightbackground=colors["border"],
            highlightthickness=1,
        )

        header = tk.Frame(master, bg=colors["background"])
        header.pack(fill="x", pady=(0, 22))

        icon_box = tk.Frame(header, bg=colors["primary"], width=52, height=52)
        icon_box.pack(side="left", padx=(0, 14))
        icon_box.pack_propagate(False)
        tk.Label(
            icon_box,
            text="",
            bg=colors["primary"],
            fg="#FFFFFF",
            font=("Segoe UI", 13, "bold"),
        ).pack(expand=True)

        title_box = tk.Frame(header, bg=colors["background"])
        title_box.pack(side="left", fill="x", expand=True)

        tk.Label(
            title_box,
            text="Extrator de Folha de Pagamento",
            bg=colors["background"],
            fg=colors["text"],
            font=("Segoe UI", 20, "bold"),
            anchor="w",
        ).pack(anchor="w")
    
       