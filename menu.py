import tkinter as tk
from tkinter import ttk, messagebox


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Meu Sistema - Tela Principal")
        self.minsize(980, 560)

        # ====== Estilo (opcional) ======
        self.style = ttk.Style(self)
        # self.style.theme_use("clam")  # se quiser testar outro tema

        # ====== Layout principal ======
        self._build_layout()
        self._build_left_panel()
        self._build_right_panel()

        # Carrega dados iniciais
        self.load_data()

    def _build_layout(self):
        # Container principal
        self.container = ttk.Frame(self, padding=8)
        self.container.grid(row=0, column=0, sticky="nsew")

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)

        # PanedWindow para dividir esquerda/direita (redimensionável)
        self.paned = ttk.Panedwindow(self.container, orient=tk.HORIZONTAL)
        self.paned.grid(row=0, column=0, sticky="nsew")

        # Frames base
        self.left = ttk.Frame(self.paned, padding=(0, 0, 8, 0))
        self.right = ttk.Frame(self.paned)

        # Importante: adicionar com pesos
        self.paned.add(self.left, weight=1)   # esquerda
        self.paned.add(self.right, weight=2)  # direita

        # Config grid do frame esquerdo/direito
        self.left.rowconfigure(0, weight=1)   # frame_tree ocupa
        self.left.rowconfigure(1, weight=0)   # frame_buttons fica fixo
        self.left.columnconfigure(0, weight=1)

        self.right.rowconfigure(0, weight=1)
        self.right.columnconfigure(0, weight=1)

    def _build_left_panel(self):
        # ====== Frame de Filtros (opcional) ======
        self.frame_filters = ttk.LabelFrame(
            self.left, text="Filtros", padding=8)
        self.frame_filters.grid(row=0, column=0, sticky="new", pady=(0, 8))
        self.frame_filters.columnconfigure(1, weight=1)

        ttk.Label(self.frame_filters, text="Buscar:").grid(
            row=0, column=0, sticky="w")
        self.var_search = tk.StringVar()
        self.entry_search = ttk.Entry(
            self.frame_filters, textvariable=self.var_search)
        self.entry_search.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.entry_search.bind("<Return>", lambda e: self.apply_filter())

        ttk.Button(self.frame_filters, text="Filtrar", command=self.apply_filter).grid(
            row=0, column=2, padx=(8, 0), sticky="e"
        )

        # ====== Frame do Treeview (Tree + Scrollbar) ======
        self.frame_tree = ttk.Frame(self.left)
        self.frame_tree.grid(row=1, column=0, sticky="nsew")

        # Agora o frame_tree sim cresce
        self.left.rowconfigure(1, weight=1)

        self.frame_tree.rowconfigure(0, weight=1)
        self.frame_tree.columnconfigure(0, weight=1)

        columns = ("id", "nome", "status")
        self.tree = ttk.Treeview(
            self.frame_tree,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Cabeçalhos
        self.tree.heading("id", text="ID")
        self.tree.heading("nome", text="Nome")
        self.tree.heading("status", text="Status")

        # Larguras (ajuste como quiser)
        self.tree.column("id", width=80, anchor="center", stretch=False)
        self.tree.column("nome", width=240, anchor="w", stretch=True)
        self.tree.column("status", width=100, anchor="center", stretch=False)

        # Scrollbar vertical
        self.vsb = ttk.Scrollbar(
            self.frame_tree, orient="vertical", command=self.tree.yview)
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=self.vsb.set)

        # (Opcional) Scrollbar horizontal
        self.hsb = ttk.Scrollbar(
            self.frame_tree, orient="horizontal", command=self.tree.xview)
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=self.hsb.set)

        # Seleção
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # ====== Frame dos botões (fixo embaixo) ======
        self.frame_buttons = ttk.Frame(self.left)
        self.frame_buttons.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.frame_buttons.columnconfigure(0, weight=1)

        # Botões alinhados à direita (sem pack no mesmo container: aqui usamos grid)
        btns = ttk.Frame(self.frame_buttons)
        btns.grid(row=0, column=0, sticky="e")

        ttk.Button(btns, text="Novo", command=self.new_item).grid(
            row=0, column=0, padx=4)
        ttk.Button(btns, text="Editar", command=self.edit_item).grid(
            row=0, column=1, padx=4)
        ttk.Button(btns, text="Excluir", command=self.delete_item).grid(
            row=0, column=2, padx=4)
        ttk.Button(btns, text="Atualizar", command=self.load_data).grid(
            row=0, column=3, padx=4)

    def _build_right_panel(self):
        # Painel direito com detalhes / edição
        self.details = ttk.LabelFrame(self.right, text="Detalhes", padding=12)
        self.details.grid(row=0, column=0, sticky="nsew")
        self.details.columnconfigure(1, weight=1)

        ttk.Label(self.details, text="ID:").grid(row=0, column=0, sticky="w")
        self.var_id = tk.StringVar()
        self.entry_id = ttk.Entry(
            self.details, textvariable=self.var_id, state="readonly")
        self.entry_id.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=4)

        ttk.Label(self.details, text="Nome:").grid(row=1, column=0, sticky="w")
        self.var_nome = tk.StringVar()
        self.entry_nome = ttk.Entry(self.details, textvariable=self.var_nome)
        self.entry_nome.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=4)

        ttk.Label(self.details, text="Status:").grid(
            row=2, column=0, sticky="w")
        self.var_status = tk.StringVar(value="Ativo")
        self.combo_status = ttk.Combobox(
            self.details, textvariable=self.var_status, values=["Ativo", "Inativo"], state="readonly"
        )
        self.combo_status.grid(
            row=2, column=1, sticky="w", padx=(8, 0), pady=4)

        # Ações do formulário
        actions = ttk.Frame(self.details)
        actions.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))

        ttk.Button(actions, text="Salvar", command=self.save_item).grid(
            row=0, column=0, padx=4)
        ttk.Button(actions, text="Cancelar", command=self.clear_form).grid(
            row=0, column=1, padx=4)

        # Uma área inferior (ex.: logs/observações)
        self.notes = ttk.LabelFrame(self.right, text="Observações", padding=12)
        self.notes.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        self.right.rowconfigure(1, weight=1)
        self.notes.rowconfigure(0, weight=1)
        self.notes.columnconfigure(0, weight=1)

        self.txt_notes = tk.Text(self.notes, height=6, wrap="word")
        self.txt_notes.grid(row=0, column=0, sticky="nsew")

    # =========================
    # Dados / Regras de negócio
    # =========================

    def load_data(self):
        """Recarrega lista no Treeview."""
        # TODO: buscar no seu banco/arquivo/api
        sample = [
            (1, "Produto A", "Ativo"),
            (2, "Produto B", "Inativo"),
            (3, "Produto C", "Ativo"),
        ]

        self.tree.delete(*self.tree.get_children())
        for row in sample:
            self.tree.insert("", "end", values=row)

        self.clear_form()

    def apply_filter(self):
        """Filtra dados conforme campo buscar."""
        term = (self.var_search.get() or "").strip().lower()

        # Exemplo simples filtrando o que já está no tree (você pode filtrar na consulta do banco)
        all_rows = []
        for iid in self.tree.get_children():
            all_rows.append(self.tree.item(iid)["values"])

        self.tree.delete(*self.tree.get_children())

        for r in all_rows:
            if not term or term in str(r[1]).lower():
                self.tree.insert("", "end", values=r)

    def on_select(self, _event=None):
        """Quando seleciona uma linha, carrega no formulário."""
        iid = self.tree.focus()
        if not iid:
            return
        values = self.tree.item(iid)["values"]
        if not values:
            return

        self.var_id.set(values[0])
        self.var_nome.set(values[1])
        self.var_status.set(values[2])

    def clear_form(self):
        self.var_id.set("")
        self.var_nome.set("")
        self.var_status.set("Ativo")
        self.txt_notes.delete("1.0", "end")

    # ============
    # Ações CRUD
    # ============

    def new_item(self):
        """Modo novo registro."""
        self.tree.selection_remove(self.tree.selection())
        self.clear_form()
        self.entry_nome.focus_set()

    def edit_item(self):
        """Garante que algo esteja selecionado."""
        iid = self.tree.focus()
        if not iid:
            messagebox.showinfo("Editar", "Selecione um registro na lista.")
            return
        self.entry_nome.focus_set()

    def delete_item(self):
        iid = self.tree.focus()
        if not iid:
            messagebox.showinfo("Excluir", "Selecione um registro na lista.")
            return

        values = self.tree.item(iid)["values"]
        if not values:
            return

        if not messagebox.askyesno("Confirmar", f"Excluir o registro ID {values[0]}?"):
            return

        # TODO: excluir no banco
        self.tree.delete(iid)
        self.clear_form()

    def save_item(self):
        """Salva (novo ou edição)."""
        nome = (self.var_nome.get() or "").strip()
        status = (self.var_status.get() or "").strip()

        if not nome:
            messagebox.showwarning("Validação", "Informe o nome.")
            self.entry_nome.focus_set()
            return

        current_id = (self.var_id.get() or "").strip()

        if current_id:
            # Edição
            iid = self.tree.focus()
            if iid:
                # TODO: atualizar no banco usando current_id
                self.tree.item(iid, values=(int(current_id), nome, status))
        else:
            # Novo
            # TODO: gerar ID real via banco
            new_id = self._next_id()
            self.tree.insert("", "end", values=(new_id, nome, status))

        messagebox.showinfo("Salvar", "Registro salvo com sucesso.")
        self.clear_form()

    def _next_id(self):
        """Gera um ID simples baseado no Tree (substitua por seu banco)."""
        ids = []
        for iid in self.tree.get_children():
            v = self.tree.item(iid)["values"]
            if v:
                try:
                    ids.append(int(v[0]))
                except Exception:
                    pass
        return (max(ids) + 1) if ids else 1


if __name__ == "__main__":
    App().mainloop()
