from __future__ import annotations

"""
tela_arranjo.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Tabela arranjo (conforme imagem):
  sku, nomeproduto, quantidade, chapa, material

Requisito:
- Ter botão que busque SKU e Nome do Produto na tabela produtos.

Ao fechar:
- reabre menu_principal.py com --skip-entrada (para não abrir tela de entrada).
"""

import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Optional, Any, List, Tuple

import psycopg2


# ============================================================
# TABELAS (AJUSTE SE NECESSÁRIO)
# ============================================================

ARRANJO_TABLE = '"Ekenox"."arranjo"'
PRODUTOS_TABLE = '"Ekenox"."produtos"'


# ============================================================
# PASTAS / BASE DIR
# ============================================================

def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()

BASE_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
os.makedirs(BASE_DIR, exist_ok=True)


# ============================================================
# LOG
# ============================================================

def _log_write(filename: str, msg: str) -> None:
    try:
        log_dir = os.path.join(BASE_DIR, "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, filename)
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def log_arranjo(msg: str) -> None:
    _log_write("tela_arranjo.log", msg)


# ============================================================
# ÍCONE
# ============================================================

def find_icon_path() -> Optional[str]:
    candidates = [
        os.path.join(BASE_DIR, "imagens", "favicon.ico"),
        os.path.join(BASE_DIR, "favicon.ico"),
        os.path.join(APP_DIR, "imagens", "favicon.ico"),
        os.path.join(APP_DIR, "favicon.ico"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def apply_window_icon(win) -> None:
    try:
        ico = find_icon_path()
        if ico:
            try:
                win.iconbitmap(default=ico)
            except Exception:
                pass

        png_candidates = [
            os.path.join(BASE_DIR, "imagens", "favicon.png"),
            os.path.join(APP_DIR, "imagens", "favicon.png"),
        ]
        png = next((p for p in png_candidates if os.path.isfile(p)), None)
        if png:
            img = tk.PhotoImage(file=png)
            win.iconphoto(True, img)
            win._icon_img = img  # mantém referência
    except Exception:
        pass


# ============================================================
# CONFIG DO APP (BANCO)
# ============================================================

@dataclass
class AppConfig:
    db_host: str = "10.0.0.154"
    db_database: str = "postgresekenox"
    db_user: str = "postgresekenox"
    db_password: str = "Ekenox5426"
    db_port: int = 55432


def config_path() -> str:
    return os.path.join(BASE_DIR, "config_op.json")


def load_config() -> AppConfig:
    p = config_path()
    if not os.path.exists(p):
        return AppConfig()
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        return AppConfig(**data)
    except Exception:
        return AppConfig()


def env_override(cfg: AppConfig) -> AppConfig:
    host = (os.getenv("DB_HOST") or "").strip() or cfg.db_host
    port_s = (os.getenv("DB_PORT") or "").strip()
    dbname = (os.getenv("DB_NAME") or os.getenv(
        "DB_DATABASE") or "").strip() or cfg.db_database
    user = (os.getenv("DB_USER") or "").strip() or cfg.db_user
    password = (os.getenv("DB_PASSWORD") or "").strip() or cfg.db_password

    try:
        port = int(port_s) if port_s else int(cfg.db_port)
    except ValueError:
        port = int(cfg.db_port)

    return AppConfig(
        db_host=host,
        db_port=port,
        db_database=dbname,
        db_user=user,
        db_password=password,
    )


# ============================================================
# DB
# ============================================================

class Database:
    def __init__(self, cfg: AppConfig) -> None:
        self.cfg = cfg
        self.conn = None
        self.cursor = None
        self.ultimo_erro: Optional[str] = None

    def conectar(self) -> bool:
        self.ultimo_erro = None
        try:
            self.conn = psycopg2.connect(
                host=self.cfg.db_host,
                database=self.cfg.db_database,
                user=self.cfg.db_user,
                password=self.cfg.db_password,
                port=int(self.cfg.db_port),
                connect_timeout=5,
            )
            self.cursor = self.conn.cursor()
            return True
        except Exception as e:
            self.ultimo_erro = f"{type(e).__name__}: {e}"
            return False

    def commit(self) -> None:
        if self.conn:
            self.conn.commit()

    def desconectar(self) -> None:
        try:
            if self.cursor:
                try:
                    self.cursor.close()
                except Exception:
                    pass
            if self.conn:
                try:
                    self.conn.close()
                except Exception:
                    pass
        finally:
            self.cursor = None
            self.conn = None


# ============================================================
# MENU PRINCIPAL (ao fechar)
# ============================================================

MENU_FILENAMES = [
    "menu_principal.py",
    "Menu_Principal.py",
    "MenuPrincipal.py",
    "menu.py",
    "menu_principal.exe",
    "Menu_Principal.exe",
    "MenuPrincipal.exe",
    "menu.exe",
]


def localizar_menu_principal() -> Optional[str]:
    pastas = [APP_DIR, BASE_DIR]
    for pasta in pastas:
        for nome in MENU_FILENAMES:
            p = os.path.join(pasta, nome)
            if os.path.isfile(p):
                return os.path.abspath(p)
    return None


def _python_gui_windows() -> str:
    py = sys.executable
    if os.name == "nt":
        base = os.path.basename(py).lower()
        if base == "python.exe":
            pyw = os.path.join(os.path.dirname(py), "pythonw.exe")
            if os.path.isfile(pyw):
                return pyw
    return py


def abrir_menu_principal_skip_entrada() -> None:
    menu_path = localizar_menu_principal()
    if not menu_path:
        log_arranjo(
            f"MENU: não encontrado. APP_DIR={APP_DIR} BASE_DIR={BASE_DIR}")
        return

    try:
        cwd = os.path.dirname(menu_path) or APP_DIR

        if menu_path.lower().endswith(".exe"):
            if os.name == "nt":
                os.startfile(menu_path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen([menu_path], cwd=cwd)
            return

        py = _python_gui_windows() if os.name == "nt" else sys.executable
        cmd = [py, menu_path, "--skip-entrada"]

        popen_kwargs: dict[str, Any] = {"cwd": cwd}
        if os.name == "nt":
            popen_kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        else:
            popen_kwargs["start_new_session"] = True

        subprocess.Popen(cmd, **popen_kwargs)
        log_arranjo(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_arranjo(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# MODEL
# ============================================================

@dataclass
class Arranjo:
    sku: str
    nomeproduto: Optional[str]
    quantidade: Decimal
    chapa: Optional[str]
    material: Optional[str]


# ============================================================
# CONVERSÕES
# ============================================================

def _clean_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def _to_decimal(v: Any, field_name: str) -> Decimal:
    s = "" if v is None else str(v).strip()
    if s == "":
        return Decimal("0")

    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")

    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(f"{field_name} inválido: {v!r}")


# ============================================================
# REPOSITORY
# ============================================================

class ArranjoRepo:
    def __init__(self, db: Database) -> None:
        self.db = db

    def listar(self, termo: Optional[str] = None, limit: int = 500) -> list[Arranjo]:
        like = f"%{termo}%" if termo else None

        sql = f"""
            SELECT
                a."sku", a."nomeproduto", a."quantidade", a."chapa", a."material"
            FROM {ARRANJO_TABLE} AS a
            WHERE (%s IS NULL)
               OR (COALESCE(a."sku",'') ILIKE %s)
               OR (COALESCE(a."nomeproduto",'') ILIKE %s)
               OR (COALESCE(a."chapa",'') ILIKE %s)
               OR (COALESCE(a."material",'') ILIKE %s)
            ORDER BY a."sku"
            LIMIT %s
        """
        params = (termo, like, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            return [Arranjo(*row) for row in rows]
        finally:
            self.db.desconectar()

    def exists(self, sku: str) -> bool:
        sql = f'SELECT 1 FROM {ARRANJO_TABLE} WHERE "sku" = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (sku,))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, a: Arranjo) -> None:
        sql = f"""
            INSERT INTO {ARRANJO_TABLE}
            ("sku","nomeproduto","quantidade","chapa","material")
            VALUES (%s,%s,%s,%s,%s)
        """
        params = (a.sku, a.nomeproduto, a.quantidade, a.chapa, a.material)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            self.db.commit()
        finally:
            self.db.desconectar()

    def atualizar(self, a: Arranjo, sku_original: str) -> None:
        # Se o usuário mudar o SKU, atualizamos usando sku_original
        sql = f"""
            UPDATE {ARRANJO_TABLE}
               SET "sku" = %s,
                   "nomeproduto" = %s,
                   "quantidade" = %s,
                   "chapa" = %s,
                   "material" = %s
             WHERE "sku" = %s
        """
        params = (a.sku, a.nomeproduto, a.quantidade,
                  a.chapa, a.material, sku_original)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            self.db.commit()
        finally:
            self.db.desconectar()

    def excluir(self, sku: str) -> None:
        sql = f'DELETE FROM {ARRANJO_TABLE} WHERE "sku" = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (sku,))
            self.db.commit()
        finally:
            self.db.desconectar()

    # -------- lookup produtos --------

    def produto_por_sku(self, sku: str) -> Optional[Tuple[str, str]]:
        # retorna (sku, nomeProduto)
        sql = f"""
            SELECT p."sku", p."nomeProduto"
            FROM {PRODUTOS_TABLE} AS p
            WHERE p."sku" = %s
            LIMIT 1
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (sku,))
            row = self.db.cursor.fetchone()
            if not row:
                return None
            return str(row[0] or ""), str(row[1] or "")
        finally:
            self.db.desconectar()

    def buscar_produtos(self, termo: Optional[str], limit: int = 300) -> List[Tuple[str, str]]:
        like = f"%{termo}%" if termo else None
        sql = f"""
            SELECT p."sku", p."nomeProduto"
            FROM {PRODUTOS_TABLE} AS p
            WHERE (%s IS NULL)
               OR (COALESCE(p."sku",'') ILIKE %s)
               OR (COALESCE(p."nomeProduto",'') ILIKE %s)
            ORDER BY p."nomeProduto"
            LIMIT %s
        """
        params = (termo, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            return [(str(r[0] or ""), str(r[1] or "")) for r in rows]
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class ArranjoService:
    def __init__(self, repo: ArranjoRepo) -> None:
        self.repo = repo

    def listar(self, termo: Optional[str]) -> list[Arranjo]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo=termo)

    def salvar_from_form(self, form: dict[str, Any], sku_original: Optional[str]) -> str:
        sku = (form.get("sku") or "").strip()
        if not sku:
            raise ValueError("SKU é obrigatório.")

        a = Arranjo(
            sku=sku,
            nomeproduto=_clean_text(form.get("nomeproduto")),
            quantidade=_to_decimal(form.get("quantidade"), "Quantidade"),
            chapa=_clean_text(form.get("chapa")),
            material=_clean_text(form.get("material")),
        )

        if a.quantidade < 0:
            raise ValueError("Quantidade não pode ser negativa.")

        if sku_original and self.repo.exists(sku_original):
            # update
            self.repo.atualizar(a, sku_original=sku_original)
            return "atualizado"
        else:
            # insert
            if self.repo.exists(a.sku):
                # evita duplicar quando usuário esquece selecionado
                self.repo.atualizar(a, sku_original=a.sku)
                return "atualizado"
            self.repo.inserir(a)
            return "inserido"

    def excluir(self, sku: str) -> None:
        self.repo.excluir(sku)

    def preencher_nome_por_sku(self, sku: str) -> Optional[Tuple[str, str]]:
        sku = (sku or "").strip()
        if not sku:
            return None
        return self.repo.produto_por_sku(sku)

    def buscar_produtos(self, termo: Optional[str]) -> List[Tuple[str, str]]:
        termo = (termo or "").strip() or None
        return self.repo.buscar_produtos(termo)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1100x700"
APP_TITLE = "Tela de Arranjo"

CAMPOS = [
    ("sku", "SKU"),
    ("nomeproduto", "Nome Produto"),
    ("quantidade", "Quantidade"),
    ("chapa", "Chapa"),
    ("material", "Material"),
]

TREE_COLS = ["sku", "nomeproduto", "quantidade", "chapa", "material"]


class ProdutoPicker(tk.Toplevel):
    """
    Popup para pesquisar e escolher SKU + NomeProduto na tabela produtos.
    """

    def __init__(self, master: tk.Misc, service: ArranjoService, on_pick):
        super().__init__(master)
        self.service = service
        self.on_pick = on_pick

        self.title("Buscar Produto")
        self.geometry("720x450")
        self.minsize(650, 380)
        apply_window_icon(self)

        self.var_busca = tk.StringVar()

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (SKU/Nome):").grid(row=0,
                                                       column=0, sticky="w")
        ent = ttk.Entry(top, textvariable=self.var_busca)
        ent.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ent.bind("<Return>", lambda e: self._load())

        ttk.Button(top, text="Pesquisar",
                   command=self._load).grid(row=0, column=2)

        lst = ttk.Frame(self, padding=(10, 0, 10, 10))
        lst.grid(row=1, column=0, sticky="nsew")
        lst.rowconfigure(0, weight=1)
        lst.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(lst, columns=(
            "sku", "nome"), show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.heading("sku", text="SKU")
        self.tree.heading("nome", text="Nome Produto")
        self.tree.column("sku", width=150, anchor="w", stretch=False)
        self.tree.column("nome", width=480, anchor="w", stretch=True)

        self.tree.bind("<Double-1>", lambda e: self._pick())
        self.tree.bind("<Return>", lambda e: self._pick())

        bottom = ttk.Frame(self, padding=10)
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        ttk.Button(bottom, text="Selecionar", command=self._pick).pack(
            side="right", padx=(8, 0))
        ttk.Button(bottom, text="Fechar",
                   command=self.destroy).pack(side="right")

        self._load()
        ent.focus_set()

        self.transient(master)
        self.grab_set()

    def _load(self) -> None:
        termo = self.var_busca.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.buscar_produtos(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao buscar produtos:\n{e}")
            return

        for sku, nome in rows:
            self.tree.insert("", "end", values=(sku, nome))

    def _pick(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selecionar", "Selecione um produto.")
            return
        sku, nome = self.tree.item(sel[0], "values")
        try:
            self.on_pick(str(sku), str(nome))
        finally:
            self.destroy()


class TelaArranjo(ttk.Frame):
    def __init__(self, master: tk.Misc, service: ArranjoService):
        super().__init__(master)
        self.service = service

        self.vars: dict[str, tk.StringVar] = {
            k: tk.StringVar() for k, _ in CAMPOS}
        self.var_filtro = tk.StringVar()

        # guarda chave do registro selecionado (para update mesmo se SKU mudar)
        self._sku_original: Optional[str] = None

        self._build_ui()
        self.atualizar_lista()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Topo (filtro + botões)
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(
            top, text="Buscar (SKU / Nome / Chapa / Material):").grid(row=0, column=0, sticky="w")
        ent_busca = ttk.Entry(top, textvariable=self.var_filtro)
        ent_busca.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent_busca.bind("<Return>", lambda e: self.atualizar_lista())

        ttk.Button(top, text="Atualizar", command=self.atualizar_lista).grid(
            row=0, column=2, padx=(0, 6))
        ttk.Button(top, text="Novo", command=self.novo).grid(
            row=0, column=3, padx=(0, 6))
        ttk.Button(top, text="Salvar", command=self.salvar).grid(
            row=0, column=4, padx=(0, 6))
        ttk.Button(top, text="Excluir", command=self.excluir).grid(
            row=0, column=5, padx=(0, 6))
        ttk.Button(top, text="Limpar", command=self.limpar_form).grid(
            row=0, column=6)

        # Formulário
        form = ttk.LabelFrame(self, text="Arranjo")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(6):
            form.columnconfigure(c, weight=1)

        # Linha 0: SKU + botões
        ttk.Label(form, text="SKU:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_sku = ttk.Entry(form, textvariable=self.vars["sku"])
        ent_sku.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=6)

        ttk.Button(form, text="Buscar Produto...", command=self.buscar_produto_popup).grid(
            row=0, column=2, sticky="w", padx=(0, 6), pady=6
        )
        ttk.Button(form, text="Preencher", command=self.preencher_nome_por_sku).grid(
            row=0, column=3, sticky="w", padx=(0, 6), pady=6
        )

        ttk.Label(form, text="Nome Produto:").grid(
            row=0, column=4, sticky="w", padx=(10, 6), pady=6)
        ent_nome = ttk.Entry(form, textvariable=self.vars["nomeproduto"])
        ent_nome.grid(row=0, column=5, sticky="ew", padx=(0, 10), pady=6)

        # Linha 1
        self._add_field(form, 1, 0, "quantidade", width=14)
        self._add_field(form, 1, 2, "chapa", width=22)
        self._add_field(form, 1, 4, "material", width=22)

        # Lista
        lst_frame = ttk.Frame(self)
        lst_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        lst_frame.rowconfigure(0, weight=1)
        lst_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            lst_frame, columns=TREE_COLS, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst_frame, orient="vertical",
                            command=self.tree.yview)
        hsb = ttk.Scrollbar(lst_frame, orient="horizontal",
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        headings = {
            "sku": "SKU",
            "nomeproduto": "Nome Produto",
            "quantidade": "Quantidade",
            "chapa": "Chapa",
            "material": "Material",
        }
        for col in TREE_COLS:
            self.tree.heading(col, text=headings.get(col, col))

        self.tree.column("sku", width=140, anchor="w", stretch=False)
        self.tree.column("nomeproduto", width=380, anchor="w", stretch=True)
        self.tree.column("quantidade", width=100, anchor="e", stretch=False)
        self.tree.column("chapa", width=150, anchor="w", stretch=False)
        self.tree.column("material", width=180, anchor="w", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # atalhos
        ent_sku.bind("<F3>", lambda e: self.buscar_produto_popup())
        ent_sku.bind("<Return>", lambda e: self.preencher_nome_por_sku())

    def _add_field(
        self,
        parent: ttk.Frame,
        row: int,
        col: int,
        key: str,
        readonly: bool = False,
        colspan: int = 2,
        width: int | None = None,
    ) -> None:
        label = dict(CAMPOS)[key]
        ttk.Label(parent, text=f"{label}:").grid(
            row=row, column=col, sticky="w", padx=(10, 6), pady=6)

        state = "readonly" if readonly else "normal"
        ent = ttk.Entry(parent, textvariable=self.vars[key], state=state)
        if width is not None:
            ent.configure(width=width)

        ent.grid(row=row, column=col + 1, sticky="ew",
                 padx=(0, 10), pady=6, columnspan=colspan - 1)

    # ---------- ações ----------
    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None

        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            itens = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar Arranjo:\n{e}")
            return

        for a in itens:
            values = [
                a.sku,
                a.nomeproduto or "",
                str(a.quantidade),
                a.chapa or "",
                a.material or "",
            ]
            self.tree.insert("", "end", values=values)

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")

        self.vars["sku"].set(str(vals[0] or ""))
        self.vars["nomeproduto"].set(str(vals[1] or ""))
        self.vars["quantidade"].set(str(vals[2] or "0"))
        self.vars["chapa"].set(str(vals[3] or ""))
        self.vars["material"].set(str(vals[4] or ""))

        self._sku_original = str(vals[0] or "").strip() or None

    def novo(self) -> None:
        self.limpar_form()
        self.vars["quantidade"].set("0")

    def limpar_form(self) -> None:
        for k in self.vars:
            self.vars[k].set("")
        self._sku_original = None
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        form = {k: self.vars[k].get() for k, _ in CAMPOS}

        try:
            status = self.service.salvar_from_form(
                form, sku_original=self._sku_original)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        messagebox.showinfo("OK", f"Arranjo {status} com sucesso.")
        self.atualizar_lista()

        # se mudou sku, atualiza sku_original para o novo
        self._sku_original = (self.vars["sku"].get().strip() or None)

    def excluir(self) -> None:
        sku = self.vars["sku"].get().strip()
        if not sku:
            messagebox.showwarning(
                "Atenção", "Informe/Selecione um SKU para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Arranjo do SKU {sku}?"):
            return

        try:
            self.service.excluir(sku)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Arranjo excluído.")
        self.limpar_form()
        self.atualizar_lista()

    # ---------- lookup produtos ----------
    def buscar_produto_popup(self) -> None:
        def on_pick(sku: str, nome: str) -> None:
            self.vars["sku"].set(sku)
            self.vars["nomeproduto"].set(nome)
            if not self.vars["quantidade"].get().strip():
                self.vars["quantidade"].set("0")

        ProdutoPicker(self.winfo_toplevel(), self.service, on_pick)

    def preencher_nome_por_sku(self) -> None:
        sku = self.vars["sku"].get().strip()
        if not sku:
            messagebox.showwarning("SKU", "Informe um SKU para buscar.")
            return
        try:
            r = self.service.preencher_nome_por_sku(sku)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao buscar produto:\n{e}")
            return

        if not r:
            messagebox.showinfo(
                "Produto", "SKU não encontrado na tabela de produtos.")
            return

        sku_db, nome_db = r
        self.vars["sku"].set(sku_db)
        self.vars["nomeproduto"].set(nome_db)


# ============================================================
# STARTUP
# ============================================================

def test_connection_or_die(cfg: AppConfig) -> None:
    try:
        conn = psycopg2.connect(
            host=cfg.db_host,
            database=cfg.db_database,
            user=cfg.db_user,
            password=cfg.db_password,
            port=int(cfg.db_port),
            connect_timeout=5,
        )
        cur = conn.cursor()
        cur.execute("SELECT 1")
        cur.fetchone()
        cur.close()
        conn.close()
    except Exception as e:
        raise RuntimeError(f"{type(e).__name__}: {e}")


def main() -> None:
    cfg = env_override(load_config())

    root = tk.Tk()
    root.title("Tela de Arranjo")
    root.geometry(DEFAULT_GEOMETRY)
    apply_window_icon(root)

    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    try:
        test_connection_or_die(cfg)
    except Exception as e:
        messagebox.showerror(
            "Erro de conexão",
            "Não foi possível conectar ao banco.\n\n"
            f"Host: {cfg.db_host}\n"
            f"Porta: {cfg.db_port}\n"
            f"Banco: {cfg.db_database}\n"
            f"Usuário: {cfg.db_user}\n\n"
            f"Erro:\n{e}"
        )
        root.destroy()
        return

    db = Database(cfg)
    repo = ArranjoRepo(db)
    service = ArranjoService(repo)

    tela = TelaArranjo(root, service)
    tela.pack(fill="both", expand=True)

    closing = {"done": False}

    def open_menu_then_close():
        if closing["done"]:
            return
        closing["done"] = True
        try:
            abrir_menu_principal_skip_entrada()
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def on_close():
        try:
            root.withdraw()
            root.update_idletasks()
        except Exception:
            pass
        root.after(50, open_menu_then_close)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
