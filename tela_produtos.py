"""
tela_produtos.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Pré-requisitos:
  pip install psycopg2-binary

Ao fechar a janela:
  tenta abrir o menu principal (.py/.exe) em APP_DIR e BASE_DIR,
  e grava logs em:
    <BASE_DIR>/logs/tela_produtos.log
    <BASE_DIR>/logs/menu_principal_run.log
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tkinter as tk
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from tkinter import messagebox, ttk
from typing import Optional

import psycopg2


# ============================================================
# PASTAS / DIRS
# ============================================================

def get_app_dir() -> str:
    # quando empacotado (pyinstaller), __file__ não aponta para pasta real
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()

# Ajuste aqui
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


def log_app(msg: str) -> None:
    _log_write("tela_produtos.log", msg)


def log_menu_run(msg: str) -> None:
    _log_write("menu_principal_run.log", msg)


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
# MENU PRINCIPAL (ao fechar)
# ============================================================

# Ajuste a lista conforme o nome REAL do seu menu
MENU_FILENAMES = [
    "menu_principal.py",
    "menu.py",
    "Menu_Principal.py",
    "MenuPrincipal.py",
    "Ordem_Producao.py",
    "OrdemProducao.py",
    "menu_principal.exe",
    "menu.exe",
    "Menu_Principal.exe",
    "MenuPrincipal.exe",
    "Ordem_Producao.exe",
    "OrdemProducao.exe",
]


def _this_script_abspath() -> str:
    try:
        return os.path.abspath(sys.argv[0])
    except Exception:
        return ""


def localizar_menu_principal() -> str | None:
    """
    Procura o menu principal em APP_DIR e BASE_DIR.
    NUNCA devolve o próprio tela_produtos.py.
    """
    pastas = [APP_DIR, BASE_DIR]
    this_file = _this_script_abspath()

    # prioridade: se empacotado -> .exe primeiro; senão -> .py primeiro
    if getattr(sys, "frozen", False):
        nomes = [n for n in MENU_FILENAMES if n.lower().endswith(".exe")] + \
                [n for n in MENU_FILENAMES if n.lower().endswith(".py")]
    else:
        nomes = [n for n in MENU_FILENAMES if n.lower().endswith(".py")] + \
                [n for n in MENU_FILENAMES if n.lower().endswith(".exe")]

    for pasta in pastas:
        for nome in nomes:
            p = os.path.abspath(os.path.join(pasta, nome))
            if not os.path.isfile(p):
                continue
            # evita voltar para o próprio arquivo
            try:
                if this_file and os.path.samefile(p, this_file):
                    continue
            except Exception:
                # se samefile falhar, ignora
                pass
            return p
    return None


def _pick_python_launcher_windows() -> list[str]:
    """
    Retorna um comando (lista) para executar .py no Windows de forma confiável.
    Preferências:
      1) pyw -3.12
      2) pyw
      3) pythonw
      4) python
    """
    for cmd in (["pyw", "-3.12"], ["pyw"], ["pythonw"], ["python"]):
        try:
            subprocess.run(cmd + ["-c", "print('ok')"],
                           capture_output=True, text=True, timeout=2)
            return cmd
        except Exception:
            pass
    return ["python"]


def abrir_menu_principal() -> bool:
    """
    Abre o menu principal em um processo separado.
    Captura stdout/stderr do menu em logs/menu_principal_run.log
    """
    menu_path = localizar_menu_principal()

    log_menu_run("=== abrir_menu_principal() ===")
    log_menu_run(f"APP_DIR={APP_DIR}")
    log_menu_run(f"BASE_DIR={BASE_DIR}")
    log_menu_run(f"sys.executable={sys.executable}")
    log_menu_run(f"argv0={_this_script_abspath()}")
    log_menu_run(f"menu_path={menu_path}")

    if not menu_path:
        log_app("MENU: não encontrado.")
        return False

    try:
        cwd = os.path.dirname(menu_path) or APP_DIR
        is_exe = menu_path.lower().endswith(".exe")
        is_py = menu_path.lower().endswith(".py")

        if is_exe:
            cmd = [menu_path]
        elif is_py:
            if os.name == "nt":
                # evita WindowsApps; usa launcher do PATH
                cmd = _pick_python_launcher_windows() + [menu_path]
            else:
                cmd = [sys.executable, menu_path]
        else:
            cmd = [menu_path]

        log_menu_run(f"cwd={cwd}")
        log_menu_run(f"cmd={cmd}")

        popen_kwargs: dict = {"cwd": cwd}

        if os.name == "nt":
            popen_kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        else:
            popen_kwargs["start_new_session"] = True

        child_log_path = os.path.join(
            BASE_DIR, "logs", "menu_principal_run.log")
        with open(child_log_path, "a", encoding="utf-8") as out:
            out.write("\n\n=== START MENU PROCESS ===\n")
            out.write(f"cwd={cwd}\n")
            out.write(f"cmd={cmd}\n")
            out.flush()

            p = subprocess.Popen(
                cmd,
                stdout=out,
                stderr=out,
                close_fds=(os.name != "nt"),
                **popen_kwargs,
            )

        log_menu_run(f"popen_ok pid={getattr(p, 'pid', None)}")
        log_app(f"MENU: iniciado -> {menu_path}")
        return True

    except Exception as e:
        log_menu_run(f"ERRO: {type(e).__name__}: {e}")
        log_app(f"MENU: erro ao abrir {menu_path}: {type(e).__name__}: {e}")
        return False


# ============================================================
# DB (conexão direta)
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
# UI / MODEL
# ============================================================

DEFAULT_GEOMETRY = "1200x700"
APP_TITLE = "Tela de Produtos"


@dataclass
class Produto:
    produtoId: Optional[int]
    nomeProduto: str
    sku: Optional[str]
    preco: Decimal
    custo: Decimal
    tipo: Optional[str]
    formato: Optional[str]
    descricaoCurta: Optional[str]
    idProdutoPai: Optional[int]
    descImetro: Optional[str]


class ProdutosRepo:
    def __init__(self, db: Database) -> None:
        self.db = db

    def listar(self, termo: str | None = None, limit: int = 300) -> list[Produto]:
        sql = """
            SELECT p."produtoId", p."nomeProduto", p."sku", p."preco", p."custo", p."tipo", p."formato",
                   p."descricaoCurta", p."idProdutoPai", p."descImetro"
            FROM "Ekenox"."produtos" AS p
            WHERE (%s IS NULL)
               OR (p."nomeProduto" ILIKE %s)
               OR (p."sku" ILIKE %s)
               OR (CAST(p."produtoId" AS TEXT) ILIKE %s)
            ORDER BY p."nomeProduto"
            LIMIT %s
        """
        like = f"%{termo}%" if termo else None
        params = (termo, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(
                f"Falha ao conectar no banco: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            return [Produto(*row) for row in rows]
        finally:
            self.db.desconectar()

    def inserir(self, p: Produto) -> int:
        sql = """
            INSERT INTO "Ekenox"."produtos"
                ("nomeProduto", "sku", "preco", "custo", "tipo", "formato", "descricaoCurta", "idProdutoPai", "descImetro")
            VALUES
                (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING "produtoId"
        """
        params = (
            p.nomeProduto, p.sku, p.preco, p.custo, p.tipo, p.formato,
            p.descricaoCurta, p.idProdutoPai, p.descImetro
        )

        if not self.db.conectar():
            raise RuntimeError(
                f"Falha ao conectar no banco: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            new_id = self.db.cursor.fetchone()[0]
            assert self.db.conn is not None
            self.db.conn.commit()
            return int(new_id)
        finally:
            self.db.desconectar()

    def atualizar(self, p: Produto) -> None:
        if p.produtoId is None:
            raise ValueError("Código é obrigatório para atualizar.")

        sql = """
            UPDATE "Ekenox"."produtos"
               SET "nomeProduto" = %s,
                   "sku" = %s,
                   "preco" = %s,
                   "custo" = %s,
                   "tipo" = %s,
                   "formato" = %s,
                   "descricaoCurta" = %s,
                   "idProdutoPai" = %s,
                   "descImetro" = %s
             WHERE "produtoId" = %s
        """
        params = (
            p.nomeProduto, p.sku, p.preco, p.custo, p.tipo, p.formato,
            p.descricaoCurta, p.idProdutoPai, p.descImetro, p.produtoId
        )

        if not self.db.conectar():
            raise RuntimeError(
                f"Falha ao conectar no banco: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            assert self.db.conn is not None
            self.db.conn.commit()
        finally:
            self.db.desconectar()

    def excluir(self, produto_id: int) -> None:
        sql = 'DELETE FROM "Ekenox"."produtos" WHERE "produtoId" = %s'

        if not self.db.conectar():
            raise RuntimeError(
                f"Falha ao conectar no banco: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (produto_id,))
            assert self.db.conn is not None
            self.db.conn.commit()
        finally:
            self.db.desconectar()


def _clean_text(v: object) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def _to_int_or_none(v: object) -> int | None:
    s = "" if v is None else str(v).strip()
    if s == "":
        return None
    try:
        return int(s)
    except ValueError:
        raise ValueError("Código Pai/Código deve ser um inteiro (ou vazio).")


def _to_decimal(v: object) -> Decimal:
    s = "" if v is None else str(v).strip()
    if s == "":
        return Decimal("0")

    # aceita 1.234,56 e 1234,56 e 1234.56
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")

    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Valor numérico inválido: {v!r}")


class ProdutosService:
    def __init__(self, repo: ProdutosRepo) -> None:
        self.repo = repo

    def listar(self, termo: str | None = None) -> list[Produto]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo=termo)

    def salvar_from_form(self, form: dict) -> int | None:
        produtoId = _to_int_or_none(form.get("produtoId"))
        nomeProduto = (form.get("nomeProduto") or "").strip()
        if not nomeProduto:
            raise ValueError("Nome é obrigatório.")

        p = Produto(
            produtoId=produtoId,
            nomeProduto=nomeProduto,
            sku=_clean_text(form.get("sku")),
            preco=_to_decimal(form.get("preco")),
            custo=_to_decimal(form.get("custo")),
            tipo=_clean_text(form.get("tipo")),
            formato=_clean_text(form.get("formato")),
            descricaoCurta=_clean_text(form.get("descricaoCurta")),
            idProdutoPai=_to_int_or_none(form.get("idProdutoPai")),
            descImetro=_clean_text(form.get("descImetro")),
        )

        if p.preco < 0 or p.custo < 0:
            raise ValueError("preço/custo não podem ser negativos.")

        if p.produtoId is None:
            return self.repo.inserir(p)

        self.repo.atualizar(p)
        return None

    def excluir(self, produto_id: int) -> None:
        self.repo.excluir(produto_id)


CAMPOS = [
    ("produtoId", "Código"),
    ("nomeProduto", "Nome"),
    ("sku", "SKU"),
    ("preco", "Preço"),
    ("custo", "Custo"),
    ("tipo", "Tipo"),
    ("formato", "Formato"),
    ("descricaoCurta", "Descrição Curta"),
    ("idProdutoPai", "Código Pai"),
    ("descImetro", "Desc iMetro"),
]


class TelaProdutos(ttk.Frame):
    def __init__(self, master: tk.Misc, service: ProdutosService):
        super().__init__(master)
        self.service = service

        self.vars: dict[str, tk.StringVar] = {
            k: tk.StringVar() for k, _ in CAMPOS}
        self.var_filtro = tk.StringVar()

        self._build_ui()
        self.atualizar_lista()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (nome/sku/código):").grid(row=0,
                                                              column=0, sticky="w")
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

        form = ttk.LabelFrame(self, text="Produto")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(6):
            form.columnconfigure(c, weight=1)

        self._add_field(form, 0, 0, "produtoId", readonly=True, width=12)
        self._add_field(form, 0, 2, "nomeProduto", width=40)
        self._add_field(form, 0, 4, "sku", width=22)

        self._add_field(form, 1, 0, "preco", width=12)
        self._add_field(form, 1, 2, "custo", width=12)
        self._add_field(form, 1, 4, "idProdutoPai", width=14)

        self._add_field(form, 2, 0, "tipo", width=6)
        self._add_field(form, 2, 2, "formato", width=8)
        self._add_field(form, 2, 4, "descricaoCurta", width=28)

        self._add_field(form, 3, 0, "descImetro", colspan=6, width=80)

        lst_frame = ttk.Frame(self)
        lst_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        lst_frame.rowconfigure(0, weight=1)
        lst_frame.columnconfigure(0, weight=1)

        cols = [k for k, _ in CAMPOS]
        self.tree = ttk.Treeview(
            lst_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst_frame, orient="vertical",
                            command=self.tree.yview)
        hsb = ttk.Scrollbar(lst_frame, orient="horizontal",
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for key, label in CAMPOS:
            self.tree.heading(key, text=label)

        self.tree.column("produtoId", width=90, stretch=False, anchor="e")
        self.tree.column("nomeProduto", width=320, stretch=True, anchor="w")
        self.tree.column("sku", width=160, stretch=False, anchor="w")
        self.tree.column("preco", width=80, stretch=False, anchor="e")
        self.tree.column("custo", width=80, stretch=False, anchor="e")
        self.tree.column("tipo", width=40, stretch=False, anchor="center")
        self.tree.column("formato", width=55, stretch=False, anchor="center")
        self.tree.column("descricaoCurta", width=120, stretch=True, anchor="w")
        self.tree.column("idProdutoPai", width=90, stretch=False, anchor="e")
        self.tree.column("descImetro", width=160, stretch=True, anchor="w")

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def _add_field(self, parent: ttk.Frame, row: int, col: int, key: str,
                   readonly: bool = False, colspan: int = 2, width: int | None = None):
        label = dict(CAMPOS)[key]
        ttk.Label(parent, text=f"{label}:").grid(
            row=row, column=col, sticky="w", padx=(10, 6), pady=6)

        state = "readonly" if readonly else "normal"
        ent = ttk.Entry(parent, textvariable=self.vars[key], state=state)
        if width is not None:
            ent.configure(width=width)

        ent.grid(row=row, column=col + 1, sticky="ew",
                 padx=(0, 10), pady=6, columnspan=colspan - 1)

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            produtos = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar produtos:\n{e}")
            return

        for p in produtos:
            values = [
                p.produtoId,
                p.nomeProduto,
                p.sku or "",
                str(p.preco),
                str(p.custo),
                p.tipo or "",
                p.formato or "",
                p.descricaoCurta or "",
                "" if p.idProdutoPai is None else p.idProdutoPai,
                p.descImetro or "",
            ]
            self.tree.insert("", "end", values=values)

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        for i, (k, _) in enumerate(CAMPOS):
            self.vars[k].set("" if vals[i] in (None, "None") else str(vals[i]))

    def novo(self) -> None:
        self.limpar_form()
        self.vars["preco"].set("0")
        self.vars["custo"].set("0")

    def limpar_form(self) -> None:
        for k in self.vars:
            self.vars[k].set("")
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        form = {k: self.vars[k].get() for k, _ in CAMPOS}
        try:
            new_id = self.service.salvar_from_form(form)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        if new_id is not None:
            messagebox.showinfo("OK", f"Produto inserido com Código {new_id}.")
            self.vars["produtoId"].set(str(new_id))
        else:
            messagebox.showinfo("OK", "Produto atualizado.")

        self.atualizar_lista()

    def excluir(self) -> None:
        produto_id_str = self.vars["produtoId"].get().strip()
        if not produto_id_str:
            messagebox.showwarning(
                "Atenção", "Selecione um produto para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir o produto Código {produto_id_str}?"):
            return

        try:
            self.service.excluir(int(produto_id_str))
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Produto excluído.")
        self.limpar_form()
        self.atualizar_lista()


# ============================================================
# STARTUP
# ============================================================

def test_connection_or_die(cfg: AppConfig) -> None:
    db = Database(cfg)
    if not db.conectar():
        raise RuntimeError(db.ultimo_erro or "Erro desconhecido")
    try:
        assert db.cursor is not None
        db.cursor.execute("SELECT 1")
        db.cursor.fetchone()
    finally:
        db.desconectar()


def main():
    log_app("=== START tela_produtos ===")
    log_app(f"APP_DIR={APP_DIR}")
    log_app(f"BASE_DIR={BASE_DIR}")
    log_app(f"sys.executable={sys.executable}")
    log_app(f"argv0={_this_script_abspath()}")

    cfg = env_override(load_config())

    root = tk.Tk()
    root.title(APP_TITLE)
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
            f"Erro:\n{type(e).__name__}: {e}"
        )
        root.destroy()
        return

    db = Database(cfg)
    repo = ProdutosRepo(db)
    service = ProdutosService(repo)

    tela = TelaProdutos(root, service)
    tela.pack(fill="both", expand=True)

    closing = {"done": False}

    def open_menu_then_close():
        if closing["done"]:
            return
        closing["done"] = True

        log_app("WM_DELETE_WINDOW: acionado")
        ok = abrir_menu_principal()
        log_app(f"WM_DELETE_WINDOW: abrir_menu_principal() -> {ok}")

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
