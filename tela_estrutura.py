from __future__ import annotations

"""
tela_estrutura.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Tabela estrutura (conforme imagem):
  fkproduto, componente, quantidade, dados

Tabela produtos (conforme imagem):
  produtoId (PK), nomeProduto, ...

Recursos:
- CRUD completo
- Lista com nomes do Produto (fkproduto) e do Componente via JOIN em produtos
- Campos de nomes (readonly) preenchidos ao digitar IDs
- Config: BASE_DIR\\config_op.json + override DB_*
- Log em BASE_DIR\\logs\\tela_estrutura.log
- Ícone favicon.ico/png (se existir)
- Ao fechar: reabre menu_principal.py com --skip-entrada
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


def log_estrutura(msg: str) -> None:
    _log_write("tela_estrutura.log", msg)


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
        log_estrutura(
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
        log_estrutura(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_estrutura(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# AUTO-DETECT TABELAS (schema ou não)
# ============================================================

ESTRUTURA_TABLES = [
    '"Ekenox"."estrutura"',
    '"estrutura"',
]

PRODUTOS_TABLES = [
    '"Ekenox"."produtos"',
    '"produtos"',
]


def _table_exists(cfg: AppConfig, table_name: str) -> bool:
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
        cur.execute(f"SELECT 1 FROM {table_name} LIMIT 1")
        cur.fetchone()
        cur.close()
        conn.close()
        return True
    except Exception:
        try:
            conn.close()  # type: ignore[name-defined]
        except Exception:
            pass
        return False


def detectar_tabela(cfg: AppConfig, candidates: List[str], fallback: str) -> str:
    for t in candidates:
        ok = _table_exists(cfg, t)
        log_estrutura(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


# ============================================================
# MODEL
# ============================================================

@dataclass
class EstruturaRow:
    fkproduto: int
    produto_nome: str
    componente: int
    componente_nome: str
    quantidade: Decimal
    dados: Optional[str]


def _clean_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def _to_decimal(v: str, field_name: str) -> Decimal:
    s = (v or "").strip()
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

class EstruturaRepo:
    def __init__(self, db: Database, estrutura_table: str, produtos_table: str) -> None:
        self.db = db
        self.estrutura_table = estrutura_table
        self.produtos_table = produtos_table

    def buscar_produtos(self, termo: Optional[str], limit: int = 500) -> List[Tuple[int, str]]:
        termo = (termo or "").strip() or None
        like = f"%{termo}%" if termo else None

        sql = f"""
            SELECT p."produtoId", COALESCE(p."nomeProduto",'')
            FROM {self.produtos_table} AS p
            WHERE (%s IS NULL)
               OR (CAST(p."produtoId" AS TEXT) ILIKE %s)
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
            return [(int(r[0]), str(r[1] or "")) for r in rows]
        finally:
            self.db.desconectar()

    def listar(self, termo: Optional[str] = None, limit: int = 1500) -> List[EstruturaRow]:
        like = f"%{termo}%" if termo else None
        sql = f"""
            SELECT
                e."fkproduto",
                COALESCE(p1."nomeProduto",'') AS "produto_nome",
                e."componente",
                COALESCE(p2."nomeProduto",'') AS "componente_nome",
                e."quantidade",
                e."dados"
            FROM {self.estrutura_table} AS e
            LEFT JOIN {self.produtos_table} AS p1
                   ON p1."produtoId" = e."fkproduto"
            LEFT JOIN {self.produtos_table} AS p2
                   ON p2."produtoId" = e."componente"
            WHERE (%s IS NULL)
               OR (CAST(e."fkproduto" AS TEXT) ILIKE %s)
               OR (CAST(e."componente" AS TEXT) ILIKE %s)
               OR (COALESCE(p1."nomeProduto",'') ILIKE %s)
               OR (COALESCE(p2."nomeProduto",'') ILIKE %s)
            ORDER BY e."fkproduto", e."componente"
            LIMIT %s
        """
        params = (termo, like, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            out: List[EstruturaRow] = []
            for r in rows:
                out.append(EstruturaRow(
                    fkproduto=int(r[0]),
                    produto_nome=str(r[1] or ""),
                    componente=int(r[2]),
                    componente_nome=str(r[3] or ""),
                    quantidade=Decimal(str(r[4] if r[4] is not None else "0")),
                    dados=_clean_text(r[5]),
                ))
            return out
        finally:
            self.db.desconectar()

    def existe(self, fkproduto: int, componente: int) -> bool:
        sql = f'SELECT 1 FROM {self.estrutura_table} WHERE "fkproduto"=%s AND "componente"=%s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fkproduto, componente))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, fkproduto: int, componente: int, quantidade: Decimal, dados: Optional[str]) -> None:
        sql = f"""
            INSERT INTO {self.estrutura_table} ("fkproduto","componente","quantidade","dados")
            VALUES (%s,%s,%s,%s)
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                sql, (fkproduto, componente, quantidade, dados))
            self.db.commit()
        finally:
            self.db.desconectar()

    def atualizar(self, fkproduto: int, componente: int, quantidade: Decimal, dados: Optional[str]) -> None:
        sql = f"""
            UPDATE {self.estrutura_table}
               SET "quantidade"=%s,
                   "dados"=%s
             WHERE "fkproduto"=%s AND "componente"=%s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                sql, (quantidade, dados, fkproduto, componente))
            self.db.commit()
        finally:
            self.db.desconectar()

    def excluir(self, fkproduto: int, componente: int) -> None:
        sql = f'DELETE FROM {self.estrutura_table} WHERE "fkproduto"=%s AND "componente"=%s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fkproduto, componente))
            self.db.commit()
        finally:
            self.db.desconectar()

    def nome_produto(self, produto_id: int) -> str:
        sql = f'SELECT COALESCE("nomeProduto", \'\') FROM {self.produtos_table} WHERE "produtoId"=%s LIMIT 1'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (produto_id,))
            r = self.db.cursor.fetchone()
            return str(r[0] or "") if r else ""
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class EstruturaService:
    def __init__(self, repo: EstruturaRepo) -> None:
        self.repo = repo

    def buscar_produtos(self, termo: Optional[str]) -> List[Tuple[int, str]]:
        termo = (termo or "").strip() or None
        return self.repo.buscar_produtos(termo)

    def listar(self, termo: Optional[str]) -> List[EstruturaRow]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def preencher_nome(self, produto_id_txt: str) -> str:
        produto_id_txt = (produto_id_txt or "").strip()
        if not produto_id_txt:
            return ""
        try:
            pid = int(produto_id_txt)
        except ValueError:
            return ""
        nome = self.repo.nome_produto(pid)
        return nome if nome else "(não encontrado)"

    def salvar(self, fkproduto_txt: str, componente_txt: str, quantidade_txt: str, dados_txt: str) -> str:
        fkproduto_txt = (fkproduto_txt or "").strip()
        componente_txt = (componente_txt or "").strip()

        if not fkproduto_txt:
            raise ValueError("fkproduto é obrigatório.")
        if not componente_txt:
            raise ValueError("componente é obrigatório.")

        try:
            fkproduto = int(fkproduto_txt)
        except ValueError:
            raise ValueError("produto deve ser número inteiro.")

        try:
            componente = int(componente_txt)
        except ValueError:
            raise ValueError("componente deve ser número inteiro.")

        if fkproduto == componente:
            raise ValueError("componente não pode ser igual ao produto.")

        quantidade = _to_decimal(quantidade_txt, "quantidade")
        if quantidade <= 0:
            raise ValueError("quantidade deve ser maior que 0.")

        dados = _clean_text(dados_txt)

        if self.repo.existe(fkproduto, componente):
            self.repo.atualizar(fkproduto, componente, quantidade, dados)
            return "atualizado"
        else:
            self.repo.inserir(fkproduto, componente, quantidade, dados)
            return "inserido"

    def excluir(self, fkproduto: int, componente: int) -> None:
        self.repo.excluir(fkproduto, componente)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1250x700"
APP_TITLE = "Tela de Estrutura"

TREE_COLS = [
    "fkproduto", "produto_nome",
    "componente", "componente_nome",
    "quantidade", "dados",
]


class ProdutoPicker(tk.Toplevel):
    """
    Popup para escolher Produto/Componente.
    Retorna (produtoId, nomeProduto).
    """

    def __init__(self, master: tk.Misc, service: EstruturaService, titulo: str, on_pick):
        super().__init__(master)
        self.service = service
        self.on_pick = on_pick

        self.title(titulo)
        self.geometry("820x520")
        self.minsize(700, 420)
        apply_window_icon(self)

        self.var_busca = tk.StringVar()

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (ID/Nome):").grid(row=0,
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
            "id", "nome"), show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.heading("id", text="ID")
        self.tree.heading("nome", text="Nome")
        self.tree.column("id", width=130, anchor="e", stretch=False)
        self.tree.column("nome", width=620, anchor="w", stretch=True)

        self.tree.bind("<Double-1>", lambda e: self._pick())
        self.tree.bind("<Return>", lambda e: self._pick())

        bottom = ttk.Frame(self, padding=10)
        bottom.grid(row=2, column=0, sticky="ew")
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

        for pid, nome in rows:
            self.tree.insert("", "end", values=(pid, nome))

    def _pick(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selecionar", "Selecione um item.")
            return
        pid, nome = self.tree.item(sel[0], "values")
        try:
            self.on_pick(int(pid), str(nome))
        finally:
            self.destroy()


class TelaEstrutura(ttk.Frame):
    def __init__(self, master: tk.Misc, service: EstruturaService):
        super().__init__(master)
        self.service = service

        self.var_filtro = tk.StringVar()

        self.var_fkproduto = tk.StringVar()
        self.var_fk_nome = tk.StringVar()

        self.var_componente = tk.StringVar()
        self.var_comp_nome = tk.StringVar()

        self.var_quantidade = tk.StringVar(value="1")
        self.var_dados = tk.StringVar()

        self._build_ui()
        self.atualizar_lista()

    def buscar_produto_popup(self) -> None:
        def on_pick(pid: int, nome: str) -> None:
            self.var_fkproduto.set(str(pid))
            self.var_fk_nome.set(nome)

        ProdutoPicker(self.winfo_toplevel(), self.service,
                      "Buscar Produto", on_pick)

    def buscar_componente_popup(self) -> None:
        def on_pick(pid: int, nome: str) -> None:
            self.var_componente.set(str(pid))
            self.var_comp_nome.set(nome)

        ProdutoPicker(self.winfo_toplevel(), self.service,
                      "Buscar Componente", on_pick)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Topo
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (ID/Nome):").grid(row=0,
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

        # Form
        form = ttk.LabelFrame(self, text="Estrutura")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(12):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="Código:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_fk = ttk.Entry(form, textvariable=self.var_fkproduto, width=14)
        ent_fk.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=6)
        ent_fk.bind("<Return>", lambda e: self._preencher_nome_fk())

        ttk.Button(form, text="Buscar Produto...", command=self.buscar_produto_popup).grid(
            row=0, column=2, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Produto (Nome):").grid(
            row=0, column=3, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_fk_nome, state="readonly").grid(
            row=0, column=4, sticky="ew", padx=(0, 10), pady=6, columnspan=8
        )

        ttk.Label(form, text="componente:").grid(
            row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_comp = ttk.Entry(form, textvariable=self.var_componente, width=14)
        ent_comp.grid(row=1, column=1, sticky="w", padx=(0, 10), pady=6)
        ent_comp.bind("<Return>", lambda e: self._preencher_nome_comp())

        ttk.Button(form, text="Buscar Componente...", command=self.buscar_componente_popup).grid(
            row=1, column=2, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Componente (Nome):").grid(
            row=1, column=3, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_comp_nome, state="readonly").grid(
            row=1, column=4, sticky="ew", padx=(0, 10), pady=6, columnspan=8
        )

        ttk.Label(form, text="Quantidade:").grid(
            row=2, column=0, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_quantidade, width=16).grid(
            row=2, column=1, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Dados:").grid(
            row=2, column=3, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_dados).grid(
            row=2, column=4, sticky="ew", padx=(0, 10), pady=6, columnspan=8
        )

        # Lista
        lst = ttk.Frame(self)
        lst.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        lst.rowconfigure(0, weight=1)
        lst.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            lst, columns=TREE_COLS, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(lst, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.heading("fkproduto", text="Código produto")
        self.tree.heading("produto_nome", text="Produto (Nome)")
        self.tree.heading("componente", text="Código componente")
        self.tree.heading("componente_nome", text="Componente (Nome)")
        self.tree.heading("quantidade", text="Quantidade")
        self.tree.heading("dados", text="Dados")

        self.tree.column("fkproduto", width=120, anchor="e", stretch=False)
        self.tree.column("produto_nome", width=340, anchor="w", stretch=True)
        self.tree.column("componente", width=120, anchor="e", stretch=False)
        self.tree.column("componente_nome", width=340,
                         anchor="w", stretch=True)
        self.tree.column("quantidade", width=120, anchor="e", stretch=False)
        self.tree.column("dados", width=300, anchor="w", stretch=True)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def _preencher_nome_fk(self) -> None:
        self.var_fk_nome.set(self.service.repo.nome_produto(int(self.var_fkproduto.get() or "0"))
                             if (self.var_fkproduto.get().strip().isdigit()) else "")

    def _preencher_nome_comp(self) -> None:
        self.var_comp_nome.set(self.service.repo.nome_produto(int(self.var_componente.get() or "0"))
                               if (self.var_componente.get().strip().isdigit()) else "")

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar estrutura:\n{e}")
            return

        for r in rows:
            self.tree.insert("", "end", values=(
                r.fkproduto, r.produto_nome,
                r.componente, r.componente_nome,
                str(r.quantidade),
                r.dados or "",
            ))

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        fk, fk_nome, comp, comp_nome, qtd, dados = self.tree.item(
            sel[0], "values")

        self.var_fkproduto.set(str(fk))
        self.var_fk_nome.set(str(fk_nome or ""))

        self.var_componente.set(str(comp))
        self.var_comp_nome.set(str(comp_nome or ""))

        self.var_quantidade.set(str(qtd or "1"))
        self.var_dados.set(str(dados or ""))

    def novo(self) -> None:
        self.limpar_form()
        self.var_quantidade.set("1")

    def limpar_form(self) -> None:
        self.var_fkproduto.set("")
        self.var_fk_nome.set("")
        self.var_componente.set("")
        self.var_comp_nome.set("")
        self.var_quantidade.set("1")
        self.var_dados.set("")
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        try:
            status = self.service.salvar(
                self.var_fkproduto.get(),
                self.var_componente.get(),
                self.var_quantidade.get(),
                self.var_dados.get(),
            )
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        self._preencher_nome_fk()
        self._preencher_nome_comp()

        messagebox.showinfo("OK", f"Estrutura {status} com sucesso.")
        self.atualizar_lista()

    def excluir(self) -> None:
        fk_txt = (self.var_fkproduto.get() or "").strip()
        comp_txt = (self.var_componente.get() or "").strip()

        if not fk_txt.isdigit() or not comp_txt.isdigit():
            messagebox.showwarning(
                "Atenção", "Informe fkproduto e componente para excluir.")
            return

        fk = int(fk_txt)
        comp = int(comp_txt)

        if not messagebox.askyesno("Confirmar", f"Excluir Estrutura produto={fk} / componente={comp}?"):
            return

        try:
            self.service.excluir(fk, comp)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Registro excluído.")
        self.limpar_form()
        self.atualizar_lista()


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
            f"Erro:\n{e}"
        )
        root.destroy()
        return

    estrutura_table = detectar_tabela(cfg, ESTRUTURA_TABLES, '"estrutura"')
    produtos_table = detectar_tabela(cfg, PRODUTOS_TABLES, '"produtos"')

    db = Database(cfg)
    repo = EstruturaRepo(db, estrutura_table=estrutura_table,
                         produtos_table=produtos_table)
    service = EstruturaService(repo)

    tela = TelaEstrutura(root, service)
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
    try:
        main()
    except Exception as e:
        log_estrutura(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_estrutura:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
