from __future__ import annotations

"""
tela_info_produto.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2

- CRUD para infoProduto
- Lookups (Produto, Categoria, Fornecedor, Localização, Unidade, Tipo Produção)
- Fecha corretamente mesmo com popups abertos e reabre menu_principal com --skip-entrada
- Lista com nomes (produto/categoria/fornecedor) quando possível, sem quebrar se não existir
"""

import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Optional, Any, Callable, Sequence

import psycopg2


# ============================================================
# AJUSTE TABELAS (auto-tenta em ordem)
# ============================================================

INFO_TABLE_CANDIDATES = [
    '"Ekenox"."infoProduto"',
    '"infoProduto"',
    '"infoproduto"',
]

PRODUTOS_LOOKUP_CANDIDATES = [
    ('"Ekenox"."produtos"', '"produtoId"', '"nomeProduto"'),
    ('"produtos"', '"produtoId"', '"nomeProduto"'),
    ('"Ekenox"."produtos"', '"id"', '"nomeProduto"'),
    ('"produtos"', '"id"', '"nomeProduto"'),
]

CATEGORIA_LOOKUP_CANDIDATES = [
    ('"Ekenox"."categorias"', '"categoriaId"', '"nomeCategoria"'),
    ('"Ekenox"."categoria"', '"categoriaId"', '"nomeCategoria"'),
    ('"Ekenox"."categoria"', '"id"', '"nome"'),
    ('"Ekenox"."categorias"', '"id"', '"nome"'),
    ('"categoria"', '"categoriaId"', '"nomeCategoria"'),
    ('"categorias"', '"categoriaId"', '"nomeCategoria"'),
]

FORNECEDOR_LOOKUP_CANDIDATES = [
    ('"Ekenox"."fornecedores"', '"fornecedorId"', '"nomeFornecedor"'),
    ('"Ekenox"."fornecedor"', '"fornecedorId"', '"nomeFornecedor"'),
    ('"Ekenox"."fornecedor"', '"id"', '"nome"'),
    ('"Ekenox"."fornecedores"', '"id"', '"nome"'),
    ('"fornecedor"', '"fornecedorId"', '"nomeFornecedor"'),
    ('"fornecedores"', '"fornecedorId"', '"nomeFornecedor"'),
]

UNIDADE_OPCOES = ["UN", "PÇ", "JOGO"]

TIPO_OPCOES_VIEW = [
    "F - Simples",
    "V - Produto Pai",
    "E - Insumo (Estrutura)",
]
TIPO_TO_VALUE = {"F - Simples": "F",
                 "V - Produto Pai": "V", "E - Insumo (Estrutura)": "E"}
VALUE_TO_TIPO = {"F": "F - Simples",
                 "V": "V - Produto Pai", "E": "E - Insumo (Estrutura)"}


# ============================================================
# PATHS
# ============================================================

def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
BASE_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
os.makedirs(BASE_DIR, exist_ok=True)


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
# CONFIG BANCO
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
    return AppConfig(db_host=host, db_port=port, db_database=dbname, db_user=user, db_password=password)


# ============================================================
# DB (simples)
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

    def rollback(self) -> None:
        try:
            if self.conn:
                self.conn.rollback()
        except Exception:
            pass

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
    "menu_principal.py", "Menu_Principal.py", "MenuPrincipal.py", "menu.py",
    "menu_principal.exe", "Menu_Principal.exe", "MenuPrincipal.exe", "menu.exe",
]


def localizar_menu_principal() -> Optional[str]:
    for pasta in (APP_DIR, BASE_DIR):
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
        return
    cwd = os.path.dirname(menu_path) or APP_DIR

    try:
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
    except Exception:
        pass


# ============================================================
# HELPERS: detectar tabela existente
# ============================================================

def _table_exists(cfg: AppConfig, table_name: str) -> bool:
    conn = None
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
            if conn:
                conn.close()
        except Exception:
            pass
        return False


def detectar_tabela(cfg: AppConfig, candidates: list[str], fallback: str) -> str:
    for t in candidates:
        if _table_exists(cfg, t):
            return t
    return fallback


# ============================================================
# MODEL
# ============================================================

@dataclass
class InfoProduto:
    estoqueMinimo: Optional[int]
    estoqueMaximo: Optional[int]
    estoqueLocalizacao: Optional[str]
    unidade: Optional[str]
    pesoLiquido: Decimal
    pesoBruto: Decimal
    volumes: Optional[int]
    itensPorCaixa: Optional[int]
    gtin: Optional[str]
    tipoProducao: Optional[str]   # F/V/E no banco
    marca: Optional[str]
    precoCompra: Decimal
    largura: Decimal
    altura: Decimal
    profundidade: Decimal
    unidadeMedida: Optional[str]
    fkFornecedor: Optional[int]
    fkCategoria: Optional[int]
    fkProduto: int

    nomeProduto: Optional[str] = None
    nomeCategoria: Optional[str] = None
    nomeFornecedor: Optional[str] = None


# ============================================================
# CONVERSÕES
# ============================================================

def _clean_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _to_int_or_none(v: Any, field_name: str) -> Optional[int]:
    s = "" if v is None else str(v).strip()
    if s == "":
        return None
    try:
        return int(s)
    except ValueError:
        raise ValueError(f"{field_name} deve ser inteiro (ou vazio).")


def _to_int_required(v: Any, field_name: str) -> int:
    s = "" if v is None else str(v).strip()
    if s == "":
        raise ValueError(f"{field_name} é obrigatório.")
    try:
        return int(s)
    except ValueError:
        raise ValueError(f"{field_name} deve ser inteiro.")


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
# POPUP DE BUSCA (genérico) - com fechamento seguro (release grab)
# ============================================================

class LookupDialog(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        title: str,
        columns: Sequence[str],
        fetch_fn: Callable[[Optional[str]], list[tuple]],
        on_pick: Callable[[tuple], None],
        width: int = 740,
        height: int = 440,
    ):
        super().__init__(master)
        self.title(title)
        self.resizable(True, True)
        apply_window_icon(self)

        self._fetch_fn = fetch_fn
        self._on_pick = on_pick
        self._columns = list(columns)

        self.geometry(f"{width}x{height}")
        self.transient(master)

        # grab_set pode travar se o root fechar: a gente sempre libera no _safe_close
        try:
            self.grab_set()
        except Exception:
            pass

        self.protocol("WM_DELETE_WINDOW", self._safe_close)

        self.var = tk.StringVar()

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Buscar:").pack(side="left")
        ent = ttk.Entry(top, textvariable=self.var)
        ent.pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(top, text="Atualizar",
                   command=self.refresh).pack(side="left")

        body = ttk.Frame(self, padding=(10, 0, 10, 10))
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            body, columns=self._columns, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        for c in self._columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=220, anchor="w", stretch=True)

        btm = ttk.Frame(self, padding=10)
        btm.pack(fill="x")
        ttk.Button(btm, text="Selecionar",
                   command=self.pick).pack(side="right")
        ttk.Button(btm, text="Cancelar", command=self._safe_close).pack(
            side="right", padx=(0, 8))

        self.tree.bind("<Double-1>", lambda e: self.pick())
        self.bind("<Return>", lambda e: self.pick())
        self.bind("<Escape>", lambda e: self._safe_close())
        ent.bind("<Return>", lambda e: self.refresh())
        ent.focus_set()

        self.refresh()
        self._center()

    def _safe_close(self) -> None:
        try:
            self.grab_release()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def _center(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def refresh(self) -> None:
        term = self.var.get().strip() or None
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = self._fetch_fn(term)
        for r in rows:
            self.tree.insert("", "end", values=r)

    def pick(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning(
                "Atenção", "Selecione um item.", parent=self)
            return
        values = self.tree.item(sel[0], "values")
        self._on_pick(tuple(values))
        self._safe_close()


# ============================================================
# REPOSITORY
# ============================================================

class InfoProdutoRepo:
    """
    Correções importantes:
    - JOIN dinâmico (produtos/categorias/fornecedores): tenta candidates e cacheia o que funcionar.
    - lista retorna também nomes (quando disponível) sem quebrar caso não exista.
    """

    def __init__(self, db: Database, info_table: str) -> None:
        self.db = db
        self.info_table = info_table

        self._prod_join: Optional[tuple[str, str, str]] = None
        self._cat_join: Optional[tuple[str, str, str]] = None
        self._for_join: Optional[tuple[str, str, str]] = None

    def _auto_lookup(self, termo: Optional[str], limit: int, kind: str) -> list[tuple]:
        """
        kind: 'fornecedor' ou 'categoria' ou 'produto'
        Tenta descobrir automaticamente tabela/colunas via information_schema.
        Retorna [(id, nome), ...]
        """
        like = f"%{termo}%" if termo else None

        # padrões de nome de tabela
        if kind == "fornecedor":
            table_patterns = ["%fornec%", "%supplier%"]
            id_patterns = ["%fornecedor%id%", "%id%fornecedor%", "%id%"]
            name_patterns = ["%nome%fornec%",
                             "%razao%", "%fantasia%", "%nome%"]
        elif kind == "categoria":
            table_patterns = ["%categ%", "%category%"]
            id_patterns = ["%categoria%id%", "%id%categoria%", "%id%"]
            name_patterns = ["%nome%categ%", "%descricao%", "%nome%"]
        else:  # produto
            table_patterns = ["%produt%", "%product%"]
            id_patterns = ["%produto%id%", "%id%produto%", "%id%"]
            name_patterns = ["%nome%produt%", "%descricao%", "%nome%"]

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None

            # acha tabelas candidatas
            self.db.cursor.execute(
                """
                SELECT table_schema, table_name
                FROM information_schema.tables
                WHERE table_type='BASE TABLE'
                AND (
                        lower(table_name) LIKE %s
                    OR lower(table_name) LIKE %s
                )
                ORDER BY table_schema, table_name
                """,
                (table_patterns[0], table_patterns[1]),
            )
            tables = self.db.cursor.fetchall()

            last_err = None

            for schema, table in tables:
                # acha colunas candidatas (id e nome)
                self.db.cursor.execute(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_schema=%s AND table_name=%s
                    """,
                    (schema, table),
                )
                cols = [r[0] for r in self.db.cursor.fetchall()]
                cols_l = [c.lower() for c in cols]

                def pick_col(patterns: list[str]) -> Optional[str]:
                    for pat in patterns:
                        p = pat.replace("%", "")
                        # tenta match simples por "contém"
                        for c, cl in zip(cols, cols_l):
                            if p in cl:
                                return c
                    return None

                id_col = pick_col([p.replace("%", "") for p in id_patterns]) or (
                    cols[0] if cols else None)
                name_col = pick_col([p.replace("%", "")
                                    for p in name_patterns])

                if not id_col or not name_col:
                    continue

                full_table = f'"{schema}"."{table}"'
                id_q = f'"{id_col}"'
                name_q = f'"{name_col}"'

                sql = f"""
                    SELECT {id_q}::text AS id, COALESCE({name_q}::text,'') AS nome
                    FROM {full_table}
                    WHERE (%s IS NULL)
                    OR ({id_q}::text ILIKE %s)
                    OR (COALESCE({name_q}::text,'') ILIKE %s)
                    ORDER BY nome, id
                    LIMIT %s
                """

                try:
                    self.db.cursor.execute(sql, (termo, like, like, limit))
                    rows = self.db.cursor.fetchall()
                    return [(r[0], r[1]) for r in rows]
                except Exception as e:
                    self.db.rollback()
                    last_err = e
                    continue

            raise RuntimeError(
                f"Auto-lookup não encontrou tabela/colunas para {kind}. "
                f"Último erro: {type(last_err).__name__}: {last_err}"
            )
        finally:
            self.db.desconectar()

    def _pick_first_working_join(self, candidates: list[tuple[str, str, str]]) -> Optional[tuple[str, str, str]]:
        """
        Retorna (table, id_col, nome_col) que executa um SELECT simples com sucesso.
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            for table, id_col, nome_col in candidates:
                try:
                    self.db.cursor.execute(
                        f"SELECT {id_col}, {nome_col} FROM {table} LIMIT 1")
                    self.db.cursor.fetchone()
                    return (table, id_col, nome_col)
                except Exception:
                    continue
            return None
        finally:
            self.db.desconectar()

    def _ensure_joins_cached(self) -> None:
        if self._prod_join is None:
            self._prod_join = self._pick_first_working_join(
                PRODUTOS_LOOKUP_CANDIDATES)
        if self._cat_join is None:
            self._cat_join = self._pick_first_working_join(
                CATEGORIA_LOOKUP_CANDIDATES)
        if self._for_join is None:
            self._for_join = self._pick_first_working_join(
                FORNECEDOR_LOOKUP_CANDIDATES)

    def listar(self, termo: Optional[str] = None, limit: int = 300) -> list[InfoProduto]:
        like = f"%{termo}%" if termo else None

        self._ensure_joins_cached()

        # Campos base do InfoProduto (19 campos obrigatórios no SELECT)
        base_select = """
            i."estoqueMinimo", i."estoqueMaximo", i."estoqueLocalizacao", i."unidade",
            i."pesoLiquido", i."pesoBruto", i."volumes", i."itensPorCaixa",
            i."gtin", i."tipoProducao", i."marca", i."precoCompra",
            i."largura", i."altura", i."profundidade",
            i."unidadeMedida", i."fkFornecedor", i."fkCategoria", i."fkProduto"
        """

        joins = []
        extra_select = []

        if self._prod_join:
            pt, pid, pnome = self._prod_join
            joins.append(f'LEFT JOIN {pt} AS p ON p.{pid} = i."fkProduto"')
            extra_select.append(f"p.{pnome}::text AS nomeProduto")
        else:
            extra_select.append("NULL::text AS nomeProduto")

        if self._cat_join:
            ct, cid, cnome = self._cat_join
            joins.append(f'LEFT JOIN {ct} AS c ON c.{cid} = i."fkCategoria"')
            extra_select.append(f"c.{cnome}::text AS nomeCategoria")
        else:
            extra_select.append("NULL::text AS nomeCategoria")

        if self._for_join:
            ft, fid, fnome = self._for_join
            joins.append(f'LEFT JOIN {ft} AS f ON f.{fid} = i."fkFornecedor"')
            extra_select.append(f"f.{fnome}::text AS nomeFornecedor")
        else:
            extra_select.append("NULL::text AS nomeFornecedor")

        sql = f"""
            SELECT
                {base_select},
                {", ".join(extra_select)}
            FROM {self.info_table} AS i
            {" ".join(joins)}
            WHERE (%s IS NULL)
               OR (CAST(i."fkProduto" AS TEXT) ILIKE %s)
               OR (COALESCE(i."gtin",'') ILIKE %s)
               OR (COALESCE(i."marca",'') ILIKE %s)
            ORDER BY i."fkProduto"
            LIMIT %s
        """

        params = (termo, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()

            out: list[InfoProduto] = []
            for row in rows:
                # base 19 + 3 extras = 22
                base = row[:19]
                nome_prod = row[19] if len(row) > 19 else None
                nome_cat = row[20] if len(row) > 20 else None
                nome_for = row[21] if len(row) > 21 else None
                out.append(InfoProduto(*base, nomeProduto=nome_prod,
                           nomeCategoria=nome_cat, nomeFornecedor=nome_for))
            return out
        finally:
            self.db.desconectar()

    def exists(self, fkProduto: int) -> bool:
        sql = f'SELECT 1 FROM {self.info_table} WHERE "fkProduto" = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fkProduto,))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, i: InfoProduto) -> None:
        sql = f"""
            INSERT INTO {self.info_table}
            ("estoqueMinimo","estoqueMaximo","estoqueLocalizacao","unidade",
             "pesoLiquido","pesoBruto","volumes","itensPorCaixa",
             "gtin","tipoProducao","marca","precoCompra",
             "largura","altura","profundidade",
             "unidadeMedida","fkFornecedor","fkCategoria","fkProduto")
            VALUES
            (%s,%s,%s,%s,
             %s,%s,%s,%s,
             %s,%s,%s,%s,
             %s,%s,%s,
             %s,%s,%s,%s)
        """
        params = (
            i.estoqueMinimo, i.estoqueMaximo, i.estoqueLocalizacao, i.unidade,
            i.pesoLiquido, i.pesoBruto, i.volumes, i.itensPorCaixa,
            i.gtin, i.tipoProducao, i.marca, i.precoCompra,
            i.largura, i.altura, i.profundidade,
            i.unidadeMedida, i.fkFornecedor, i.fkCategoria, i.fkProduto
        )
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def atualizar(self, i: InfoProduto) -> None:
        sql = f"""
            UPDATE {self.info_table}
               SET "estoqueMinimo" = %s,
                   "estoqueMaximo" = %s,
                   "estoqueLocalizacao" = %s,
                   "unidade" = %s,
                   "pesoLiquido" = %s,
                   "pesoBruto" = %s,
                   "volumes" = %s,
                   "itensPorCaixa" = %s,
                   "gtin" = %s,
                   "tipoProducao" = %s,
                   "marca" = %s,
                   "precoCompra" = %s,
                   "largura" = %s,
                   "altura" = %s,
                   "profundidade" = %s,
                   "unidadeMedida" = %s,
                   "fkFornecedor" = %s,
                   "fkCategoria" = %s
             WHERE "fkProduto" = %s
        """
        params = (
            i.estoqueMinimo, i.estoqueMaximo, i.estoqueLocalizacao, i.unidade,
            i.pesoLiquido, i.pesoBruto, i.volumes, i.itensPorCaixa,
            i.gtin, i.tipoProducao, i.marca, i.precoCompra,
            i.largura, i.altura, i.profundidade,
            i.unidadeMedida, i.fkFornecedor, i.fkCategoria, i.fkProduto
        )
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def excluir(self, fkProduto: int) -> None:
        sql = f'DELETE FROM {self.info_table} WHERE "fkProduto" = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fkProduto,))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    # ---------- LOOKUPS ----------

    def lookup_localizacoes(self, termo: Optional[str], limit: int = 200) -> list[tuple]:
        like = f"%{termo}%" if termo else None
        sql = f"""
            SELECT DISTINCT COALESCE(i."estoqueLocalizacao",'') AS local
            FROM {self.info_table} i
            WHERE COALESCE(i."estoqueLocalizacao",'') <> ''
              AND (%s IS NULL OR i."estoqueLocalizacao" ILIKE %s)
            ORDER BY local
            LIMIT %s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (termo, like, limit))
            rows = self.db.cursor.fetchall()
            return [(r[0],) for r in rows]
        finally:
            self.db.desconectar()

    def _try_lookup_table(
        self,
        candidates: list[tuple[str, str, str]],
        termo: Optional[str],
        limit: int = 300,
    ) -> list[tuple]:
        like = f"%{termo}%" if termo else None

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            last_err: Optional[str] = None

            for table, id_col, nome_col in candidates:
                sql = f"""
                    SELECT {id_col}::text AS id, COALESCE({nome_col}::text,'') AS nome
                    FROM {table}
                    WHERE (%s IS NULL)
                    OR ({id_col}::text ILIKE %s)
                    OR (COALESCE({nome_col}::text,'') ILIKE %s)
                    ORDER BY nome, id
                    LIMIT %s
                """
                try:
                    self.db.cursor.execute(sql, (termo, like, like, limit))
                    rows = self.db.cursor.fetchall()
                    return [(r[0], r[1]) for r in rows]

                except Exception as e:
                    # ✅ MUITO IMPORTANTE: limpar transação abortada
                    self.db.rollback()
                    last_err = f"{type(e).__name__}: {e}"
                    continue

            raise RuntimeError(
                "Lookup não configurado (tabelas/colunas não encontradas). "
                "Ajuste candidates. Detalhe: " + (last_err or "")
            )
        finally:
            self.db.desconectar()

    def lookup_produtos(self, termo: Optional[str], limit: int = 400) -> list[tuple]:
        return self._try_lookup_table(PRODUTOS_LOOKUP_CANDIDATES, termo, limit)

    def lookup_categorias(self, termo: Optional[str], limit: int = 300) -> list[tuple]:
        try:
            return self._try_lookup_table(CATEGORIA_LOOKUP_CANDIDATES, termo, limit)
        except Exception:
            return self._auto_lookup(termo, limit, kind="categoria")

    def lookup_fornecedores(self, termo: Optional[str], limit: int = 300) -> list[tuple]:
        try:
            return self._try_lookup_table(FORNECEDOR_LOOKUP_CANDIDATES, termo, limit)
        except Exception:
            return self._auto_lookup(termo, limit, kind="fornecedor")


# ============================================================
# SERVICE
# ============================================================

class InfoProdutoService:
    def __init__(self, repo: InfoProdutoRepo) -> None:
        self.repo = repo

    def listar(self, termo: Optional[str]) -> list[InfoProduto]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo=termo)

    def salvar_from_form(self, form: dict[str, Any]) -> str:
        fkProduto = _to_int_required(form.get("fkProduto"), "fkProduto")

        tipo_view = (form.get("tipoProducao") or "").strip()
        if tipo_view in TIPO_TO_VALUE:
            tipo_db = TIPO_TO_VALUE[tipo_view]
        else:
            tipo_db = (tipo_view[:1].upper() if tipo_view else None)

        info = InfoProduto(
            estoqueMinimo=_to_int_or_none(
                form.get("estoqueMinimo"), "estoqueMinimo"),
            estoqueMaximo=_to_int_or_none(
                form.get("estoqueMaximo"), "estoqueMaximo"),
            estoqueLocalizacao=_clean_text(form.get("estoqueLocalizacao")),
            unidade=_clean_text(form.get("unidade")),
            pesoLiquido=_to_decimal(form.get("pesoLiquido"), "pesoLiquido"),
            pesoBruto=_to_decimal(form.get("pesoBruto"), "pesoBruto"),
            volumes=_to_int_or_none(form.get("volumes"), "volumes"),
            itensPorCaixa=_to_int_or_none(
                form.get("itensPorCaixa"), "itensPorCaixa"),
            gtin=_clean_text(form.get("gtin")),
            tipoProducao=_clean_text(tipo_db),
            marca=_clean_text(form.get("marca")),
            precoCompra=_to_decimal(form.get("precoCompra"), "precoCompra"),
            largura=_to_decimal(form.get("largura"), "largura"),
            altura=_to_decimal(form.get("altura"), "altura"),
            profundidade=_to_decimal(form.get("profundidade"), "profundidade"),
            unidadeMedida=_clean_text(form.get("unidadeMedida")),
            fkFornecedor=_to_int_or_none(
                form.get("fkFornecedor"), "fkFornecedor"),
            fkCategoria=_to_int_or_none(
                form.get("fkCategoria"), "fkCategoria"),
            fkProduto=fkProduto,
        )

        if info.estoqueMinimo is not None and info.estoqueMinimo < 0:
            raise ValueError("estoqueMinimo não pode ser negativo.")
        if info.estoqueMaximo is not None and info.estoqueMaximo < 0:
            raise ValueError("estoqueMaximo não pode ser negativo.")

        if self.repo.exists(fkProduto):
            self.repo.atualizar(info)
            return "atualizado"
        else:
            self.repo.inserir(info)
            return "inserido"

    def excluir(self, fkProduto: int) -> None:
        self.repo.excluir(fkProduto)

    def lookup_produtos(self, termo: Optional[str]) -> list[tuple]:
        return self.repo.lookup_produtos(termo)

    def lookup_categorias(self, termo: Optional[str]) -> list[tuple]:
        return self.repo.lookup_categorias(termo)

    def lookup_fornecedores(self, termo: Optional[str]) -> list[tuple]:
        return self.repo.lookup_fornecedores(termo)

    def lookup_localizacoes(self, termo: Optional[str]) -> list[tuple]:
        return self.repo.lookup_localizacoes(termo)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1200x720"
APP_TITLE = "Tela de InfoProduto"

CAMPOS = [
    ("fkProduto", "Produto (fkProduto)"),
    ("fkCategoria", "Categoria (fkCategoria)"),
    ("fkFornecedor", "Fornecedor (fkFornecedor)"),
    ("gtin", "GTIN"),
    ("marca", "Marca"),
    ("tipoProducao", "Tipo Produção"),
    ("estoqueMinimo", "Estoque Mínimo"),
    ("estoqueMaximo", "Estoque Máximo"),
    ("estoqueLocalizacao", "Localização"),
    ("unidade", "Unidade"),
    ("unidadeMedida", "Unidade Medida"),
    ("itensPorCaixa", "Itens por Caixa"),
    ("volumes", "Volumes"),
    ("precoCompra", "Preço Compra"),
    ("pesoLiquido", "Peso Líquido"),
    ("pesoBruto", "Peso Bruto"),
    ("largura", "Largura"),
    ("altura", "Altura"),
    ("profundidade", "Profundidade"),
]

# Agora a grade mostra nomes também (quando houver)
TREE_COLS = [
    "fkProduto", "nomeProduto",
    "gtin", "marca", "tipoProducao",
    "estoqueMinimo", "estoqueMaximo", "estoqueLocalizacao",
    "unidade", "unidadeMedida",
    "itensPorCaixa", "volumes",
    "precoCompra", "pesoLiquido", "pesoBruto",
    "largura", "altura", "profundidade",
    "fkFornecedor", "nomeFornecedor",
    "fkCategoria", "nomeCategoria",
]


class TelaInfoProduto(ttk.Frame):
    def __init__(self, master: tk.Misc, service: InfoProdutoService):
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

        ttk.Label(
            top, text="Buscar (fkProduto / GTIN / Marca):").grid(row=0, column=0, sticky="w")
        ent = ttk.Entry(top, textvariable=self.var_filtro)
        ent.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent.bind("<Return>", lambda e: self.atualizar_lista())

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

        form = ttk.LabelFrame(self, text="InfoProduto")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))

        for c in range(6):
            form.columnconfigure(c, weight=1)

        # Linha 0
        self._add_field(form, 0, 0, "fkProduto", width=14,
                        with_button=self._buscar_produto)
        self._add_field(form, 0, 2, "fkCategoria", width=14,
                        with_button=self._buscar_categoria)
        self._add_field(form, 0, 4, "fkFornecedor", width=14,
                        with_button=self._buscar_fornecedor)

        # Linha 1
        self._add_field(form, 1, 0, "gtin", width=18)
        self._add_field(form, 1, 2, "marca", width=18)
        self._add_field(form, 1, 4, "tipoProducao", widget="combo", combo_values=TIPO_OPCOES_VIEW,
                        width=18, with_button=self._buscar_tipo)

        # Linha 2
        self._add_field(form, 2, 0, "estoqueMinimo", width=14)
        self._add_field(form, 2, 2, "estoqueMaximo", width=14)
        self._add_field(form, 2, 4, "estoqueLocalizacao",
                        width=22, with_button=self._buscar_localizacao)

        # Linha 3
        self._add_field(form, 3, 0, "unidade", widget="combo", combo_values=UNIDADE_OPCOES,
                        width=14, with_button=self._buscar_unidade)
        self._add_field(form, 3, 2, "unidadeMedida", width=14)
        self._add_field(form, 3, 4, "itensPorCaixa", width=14)

        # Linha 4
        self._add_field(form, 4, 0, "volumes", width=14)
        self._add_field(form, 4, 2, "precoCompra", width=14)
        self._add_field(form, 4, 4, "pesoLiquido", width=14)

        # Linha 5
        self._add_field(form, 5, 0, "pesoBruto", width=14)
        self._add_field(form, 5, 2, "largura", width=14)
        self._add_field(form, 5, 4, "altura", width=14)

        # Linha 6
        self._add_field(form, 6, 0, "profundidade", width=14, colspan=2)

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

        headings = {
            "fkProduto": "fkProduto",
            "nomeProduto": "Nome Produto",
            "gtin": "GTIN",
            "marca": "Marca",
            "tipoProducao": "Tipo",
            "estoqueMinimo": "Min",
            "estoqueMaximo": "Max",
            "estoqueLocalizacao": "Local",
            "unidade": "Unidade",
            "unidadeMedida": "Unid Med",
            "itensPorCaixa": "Itens/Cx",
            "volumes": "Vol",
            "precoCompra": "Preço",
            "pesoLiquido": "Peso L",
            "pesoBruto": "Peso B",
            "largura": "Larg",
            "altura": "Alt",
            "profundidade": "Prof",
            "fkFornecedor": "Fk Forn",
            "nomeFornecedor": "Fornecedor",
            "fkCategoria": "Fk Cat",
            "nomeCategoria": "Categoria",
        }
        for col in TREE_COLS:
            self.tree.heading(col, text=headings.get(col, col))

        self.tree.column("fkProduto", width=95, anchor="e", stretch=False)
        self.tree.column("nomeProduto", width=300, anchor="w", stretch=True)
        self.tree.column("gtin", width=150, anchor="w", stretch=False)
        self.tree.column("marca", width=140, anchor="w", stretch=False)
        self.tree.column("tipoProducao", width=120, anchor="w", stretch=False)

        self.tree.column("fkFornecedor", width=95, anchor="e", stretch=False)
        self.tree.column("nomeFornecedor", width=220, anchor="w", stretch=True)
        self.tree.column("fkCategoria", width=95, anchor="e", stretch=False)
        self.tree.column("nomeCategoria", width=220, anchor="w", stretch=True)

        for col in TREE_COLS:
            if col in ("fkProduto", "nomeProduto", "gtin", "marca", "tipoProducao",
                       "fkFornecedor", "nomeFornecedor", "fkCategoria", "nomeCategoria"):
                continue
            self.tree.column(col, width=100, anchor="e", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def _add_field(
        self,
        parent: ttk.Frame,
        row: int,
        col: int,
        key: str,
        readonly: bool = False,
        colspan: int = 2,
        width: int | None = None,
        widget: str = "entry",
        combo_values: Optional[list[str]] = None,
        with_button: Optional[Callable[[], None]] = None,
    ) -> None:
        label = dict(CAMPOS)[key]
        ttk.Label(parent, text=f"{label}:").grid(
            row=row, column=col, sticky="w", padx=(10, 6), pady=6)

        wrap = ttk.Frame(parent)
        wrap.grid(row=row, column=col + 1, sticky="ew",
                  padx=(0, 10), pady=6, columnspan=colspan - 1)
        wrap.columnconfigure(0, weight=1)

        if widget == "combo":
            cb = ttk.Combobox(
                wrap, textvariable=self.vars[key], state="readonly" if readonly else "normal")
            if combo_values:
                cb["values"] = combo_values
            if width is not None:
                cb.configure(width=width)
            cb.grid(row=0, column=0, sticky="ew")
            field_widget = cb
        else:
            ent = ttk.Entry(
                wrap, textvariable=self.vars[key], state="readonly" if readonly else "normal")
            if width is not None:
                ent.configure(width=width)
            ent.grid(row=0, column=0, sticky="ew")
            field_widget = ent

        if with_button is not None:
            ttk.Button(wrap, text="...", width=3, command=with_button).grid(
                row=0, column=1, padx=(6, 0))

        try:
            field_widget.bind("<Return>", lambda e: parent.focus_set())
        except Exception:
            pass

    # -------- botões de busca --------

    def _buscar_produto(self) -> None:
        def fetch(term: Optional[str]) -> list[tuple]:
            return self.service.lookup_produtos(term)

        def pick(values: tuple) -> None:
            self.vars["fkProduto"].set(str(values[0]))

        LookupDialog(self, "Selecionar Produto", [
                     "Identificador", "Nome"], fetch, pick, width=860, height=520)

    def _buscar_categoria(self) -> None:
        def fetch(term: Optional[str]) -> list[tuple]:
            return self.service.lookup_categorias(term)

        def pick(values: tuple) -> None:
            self.vars["fkCategoria"].set(str(values[0]))

        LookupDialog(self, "Selecionar Categoria", [
                     "Identificador", "Nome"], fetch, pick, width=820, height=500)

    def _buscar_fornecedor(self) -> None:
        def fetch(term: Optional[str]) -> list[tuple]:
            return self.service.lookup_fornecedores(term)

        def pick(values: tuple) -> None:
            self.vars["fkFornecedor"].set(str(values[0]))

        LookupDialog(self, "Selecionar Fornecedor", [
                     "Identificador", "Nome"], fetch, pick, width=820, height=500)

    def _buscar_localizacao(self) -> None:
        def fetch(term: Optional[str]) -> list[tuple]:
            return self.service.lookup_localizacoes(term)

        def pick(values: tuple) -> None:
            self.vars["estoqueLocalizacao"].set(str(values[0]))

        LookupDialog(self, "Selecionar Localização", [
                     "Localização"], fetch, pick, width=600, height=450)

    def _buscar_unidade(self) -> None:
        def fetch(_term: Optional[str]) -> list[tuple]:
            return [(u,) for u in UNIDADE_OPCOES]

        def pick(values: tuple) -> None:
            self.vars["unidade"].set(str(values[0]))

        LookupDialog(self, "Selecionar Unidade", [
                     "Unidade"], fetch, pick, width=460, height=360)

    def _buscar_tipo(self) -> None:
        def fetch(_term: Optional[str]) -> list[tuple]:
            return [(t,) for t in TIPO_OPCOES_VIEW]

        def pick(values: tuple) -> None:
            self.vars["tipoProducao"].set(str(values[0]))

        LookupDialog(self, "Selecionar Tipo Produção", [
                     "Tipo"], fetch, pick, width=560, height=380)

    # -------- ações --------

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            itens = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar InfoProduto:\n{e}")
            return

        for it in itens:
            tipo_view = VALUE_TO_TIPO.get(
                (it.tipoProducao or "").strip().upper(), it.tipoProducao or "")
            values = [
                it.fkProduto,
                it.nomeProduto or "",
                it.gtin or "",
                it.marca or "",
                tipo_view,
                "" if it.estoqueMinimo is None else it.estoqueMinimo,
                "" if it.estoqueMaximo is None else it.estoqueMaximo,
                it.estoqueLocalizacao or "",
                it.unidade or "",
                it.unidadeMedida or "",
                "" if it.itensPorCaixa is None else it.itensPorCaixa,
                "" if it.volumes is None else it.volumes,
                str(it.precoCompra),
                str(it.pesoLiquido),
                str(it.pesoBruto),
                str(it.largura),
                str(it.altura),
                str(it.profundidade),
                "" if it.fkFornecedor is None else it.fkFornecedor,
                it.nomeFornecedor or "",
                "" if it.fkCategoria is None else it.fkCategoria,
                it.nomeCategoria or "",
            ]
            self.tree.insert("", "end", values=values)

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")

        # índices conforme TREE_COLS
        mapping = {
            "fkProduto": vals[0],
            "gtin": vals[2],
            "marca": vals[3],
            "tipoProducao": vals[4],
            "estoqueMinimo": vals[5],
            "estoqueMaximo": vals[6],
            "estoqueLocalizacao": vals[7],
            "unidade": vals[8],
            "unidadeMedida": vals[9],
            "itensPorCaixa": vals[10],
            "volumes": vals[11],
            "precoCompra": vals[12],
            "pesoLiquido": vals[13],
            "pesoBruto": vals[14],
            "largura": vals[15],
            "altura": vals[16],
            "profundidade": vals[17],
            "fkFornecedor": vals[18],
            "fkCategoria": vals[20],
        }

        for k, _ in CAMPOS:
            v = mapping.get(k, "")
            self.vars[k].set("" if v in (None, "None") else str(v))

    def novo(self) -> None:
        self.limpar_form()
        for k in ("precoCompra", "pesoLiquido", "pesoBruto", "largura", "altura", "profundidade"):
            self.vars[k].set("0")

    def limpar_form(self) -> None:
        for k in self.vars:
            self.vars[k].set("")
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        form = {k: self.vars[k].get() for k, _ in CAMPOS}
        try:
            status = self.service.salvar_from_form(form)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return
        messagebox.showinfo("OK", f"InfoProduto {status} com sucesso.")
        self.atualizar_lista()

    def excluir(self) -> None:
        fk_str = self.vars["fkProduto"].get().strip()
        if not fk_str:
            messagebox.showwarning(
                "Atenção", "Informe/Selecione um fkProduto para excluir.")
            return
        if not messagebox.askyesno("Confirmar", f"Excluir InfoProduto do Produto {fk_str}?"):
            return
        try:
            self.service.excluir(int(fk_str))
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return
        messagebox.showinfo("OK", "InfoProduto excluído.")
        self.limpar_form()
        self.atualizar_lista()


# ============================================================
# STARTUP
# ============================================================

def test_connection_or_die(cfg: AppConfig) -> None:
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


def main() -> None:
    cfg = env_override(load_config())

    info_table = detectar_tabela(
        cfg, INFO_TABLE_CANDIDATES, '"Ekenox"."infoProduto"')

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
            f"Host: {cfg.db_host}\nPorta: {cfg.db_port}\nBanco: {cfg.db_database}\nUsuário: {cfg.db_user}\n\n"
            f"Erro:\n{e}"
        )
        try:
            root.destroy()
        except Exception:
            pass
        return

    db = Database(cfg)
    repo = InfoProdutoRepo(db, info_table=info_table)
    service = InfoProdutoService(repo)

    tela = TelaInfoProduto(root, service)
    tela.pack(fill="both", expand=True)

    # ============================================================
    # FECHAMENTO: fecha popups (release grab), reabre menu principal e fecha root
    # ============================================================
    closing = {"done": False}

    def force_close():
        if closing["done"]:
            return
        closing["done"] = True

        # fecha Toplevels e libera grab
        try:
            for w in list(root.winfo_children()):
                if isinstance(w, tk.Toplevel):
                    try:
                        w.grab_release()
                    except Exception:
                        pass
                    try:
                        w.destroy()
                    except Exception:
                        pass
        except Exception:
            pass

        # abre menu antes de destruir a janela
        try:
            abrir_menu_principal_skip_entrada()
        except Exception:
            pass

        # fecha root normalmente (NÃO use os._exit)
        try:
            root.destroy()
        except Exception:
            pass

    # IMPORTANTE: protocol definido UMA VEZ, fora do callback
    root.protocol("WM_DELETE_WINDOW", force_close)
    root.mainloop()


if __name__ == "__main__":
    main()
