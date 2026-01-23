from __future__ import annotations

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


def log_estoque(msg: str) -> None:
    _log_write("tela_estoque.log", msg)


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
            win._icon_img = img
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
        log_estoque(
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
        log_estoque(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_estoque(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# AUTO-DETECT TABELAS (schema ou não)
# ============================================================

ESTOQUE_TABLES = [
    '"Ekenox"."estoque"',
    '"estoque"',
]

PRODUTOS_TABLES = [
    '"Ekenox"."produtos"',
    '"produtos"',
]


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


def detectar_tabela(cfg: AppConfig, candidates: List[str], fallback: str) -> str:
    for t in candidates:
        ok = _table_exists(cfg, t)
        log_estoque(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


# ============================================================
# MODEL
# ============================================================

@dataclass
class Estoque:
    fkProduto: int
    nomeProduto: str
    saldoFisico: Decimal
    saldoVirtual: Decimal


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


def _qident(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def _is_numeric_pg_type(typname: str) -> bool:
    t = (typname or "").lower()
    return t in {"int2", "int4", "int8", "numeric", "float4", "float8"}


# ============================================================
# REPOSITORY
# ============================================================

class EstoqueRepo:
    def __init__(self, db: Database, estoque_table: str, produtos_table: str) -> None:
        self.db = db
        self.estoque_table = estoque_table
        self.produtos_table = produtos_table

        self.pk_col = "fkProduto"
        self.produto_id_col = "produtoId"

        # Detecta tipo real das colunas (evita text = bigint / insert errado)
        self.fk_is_numeric = self._col_is_numeric(
            self.estoque_table, self.pk_col)
        self.produtoid_is_numeric = self._col_is_numeric(
            self.produtos_table, self.produto_id_col)

    def _col_typname(self, table_reg: str, col: str) -> str:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT t.typname
                  FROM pg_attribute a
                  JOIN pg_type t ON t.oid = a.atttypid
                 WHERE a.attrelid = %s::regclass
                   AND a.attname = %s
                   AND a.attnum > 0
                   AND NOT a.attisdropped
                """,
                (table_reg, col),
            )
            r = self.db.cursor.fetchone()
            return str(r[0]) if r and r[0] else ""
        finally:
            self.db.desconectar()

    def _col_is_numeric(self, table_reg: str, col: str) -> bool:
        try:
            return _is_numeric_pg_type(self._col_typname(table_reg, col))
        except Exception:
            # se falhar a detecção, assume TEXT (mais seguro para evitar cast implícito)
            return False

    def _fk_param(self, fk: int) -> Any:
        return fk if self.fk_is_numeric else str(fk)

    def _produtoid_param(self, pid: int) -> Any:
        return pid if self.produtoid_is_numeric else str(pid)

    # ---------- NEXTVAL / SEQUENCE ----------

    def _get_schema_and_table_real(self) -> Tuple[str, str]:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT ns.nspname, c.relname
                  FROM pg_class c
                  JOIN pg_namespace ns ON ns.oid = c.relnamespace
                 WHERE c.oid = %s::regclass
                """,
                (self.estoque_table,),
            )
            r = self.db.cursor.fetchone()
            if not r:
                raise RuntimeError(
                    f"Não consegui resolver a tabela via regclass: {self.estoque_table}")
            return str(r[0]), str(r[1])
        finally:
            self.db.desconectar()

    def _find_sequence_via_pg_get_serial_sequence(self) -> Optional[str]:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                "SELECT pg_get_serial_sequence(%s, %s)",
                (self.estoque_table, self.pk_col),
            )
            r = self.db.cursor.fetchone()
            if r and r[0]:
                return str(r[0])
            return None
        finally:
            self.db.desconectar()

    def _find_sequence_owned_by_column(self) -> Optional[str]:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT quote_ident(ns_seq.nspname) || '.' || quote_ident(seq.relname) AS seq_qual
                  FROM pg_class seq
                  JOIN pg_namespace ns_seq ON ns_seq.oid = seq.relnamespace
                  JOIN pg_depend dep ON dep.objid = seq.oid
                  JOIN pg_class tbl ON tbl.oid = dep.refobjid
                  JOIN pg_attribute att
                    ON att.attrelid = tbl.oid
                   AND att.attnum = dep.refobjsubid
                 WHERE seq.relkind = 'S'
                   AND tbl.oid = %s::regclass
                   AND att.attname = %s
                   AND dep.deptype IN ('a','n')
                 LIMIT 1
                """,
                (self.estoque_table, self.pk_col),
            )
            r = self.db.cursor.fetchone()
            if r and r[0]:
                return str(r[0])
            return None
        finally:
            self.db.desconectar()

    def _ensure_sequence_and_default(self) -> str:
        seq = self._find_sequence_via_pg_get_serial_sequence()
        if seq:
            return seq

        seq2 = self._find_sequence_owned_by_column()
        if seq2:
            return seq2

        schema, table = self._get_schema_and_table_real()
        seq_name = f"{table}_{self.pk_col}_seq"

        seq_qual = f"{_qident(schema)}.{_qident(seq_name)}"
        table_qual = f"{_qident(schema)}.{_qident(table)}"
        col_qual = _qident(self.pk_col)

        # max para setval (sempre numérico)
        if self.fk_is_numeric:
            max_expr_sql = f"COALESCE(MAX({col_qual})::bigint, 0)"
        else:
            max_expr_sql = (
                f"COALESCE(MAX((NULLIF(regexp_replace({col_qual}::text, '[^0-9]', '', 'g'), '') )::bigint), 0)"
            )

        # default: se fkProduto for TEXT, default vira nextval(...)::text
        default_expr_sql = "nextval(%s::regclass)" if self.fk_is_numeric else "nextval(%s::regclass)::text"

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None

            self.db.cursor.execute(f"CREATE SEQUENCE IF NOT EXISTS {seq_qual}")

            try:
                self.db.cursor.execute(
                    f"ALTER SEQUENCE {seq_qual} OWNED BY {table_qual}.{col_qual}")
            except Exception:
                self.db.rollback()

            self.db.cursor.execute(
                f"ALTER TABLE {table_qual} ALTER COLUMN {col_qual} SET DEFAULT {default_expr_sql}",
                (seq_qual,),
            )

            self.db.cursor.execute(
                f"""
                SELECT setval(
                    %s::regclass,
                    GREATEST((SELECT {max_expr_sql} FROM {table_qual}) + 1, 1),
                    false
                )
                """,
                (seq_qual,),
            )

            self.db.commit()
            return seq_qual
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def proximo_fk_nextval(self) -> int:
        seq_qual = self._ensure_sequence_and_default()
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute("SELECT nextval(%s::regclass)", (seq_qual,))
            r = self.db.cursor.fetchone()
            return int(r[0])
        finally:
            self.db.desconectar()

    # ---------- PRODUTOS ----------

    def produto_existe(self, produto_id: int) -> bool:
        # ✅ comparação por TEXT evita text=bigint
        sql = f'SELECT 1 FROM {self.produtos_table} WHERE CAST("produtoId" AS TEXT) = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(produto_id),))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir_produto(self, produto_id: int, nome_produto: str) -> None:
        # ✅ insere com tipo correto (se produtoId for TEXT, manda str; se for numérico, manda int)
        sql = f'INSERT INTO {self.produtos_table} ("produtoId","nomeProduto") VALUES (%s,%s)'
        pid_param = self._produtoid_param(produto_id)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (pid_param, nome_produto))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def nome_produto_por_fk(self, fk: int) -> str:
        # ✅ comparação por TEXT evita text=bigint
        sql = f"""
            SELECT COALESCE(p."nomeProduto",'')
            FROM {self.produtos_table} AS p
            WHERE CAST(p."produtoId" AS TEXT) = %s
            LIMIT 1
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(fk),))
            r = self.db.cursor.fetchone()
            return str(r[0] or "") if r else ""
        finally:
            self.db.desconectar()

    # ---------- CRUD ESTOQUE ----------

    def listar(self, termo: Optional[str] = None, limit: int = 1200) -> List[Estoque]:
        like = f"%{termo}%" if termo else None
        # ✅ join por TEXT evita mismatch se um lado for text e outro bigint
        sql = f"""
            SELECT
                e."fkProduto",
                COALESCE(p."nomeProduto",'') AS "nomeProduto",
                e."saldoFisico",
                e."saldoVirtual"
            FROM {self.estoque_table} AS e
            LEFT JOIN {self.produtos_table} AS p
                   ON CAST(p."produtoId" AS TEXT) = CAST(e."fkProduto" AS TEXT)
            WHERE (%s IS NULL)
               OR (CAST(e."fkProduto" AS TEXT) ILIKE %s)
               OR (COALESCE(p."nomeProduto",'') ILIKE %s)
            ORDER BY CAST(e."fkProduto" AS TEXT)
            LIMIT %s
        """
        params = (termo, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            out: List[Estoque] = []
            for r in rows:
                out.append(Estoque(
                    fkProduto=int(str(r[0] or "0")),
                    nomeProduto=str(r[1] or ""),
                    saldoFisico=Decimal(
                        str(r[2] if r[2] is not None else "0")),
                    saldoVirtual=Decimal(
                        str(r[3] if r[3] is not None else "0")),
                ))
            return out
        finally:
            self.db.desconectar()

    def existe_fk(self, fk: int) -> bool:
        # ✅ comparação por TEXT evita mismatch
        sql = f'SELECT 1 FROM {self.estoque_table} WHERE CAST("fkProduto" AS TEXT) = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(fk),))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, fk: int, fisico: Decimal, virtual: Decimal) -> int:
        sql = f"""
            INSERT INTO {self.estoque_table} ("fkProduto","saldoFisico","saldoVirtual")
            VALUES (%s,%s,%s)
            RETURNING "fkProduto"
        """
        fk_param = self._fk_param(fk)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fk_param, fisico, virtual))
            new_fk = self.db.cursor.fetchone()[0]
            self.db.commit()
            return int(str(new_fk))
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def atualizar(self, fk: int, fisico: Decimal, virtual: Decimal) -> None:
        # atualiza por TEXT no WHERE, mas grava com tipo correto em fk_param
        sql = f"""
            UPDATE {self.estoque_table}
               SET "saldoFisico" = %s,
                   "saldoVirtual" = %s
             WHERE CAST("fkProduto" AS TEXT) = %s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fisico, virtual, str(fk)))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def excluir(self, fk: int) -> None:
        sql = f'DELETE FROM {self.estoque_table} WHERE CAST("fkProduto" AS TEXT) = %s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(fk),))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class EstoqueService:
    def __init__(self, repo: EstoqueRepo) -> None:
        self.repo = repo

    def proximo_fk_nextval(self) -> int:
        return self.repo.proximo_fk_nextval()

    def listar(self, termo: Optional[str]) -> List[Estoque]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def preencher_nome_produto(self, fk_txt: str) -> Tuple[str, bool]:
        fk_txt = (fk_txt or "").strip()
        if not fk_txt:
            return ("", False)
        try:
            fk = int(fk_txt)
        except ValueError:
            return ("", False)

        nome = self.repo.nome_produto_por_fk(fk)
        return (nome, bool(nome.strip()))

    def salvar(self, fk_txt: str, nome_txt: str, fisico_txt: str, virtual_txt: str) -> tuple[str, int]:
        fk_txt = (fk_txt or "").strip()
        if not fk_txt:
            fk_prod = self.repo.proximo_fk_nextval()
        else:
            try:
                fk_prod = int(fk_txt)
            except ValueError:
                raise ValueError("fkProduto deve ser número inteiro.")

        nome_txt = (nome_txt or "").strip()
        fisico = _to_decimal(fisico_txt, "saldoFisico")
        virtual = _to_decimal(virtual_txt, "saldoVirtual")

        # Se digitou nome e ainda não existe produto, cria na tabela produtos
        if nome_txt and not self.repo.produto_existe(fk_prod):
            self.repo.inserir_produto(fk_prod, nome_txt)

        if self.repo.existe_fk(fk_prod):
            self.repo.atualizar(fk_prod, fisico, virtual)
            return ("atualizado", fk_prod)

        new_fk = self.repo.inserir(fk_prod, fisico, virtual)
        return ("inserido", new_fk)

    def excluir(self, fk: int) -> None:
        self.repo.excluir(fk)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1100x650"
APP_TITLE = "Tela de Estoque"

TREE_COLS = ["fkProduto", "nomeProduto", "saldoFisico", "saldoVirtual"]


class TelaEstoque(ttk.Frame):
    def __init__(self, master: tk.Misc, service: EstoqueService):
        super().__init__(master)
        self.service = service

        self.var_filtro = tk.StringVar()

        self.var_fk = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_fisico = tk.StringVar(value="0")
        self.var_virtual = tk.StringVar(value="0")

        self.ent_nome: Optional[ttk.Entry] = None

        self._build_ui()
        self.atualizar_lista()

    def _set_nome_editavel(self, editavel: bool) -> None:
        if not self.ent_nome:
            return
        self.ent_nome.configure(state=("normal" if editavel else "readonly"))
        if editavel:
            self.ent_nome.focus_set()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(
            top, text="Buscar (fkProduto / Nome Produto):").grid(row=0, column=0, sticky="w")
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

        form = ttk.LabelFrame(self, text="Estoque")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(10):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="fkProduto:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_fk = ttk.Entry(form, textvariable=self.var_fk, width=14)
        ent_fk.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=6)
        ent_fk.bind("<Return>", lambda e: self._preencher_nome())

        ttk.Button(form, text="Preencher Nome", command=self._preencher_nome).grid(
            row=0, column=2, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Nome Produto:").grid(
            row=0, column=3, sticky="w", padx=(10, 6), pady=6)

        self.ent_nome = ttk.Entry(
            form, textvariable=self.var_nome, state="readonly")
        self.ent_nome.grid(row=0, column=4, sticky="ew",
                           padx=(0, 10), pady=6, columnspan=6)

        ttk.Label(form, text="Saldo Físico:").grid(
            row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_f = ttk.Entry(form, textvariable=self.var_fisico, width=18)
        ent_f.grid(row=1, column=1, sticky="w", padx=(0, 10), pady=6)

        ttk.Label(form, text="Saldo Virtual:").grid(
            row=1, column=3, sticky="w", padx=(10, 6), pady=6)
        ent_v = ttk.Entry(form, textvariable=self.var_virtual, width=18)
        ent_v.grid(row=1, column=4, sticky="w", padx=(0, 10), pady=6)

        ent_f.bind("<Return>", lambda e: self.salvar())
        ent_v.bind("<Return>", lambda e: self.salvar())

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

        self.tree.heading("fkProduto", text="fkProduto")
        self.tree.heading("nomeProduto", text="Nome Produto")
        self.tree.heading("saldoFisico", text="Saldo Físico")
        self.tree.heading("saldoVirtual", text="Saldo Virtual")

        self.tree.column("fkProduto", width=130, anchor="e", stretch=False)
        self.tree.column("nomeProduto", width=520, anchor="w", stretch=True)
        self.tree.column("saldoFisico", width=160, anchor="e", stretch=False)
        self.tree.column("saldoVirtual", width=160, anchor="e", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def _preencher_nome(self) -> None:
        nome, encontrado = self.service.preencher_nome_produto(
            self.var_fk.get())
        if encontrado:
            self.var_nome.set(nome)
            self._set_nome_editavel(False)
        else:
            # Não achou -> libera para digitar
            self.var_nome.set("")
            self._set_nome_editavel(True)

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar estoque:\n{e}")
            return

        for e in rows:
            self.tree.insert("", "end", values=(
                e.fkProduto,
                e.nomeProduto,
                str(e.saldoFisico),
                str(e.saldoVirtual),
            ))

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        fk, nome, fis, vir = self.tree.item(sel[0], "values")

        self.var_fk.set(str(fk))
        self.var_nome.set(str(nome or ""))
        self.var_fisico.set(str(fis or "0"))
        self.var_virtual.set(str(vir or "0"))

        self._set_nome_editavel(False)

    def novo(self) -> None:
        self.limpar_form()
        self.var_fisico.set("0")
        self.var_virtual.set("0")
        try:
            next_fk = self.service.proximo_fk_nextval()
            self.var_fk.set(str(next_fk))
            # para novo normalmente não existe: libera para digitar
            self._preencher_nome()
        except Exception as e:
            messagebox.showerror(
                "Erro",
                "Falha ao obter próximo fkProduto (nextval).\n"
                "Obs: o programa tenta criar/vincular sequence automaticamente.\n"
                "Se falhar, pode ser falta de permissão no banco.\n\n"
                f"Detalhe:\n{e}"
            )

    def limpar_form(self) -> None:
        self.var_fk.set("")
        self.var_nome.set("")
        self.var_fisico.set("0")
        self.var_virtual.set("0")
        self.tree.selection_remove(self.tree.selection())
        self._set_nome_editavel(False)

    def salvar(self) -> None:
        try:
            status, fk = self.service.salvar(
                self.var_fk.get(),
                self.var_nome.get(),
                self.var_fisico.get(),
                self.var_virtual.get()
            )
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        self.var_fk.set(str(fk))
        self._preencher_nome()
        messagebox.showinfo("OK", f"Estoque {status} com sucesso.\nFK: {fk}")
        self.atualizar_lista()

    def excluir(self) -> None:
        fk_txt = (self.var_fk.get() or "").strip()
        if not fk_txt:
            messagebox.showwarning(
                "Atenção", "Selecione um registro para excluir.")
            return
        try:
            fk = int(fk_txt)
        except ValueError:
            messagebox.showerror("Validação", "fkProduto inválido.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Estoque fkProduto {fk}?"):
            return

        try:
            self.service.excluir(fk)
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
        cur.execute("SELECT 1")
        cur.fetchone()
        cur.close()
        conn.close()
    except Exception as e:
        try:
            if conn:
                conn.close()
        except Exception:
            pass
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

    estoque_table = detectar_tabela(cfg, ESTOQUE_TABLES, '"estoque"')
    produtos_table = detectar_tabela(cfg, PRODUTOS_TABLES, '"produtos"')

    db = Database(cfg)
    repo = EstoqueRepo(db, estoque_table=estoque_table,
                       produtos_table=produtos_table)
    service = EstoqueService(repo)

    tela = TelaEstoque(root, service)
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
        log_estoque(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_estoque:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
