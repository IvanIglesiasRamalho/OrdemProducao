from __future__ import annotations

import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass
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


def log_deposito(msg: str) -> None:
    _log_write("tela_deposito.log", msg)


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
        log_deposito(
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
        log_deposito(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_deposito(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# AUTO-DETECT TABELAS
# ============================================================

DEPOSITO_TABLES = [
    '"Ekenox"."deposito"',
    '"deposito"',
    '"Ekenox"."depositos"',
    '"depositos"',
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
        log_deposito(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


# ============================================================
# HELPERS SQL
# ============================================================

def _qident(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def _is_numeric_pg_type(typname: str) -> bool:
    t = (typname or "").lower()
    return t in {"int2", "int4", "int8", "numeric", "float4", "float8"}


def _to_bool(val: Any) -> bool:
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    s = str(val).strip().upper()
    return s in {"1", "T", "TRUE", "S", "SIM", "A", "ATIVO", "Y", "YES"}


# ============================================================
# MODEL
# ============================================================

@dataclass
class Deposito:
    codigo: str
    descricao: str
    ativo: bool
    padrao: bool
    desconsidera_saldo: bool


# ============================================================
# REPOSITORY
# ============================================================

class DepositoRepo:
    def __init__(self, db: Database, deposito_table: str) -> None:
        self.db = db
        self.deposito_table = deposito_table

        # tenta achar PK real; se não achar, assume "codigo"
        self.pk_col = self._find_primary_key_column() or "codigo"
        self.pk_typname = self._col_typname(self.deposito_table, self.pk_col)
        self.pk_is_numeric = _is_numeric_pg_type(self.pk_typname)

        # colunas prováveis
        self.col_descricao = self._find_existing_column(
            ["descricao", "descricaoDeposito", "desc", "Descrição"])
        self.col_ativo = self._find_existing_column(
            ["ativo", "situacao", "status", "Ativo"])
        self.col_padrao = self._find_existing_column(
            ["padrao", "default", "Padrao"])
        self.col_desconsidera = self._find_existing_column(
            ["desconsiderarsaldo", "desconsideraSaldo", "ignorasaldo"])

        # fallback mínimo
        self.col_descricao = self.col_descricao or "descricao"
        self.col_ativo = self.col_ativo or "ativo"
        self.col_padrao = self.col_padrao or "padrao"
        self.col_desconsidera = self.col_desconsidera or "desconsideraSaldo"

        # tipos (pra gravar sem dar mismatch)
        self.ativo_is_bool = _is_numeric_pg_type("bool") and False  # dummy
        self.ativo_is_bool = (self._col_typname(
            self.deposito_table, self.col_ativo).lower() == "bool")
        self.padrao_is_bool = (self._col_typname(
            self.deposito_table, self.col_padrao).lower() == "bool")
        self.desconsidera_is_bool = (self._col_typname(
            self.deposito_table, self.col_desconsidera).lower() == "bool")

    def _list_columns(self) -> List[str]:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT a.attname
                  FROM pg_attribute a
                 WHERE a.attrelid = %s::regclass
                   AND a.attnum > 0
                   AND NOT a.attisdropped
                 ORDER BY a.attnum
                """,
                (self.deposito_table,),
            )
            return [str(r[0]) for r in self.db.cursor.fetchall()]
        finally:
            self.db.desconectar()

    def _find_existing_column(self, candidates: List[str]) -> Optional[str]:
        try:
            cols = {c.lower(): c for c in self._list_columns()}
        except Exception:
            cols = {}
        for cand in candidates:
            key = str(cand).lower()
            if key in cols:
                return cols[key]
        return None

    def _find_primary_key_column(self) -> Optional[str]:
        if not self.db.conectar():
            return None
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT a.attname
                  FROM pg_index i
                  JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)
                 WHERE i.indrelid = %s::regclass
                   AND i.indisprimary
                 ORDER BY a.attnum
                 LIMIT 1
                """,
                (self.deposito_table,),
            )
            r = self.db.cursor.fetchone()
            return str(r[0]) if r and r[0] else None
        except Exception:
            return None
        finally:
            self.db.desconectar()

    def _col_typname(self, table_reg: str, col: str) -> str:
        if not self.db.conectar():
            return ""
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
        except Exception:
            return ""
        finally:
            self.db.desconectar()

    def _pk_param(self, codigo_int: int) -> Any:
        return codigo_int if self.pk_is_numeric else str(codigo_int)

    # ----------------- GERAR PRÓXIMO CÓDIGO -----------------

    def _find_sequence(self) -> Optional[str]:
        # 1) pg_get_serial_sequence
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                "SELECT pg_get_serial_sequence(%s, %s)", (self.deposito_table, self.pk_col))
            r = self.db.cursor.fetchone()
            if r and r[0]:
                return str(r[0])
        finally:
            self.db.desconectar()

        # 2) owned by
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
                (self.deposito_table, self.pk_col),
            )
            r = self.db.cursor.fetchone()
            return str(r[0]) if r and r[0] else None
        finally:
            self.db.desconectar()

    def _ensure_sequence(self) -> Optional[str]:
        seq = self._find_sequence()
        if seq:
            return seq

        # tenta criar (se tiver permissão); se falhar, retorna None e vamos de MAX+1
        conn_ok = self.db.conectar()
        if not conn_ok:
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
                (self.deposito_table,),
            )
            r = self.db.cursor.fetchone()
            if not r:
                return None

            schema, table = str(r[0]), str(r[1])
            seq_name = f"{table}_{self.pk_col}_seq"

            seq_qual = f"{_qident(schema)}.{_qident(seq_name)}"
            table_qual = f"{_qident(schema)}.{_qident(table)}"
            col_qual = _qident(self.pk_col)

            self.db.cursor.execute(f"CREATE SEQUENCE IF NOT EXISTS {seq_qual}")

            try:
                self.db.cursor.execute(
                    f"ALTER SEQUENCE {seq_qual} OWNED BY {table_qual}.{col_qual}")
            except Exception:
                self.db.rollback()

            # default com cast se PK for TEXT
            default_expr = "nextval(%s::regclass)" if self.pk_is_numeric else "nextval(%s::regclass)::text"
            self.db.cursor.execute(
                f"ALTER TABLE {table_qual} ALTER COLUMN {col_qual} SET DEFAULT {default_expr}",
                (seq_qual,),
            )

            # setval com base no maior existente
            if self.pk_is_numeric:
                max_expr = f"COALESCE(MAX({col_qual})::bigint, 0)"
            else:
                max_expr = f"COALESCE(MAX((NULLIF(regexp_replace({col_qual}::text,'[^0-9]','','g'),'') )::bigint),0)"

            self.db.cursor.execute(
                f"""
                SELECT setval(
                    %s::regclass,
                    GREATEST((SELECT {max_expr} FROM {table_qual}) + 1, 1),
                    false
                )
                """,
                (seq_qual,),
            )

            self.db.commit()
            return seq_qual

        except Exception as e:
            self.db.rollback()
            log_deposito(
                f"SEQ: falhou criar/vincular -> {type(e).__name__}: {e}")
            return None
        finally:
            self.db.desconectar()

    def proximo_codigo(self) -> int:
        # 1) tenta sequence
        seq = self._ensure_sequence()
        if seq:
            if not self.db.conectar():
                raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
            try:
                assert self.db.cursor is not None
                self.db.cursor.execute("SELECT nextval(%s::regclass)", (seq,))
                r = self.db.cursor.fetchone()
                return int(r[0])
            finally:
                self.db.desconectar()

        # 2) fallback MAX+1 (sem depender de permission)
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            pk = _qident(self.pk_col)
            if self.pk_is_numeric:
                sql = f"SELECT COALESCE(MAX({pk})::bigint, 0) + 1 FROM {self.deposito_table}"
                self.db.cursor.execute(sql)
            else:
                sql = f"""
                    SELECT COALESCE(
                        MAX((NULLIF(regexp_replace({pk}::text,'[^0-9]','','g'),'') )::bigint),
                        0
                    ) + 1
                    FROM {self.deposito_table}
                """
                self.db.cursor.execute(sql)
            r = self.db.cursor.fetchone()
            return int(r[0] or 1)
        finally:
            self.db.desconectar()

    # ----------------- CRUD -----------------

    def listar(self, termo: Optional[str] = None, limit: int = 2000) -> List[Deposito]:
        termo = (termo or "").strip()
        like = f"%{termo}%" if termo else None

        pk = _qident(self.pk_col)
        desc = _qident(self.col_descricao)
        ativo = _qident(self.col_ativo)
        padrao = _qident(self.col_padrao)
        descons = _qident(self.col_desconsidera)

        sql = f"""
            SELECT
                {pk}::text AS codigo,
                COALESCE({desc}::text,'') AS descricao,
                {ativo},
                {padrao},
                {descons}
            FROM {self.deposito_table}
            WHERE (%s IS NULL)
               OR ({pk}::text ILIKE %s)
               OR (COALESCE({desc}::text,'') ILIKE %s)
            ORDER BY {pk}::text
            LIMIT %s
        """
        params = (termo or None, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            out: List[Deposito] = []
            for (codigo, descricao, v_ativo, v_padrao, v_descons) in self.db.cursor.fetchall():
                out.append(
                    Deposito(
                        codigo=str(codigo or ""),
                        descricao=str(descricao or ""),
                        ativo=_to_bool(v_ativo),
                        padrao=_to_bool(v_padrao),
                        desconsidera_saldo=_to_bool(v_descons),
                    )
                )
            return out
        finally:
            self.db.desconectar()

    def existe(self, codigo: str) -> bool:
        pk = _qident(self.pk_col)
        sql = f"SELECT 1 FROM {self.deposito_table} WHERE {pk}::text = %s"
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(codigo),))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def _encode_flag(self, col_name: str, is_bool: bool, value: bool) -> Any:
        if is_bool:
            return bool(value)

        # se for coluna de "ativo/situacao", normalmente é 'A'/'I'
        if col_name.lower() in {"ativo", "situacao", "status"}:
            return "A" if value else "I"

        # padrão para outras flags: 'S'/'N'
        return "S" if value else "N"

    def inserir(self, codigo_int: int, descricao: str, ativo: bool, padrao: bool, descons: bool) -> str:
        pk = _qident(self.pk_col)
        desc = _qident(self.col_descricao)
        c_ativo = _qident(self.col_ativo)
        c_padrao = _qident(self.col_padrao)
        c_descons = _qident(self.col_desconsidera)

        codigo_param = self._pk_param(codigo_int)
        ativo_param = self._encode_flag(
            self.col_ativo, self.ativo_is_bool, ativo)
        padrao_param = self._encode_flag(
            self.col_padrao, self.padrao_is_bool, padrao)
        descons_param = self._encode_flag(
            self.col_desconsidera, self.desconsidera_is_bool, descons)

        sql = f"""
            INSERT INTO {self.deposito_table} ({pk},{desc},{c_ativo},{c_padrao},{c_descons})
            VALUES (%s,%s,%s,%s,%s)
            RETURNING {pk}::text
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                sql, (codigo_param, descricao, ativo_param, padrao_param, descons_param))
            new_code = self.db.cursor.fetchone()[0]
            self.db.commit()
            return str(new_code)
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def atualizar(self, codigo_txt: str, descricao: str, ativo: bool, padrao: bool, descons: bool) -> None:
        pk = _qident(self.pk_col)
        desc = _qident(self.col_descricao)
        c_ativo = _qident(self.col_ativo)
        c_padrao = _qident(self.col_padrao)
        c_descons = _qident(self.col_desconsidera)

        ativo_param = self._encode_flag(
            self.col_ativo, self.ativo_is_bool, ativo)
        padrao_param = self._encode_flag(
            self.col_padrao, self.padrao_is_bool, padrao)
        descons_param = self._encode_flag(
            self.col_desconsidera, self.desconsidera_is_bool, descons)

        sql = f"""
            UPDATE {self.deposito_table}
               SET {desc} = %s,
                   {c_ativo} = %s,
                   {c_padrao} = %s,
                   {c_descons} = %s
             WHERE {pk}::text = %s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                sql, (descricao, ativo_param, padrao_param, descons_param, str(codigo_txt)))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def excluir(self, codigo_txt: str) -> None:
        pk = _qident(self.pk_col)
        sql = f"DELETE FROM {self.deposito_table} WHERE {pk}::text = %s"
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(codigo_txt),))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class DepositoService:
    def __init__(self, repo: DepositoRepo) -> None:
        self.repo = repo

    def listar(self, termo: Optional[str]) -> List[Deposito]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def proximo_codigo(self) -> int:
        return self.repo.proximo_codigo()

    def salvar(self, codigo_txt: str, descricao: str, ativo: bool, padrao: bool, descons: bool) -> Tuple[str, str]:
        descricao = (descricao or "").strip()
        if not descricao:
            raise ValueError("Descrição é obrigatória.")

        codigo_txt = (codigo_txt or "").strip()
        if not codigo_txt:
            codigo_int = self.repo.proximo_codigo()
            codigo_salvo = self.repo.inserir(
                codigo_int, descricao, ativo, padrao, descons)
            return ("inserido", codigo_salvo)

        # se usuário digitou, tentamos tratar como int; se der, ok; senão, salva como texto
        try:
            codigo_int = int(codigo_txt)
            if self.repo.existe(codigo_txt):
                self.repo.atualizar(codigo_txt, descricao,
                                    ativo, padrao, descons)
                return ("atualizado", codigo_txt)
            codigo_salvo = self.repo.inserir(
                codigo_int, descricao, ativo, padrao, descons)
            return ("inserido", codigo_salvo)
        except ValueError:
            # código não numérico: só atualiza (se existir) ou tenta inserir como sequência numérica não faz sentido
            if self.repo.existe(codigo_txt):
                self.repo.atualizar(codigo_txt, descricao,
                                    ativo, padrao, descons)
                return ("atualizado", codigo_txt)
            raise ValueError(
                "Código inválido (não numérico). Use o botão NOVO para gerar automaticamente.")

    def excluir(self, codigo_txt: str) -> None:
        self.repo.excluir(codigo_txt)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1100x650"
APP_TITLE = "Tela de Depósito"

TREE_COLS = ["codigo", "descricao", "ativo", "padrao", "descons"]


class TelaDeposito(ttk.Frame):
    def __init__(self, master: tk.Misc, service: DepositoService):
        super().__init__(master)
        self.service = service

        self.var_filtro = tk.StringVar()

        self.var_codigo = tk.StringVar()
        self.var_descricao = tk.StringVar()
        self.var_ativo = tk.BooleanVar(value=True)
        self.var_padrao = tk.BooleanVar(value=False)
        self.var_descons = tk.BooleanVar(value=False)

        self._build_ui()
        self.atualizar_lista()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (ID/Descrição):").grid(row=0,
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

        form = ttk.LabelFrame(self, text="Depósito")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(10):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="Código:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_codigo = ttk.Entry(form, textvariable=self.var_codigo, width=18)
        ent_codigo.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=6)

        ttk.Label(form, text="Descrição:").grid(
            row=0, column=3, sticky="w", padx=(10, 6), pady=6)
        ent_desc = ttk.Entry(form, textvariable=self.var_descricao)
        ent_desc.grid(row=0, column=4, sticky="ew",
                      padx=(0, 10), pady=6, columnspan=6)

        chk_ativo = ttk.Checkbutton(
            form, text="Ativo (situação)", variable=self.var_ativo)
        chk_padrao = ttk.Checkbutton(
            form, text="Padrão", variable=self.var_padrao)
        chk_des = ttk.Checkbutton(
            form, text="Desconsiderar saldo", variable=self.var_descons)

        chk_ativo.grid(row=1, column=0, sticky="w", padx=(
            10, 10), pady=(0, 6), columnspan=2)
        chk_padrao.grid(row=1, column=2, sticky="w",
                        padx=(10, 10), pady=(0, 6))
        chk_des.grid(row=1, column=3, sticky="w", padx=(
            10, 10), pady=(0, 6), columnspan=3)

        ent_desc.bind("<Return>", lambda e: self.salvar())

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

        self.tree.heading("codigo", text="Código")
        self.tree.heading("descricao", text="Descrição")
        self.tree.heading("ativo", text="Ativo")
        self.tree.heading("padrao", text="Padrão")
        self.tree.heading("descons", text="Desconsidera Saldo")

        self.tree.column("codigo", width=140, anchor="e", stretch=False)
        self.tree.column("descricao", width=520, anchor="w", stretch=True)
        self.tree.column("ativo", width=90, anchor="center", stretch=False)
        self.tree.column("padrao", width=90, anchor="center", stretch=False)
        self.tree.column("descons", width=140, anchor="center", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar depósito:\n{e}")
            return

        def simnao(b: bool) -> str:
            return "Sim" if b else "Não"

        for d in rows:
            self.tree.insert("", "end", values=(
                d.codigo,
                d.descricao,
                simnao(d.ativo),
                simnao(d.padrao),
                simnao(d.desconsidera_saldo),
            ))

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        codigo, descricao, ativo, padrao, descons = self.tree.item(
            sel[0], "values")
        self.var_codigo.set(str(codigo or ""))
        self.var_descricao.set(str(descricao or ""))
        self.var_ativo.set(str(ativo).strip().lower() == "sim")
        self.var_padrao.set(str(padrao).strip().lower() == "sim")
        self.var_descons.set(str(descons).strip().lower() == "sim")

    def limpar_form(self) -> None:
        self.var_codigo.set("")
        self.var_descricao.set("")
        self.var_ativo.set(True)
        self.var_padrao.set(False)
        self.var_descons.set(False)
        self.tree.selection_remove(self.tree.selection())

    def novo(self) -> None:
        # ✅ AQUI está a correção pedida: gera código automaticamente
        self.limpar_form()
        try:
            prox = self.service.proximo_codigo()
            self.var_codigo.set(str(prox))
            # mantém descrição em branco para digitar
        except Exception as e:
            messagebox.showerror(
                "Erro",
                "Falha ao gerar próximo Código.\n"
                "Obs: o sistema tenta usar sequence/nextval; se não houver permissão, usa MAX+1.\n\n"
                f"Detalhe:\n{e}"
            )

    def salvar(self) -> None:
        try:
            status, codigo = self.service.salvar(
                self.var_codigo.get(),
                self.var_descricao.get(),
                bool(self.var_ativo.get()),
                bool(self.var_padrao.get()),
                bool(self.var_descons.get()),
            )
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        self.var_codigo.set(str(codigo))
        messagebox.showinfo(
            "OK", f"Depósito {status} com sucesso.\nCódigo: {codigo}")
        self.atualizar_lista()

    def excluir(self) -> None:
        codigo = (self.var_codigo.get() or "").strip()
        if not codigo:
            messagebox.showwarning(
                "Atenção", "Selecione um registro para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Depósito {codigo}?"):
            return

        try:
            self.service.excluir(codigo)
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

    deposito_table = detectar_tabela(cfg, DEPOSITO_TABLES, '"deposito"')

    db = Database(cfg)
    repo = DepositoRepo(db, deposito_table=deposito_table)
    service = DepositoService(repo)

    tela = TelaDeposito(root, service)
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
        log_deposito(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_deposito:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
