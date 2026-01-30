from __future__ import annotations

import argparse
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


def db_connect(cfg: AppConfig):
    return psycopg2.connect(
        host=cfg.db_host,
        database=cfg.db_database,
        user=cfg.db_user,
        password=cfg.db_password,
        port=int(cfg.db_port),
        connect_timeout=5,
    )


# ============================================================
# DB "mini helper"
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
            self.conn = db_connect(self.cfg)
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
# MENU PRINCIPAL
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
# CONTROLE DE ACESSO (igual Arranjo)
# ============================================================

# ajuste aqui se no seu cadastro o nome/código for diferente
THIS_PROGRAMA_TERMO = "Depósito"

NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Edição",
}


def _deny_and_exit(msg: str) -> None:
    r = tk.Tk()
    try:
        r.withdraw()
        messagebox.showerror("Acesso negado", msg, parent=r)
    finally:
        try:
            r.destroy()
        except Exception:
            pass
    raise SystemExit(1)


def fetch_user_nome(cfg: AppConfig, usuario_id: int) -> str:
    sql = """
        SELECT COALESCE(u."nome",'')
          FROM "Ekenox"."usuarios" u
         WHERE u."usuarioId"=%s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
        return str(r[0] or "").strip() if r else ""
    finally:
        conn.close()


def _user_esta_ativo(cfg: AppConfig, usuario_id: int) -> bool:
    sql = 'SELECT COALESCE(u."ativo", true) FROM "Ekenox"."usuarios" u WHERE u."usuarioId"=%s LIMIT 1'
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
        return bool(r[0]) if r else False
    finally:
        conn.close()


def _fetch_programa_id_por_termo(cfg: AppConfig, termo: str) -> Optional[int]:
    like = f"%{(termo or '').strip()}%"
    if like == "%%":
        return None

    sql = """
        SELECT pr."programaId"
          FROM "Ekenox"."programas" pr
         WHERE COALESCE(pr."nome",'') ILIKE %s
            OR COALESCE(pr."codigo",'') ILIKE %s
         ORDER BY pr."programaId" DESC
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (like, like))
            r = cur.fetchone()
        return int(r[0]) if r else None
    finally:
        conn.close()


def _fetch_user_nivel(cfg: AppConfig, usuario_id: int, programa_id: int) -> int:
    sql = """
        SELECT COALESCE(up."nivel", 0)
          FROM "Ekenox"."usuario_programa" up
         WHERE up."usuarioId"=%s AND up."programaId"=%s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id), int(programa_id)))
            r = cur.fetchone()
        return int(r[0] or 0) if r else 0
    finally:
        conn.close()


def get_access_level_for_this_screen(cfg: AppConfig, usuario_id: int) -> Tuple[int, str]:
    if not _user_esta_ativo(cfg, usuario_id):
        return 0, "Usuário inativo ou não encontrado."

    pid = _fetch_programa_id_por_termo(cfg, THIS_PROGRAMA_TERMO)
    if pid is None:
        return 1, (
            f'Atenção: não encontrei este programa na tabela "Ekenox"."programas".\n\n'
            f'Termo usado: "{THIS_PROGRAMA_TERMO}"\n\n'
            "Abrindo em NÍVEL 1 (Leitura). Cadastre o programa ou ajuste o termo."
        )

    nivel = _fetch_user_nivel(cfg, usuario_id, pid)
    if nivel <= 0:
        return 1, (
            "Atenção: não existe permissão cadastrada para este usuário neste programa.\n\n"
            f"Usuário ID: {usuario_id}\nPrograma ID: {pid}\n\n"
            "Abrindo em NÍVEL 1 (Leitura)."
        )

    if nivel not in (1, 2, 3):
        nivel = 1
    return nivel, ""


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
        conn = db_connect(cfg)
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

        self.pk_col = self._find_primary_key_column() or "codigo"
        self.pk_typname = self._col_typname(self.deposito_table, self.pk_col)
        self.pk_is_numeric = _is_numeric_pg_type(self.pk_typname)

        self.col_descricao = self._find_existing_column(
            ["descricao", "descricaoDeposito", "desc", "Descrição"]) or "descricao"
        self.col_ativo = self._find_existing_column(
            ["ativo", "situacao", "status", "Ativo"]) or "ativo"
        self.col_padrao = self._find_existing_column(
            ["padrao", "default", "Padrao"]) or "padrao"
        self.col_desconsidera = self._find_existing_column(
            ["desconsiderarsaldo", "desconsideraSaldo", "ignorasaldo"]) or "desconsideraSaldo"

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

    def _find_sequence(self) -> Optional[str]:
        if not self.db.conectar():
            return None
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                "SELECT pg_get_serial_sequence(%s, %s)", (self.deposito_table, self.pk_col))
            r = self.db.cursor.fetchone()
            return str(r[0]) if r and r[0] else None
        except Exception:
            return None
        finally:
            self.db.desconectar()

    def proximo_codigo(self) -> int:
        seq = self._find_sequence()
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

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            pk = _qident(self.pk_col)
            if self.pk_is_numeric:
                self.db.cursor.execute(
                    f"SELECT COALESCE(MAX({pk})::bigint, 0) + 1 FROM {self.deposito_table}")
            else:
                self.db.cursor.execute(
                    f"""
                    SELECT COALESCE(
                        MAX((NULLIF(regexp_replace({pk}::text,'[^0-9]','','g'),'') )::bigint),
                        0
                    ) + 1
                    FROM {self.deposito_table}
                    """
                )
            r = self.db.cursor.fetchone()
            return int(r[0] or 1)
        finally:
            self.db.desconectar()

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
        if col_name.lower() in {"ativo", "situacao", "status"}:
            return "A" if value else "I"
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
            if self.repo.existe(codigo_txt):
                self.repo.atualizar(codigo_txt, descricao,
                                    ativo, padrao, descons)
                return ("atualizado", codigo_txt)
            raise ValueError(
                "Código inválido (não numérico). Use o botão NOVO para gerar automaticamente.")

    def excluir(self, codigo_txt: str) -> None:
        self.repo.excluir(codigo_txt)


# ============================================================
# UI / APP
# ============================================================

DEFAULT_GEOMETRY = "1100x650"
APP_TITLE = "Tela de Depósito"
TREE_COLS = ["codigo", "descricao", "ativo", "padrao", "descons"]


class TelaDepositoApp(tk.Tk):
    def __init__(self, cfg: AppConfig, service: DepositoService, *, usuario_id: int, nivel: int, aviso: str, from_menu: bool) -> None:
        super().__init__()
        self.cfg = cfg
        self.service = service

        self.usuario_id = int(usuario_id)
        self.nivel = int(nivel)
        self.aviso = aviso or ""
        self.from_menu = bool(from_menu)

        self.usuario_nome = fetch_user_nome(
            cfg, self.usuario_id) if self.usuario_id else ""

        self.title(APP_TITLE)
        self.geometry(DEFAULT_GEOMETRY)
        apply_window_icon(self)

        try:
            style = ttk.Style()
            if "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

        self._build_ui()
        self._apply_access_rules()
        self._load()

        if self.aviso:
            self.after(200, lambda: messagebox.showwarning(
                "Aviso", self.aviso, parent=self))

        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _can_edit(self) -> bool:
        return self.nivel >= 2  # 2 e 3 editam

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        # Topbar (igual Arranjo/Usuários)
        topbar = ttk.Frame(root)
        topbar.pack(fill="x", pady=(0, 8))
        topbar.columnconfigure(0, weight=1)

        self.lbl_access = ttk.Label(
            topbar, text="", foreground="gray", font=("Segoe UI", 9, "bold"))
        self.lbl_access.pack(side="left", anchor="w")

        self.btn_voltar = ttk.Button(
            topbar, text="Voltar ao Menu", command=self._voltar_ou_fechar)
        self.btn_voltar.pack(side="right")

        # Busca e botões
        top = ttk.Frame(root)
        top.pack(fill="x", pady=(0, 8))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (ID/Descrição):").grid(row=0,
                                                           column=0, sticky="w")
        self.var_filtro = tk.StringVar()
        self.ent_busca = ttk.Entry(top, textvariable=self.var_filtro)
        self.ent_busca.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        self.ent_busca.bind("<Return>", lambda e: self._load())

        ttk.Button(top, text="Atualizar", command=self._load).grid(
            row=0, column=2, padx=(0, 6))
        self.btn_novo = ttk.Button(top, text="Novo", command=self._novo)
        self.btn_novo.grid(row=0, column=3, padx=(0, 6))
        self.btn_salvar = ttk.Button(top, text="Salvar", command=self._salvar)
        self.btn_salvar.grid(row=0, column=4, padx=(0, 6))
        self.btn_excluir = ttk.Button(
            top, text="Excluir", command=self._excluir)
        self.btn_excluir.grid(row=0, column=5, padx=(0, 6))
        ttk.Button(top, text="Limpar", command=self._limpar).grid(
            row=0, column=6)

        # Form
        form = ttk.LabelFrame(root, text="Depósito", padding=10)
        form.pack(fill="x", pady=(0, 10))
        form.columnconfigure(4, weight=1)

        self.var_codigo = tk.StringVar()
        self.var_desc = tk.StringVar()
        self.var_ativo = tk.BooleanVar(value=True)
        self.var_padrao = tk.BooleanVar(value=False)
        self.var_descons = tk.BooleanVar(value=False)

        ttk.Label(form, text="Código:").grid(row=0, column=0, sticky="w")
        self.ent_codigo = ttk.Entry(
            form, textvariable=self.var_codigo, width=18)
        self.ent_codigo.grid(row=0, column=1, sticky="w", padx=(6, 14))

        ttk.Label(form, text="Descrição:").grid(row=0, column=2, sticky="w")
        self.ent_desc = ttk.Entry(form, textvariable=self.var_desc)
        self.ent_desc.grid(row=0, column=3, columnspan=2,
                           sticky="ew", padx=(6, 0))

        self.chk_ativo = ttk.Checkbutton(
            form, text="Ativo (situação)", variable=self.var_ativo)
        self.chk_padrao = ttk.Checkbutton(
            form, text="Padrão", variable=self.var_padrao)
        self.chk_des = ttk.Checkbutton(
            form, text="Desconsiderar saldo", variable=self.var_descons)
        self.chk_ativo.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.chk_padrao.grid(row=1, column=2, sticky="w", pady=(8, 0))
        self.chk_des.grid(row=1, column=3, sticky="w", pady=(8, 0))

        self.ent_desc.bind("<Return>", lambda e: self._salvar())

        # Lista
        lst = ttk.Frame(root)
        lst.pack(fill="both", expand=True)
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

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def _apply_access_rules(self) -> None:
        nome = self.usuario_nome or (
            f"ID {self.usuario_id}" if self.usuario_id else "Não informado")
        nivel_txt = NIVEL_LABEL.get(self.nivel, str(self.nivel))

        can_edit = self._can_edit()
        self.lbl_access.config(
            text=f"Logado: {nome} | Nível: {nivel_txt}",
            foreground=("green" if can_edit else "gray"),
        )

        # Botão voltar (se veio do menu, só fecha)
        self.btn_voltar.config(
            text=("Fechar" if self.from_menu else "Voltar ao Menu"))

        # Campos e checks
        self.ent_codigo.configure(state=("normal" if can_edit else "readonly"))
        self.ent_desc.configure(state=("normal" if can_edit else "readonly"))
        for chk in (self.chk_ativo, self.chk_padrao, self.chk_des):
            chk.state(["!disabled"] if can_edit else ["disabled"])

        # Botões
        for btn in (self.btn_novo, self.btn_salvar, self.btn_excluir):
            btn.state(["!disabled"] if can_edit else ["disabled"])

    def _limpar(self) -> None:
        self.var_codigo.set("")
        self.var_desc.set("")
        self.var_ativo.set(True)
        self.var_padrao.set(False)
        self.var_descons.set(False)
        self.tree.selection_remove(self.tree.selection())

    def _load(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao listar depósito:\n{e}", parent=self)
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

    def _on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        codigo, descricao, ativo, padrao, descons = self.tree.item(
            sel[0], "values")
        self.var_codigo.set(str(codigo or ""))
        self.var_desc.set(str(descricao or ""))
        self.var_ativo.set(str(ativo).strip().lower() == "sim")
        self.var_padrao.set(str(padrao).strip().lower() == "sim")
        self.var_descons.set(str(descons).strip().lower() == "sim")

    def _novo(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Seu nível é somente leitura.", parent=self)
            return
        self._limpar()
        try:
            prox = self.service.proximo_codigo()
            self.var_codigo.set(str(prox))
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao gerar próximo código:\n{e}", parent=self)

    def _salvar(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Seu nível é somente leitura.", parent=self)
            return
        try:
            status, codigo = self.service.salvar(
                self.var_codigo.get(),
                self.var_desc.get(),
                bool(self.var_ativo.get()),
                bool(self.var_padrao.get()),
                bool(self.var_descons.get()),
            )
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e), parent=self)
            return

        self.var_codigo.set(str(codigo))
        messagebox.showinfo(
            "OK", f"Depósito {status} com sucesso.\nCódigo: {codigo}", parent=self)
        self._load()

    def _excluir(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Seu nível é somente leitura.", parent=self)
            return

        codigo = (self.var_codigo.get() or "").strip()
        if not codigo:
            messagebox.showwarning(
                "Atenção", "Selecione um registro para excluir.", parent=self)
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Depósito {codigo}?", parent=self):
            return

        try:
            self.service.excluir(codigo)
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao excluir:\n{e}", parent=self)
            return

        messagebox.showinfo("OK", "Registro excluído.", parent=self)
        self._limpar()
        self._load()

    def _voltar_ou_fechar(self) -> None:
        # Se veio do menu, só fecha (não cria menu novo)
        if self.from_menu:
            self.destroy()
            return

        # Standalone: abre o menu e fecha
        try:
            abrir_menu_principal_skip_entrada()
        finally:
            self.destroy()

# ============================================================
# STARTUP
# ============================================================


def test_connection_or_die(cfg: AppConfig) -> None:
    conn = None
    try:
        conn = db_connect(cfg)
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


def _detect_from_menu_flag() -> bool:
    # Força standalone (se você quiser testar manualmente)
    if "--standalone" in sys.argv:
        return False

    # Se o menu chamou, normalmente passa usuario-id
    if "--usuario-id" in sys.argv or "--uid" in sys.argv:
        return True

    # Flags explícitas
    if "--from-menu" in sys.argv:
        return True

    # Alguns menus reaproveitam isso
    if "--skip-entrada" in sys.argv:
        return True

    # Variável de ambiente (se você usar no subprocess)
    env = (os.getenv("EKENOX_FROM_MENU") or "").strip().lower()
    return env in {"1", "true", "yes", "sim", "s"}


def _parse_usuario_id() -> Optional[int]:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--usuario-id", "--uid", dest="usuario_id", type=int)
    args, _ = parser.parse_known_args(sys.argv[1:])
    return int(args.usuario_id) if args.usuario_id else None


def main() -> None:
    cfg = env_override(load_config())

    from_menu = _detect_from_menu_flag()
    usuario_id = _parse_usuario_id()

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
        return

    deposito_table = detectar_tabela(cfg, DEPOSITO_TABLES, '"deposito"')
    db = Database(cfg)
    repo = DepositoRepo(db, deposito_table=deposito_table)
    service = DepositoService(repo)

    # Permissões (se não veio usuário, entra leitura)
    if usuario_id is None:
        nivel = 1
        aviso = (
            "Atenção: usuário não informado ao abrir a tela.\n\n"
            "Abrindo em NÍVEL 1 (Leitura).\n"
            "Para respeitar permissões reais, chame com --usuario-id <id>."
        )
        uid = 0
    else:
        uid = int(usuario_id)
        nivel, aviso = get_access_level_for_this_screen(cfg, uid)
        if nivel <= 0:
            _deny_and_exit(aviso)

    app = TelaDepositoApp(cfg, service, usuario_id=uid,
                          nivel=nivel, aviso=aviso, from_menu=from_menu)
    app.mainloop()


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
