from __future__ import annotations

import argparse
import glob
import json
import os
import subprocess
import sys
import tempfile
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from tkinter import messagebox, ttk
from typing import Any, List, Optional, Tuple

import psycopg2


# ============================================================
# PROGRAMA / PERMISSÕES
# ============================================================

PROGRAMA_CODIGO = "ARRANJO"  # deve bater com Ekenox.programas.codigo


NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Admin",
}


# ============================================================
# TABELAS (AJUSTE SE NECESSÁRIO)
# ============================================================

ARRANJO_TABLE = '"Ekenox"."arranjo"'
PRODUTOS_TABLE = '"Ekenox"."produtos"'
USUARIOS_TABLE = '"Ekenox"."usuarios"'


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
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
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
    conn = psycopg2.connect(
        host=cfg.db_host,
        database=cfg.db_database,
        user=cfg.db_user,
        password=cfg.db_password,
        port=int(cfg.db_port),
        connect_timeout=5,
    )

    # encoding automático (mantém seu comportamento)
    try:
        forced = (os.getenv("DB_CLIENT_ENCODING") or "").strip()
        if forced:
            conn.set_client_encoding(forced)
            return conn

        with conn.cursor() as cur:
            cur.execute("SHOW server_encoding")
            server_enc = str(cur.fetchone()[0] or "").strip().upper()

        if server_enc == "SQL_ASCII":
            conn.set_client_encoding("WIN1252")
        else:
            conn.set_client_encoding(server_enc or "UTF8")
    except Exception:
        pass

    return conn


# ============================================================
# DB (uso geral)
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
# MENU PRINCIPAL (igual Depósito)
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
    for pasta in (APP_DIR, BASE_DIR, os.getcwd()):
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
# SESSÃO / RESOLUÇÃO DE USUÁRIO
# ============================================================

def _extract_user_id(obj: Any) -> Optional[int]:
    key_candidates = {
        "user_id", "userid",
        "usuario_id", "usuarioid",
        "id_usuario", "idusuario",
        "logged_user_id", "usuario_logado_id",
        "user", "usuario",
        "id",
        "usuarioId", "UsuarioId", "usuarioID",
    }

    def as_int(v: Any) -> Optional[int]:
        try:
            if v is None or isinstance(v, bool):
                return None
            if isinstance(v, int):
                return v
            s = str(v).strip()
            if s.isdigit():
                return int(s)
        except Exception:
            return None
        return None

    if isinstance(obj, dict):
        for k, v in obj.items():
            kl = str(k).strip()
            kll = kl.lower()

            if kll in {x.lower() for x in key_candidates}:
                if isinstance(v, dict):
                    for kk, vv in v.items():
                        if str(kk).strip().lower() in {"id", "user_id", "usuario_id", "usuarioid", "userid", "usuarioid"}:
                            got = as_int(vv)
                            if got is not None:
                                return got
                    got = _extract_user_id(v)
                    if got is not None:
                        return got
                else:
                    got = as_int(v)
                    if got is not None:
                        return got

            if ("usuario" in kll or "user" in kll) and kll.endswith("id"):
                got = as_int(v)
                if got is not None:
                    return got

        for v in obj.values():
            got = _extract_user_id(v)
            if got is not None:
                return got

    if isinstance(obj, (list, tuple)):
        for it in obj:
            got = _extract_user_id(it)
            if got is not None:
                return got

    return None


def _candidate_session_dirs() -> List[str]:
    dirs: List[str] = []
    for d in (BASE_DIR, APP_DIR, os.getcwd(), tempfile.gettempdir()):
        try:
            if d and os.path.isdir(d) and d not in dirs:
                dirs.append(d)
        except Exception:
            continue

    try:
        user_home = os.path.expanduser("~")
        if user_home and os.path.isdir(user_home) and user_home not in dirs:
            dirs.append(user_home)
        docs = os.path.join(user_home, "Documents")
        if os.path.isdir(docs) and docs not in dirs:
            dirs.append(docs)
    except Exception:
        pass

    return dirs


def _load_user_id_from_session_files(session_file: Optional[str] = None) -> Optional[int]:
    # 1) ENV direto
    for envk in ("EKENOX_USER_ID", "USER_ID", "USUARIO_ID", "LOGGED_USER_ID"):
        v = (os.getenv(envk) or "").strip()
        if v.isdigit():
            return int(v)

    # 2) arquivo de sessão explícito
    explicit = (session_file or os.getenv("EKENOX_SESSION_FILE") or "").strip()
    if explicit and os.path.isfile(explicit):
        try:
            with open(explicit, "r", encoding="utf-8") as f:
                data = json.load(f)
            uid = _extract_user_id(data)
            if uid is not None:
                return uid
        except Exception:
            pass

    # 3) nomes conhecidos
    candidates = [
        "sessao.json", "sessao_atual.json", "sessao_usuario.json",
        "session.json", "current_session.json",
        "usuario_logado.json", "usuarioAtual.json", "usuario_atual.json",
        "login.json", "login_atual.json", "auth.json", "autenticacao.json",
        "entrada.json", "entrada_op.json", "entrada_usuario.json",
        "estado.json", "state.json",
        "contexto.json", "context.json",
    ]

    search_dirs = _candidate_session_dirs()
    for d in search_dirs:
        for name in candidates:
            path = os.path.join(d, name)
            if not os.path.isfile(path):
                continue
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                uid = _extract_user_id(data)
                if uid is not None:
                    return uid
            except Exception:
                continue

    # 4) varredura por padrões
    patterns: List[str] = []
    for d in search_dirs:
        patterns += [
            os.path.join(d, "*sess*.json"),
            os.path.join(d, "*login*.json"),
            os.path.join(d, "*usuario*.json"),
            os.path.join(d, "*auth*.json"),
            os.path.join(d, "*entrada*.json"),
            os.path.join(d, "*user*.json"),
        ]

    files: List[str] = []
    for pat in patterns:
        files.extend(glob.glob(pat))

    uniq: dict[str, float] = {}
    for p in files:
        try:
            base = os.path.basename(p).lower()
            if base in {"config_op.json"}:
                continue
            size = os.path.getsize(p)
            if size > 2_000_000:
                continue
            uniq[p] = os.path.getmtime(p)
        except Exception:
            continue

    ordered = sorted(uniq.items(), key=lambda kv: kv[1], reverse=True)
    for path, _mtime in ordered[:50]:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            uid = _extract_user_id(data)
            if uid is not None:
                return uid
        except Exception:
            continue

    return None


def _usuarios_cols(cfg: AppConfig) -> dict[str, str]:
    """
    Descobre colunas reais de Ekenox.usuarios, priorizando padrão Depósito:
      - ID: usuarioId (preferência), depois id
      - Nome: nome
      - Ativo: ativo
    Retorna dict com chaves: id_col, nome_col, ativo_col
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema='Ekenox' AND table_name='usuarios'
                """
            )
            cols = [str(r[0]) for r in cur.fetchall()]
    except Exception:
        cols = []
    finally:
        try:
            conn.close()
        except Exception:
            pass

    low = {c.lower(): c for c in cols}

    def pick(*names: str) -> Optional[str]:
        for n in names:
            c = low.get(n.lower())
            if c:
                return c
        return None

    id_col = pick("usuarioId", "usuarioid", "usuario_id",
                  "id", "userid", "user_id") or "usuarioId"
    nome_col = pick("nome", "name") or "nome"
    ativo_col = pick("ativo", "active", "status") or "ativo"

    return {"id_col": id_col, "nome_col": nome_col, "ativo_col": ativo_col}


def _resolve_user_id(cfg: AppConfig, user_id_raw: Optional[str], user_hash: Optional[str]) -> Optional[int]:
    # 1) veio id explícito
    if user_id_raw:
        s = str(user_id_raw).strip()
        if s.isdigit():
            uid = int(s)
            log_arranjo(f"RESOLVE: user_id via argumento/env = {uid}")
            return uid

    # 2) veio hash => procura no banco
    if not user_hash:
        return None

    cols = _usuarios_cols(cfg)
    id_col = cols["id_col"]

    conn = None
    try:
        conn = db_connect(cfg)
        with conn.cursor() as cur:
            # tenta várias colunas de hash comuns
            hash_cols_try = ["hash", "email_hash",
                             "hash_email", "hashEmail", "emailHash"]
            # descobre as que existem
            cur.execute(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema='Ekenox' AND table_name='usuarios'
                """
            )
            existing = {str(r[0]).lower() for r in cur.fetchall()}
            hash_cols = [hc for hc in hash_cols_try if hc.lower() in existing]

            for hc in hash_cols:
                sql = f'SELECT "{id_col}" FROM {USUARIOS_TABLE} WHERE "{hc}"=%s LIMIT 1'
                try:
                    cur.execute(sql, (user_hash,))
                    r = cur.fetchone()
                    if r and r[0] is not None:
                        uid = int(r[0])
                        log_arranjo(
                            f"RESOLVE: user_id={uid} via hash coluna={hc}")
                        return uid
                except Exception:
                    continue

    except Exception as e:
        log_arranjo(f"RESOLVE: erro via hash: {type(e).__name__}: {e}")
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass

    return None


def fetch_user_nome(cfg: AppConfig, usuario_id: int) -> str:
    cols = _usuarios_cols(cfg)
    id_col = cols["id_col"]
    nome_col = cols["nome_col"]

    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            sql = f'SELECT COALESCE(u."{nome_col}",\'\') FROM {USUARIOS_TABLE} u WHERE u."{id_col}"=%s LIMIT 1'
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
            return str(r[0] or "").strip() if r else ""
    except Exception:
        return ""
    finally:
        conn.close()


def user_esta_ativo(cfg: AppConfig, usuario_id: int) -> bool:
    cols = _usuarios_cols(cfg)
    id_col = cols["id_col"]
    ativo_col = cols["ativo_col"]

    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            sql = f'SELECT COALESCE(u."{ativo_col}", true) FROM {USUARIOS_TABLE} u WHERE u."{id_col}"=%s LIMIT 1'
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
            return bool(r[0]) if r else False
    except Exception:
        return False
    finally:
        conn.close()


# ============================================================
# PERMISSÕES (nível por programa)
# ============================================================

def _pick_col(cols: list[str], *names: str) -> Optional[str]:
    low = {c.lower(): c for c in cols}
    for n in names:
        c = low.get(n.lower())
        if c:
            return c
    return None


def _table_exists(cur, schema: str, table: str) -> bool:
    cur.execute(
        """
        SELECT 1
        FROM information_schema.tables
        WHERE table_schema=%s AND table_name=%s
        """,
        (schema, table),
    )
    return cur.fetchone() is not None


def _get_columns(cur, schema: str, table: str) -> list[str]:
    cur.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema=%s AND table_name=%s
        """,
        (schema, table),
    )
    return [r[0] for r in cur.fetchall()]


def obter_nivel_programa(cfg: AppConfig, user_id: Optional[int], programa_codigo: str) -> Optional[int]:
    if not user_id:
        return None

    programa_codigo = (programa_codigo or "").strip().upper()

    conn = None
    cur = None
    try:
        conn = db_connect(cfg)
        cur = conn.cursor()

        prog_table = None
        for cand in ("programas", "programa"):
            if _table_exists(cur, "Ekenox", cand):
                prog_table = cand
                break
        if not prog_table:
            return None

        up_table = "usuario_programa"
        if not _table_exists(cur, "Ekenox", up_table):
            return None

        prog_cols = _get_columns(cur, "Ekenox", prog_table)
        up_cols = _get_columns(cur, "Ekenox", up_table)

        prog_id_col = _pick_col(prog_cols, "programaid", "programaId", "id")
        prog_code_col = _pick_col(
            prog_cols, "codigo", "cod", "code", "sigla", "chave")

        up_user_col = _pick_col(up_cols, "usuarioid",
                                "usuarioId", "usuario_id", "user_id", "userid")
        up_prog_col = _pick_col(up_cols, "programaid",
                                "programaId", "programa_id", "program_id")
        up_nivel_col = _pick_col(up_cols, "nivel", "level")

        if not (prog_id_col and prog_code_col and up_user_col and up_prog_col and up_nivel_col):
            return None

        sql = f"""
            SELECT up."{up_nivel_col}"
            FROM "Ekenox"."{up_table}" up
            JOIN "Ekenox"."{prog_table}" p
              ON p."{prog_id_col}" = up."{up_prog_col}"
            WHERE up."{up_user_col}" = %s
              AND UPPER(p."{prog_code_col}") = %s
            LIMIT 1
        """
        cur.execute(sql, (int(user_id), programa_codigo))
        row = cur.fetchone()
        if not row:
            return 0
        return int(row[0] or 0)

    except Exception as e:
        log_arranjo(
            f"ACESSO: erro obter_nivel_programa: {type(e).__name__}: {e}")
        return None
    finally:
        try:
            if cur:
                cur.close()
        except Exception:
            pass
        try:
            if conn:
                conn.close()
        except Exception:
            pass


# ============================================================
# SESSAO ACESSO (padrão Depósito)
# ============================================================

@dataclass
class SessaoAcesso:
    nivel: int = 0
    usuario_nome: Optional[str] = None
    origem: str = "desconhecida"
    usuario_id: Optional[int] = None
    programa: str = PROGRAMA_CODIGO
    aviso: str = ""


def _build_access(cfg: AppConfig, ns) -> SessaoAcesso:
    """
    PADRÃO DEPÓSITO:
      - se não resolver usuário => nível 1 (leitura) com aviso
      - se resolver e estiver inativo => nível 0 (bloqueia)
      - se programa/permissão não existir => nível 1 com aviso
      - senão => nível conforme banco
    """
    raw = (
        getattr(ns, "user_id", None)
        or getattr(ns, "usuario_id", None)
        or getattr(ns, "uid", None)
        or os.getenv("EKENOX_USER_ID")
        or os.getenv("USER_ID")
        or os.getenv("USUARIO_ID")
        or os.getenv("LOGGED_USER_ID")
        or ""
    )
    user_id_raw = str(raw).strip() or None

    user_hash = (
        getattr(ns, "user_hash", None)
        or getattr(ns, "usuario_hash", None)
        or os.getenv("EKENOX_USER_HASH")
        or ""
    )
    user_hash = str(user_hash).strip() or None

    session_file = (getattr(ns, "session_file", None) or os.getenv(
        "EKENOX_SESSION_FILE") or "").strip() or None

    user_id = _resolve_user_id(cfg, user_id_raw, user_hash)
    if not user_id:
        user_id = _load_user_id_from_session_files(session_file=session_file)

    if user_id is None:
        aviso = (
            "Atenção: usuário não informado ao abrir a tela.\n\n"
            "Abrindo em NÍVEL 1 (Leitura).\n"
            "Para respeitar permissões reais, chame com --usuario-id <id> / --uid <id> / --user-id <id>\n"
            "ou forneça sessão (sessao.json/login.json etc)."
        )
        return SessaoAcesso(
            nivel=1,
            origem="sem_usuario",
            usuario_id=None,
            usuario_nome=None,
            programa=PROGRAMA_CODIGO,
            aviso=aviso,
        )

    uid = int(user_id)
    nome = fetch_user_nome(cfg, uid) or None

    # Usuário inativo => bloqueia
    if not user_esta_ativo(cfg, uid):
        return SessaoAcesso(
            nivel=0,
            origem="inativo",
            usuario_id=uid,
            usuario_nome=nome,
            programa=PROGRAMA_CODIGO,
            aviso="Usuário inativo ou não encontrado.",
        )

    nivel_db = obter_nivel_programa(cfg, uid, PROGRAMA_CODIGO)

    # Se não achou nível (erro) ou sem permissão => abre leitura com aviso (padrão depósito)
    if nivel_db is None or int(nivel_db) <= 0:
        aviso = (
            "Atenção: não existe permissão cadastrada para este usuário neste programa.\n\n"
            f"Usuário ID: {uid}\nPrograma: {PROGRAMA_CODIGO}\n\n"
            "Abrindo em NÍVEL 1 (Leitura)."
        )
        return SessaoAcesso(
            nivel=1,
            origem="sem_permissao",
            usuario_id=uid,
            usuario_nome=nome,
            programa=PROGRAMA_CODIGO,
            aviso=aviso,
        )

    return SessaoAcesso(
        nivel=int(nivel_db),
        origem="db",
        usuario_id=uid,
        usuario_nome=nome,
        programa=PROGRAMA_CODIGO,
        aviso="",
    )


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
            SELECT a."sku", a."nomeproduto", a."quantidade", a."chapa", a."material"
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

    def proximo_sku_numerico(self) -> str:
        sql = f"""
            SELECT COALESCE(MAX(CAST("sku" AS BIGINT)), 0)
            FROM {ARRANJO_TABLE}
            WHERE "sku" ~ '^[0-9]+$'
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql)
            mx = self.db.cursor.fetchone()[0]
            return str(int(mx) + 1)
        finally:
            self.db.desconectar()

    def inserir(self, a: Arranjo) -> None:
        sql = f"""
            INSERT INTO {ARRANJO_TABLE} ("sku","nomeproduto","quantidade","chapa","material")
            VALUES (%s,%s,%s,%s,%s)
        """
        params = (a.sku, a.nomeproduto, a.quantidade, a.chapa, a.material)

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

    def atualizar(self, a: Arranjo, sku_original: str) -> None:
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
        except Exception:
            self.db.rollback()
            raise
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
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def produto_por_sku(self, sku: str) -> Optional[Tuple[str, str]]:
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

    def proximo_sku(self) -> str:
        return self.repo.proximo_sku_numerico()

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
            self.repo.atualizar(a, sku_original=sku_original)
            return "atualizado"

        if self.repo.exists(a.sku):
            raise ValueError(
                "SKU já existe. Selecione na lista para editar ou clique em Novo.")

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
    def __init__(self, master: tk.Misc, service: ArranjoService, acesso: SessaoAcesso, *, from_menu: bool):
        super().__init__(master)
        self.service = service
        self.acesso = acesso
        self.from_menu = bool(from_menu)

        self.vars: dict[str, tk.StringVar] = {
            k: tk.StringVar() for k, _ in CAMPOS}
        self.var_filtro = tk.StringVar()

        self._sku_original: Optional[str] = None
        self.entries: dict[str, ttk.Entry] = {}

        self._build_ui()
        self._aplicar_permissoes()
        self.atualizar_lista()

        # aviso pós render
        if self.acesso.aviso:
            self.after(200, lambda: messagebox.showwarning(
                "Aviso", self.acesso.aviso, parent=self.winfo_toplevel()))

    def _can_edit(self) -> bool:
        return int(self.acesso.nivel or 0) >= 2

    def _can_delete(self) -> bool:
        return int(self.acesso.nivel or 0) >= 3

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)

        # TOPBAR (igual Depósito)
        topbar = ttk.Frame(self, padding=(10, 10, 10, 6))
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.columnconfigure(0, weight=1)

        nome = self.acesso.usuario_nome or (
            "Não informado" if not self.acesso.usuario_id else f"ID {self.acesso.usuario_id}")
        nivel_txt = NIVEL_LABEL.get(
            int(self.acesso.nivel or 0), str(self.acesso.nivel))

        self.lbl_access = ttk.Label(
            topbar,
            text=f"Logado: {nome} | Nível: {nivel_txt}",
            foreground=("green" if self._can_edit() else "gray"),
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_access.grid(row=0, column=0, sticky="w")

        self.btn_voltar = ttk.Button(topbar, text=(
            "Fechar" if self.from_menu else "Voltar ao Menu"), command=self._voltar_ou_fechar)
        self.btn_voltar.grid(row=0, column=1, sticky="e")

        # BUSCA
        top = ttk.Frame(self, padding=(10, 0, 10, 6))
        top.grid(row=1, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(
            top, text="Buscar (SKU / Nome / Chapa / Material):").grid(row=0, column=0, sticky="w")
        ent_busca = ttk.Entry(top, textvariable=self.var_filtro)
        ent_busca.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent_busca.bind("<Return>", lambda e: self.atualizar_lista())
        ttk.Button(top, text="Atualizar", command=self.atualizar_lista).grid(
            row=0, column=2, sticky="e")

        # AÇÕES
        linha_acoes = ttk.Frame(self, padding=(10, 0, 10, 6))
        linha_acoes.grid(row=2, column=0, sticky="ew")
        for i in range(4):
            linha_acoes.columnconfigure(i, weight=1)

        self.btn_novo = ttk.Button(linha_acoes, text="Novo", command=self.novo)
        self.btn_novo.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.btn_salvar = ttk.Button(
            linha_acoes, text="Salvar", command=self.salvar)
        self.btn_salvar.grid(row=0, column=1, sticky="ew", padx=(0, 6))

        self.btn_excluir = ttk.Button(
            linha_acoes, text="Excluir", command=self.excluir)
        self.btn_excluir.grid(row=0, column=2, sticky="ew", padx=(0, 6))

        self.btn_limpar = ttk.Button(
            linha_acoes, text="Limpar", command=self.limpar_form)
        self.btn_limpar.grid(row=0, column=3, sticky="ew")

        # FORM
        form = ttk.LabelFrame(self, text="Arranjo", padding=(10, 6, 10, 10))
        form.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(6):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="SKU:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        self.ent_sku = ttk.Entry(form, textvariable=self.vars["sku"])
        self.ent_sku.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=6)
        self.entries["sku"] = self.ent_sku

        self.btn_buscar_prod = ttk.Button(
            form, text="Buscar Produto...", command=self.buscar_produto_popup)
        self.btn_buscar_prod.grid(
            row=0, column=2, sticky="w", padx=(0, 6), pady=6)

        self.btn_preencher = ttk.Button(
            form, text="Preencher", command=self.preencher_nome_por_sku)
        self.btn_preencher.grid(
            row=0, column=3, sticky="w", padx=(0, 6), pady=6)

        ttk.Label(form, text="Nome Produto:").grid(
            row=0, column=4, sticky="w", padx=(10, 6), pady=6)
        ent_nome = ttk.Entry(form, textvariable=self.vars["nomeproduto"])
        ent_nome.grid(row=0, column=5, sticky="ew", padx=(0, 10), pady=6)
        self.entries["nomeproduto"] = ent_nome

        self._add_field(form, 1, 0, "quantidade", width=14)
        self._add_field(form, 1, 2, "chapa", width=22)
        self._add_field(form, 1, 4, "material", width=22)

        # LISTA
        lst_outer = ttk.Frame(self, padding=(10, 0, 10, 10))
        lst_outer.grid(row=4, column=0, sticky="nsew")
        lst_outer.rowconfigure(0, weight=1)
        lst_outer.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            lst_outer, columns=TREE_COLS, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst_outer, orient="vertical",
                            command=self.tree.yview)
        hsb = ttk.Scrollbar(lst_outer, orient="horizontal",
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

    def _add_field(self, parent: ttk.Frame, row: int, col: int, key: str, width: int | None = None) -> None:
        label = dict(CAMPOS)[key]
        ttk.Label(parent, text=f"{label}:").grid(
            row=row, column=col, sticky="w", padx=(10, 6), pady=6)

        ent = ttk.Entry(parent, textvariable=self.vars[key])
        if width is not None:
            ent.configure(width=width)

        ent.grid(row=row, column=col + 1, sticky="ew", padx=(0, 10), pady=6)
        self.entries[key] = ent

    def _aplicar_permissoes(self) -> None:
        n = int(self.acesso.nivel or 0)

        if n <= 0:
            # sem acesso -> trava tudo (mas main já deve bloquear)
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_buscar_prod.configure(state="disabled")
            self.btn_preencher.configure(state="disabled")
            self.btn_limpar.configure(state="disabled")
            for ent in self.entries.values():
                ent.configure(state="readonly")
            return

        if n == 1:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_buscar_prod.configure(state="disabled")
            self.btn_preencher.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            for ent in self.entries.values():
                ent.configure(state="readonly")
            return

        if n == 2:
            self.btn_novo.configure(state="normal")
            self.btn_salvar.configure(state="normal")
            self.btn_excluir.configure(state="disabled")
            self.btn_buscar_prod.configure(state="normal")
            self.btn_preencher.configure(state="normal")
            self.btn_limpar.configure(state="normal")
            for ent in self.entries.values():
                ent.configure(state="normal")
            return

        # n >= 3
        self.btn_novo.configure(state="normal")
        self.btn_salvar.configure(state="normal")
        self.btn_excluir.configure(state="normal")
        self.btn_buscar_prod.configure(state="normal")
        self.btn_preencher.configure(state="normal")
        self.btn_limpar.configure(state="normal")
        for ent in self.entries.values():
            ent.configure(state="normal")

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
            values = [a.sku, a.nomeproduto or "", str(
                a.quantidade), a.chapa or "", a.material or ""]
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
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para criar (somente leitura).")
            return

        self.limpar_form()
        self.vars["quantidade"].set("0")

        try:
            novo_sku = self.service.proximo_sku()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar novo SKU:\n{e}")
            return

        self.vars["sku"].set(novo_sku)
        self._sku_original = None
        self.ent_sku.focus_set()
        self.ent_sku.selection_range(0, tk.END)

    def limpar_form(self) -> None:
        for k in self.vars:
            self.vars[k].set("")
        self._sku_original = None
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para salvar (somente leitura).")
            return

        form = {k: self.vars[k].get() for k, _ in CAMPOS}

        try:
            status = self.service.salvar_from_form(
                form, sku_original=self._sku_original)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        messagebox.showinfo("OK", f"Arranjo {status} com sucesso.")
        self.atualizar_lista()
        self._sku_original = (self.vars["sku"].get().strip() or None)

    def excluir(self) -> None:
        if not self._can_delete():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para excluir (somente admin).")
            return

        sku_form = self.vars["sku"].get().strip()
        sku_target = (self._sku_original or sku_form).strip()
        if not sku_target:
            messagebox.showwarning(
                "Atenção", "Informe/Selecione um SKU para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Arranjo do SKU {sku_target}?"):
            return

        try:
            self.service.excluir(sku_target)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Arranjo excluído.")
        self.limpar_form()
        self.atualizar_lista()

    def buscar_produto_popup(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Somente leitura: não é permitido usar a busca para preencher.")
            return

        def on_pick(sku: str, nome: str) -> None:
            self.vars["sku"].set(sku)
            self.vars["nomeproduto"].set(nome)
            if not self.vars["quantidade"].get().strip():
                self.vars["quantidade"].set("0")

        ProdutoPicker(self.winfo_toplevel(), self.service, on_pick)

    def preencher_nome_por_sku(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Somente leitura: não é permitido preencher automaticamente.")
            return

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

    def _voltar_ou_fechar(self) -> None:
        if self.from_menu:
            self.winfo_toplevel().destroy()
            return
        try:
            abrir_menu_principal_skip_entrada()
        finally:
            self.winfo_toplevel().destroy()


# ============================================================
# STARTUP
# ============================================================

def test_connection_or_die(cfg: AppConfig) -> None:
    conn = None
    try:
        conn = db_connect(cfg)
        with conn.cursor() as cur:
            cur.execute("SELECT 1")
            cur.fetchone()
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def _detect_from_menu_flag() -> bool:
    if "--standalone" in sys.argv:
        return False
    if any(x in sys.argv for x in ("--usuario-id", "--uid", "--user-id", "--from-menu", "--skip-entrada")):
        return True
    env = (os.getenv("EKENOX_FROM_MENU") or "").strip().lower()
    return env in {"1", "true", "yes", "sim", "s"}


def main() -> None:
    cfg = env_override(load_config())

    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--from-menu", action="store_true")
    ap.add_argument("--standalone", action="store_true")
    ap.add_argument("--reopen-menu-on-exit", action="store_true")

    # ✅ aliases compatíveis com menu e Depósito
    ap.add_argument("--user-id", "--usuario-id", "--uid", dest="user_id")
    ap.add_argument("--user-hash", "--usuario-hash", dest="user_hash")
    ap.add_argument("--session-file", dest="session_file")

    ap.add_argument("--nivel")
    ap.add_argument(f"--nivel-{PROGRAMA_CODIGO.lower()}", dest="nivel_prog")
    ns, _ = ap.parse_known_args()

    from_menu = bool(ns.from_menu) or _detect_from_menu_flag()
    reopen_menu_on_exit = bool(ns.reopen_menu_on_exit) and (not from_menu)

    # Teste de conexão
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

    acesso = _build_access(cfg, ns)

    # Se usuário resolvido mas inativo => nega
    if int(acesso.nivel or 0) <= 0:
        root = tk.Tk()
        root.withdraw()
        try:
            messagebox.showerror(
                "Acesso negado", acesso.aviso or "Sem acesso.")
        finally:
            try:
                root.destroy()
            except Exception:
                pass

        if reopen_menu_on_exit:
            abrir_menu_principal_skip_entrada()
        return

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

    db = Database(cfg)
    repo = ArranjoRepo(db)
    service = ArranjoService(repo)

    tela = TelaArranjo(root, service, acesso, from_menu=from_menu)
    tela.pack(fill="both", expand=True)

    def on_close():
        try:
            root.destroy()
        except Exception:
            pass
        if reopen_menu_on_exit and (not from_menu):
            abrir_menu_principal_skip_entrada()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_arranjo(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_arranjo:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
