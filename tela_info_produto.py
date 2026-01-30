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
from typing import Any, Dict, List, Optional, Tuple

import psycopg2


# ============================================================
# PROGRAMA / PERMISSÕES
# ============================================================

# 🔧 AJUSTE para bater com Ekenox.programas.codigo
PROGRAMA_CODIGO = "INFO_PRODUTO"

NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Admin",
}


# ============================================================
# TABELAS
# ============================================================

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


def log_info_produto(msg: str) -> None:
    _log_write("tela_info_produto.log", msg)


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


def menu_ja_rodando(menu_path: Optional[str] = None) -> bool:
    """
    Evita duplicar Menu.
    Windows: detecta menu.exe (tasklist) ou menu_principal.py via commandline (PowerShell).
    """
    if os.name != "nt":
        return False
    try:
        menu_path = menu_path or localizar_menu_principal()
        if not menu_path:
            return False
        menu_base = os.path.basename(menu_path).lower()

        if menu_base.endswith(".exe"):
            out = subprocess.check_output(
                ["tasklist"], text=True, errors="ignore")
            return menu_base in out.lower()

        ps = r"""
        # 1) tenta achar pelo CommandLine (python rodando menu)
        $p1 = Get-CimInstance Win32_Process |
        Where-Object { $_.CommandLine -and ($_.CommandLine -match 'menu_principal\.py|Menu_Principal\.py|MenuPrincipal\.py|menu\.py') } |
        Select-Object -First 1;

        # 2) tenta achar pela janela (menu.exe / atalho / etc)
        $p2 = Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -and ($_.MainWindowTitle -match 'Menu Principal\s*-\s*Ekenox') } |
        Select-Object -First 1;

        if ($p1 -or $p2) { 'FOUND' } else { '' }
        """
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps], text=True, errors="ignore")
        return "FOUND" in out
    except Exception:
        return False


def abrir_menu_principal_skip_entrada() -> None:
    menu_path = localizar_menu_principal()
    if not menu_path:
        log_info_produto(
            f"MENU: não encontrado. APP_DIR={APP_DIR} BASE_DIR={BASE_DIR}")
        return

    # ✅ não duplica menu
    if menu_ja_rodando(menu_path):
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
        log_info_produto(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_info_produto(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


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
                        if str(kk).strip().lower() in {"id", "user_id", "usuario_id", "usuarioid", "userid"}:
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
    for envk in ("EKENOX_USER_ID", "USER_ID", "USUARIO_ID", "LOGGED_USER_ID"):
        v = (os.getenv(envk) or "").strip()
        if v.isdigit():
            return int(v)

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
    if user_id_raw:
        s = str(user_id_raw).strip()
        if s.isdigit():
            uid = int(s)
            log_info_produto(f"RESOLVE: user_id via argumento/env = {uid}")
            return uid

    if not user_hash:
        return None

    cols = _usuarios_cols(cfg)
    id_col = cols["id_col"]

    conn = None
    try:
        conn = db_connect(cfg)
        with conn.cursor() as cur:
            hash_cols_try = ["hash", "email_hash",
                             "hash_email", "hashEmail", "emailHash"]
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
                        log_info_produto(
                            f"RESOLVE: user_id={uid} via hash coluna={hc}")
                        return uid
                except Exception:
                    continue

    except Exception as e:
        log_info_produto(f"RESOLVE: erro via hash: {type(e).__name__}: {e}")
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
            log_info_produto(
                "ACESSO: tabela programas/programa não encontrada.")
            return None

        up_table = "usuario_programa"
        if not _table_exists(cur, "Ekenox", up_table):
            log_info_produto("ACESSO: tabela usuario_programa não encontrada.")
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

        # ✅ coluna "permitido" (pode variar o nome)
        up_perm_col = _pick_col(
            up_cols, "permitido", "permissao", "allowed", "acesso", "ativo", "habilitado")

        if not (prog_id_col and prog_code_col and up_user_col and up_prog_col and up_nivel_col):
            log_info_produto(
                f"ACESSO: colunas não resolvidas. prog_id={prog_id_col} prog_code={prog_code_col} "
                f"up_user={up_user_col} up_prog={up_prog_col} up_nivel={up_nivel_col}"
            )
            return None

        # Seleciona nível e, se existir, permitido
        select_cols = [f'up."{up_nivel_col}"']
        if up_perm_col:
            select_cols.append(f'up."{up_perm_col}"')

        sql = f"""
            SELECT {", ".join(select_cols)}
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
            log_info_produto(
                f"ACESSO: sem registro em usuario_programa. user_id={user_id} programa={programa_codigo}")
            return 0

        nivel_val = row[0]
        # se não existe col, assume True
        perm_val = row[1] if (up_perm_col and len(row) > 1) else True

        nivel_int = int(nivel_val or 0)
        permitido = _bool_from_db(perm_val)

        log_info_produto(
            f"ACESSO: user_id={user_id} programa={programa_codigo} "
            f"nivel={nivel_int} permitido_col={up_perm_col} permitido_val={perm_val!r} permitido={permitido}"
        )

        # ✅ BLOQUEIA se "permitido" for falso
        if up_perm_col and (not permitido):
            return 0

        return nivel_int

    except Exception as e:
        log_info_produto(
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
            "Acesso negado: usuário não informado ao abrir a tela.\n\n"
            "Esta tela exige identificação do usuário para validar permissões.\n"
            "Chame com --user-id/--usuario-id/--uid <id> ou forneça sessão (sessao.json/login.json etc)."
        )
        return SessaoAcesso(
            nivel=0,  # <- AGORA bloqueia
            origem="sem_usuario",
            usuario_id=None,
            usuario_nome=None,
            programa=PROGRAMA_CODIGO,
            aviso=aviso,
        )

    uid = int(user_id)
    nome = fetch_user_nome(cfg, uid) or None

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

    if nivel_db is None:
        return SessaoAcesso(
            nivel=0,
            origem="erro_permissao",
            usuario_id=uid,
            usuario_nome=nome,
            programa=PROGRAMA_CODIGO,
            aviso="Acesso negado: não foi possível validar a permissão no banco.",
        )

    if int(nivel_db) <= 0:
        return SessaoAcesso(
            nivel=0,  # <- AGORA bloqueia
            origem="sem_permissao",
            usuario_id=uid,
            usuario_nome=nome,
            programa=PROGRAMA_CODIGO,
            aviso=(
                "Acesso negado: usuário sem permissão para este programa.\n\n"
                f"Usuário ID: {uid}\nPrograma: {PROGRAMA_CODIGO}\n"
            ),
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
# PRODUTOS: METADADOS / CONVERSÕES
# ============================================================

def _to_decimal(v: Any) -> Decimal:
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
        raise ValueError(f"Decimal inválido: {v!r}")


def _bool_from_db(v: Any) -> bool:
    """
    Converte valores comuns do banco para bool:
    True/False, 1/0, 'Sim'/'Não', 'S'/'N', 't'/'f', etc.
    """
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return int(v) != 0
    s = str(v).strip().lower()
    return s in {"1", "true", "t", "yes", "sim", "s", "y", "on"}


def _parse_bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    s = str(v or "").strip().lower()
    return s in {"1", "true", "t", "yes", "sim", "s", "y", "on"}


def _produtos_schema(cfg: AppConfig) -> dict[str, dict[str, str]]:
    """
    Retorna dict com chaves em lower():
      {
        "nomedacoluna": {"name": "NomeReal", "data_type": "...", "udt_name": "..."}
      }
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT column_name, data_type, udt_name
                FROM information_schema.columns
                WHERE table_schema='Ekenox' AND table_name='produtos'
                ORDER BY ordinal_position
                """
            )
            rows = cur.fetchall()
    finally:
        try:
            conn.close()
        except Exception:
            pass

    out: dict[str, dict[str, str]] = {}
    for col, dt, udt in rows:
        out[str(col).lower()] = {"name": str(col),
                                 "data_type": str(dt), "udt_name": str(udt)}
    return out


def _convert_for_db(raw: str, meta: dict[str, str]) -> Any:
    s = (raw or "").strip()
    if s == "":
        return None

    dt = (meta.get("data_type") or "").lower()
    udt = (meta.get("udt_name") or "").lower()

    if dt == "boolean":
        return _parse_bool(s)

    # inteiros
    if dt in {"integer", "bigint", "smallint"} or udt in {"int2", "int4", "int8"}:
        try:
            return int(s)
        except Exception:
            raise ValueError(f"Valor inteiro inválido: {s!r}")

    # numericos
    if dt in {"numeric", "decimal"} or udt in {"numeric"}:
        return _to_decimal(s)

    if dt in {"double precision", "real"} or udt in {"float4", "float8"}:
        try:
            return float(s.replace(",", "."))
        except Exception:
            raise ValueError(f"Valor numérico inválido: {s!r}")

    # default: texto
    return s


def _format_from_db(v: Any, meta: dict[str, str]) -> str:
    if v is None:
        return ""
    dt = (meta.get("data_type") or "").lower()
    if dt == "boolean":
        return "1" if bool(v) else "0"
    return str(v)


# ============================================================
# REPOSITORY / SERVICE (INFO PRODUTO)
# ============================================================

class ProdutoRepo:
    def __init__(self, db: Database, schema: dict[str, dict[str, str]]) -> None:
        self.db = db
        self.schema = schema

        # resolve nomes reais (case-insensitive)
        self.col_sku = self._pick_col("sku") or "sku"
        self.col_nome = self._pick_any("nomeproduto", "nomeProduto", "nome") or self._pick_col(
            "nomeproduto") or "nomeProduto"

    def _pick_col(self, name: str) -> Optional[str]:
        m = self.schema.get(str(name).lower())
        return m["name"] if m else None

    def _pick_any(self, *names: str) -> Optional[str]:
        for n in names:
            m = self.schema.get(str(n).lower())
            if m:
                return m["name"]
        return None

    def listar(self, termo: Optional[str] = None, limit: int = 600) -> List[dict[str, Any]]:
        like = f"%{termo}%" if termo else None

        # colunas para busca (texto)
        text_cols = []
        for k, meta in self.schema.items():
            if (meta.get("data_type") or "").lower() in {"character varying", "text", "character"}:
                text_cols.append(meta["name"])

        # sempre inclui sku/nome se existirem
        for must in (self.col_sku, self.col_nome):
            if must and must not in text_cols:
                text_cols.insert(0, must)

        where_parts = ["(%s IS NULL)"]
        params: List[Any] = [termo]

        for c in text_cols[:6]:  # limita para não ficar pesado
            where_parts.append(f'(COALESCE(p."{c}", \'\') ILIKE %s)')
            params.append(like)

        sql = f"""
            SELECT p.*
            FROM {PRODUTOS_TABLE} p
            WHERE {" OR ".join(where_parts)}
            ORDER BY p."{self.col_nome}" NULLS LAST, p."{self.col_sku}"
            LIMIT %s
        """
        params.append(limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, tuple(params))
            rows = self.db.cursor.fetchall()

            # pega nome das colunas na ordem do SELECT p.*
            # type: ignore[union-attr]
            colnames = [d[0] for d in self.db.cursor.description]
            out = []
            for r in rows:
                out.append({str(colnames[i]): r[i]
                           for i in range(len(colnames))})
            return out
        finally:
            self.db.desconectar()

    def exists(self, sku: str) -> bool:
        sql = f'SELECT 1 FROM {PRODUTOS_TABLE} WHERE "{self.col_sku}"=%s'
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
            SELECT COALESCE(MAX(CAST("{self.col_sku}" AS BIGINT)), 0)
            FROM {PRODUTOS_TABLE}
            WHERE "{self.col_sku}" ~ '^[0-9]+$'
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

    def inserir(self, values: dict[str, Any]) -> None:
        cols = list(values.keys())
        placeholders = ", ".join(["%s"] * len(cols))
        cols_sql = ", ".join([f'"{c}"' for c in cols])
        sql = f'INSERT INTO {PRODUTOS_TABLE} ({cols_sql}) VALUES ({placeholders})'
        params = [values[c] for c in cols]

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

    def atualizar(self, values: dict[str, Any], sku_original: str) -> None:
        # não atualiza a chave pelo update
        values = dict(values)
        values.pop(self.col_sku, None)

        sets = ", ".join([f'"{c}"=%s' for c in values.keys()])
        params = [values[c] for c in values.keys()]
        params.append(sku_original)

        sql = f'UPDATE {PRODUTOS_TABLE} SET {sets} WHERE "{self.col_sku}"=%s'

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
        sql = f'DELETE FROM {PRODUTOS_TABLE} WHERE "{self.col_sku}"=%s'
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


class ProdutoService:
    def __init__(self, repo: ProdutoRepo, schema: dict[str, dict[str, str]]) -> None:
        self.repo = repo
        self.schema = schema

    def listar(self, termo: Optional[str]) -> List[dict[str, Any]]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def proximo_sku(self) -> str:
        return self.repo.proximo_sku_numerico()

    def salvar(self, form: dict[str, str], sku_original: Optional[str]) -> str:
        sku = (form.get(self.repo.col_sku) or "").strip()
        if not sku:
            raise ValueError("SKU é obrigatório.")

        values: dict[str, Any] = {}
        for col, raw in form.items():
            meta = self.schema.get(str(col).lower())
            if not meta:
                continue
            values[col] = _convert_for_db(raw, meta)

        # garante sku no insert
        values[self.repo.col_sku] = sku

        if sku_original and self.repo.exists(sku_original):
            self.repo.atualizar(values, sku_original=sku_original)
            return "atualizado"

        if self.repo.exists(sku):
            raise ValueError(
                "SKU já existe. Selecione na lista para editar ou clique em Novo.")

        self.repo.inserir(values)
        return "inserido"

    def excluir(self, sku: str) -> None:
        self.repo.excluir(sku)


# ============================================================
# UI (INFO PRODUTO)
# ============================================================

DEFAULT_GEOMETRY = "1200x720"
APP_TITLE = "Tela de Info Produto"


def _labelize(col: str) -> str:
    # rótulo simples e amigável
    s = str(col or "")
    if not s:
        return s
    return s[0].upper() + s[1:]


class TelaInfoProduto(ttk.Frame):
    def __init__(self, master: tk.Misc, service: ProdutoService, acesso: SessaoAcesso, *,
                 from_menu: bool, schema: dict[str, dict[str, str]], repo: ProdutoRepo):
        super().__init__(master)
        self.service = service
        self.acesso = acesso
        self.from_menu = bool(from_menu)
        self.schema = schema
        self.repo = repo

        # campos exibidos (auto)
        # prioriza alguns comuns; completa com os primeiros da tabela
        preferred = [
            repo.col_sku,
            repo.col_nome,
            "descricao", "descricaoProduto", "unidade", "grupo", "subgrupo",
            "ncm", "ean", "codBarras", "ativo",
        ]

        all_cols = [meta["name"] for meta in schema.values()]
        chosen: List[str] = []

        for p in preferred:
            if not p:
                continue
            real = schema.get(str(p).lower(), {}).get("name")
            if real and real in all_cols and real not in chosen:
                chosen.append(real)

        for c in all_cols:
            if c not in chosen:
                chosen.append(c)
            if len(chosen) >= 12:  # limita para não virar uma tela enorme
                break

        self.form_cols = chosen
        self.tree_cols = [repo.col_sku, repo.col_nome]
        for extra in ("unidade", "grupo", "ativo"):
            real = schema.get(extra.lower(), {}).get("name")
            if real and real not in self.tree_cols and len(self.tree_cols) < 5:
                self.tree_cols.append(real)

        # vars
        self.var_filtro = tk.StringVar()
        self._sku_original: Optional[str] = None

        self.vars: dict[str, tk.StringVar] = {}
        for c in self.form_cols:
            self.vars[c] = tk.StringVar()

        self.entries: dict[str, ttk.Entry] = {}

        self._build_ui()
        self._aplicar_permissoes()
        self.atualizar_lista()

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

        # TOPBAR
        topbar = ttk.Frame(self, padding=(10, 10, 10, 6))
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.columnconfigure(0, weight=1)

        nome = self.acesso.usuario_nome or (
            "Não informado" if not self.acesso.usuario_id else f"ID {self.acesso.usuario_id}")
        nivel_txt = NIVEL_LABEL.get(
            int(self.acesso.nivel or 0), str(self.acesso.nivel))

        ttk.Label(
            topbar,
            text=f"Logado: {nome} | Nível: {nivel_txt}",
            foreground=("green" if self._can_edit() else "gray"),
            font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            topbar,
            text=("Fechar" if self.from_menu else "Voltar ao Menu"),
            command=self._voltar_ou_fechar,
        ).grid(row=0, column=1, sticky="e")

        # BUSCA
        top = ttk.Frame(self, padding=(10, 0, 10, 6))
        top.grid(row=1, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (SKU / Nome / Texto):").grid(row=0,
                                                                 column=0, sticky="w")
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
        form = ttk.LabelFrame(self, text="Produto", padding=(10, 6, 10, 10))
        form.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))

        # grid 4 colunas de label+entry (8 colunas totais)
        for c in range(8):
            form.columnconfigure(c, weight=1)

        # distribui campos em 2 linhas (4 campos por linha)
        fields_per_row = 4
        for idx, col in enumerate(self.form_cols):
            r = idx // fields_per_row
            pos = idx % fields_per_row
            col_label = pos * 2
            col_entry = pos * 2 + 1

            ttk.Label(form, text=f"{_labelize(col)}:").grid(
                row=r, column=col_label, sticky="w", padx=(10, 6), pady=6)

            meta = self.schema.get(col.lower(), {})
            if (meta.get("data_type") or "").lower() == "boolean":
                # bool simples (0/1) em entry, para evitar complicações de widget
                ent = ttk.Entry(form, textvariable=self.vars[col], width=10)
                ent.grid(row=r, column=col_entry,
                         sticky="ew", padx=(0, 10), pady=6)
            else:
                ent = ttk.Entry(form, textvariable=self.vars[col])
                ent.grid(row=r, column=col_entry,
                         sticky="ew", padx=(0, 10), pady=6)

            self.entries[col] = ent

        # LISTA
        lst_outer = ttk.Frame(self, padding=(10, 0, 10, 10))
        lst_outer.grid(row=4, column=0, sticky="nsew")
        lst_outer.rowconfigure(0, weight=1)
        lst_outer.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            lst_outer, columns=self.tree_cols, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst_outer, orient="vertical",
                            command=self.tree.yview)
        hsb = ttk.Scrollbar(lst_outer, orient="horizontal",
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for c in self.tree_cols:
            self.tree.heading(c, text=_labelize(c))
            self.tree.column(c, width=220, anchor="w", stretch=True)

        # sku menor
        if self.repo.col_sku in self.tree_cols:
            self.tree.column(self.repo.col_sku, width=140, stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def _aplicar_permissoes(self) -> None:
        n = int(self.acesso.nivel or 0)

        if n <= 0:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="disabled")
            for ent in self.entries.values():
                ent.configure(state="readonly")
            return

        if n == 1:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            for ent in self.entries.values():
                ent.configure(state="readonly")
            return

        if n == 2:
            self.btn_novo.configure(state="normal")
            self.btn_salvar.configure(state="normal")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            for ent in self.entries.values():
                ent.configure(state="normal")
            return

        self.btn_novo.configure(state="normal")
        self.btn_salvar.configure(state="normal")
        self.btn_excluir.configure(state="normal")
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
            messagebox.showerror("Erro", f"Falha ao listar produtos:\n{e}")
            return

        for row in itens:
            values = []
            for c in self.tree_cols:
                meta = self.schema.get(c.lower(), {})
                values.append(_format_from_db(row.get(c), meta))
            self.tree.insert("", "end", values=values)

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals:
            return

        # pega sku do tree
        sku = ""
        if self.repo.col_sku in self.tree_cols:
            sku = str(vals[self.tree_cols.index(
                self.repo.col_sku)] or "").strip()

        if sku:
            # preenche form procurando na lista (já veio do listar p.*)
            # melhor: refazer busca local com base no selection do tree:
            # aqui fazemos refresh do form usando a linha atual do tree + limpar demais
            # mas como tree não contém todas colunas, vamos achar no listar novamente
            # (custo ok pois limit e local)
            termo = sku
        else:
            termo = None

        try:
            rows = self.service.listar(termo)
        except Exception:
            rows = []

        # tenta encontrar sku exato
        found = None
        if sku:
            for r in rows:
                if str(r.get(self.repo.col_sku) or "").strip() == sku:
                    found = r
                    break
        if not found and rows:
            found = rows[0]

        if not found:
            return

        for c in self.form_cols:
            meta = self.schema.get(c.lower(), {})
            self.vars[c].set(_format_from_db(found.get(c), meta))

        self._sku_original = str(
            found.get(self.repo.col_sku) or "").strip() or None

    def novo(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para criar (somente leitura).")
            return

        self.limpar_form()
        try:
            novo_sku = self.service.proximo_sku()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar novo SKU:\n{e}")
            return

        self.vars[self.repo.col_sku].set(novo_sku)
        self._sku_original = None
        self.entries[self.repo.col_sku].focus_set()
        self.entries[self.repo.col_sku].selection_range(0, tk.END)

    def limpar_form(self) -> None:
        for c in self.form_cols:
            self.vars[c].set("")
        self._sku_original = None
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para salvar (somente leitura).")
            return

        form: dict[str, str] = {c: self.vars[c].get() for c in self.form_cols}

        try:
            status = self.service.salvar(form, sku_original=self._sku_original)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        messagebox.showinfo("OK", f"Produto {status} com sucesso.")
        self.atualizar_lista()
        self._sku_original = (
            self.vars[self.repo.col_sku].get().strip() or None)

    def excluir(self) -> None:
        if not self._can_delete():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para excluir (somente admin).")
            return

        sku_form = self.vars[self.repo.col_sku].get().strip()
        sku_target = (self._sku_original or sku_form).strip()
        if not sku_target:
            messagebox.showwarning(
                "Atenção", "Informe/Selecione um SKU para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Produto do SKU {sku_target}?"):
            return

        try:
            self.service.excluir(sku_target)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Produto excluído.")
        self.limpar_form()
        self.atualizar_lista()

    def _voltar_ou_fechar(self) -> None:
        # ✅ Se o menu já está rodando, NÃO abre outro. Só fecha esta tela.
        try:
            if menu_ja_rodando():
                self.winfo_toplevel().destroy()
                return
        except Exception:
            # se falhar a detecção, ainda assim não duplicar (mais seguro)
            self.winfo_toplevel().destroy()
            return

        # Se NÃO tem menu rodando (ex.: abriu standalone), aí sim abre menu
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
    # Regra: só considera "veio do menu" se o Menu avisar explicitamente
    if "--standalone" in sys.argv:
        return False
    if "--from-menu" in sys.argv:
        return True

    env = (os.getenv("EKENOX_FROM_MENU") or "").strip().lower()
    return env in {"1", "true", "yes", "sim", "s"}


def main() -> None:
    cfg = env_override(load_config())

    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--from-menu", action="store_true")
    ap.add_argument("--standalone", action="store_true")
    ap.add_argument("--reopen-menu-on-exit", action="store_true")

    ap.add_argument("--user-id", "--usuario-id", "--uid", dest="user_id")
    ap.add_argument("--user-hash", "--usuario-hash", dest="user_hash")
    ap.add_argument("--session-file", dest="session_file")

    ap.add_argument("--nivel")
    ap.add_argument(f"--nivel-{PROGRAMA_CODIGO.lower()}", dest="nivel_prog")
    ns, _ = ap.parse_known_args()
    log_info_produto(f"ARGV: {sys.argv!r}")
    log_info_produto(
        f"FLAGS: from_menu_arg={bool(ns.from_menu)} env_from_menu={os.getenv('EKENOX_FROM_MENU')!r}")
    log_info_produto(
        f"ENV_USER: EKENOX_USER_ID={os.getenv('EKENOX_USER_ID')!r} USER_ID={os.getenv('USER_ID')!r} "
        f"USUARIO_ID={os.getenv('USUARIO_ID')!r} LOGGED_USER_ID={os.getenv('LOGGED_USER_ID')!r}"
    )

    from_menu = bool(ns.from_menu) or _detect_from_menu_flag()
    reopen_menu_on_exit = bool(ns.reopen_menu_on_exit) and (not from_menu)

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

        if reopen_menu_on_exit and (not from_menu):
            abrir_menu_principal_skip_entrada()
        return

    schema = _produtos_schema(cfg)

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
    repo = ProdutoRepo(db, schema)
    service = ProdutoService(repo, schema)

    tela = TelaInfoProduto(root, service, acesso,
                           from_menu=from_menu, schema=schema, repo=repo)
    tela.pack(fill="both", expand=True)

    def on_close():
        try:
            root.destroy()
        except Exception:
            pass

        # ✅ só reabre menu em standalone, e não duplica se já estiver rodando
        if reopen_menu_on_exit and (not from_menu):
            abrir_menu_principal_skip_entrada()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_info_produto(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_info_produto:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
