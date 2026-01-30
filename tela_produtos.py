from __future__ import annotations

"""
tela_estoque.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Compatível com tela_produtos (menu/fechar/permissões):
- Valida usuário + nível por programa (Ekenox.usuario_programa + Ekenox.programas)
- Respeita coluna Permitido (se existir) -> se "Não", bloqueia (nível 0)
- Se usuário não for identificado, bloqueia (nível 0)

✅ Estilo tela_produtos/tela_arranjo:
- Fechar no X nunca reabre menu automaticamente
- Botão vira "Fechar" quando veio do menu (ou menu já está rodando)
- Botão vira "Voltar ao Menu" quando abriu fora do menu (abre menu e fecha)
- Evita duplicar menu (checagem só no START, não no clique -> clique rápido)

Logs:
  <BASE_DIR>/logs/tela_estoque.log
  <BASE_DIR>/logs/menu_principal_run.log
"""

import argparse
import glob
import json
import os
import shutil
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

PROGRAMA_CODIGO = "ESTOQUE"

NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Admin",
}


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


def log_estoque(msg: str) -> None:
    _log_write("tela_estoque.log", msg)


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
# MENU PRINCIPAL (igual tela_produtos: rápido no botão)
# ============================================================

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
    pastas = [APP_DIR, BASE_DIR]
    this_file = _this_script_abspath()

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
            try:
                if this_file and os.path.samefile(p, this_file):
                    continue
            except Exception:
                pass
            return p
    return None


def _pick_python_launcher_windows() -> list[str]:
    """
    FAST: não testa executando subprocess. Só verifica no PATH.
    """
    if os.name != "nt":
        return [sys.executable]

    if shutil.which("pyw"):
        return ["pyw", "-3.12"]
    if shutil.which("pythonw"):
        return ["pythonw"]
    if shutil.which("python"):
        return ["python"]
    return [sys.executable]


def menu_ja_rodando(menu_path: Optional[str] = None) -> bool:
    """
    Evita duplicar Menu.
    OBS: Pode ser "pesado" (PowerShell), por isso usamos só no START,
         nunca no clique do botão.
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
        $p1 = Get-CimInstance Win32_Process |
          Where-Object { $_.CommandLine -and ($_.CommandLine -match 'menu_principal\.py|Menu_Principal\.py|MenuPrincipal\.py|menu\.py') } |
          Select-Object -First 1;

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


def abrir_menu_principal_sem_check(menu_path: str, launcher: Optional[list[str]] = None) -> bool:
    """
    Abre o menu SEM checar se já está rodando (checagem feita no START).
    Mantém o clique rápido.
    """
    try:
        if not menu_path or not os.path.isfile(menu_path):
            return False

        cwd = os.path.dirname(menu_path) or APP_DIR
        is_exe = menu_path.lower().endswith(".exe")
        is_py = menu_path.lower().endswith(".py")

        if is_exe:
            cmd = [menu_path, "--skip-entrada"]
        elif is_py:
            if os.name == "nt":
                launch = launcher or ["pythonw"]
                cmd = launch + [menu_path, "--skip-entrada"]
            else:
                cmd = [sys.executable, menu_path, "--skip-entrada"]
        else:
            cmd = [menu_path, "--skip-entrada"]

        popen_kwargs: dict = {"cwd": cwd}
        if os.name == "nt":
            popen_kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        else:
            popen_kwargs["start_new_session"] = True

        child_log_path = os.path.join(
            BASE_DIR, "logs", "menu_principal_run.log")
        os.makedirs(os.path.dirname(child_log_path), exist_ok=True)
        with open(child_log_path, "a", encoding="utf-8") as out:
            out.write("\n\n=== START MENU PROCESS (estoque no-check) ===\n")
            out.write(f"cwd={cwd}\n")
            out.write(f"cmd={cmd}\n")
            out.flush()
            subprocess.Popen(
                cmd,
                stdout=out,
                stderr=out,
                close_fds=(os.name != "nt"),
                **popen_kwargs,
            )

        return True
    except Exception as e:
        log_menu_run(
            f"ERRO abrir_menu_principal_sem_check: {type(e).__name__}: {e}")
        return False


# ============================================================
# SESSÃO / USUÁRIO / PERMISSÕES (copiado do padrão tela_produtos)
# ============================================================

def _bool_from_db(v: Any) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return int(v) != 0
    s = str(v).strip().lower()
    return s in {"1", "true", "t", "yes", "sim", "s", "y", "on"}


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
        key_low = {str(k).strip().lower(): k for k in obj.keys()}
        candidate_lows = {x.lower() for x in key_candidates}
        for klow, original_k in key_low.items():
            if klow in candidate_lows:
                v = obj.get(original_k)
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

            if ("usuario" in klow or "user" in klow) and klow.endswith("id"):
                got = as_int(obj.get(original_k))
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


def _candidate_session_dirs() -> list[str]:
    dirs: list[str] = []
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

    patterns: list[str] = []
    for d in search_dirs:
        patterns += [
            os.path.join(d, "*sess*.json"),
            os.path.join(d, "*login*.json"),
            os.path.join(d, "*usuario*.json"),
            os.path.join(d, "*auth*.json"),
            os.path.join(d, "*entrada*.json"),
            os.path.join(d, "*user*.json"),
        ]

    files: list[str] = []
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


def fetch_user_nome(cfg: AppConfig, usuario_id: int) -> str:
    cols = _usuarios_cols(cfg)
    id_col = cols["id_col"]
    nome_col = cols["nome_col"]

    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            sql = f'SELECT COALESCE(u."{nome_col}",\'\') FROM "Ekenox"."usuarios" u WHERE u."{id_col}"=%s LIMIT 1'
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
            return str(r[0] or "").strip() if r else ""
    except Exception:
        return ""
    finally:
        try:
            conn.close()
        except Exception:
            pass


def user_esta_ativo(cfg: AppConfig, usuario_id: int) -> bool:
    cols = _usuarios_cols(cfg)
    id_col = cols["id_col"]
    ativo_col = cols["ativo_col"]

    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            sql = f'SELECT COALESCE(u."{ativo_col}", true) FROM "Ekenox"."usuarios" u WHERE u."{id_col}"=%s LIMIT 1'
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
            return bool(r[0]) if r else False
    except Exception:
        return False
    finally:
        try:
            conn.close()
        except Exception:
            pass


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
            log_estoque("ACESSO: tabela programas/programa não encontrada.")
            return None

        up_table = "usuario_programa"
        if not _table_exists(cur, "Ekenox", up_table):
            log_estoque("ACESSO: tabela usuario_programa não encontrada.")
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

        up_perm_col = _pick_col(
            up_cols, "permitido", "permissao", "allowed", "acesso", "ativo", "habilitado")

        if not (prog_id_col and prog_code_col and up_user_col and up_prog_col and up_nivel_col):
            log_estoque(
                f"ACESSO: colunas não resolvidas. prog_id={prog_id_col} prog_code={prog_code_col} "
                f"up_user={up_user_col} up_prog={up_prog_col} up_nivel={up_nivel_col}"
            )
            return None

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
            log_estoque(
                f"ACESSO: sem registro em usuario_programa. user_id={user_id} programa={programa_codigo}")
            return 0

        nivel_int = int(row[0] or 0)
        permitido_val = row[1] if (up_perm_col and len(row) > 1) else True
        permitido = _bool_from_db(permitido_val)

        log_estoque(
            f"ACESSO: user_id={user_id} programa={programa_codigo} "
            f"nivel={nivel_int} permitido_col={up_perm_col} permitido_val={permitido_val!r} permitido={permitido}"
        )

        if up_perm_col and (not permitido):
            return 0

        return nivel_int

    except Exception as e:
        log_estoque(
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

    session_file = (getattr(ns, "session_file", None) or os.getenv(
        "EKENOX_SESSION_FILE") or "").strip() or None

    user_id: Optional[int] = None
    if user_id_raw and user_id_raw.isdigit():
        user_id = int(user_id_raw)
        log_estoque(f"RESOLVE: user_id via argumento/env = {user_id}")
    else:
        user_id = _load_user_id_from_session_files(session_file=session_file)

    if user_id is None:
        aviso = (
            "Acesso negado: usuário não informado ao abrir a tela.\n\n"
            "Esta tela exige identificação do usuário para validar permissões.\n"
            "Chame com --user-id/--usuario-id/--uid <id> ou forneça sessão (sessao.json/login.json etc)."
        )
        return SessaoAcesso(
            nivel=0,
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
            nivel=0,
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


def _detect_from_menu_flag() -> bool:
    if "--standalone" in sys.argv:
        return False
    if "--from-menu" in sys.argv:
        return True
    env = (os.getenv("EKENOX_FROM_MENU") or "").strip().lower()
    return env in {"1", "true", "yes", "sim", "s"}


# ============================================================
# DB BASE (igual padrão)
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


def _table_exists_quick(cfg: AppConfig, table_name: str) -> bool:
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
        ok = _table_exists_quick(cfg, t)
        log_estoque(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


# ============================================================
# MODEL / HELPERS
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

        if self.fk_is_numeric:
            max_expr_sql = f"COALESCE(MAX({col_qual})::bigint, 0)"
        else:
            max_expr_sql = (
                f"COALESCE(MAX((NULLIF(regexp_replace({col_qual}::text, '[^0-9]', '', 'g'), '') )::bigint), 0)"
            )

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
                out.append(
                    Estoque(
                        fkProduto=int(str(r[0] or "0")),
                        nomeProduto=str(r[1] or ""),
                        saldoFisico=Decimal(
                            str(r[2] if r[2] is not None else "0")),
                        saldoVirtual=Decimal(
                            str(r[3] if r[3] is not None else "0")),
                    )
                )
            return out
        finally:
            self.db.desconectar()

    def existe_fk(self, fk: int) -> bool:
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


@dataclass
class MenuContext:
    menu_path: Optional[str] = None
    menu_running: bool = False
    launcher: Optional[list[str]] = None


class TelaEstoque(ttk.Frame):
    def __init__(self, master: tk.Misc, service: EstoqueService, acesso: SessaoAcesso, *, from_menu: bool, menu_ctx: MenuContext):
        super().__init__(master)
        self.service = service
        self.acesso = acesso
        self.from_menu = bool(from_menu)
        self.menu_ctx = menu_ctx

        self.var_filtro = tk.StringVar()

        self.var_fk = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_fisico = tk.StringVar(value="0")
        self.var_virtual = tk.StringVar(value="0")

        self.ent_nome: Optional[ttk.Entry] = None

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

    def _set_nome_editavel(self, editavel: bool) -> None:
        if not self.ent_nome:
            return
        if not self._can_edit():
            self.ent_nome.configure(state="readonly")
            return
        self.ent_nome.configure(state=("normal" if editavel else "readonly"))
        if editavel:
            self.ent_nome.focus_set()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        # TOPBAR (igual tela_produtos)
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

        # botão rápido: não chama checagens pesadas no clique
        label_btn = "Fechar" if (
            self.from_menu or self.menu_ctx.menu_running) else "Voltar ao Menu"
        ttk.Button(
            topbar,
            text=label_btn,
            command=self._voltar_ou_fechar,
        ).grid(row=0, column=1, sticky="e")

        # BUSCA / AÇÕES
        top = ttk.Frame(self)
        top.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(
            top, text="Buscar (fkProduto / Nome Produto):").grid(row=0, column=0, sticky="w")
        ent_busca = ttk.Entry(top, textvariable=self.var_filtro)
        ent_busca.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent_busca.bind("<Return>", lambda e: self.atualizar_lista())

        self.btn_atualizar = ttk.Button(
            top, text="Atualizar", command=self.atualizar_lista)
        self.btn_atualizar.grid(row=0, column=2, padx=(0, 6))

        self.btn_novo = ttk.Button(top, text="Novo", command=self.novo)
        self.btn_novo.grid(row=0, column=3, padx=(0, 6))

        self.btn_salvar = ttk.Button(top, text="Salvar", command=self.salvar)
        self.btn_salvar.grid(row=0, column=4, padx=(0, 6))

        self.btn_excluir = ttk.Button(
            top, text="Excluir", command=self.excluir)
        self.btn_excluir.grid(row=0, column=5, padx=(0, 6))

        self.btn_limpar = ttk.Button(
            top, text="Limpar", command=self.limpar_form)
        self.btn_limpar.grid(row=0, column=6)

        # FORM
        form = ttk.LabelFrame(self, text="Estoque")
        form.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(10):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="fkProduto:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_fk = ttk.Entry(form, textvariable=self.var_fk, width=14)
        ent_fk.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=6)
        ent_fk.bind("<Return>", lambda e: self._preencher_nome())

        self.btn_preencher = ttk.Button(
            form, text="Preencher Nome", command=self._preencher_nome)
        self.btn_preencher.grid(
            row=0, column=2, sticky="w", padx=(0, 10), pady=6)

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

        # LISTA
        lst = ttk.Frame(self)
        lst.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
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

    def _aplicar_permissoes(self) -> None:
        n = int(self.acesso.nivel or 0)

        if n <= 0:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="disabled")
            self.btn_preencher.configure(state="disabled")
            if self.ent_nome:
                self.ent_nome.configure(state="readonly")
            return

        if n == 1:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            self.btn_preencher.configure(state="normal")
            if self.ent_nome:
                self.ent_nome.configure(state="readonly")
            return

        if n == 2:
            self.btn_novo.configure(state="normal")
            self.btn_salvar.configure(state="normal")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            self.btn_preencher.configure(state="normal")
            return

        self.btn_novo.configure(state="normal")
        self.btn_salvar.configure(state="normal")
        self.btn_excluir.configure(state="normal")
        self.btn_limpar.configure(state="normal")
        self.btn_preencher.configure(state="normal")

    def _preencher_nome(self) -> None:
        nome, encontrado = self.service.preencher_nome_produto(
            self.var_fk.get())
        if encontrado:
            self.var_nome.set(nome)
            self._set_nome_editavel(False)
        else:
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

        for r in rows:
            self.tree.insert("", "end", values=(
                r.fkProduto,
                r.nomeProduto,
                str(r.saldoFisico),
                str(r.saldoVirtual),
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
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para criar (somente leitura).")
            return

        self.limpar_form()
        self.var_fisico.set("0")
        self.var_virtual.set("0")
        try:
            next_fk = self.service.proximo_fk_nextval()
            self.var_fk.set(str(next_fk))
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
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para salvar (somente leitura).")
            return

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
        if not self._can_delete():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para excluir (somente admin).")
            return

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

    def _voltar_ou_fechar(self) -> None:
        """
        ✅ Igual tela_produtos:
        - Se veio do menu (ou menu já está rodando): apenas fecha.
        - Se abriu fora do menu: fecha primeiro (rápido) e abre menu (sem check) depois.
        """
        top = self.winfo_toplevel()

        # Se "veio do menu" OU menu já estava rodando: só fecha.
        if self.from_menu or self.menu_ctx.menu_running:
            try:
                top.destroy()
            except Exception:
                pass
            return

        # Fora do menu: fechar rápido e abrir menu em seguida
        menu_path = self.menu_ctx.menu_path
        launcher = self.menu_ctx.launcher

        try:
            top.destroy()
        except Exception:
            pass

        if menu_path:
            # depois que a janela fechar, abre o menu (sem checagem)
            try:
                top.after(0, lambda: abrir_menu_principal_sem_check(
                    menu_path, launcher))
            except Exception:
                abrir_menu_principal_sem_check(menu_path, launcher)


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


def main() -> None:
    log_estoque("=== START tela_estoque ===")
    log_estoque(f"APP_DIR={APP_DIR}")
    log_estoque(f"BASE_DIR={BASE_DIR}")
    log_estoque(f"sys.executable={sys.executable}")
    log_estoque(f"argv0={_this_script_abspath()}")
    log_estoque(f"ARGV={sys.argv!r}")

    cfg = env_override(load_config())

    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--from-menu", action="store_true")
    ap.add_argument("--standalone", action="store_true")
    ap.add_argument("--reopen-menu-on-exit",
                    action="store_true")  # opcional e seguro

    ap.add_argument("--user-id", "--usuario-id", "--uid", dest="user_id")
    ap.add_argument("--session-file", dest="session_file")
    ns, _ = ap.parse_known_args()

    menu_path = localizar_menu_principal()
    launcher = _pick_python_launcher_windows()

    # Detecta menu rodando (só aqui, não no clique)
    try:
        menu_running = menu_ja_rodando(menu_path)
    except Exception:
        menu_running = False

    # from_menu: mantém a lógica da tela_produtos
    if bool(ns.standalone):
        from_menu = False
    else:
        from_menu = bool(ns.from_menu) or _detect_from_menu_flag()
        if (not from_menu) and menu_running:
            from_menu = True

    menu_ctx = MenuContext(menu_path=menu_path, menu_running=bool(
        menu_running), launcher=launcher)

    # X nunca abre menu automaticamente. Só abre se passar flag.
    reopen_menu_on_exit = bool(ns.reopen_menu_on_exit) and (not from_menu)

    try:
        test_connection_or_die(cfg)
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        try:
            messagebox.showerror(
                "Erro de conexão",
                "Não foi possível conectar ao banco.\n\n"
                f"Host: {cfg.db_host}\n"
                f"Porta: {cfg.db_port}\n"
                f"Banco: {cfg.db_database}\n"
                f"Usuário: {cfg.db_user}\n\n"
                f"Erro:\n{type(e).__name__}: {e}"
            )
        finally:
            try:
                root.destroy()
            except Exception:
                pass
        return

    acesso = _build_access(cfg, ns)
    log_estoque(
        f"ACESSO_FINAL: nivel={acesso.nivel} origem={acesso.origem} "
        f"usuario_id={acesso.usuario_id} nome={acesso.usuario_nome!r} from_menu={from_menu} menu_running={menu_running}"
    )

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

        # opcional via flag (standalone)
        if reopen_menu_on_exit and (not from_menu) and (not menu_running) and menu_path:
            abrir_menu_principal_sem_check(menu_path, launcher)
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

    estoque_table = detectar_tabela(cfg, ESTOQUE_TABLES, '"estoque"')
    produtos_table = detectar_tabela(cfg, PRODUTOS_TABLES, '"produtos"')

    db = Database(cfg)
    repo = EstoqueRepo(db, estoque_table=estoque_table,
                       produtos_table=produtos_table)
    service = EstoqueService(repo)

    tela = TelaEstoque(root, service, acesso,
                       from_menu=from_menu, menu_ctx=menu_ctx)
    tela.pack(fill="both", expand=True)

    def on_close():
        # ✅ Fechar no X: nunca abre menu automaticamente
        try:
            root.destroy()
        except Exception:
            pass

        # se quiser forçar retorno ao menu ao fechar standalone, use a flag
        if reopen_menu_on_exit and (not from_menu) and (not menu_running) and menu_path:
            try:
                abrir_menu_principal_sem_check(menu_path, launcher)
            except Exception:
                pass

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
