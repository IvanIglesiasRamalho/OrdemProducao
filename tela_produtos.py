from __future__ import annotations

"""
tela_produtos.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Acesso (igual tela_info_produto):
- Valida usuário + nível por programa (Ekenox.usuario_programa + Ekenox.programas)
- Respeita coluna Permitido (se existir) -> se "Não", bloqueia (nível 0)
- Se usuário não for identificado, bloqueia (nível 0)

Logs:
  <BASE_DIR>/logs/tela_produtos.log
  <BASE_DIR>/logs/menu_principal_run.log
"""

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
from typing import Any, Optional

import psycopg2


# ============================================================
# PROGRAMA / PERMISSÕES
# ============================================================

# 🔧 AJUSTE para bater com Ekenox.programas.codigo (conforme tela Usuários/Programas)
PROGRAMA_CODIGO = "PRODUTOS"

NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Admin",
}


# ============================================================
# PASTAS / DIRS
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
    return conn


# ============================================================
# MENU PRINCIPAL (ao fechar)
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
    for cmd in (["pyw", "-3.12"], ["pyw"], ["pythonw"], ["python"]):
        try:
            subprocess.run(cmd + ["-c", "print('ok')"],
                           capture_output=True, text=True, timeout=2)
            return cmd
        except Exception:
            pass
    return ["python"]


def menu_ja_rodando(menu_path: Optional[str] = None) -> bool:
    """
    Evita duplicar Menu.
    - Detecta processo python rodando menu_principal.py/menu.py
    - Detecta janela "Menu Principal - Ekenox"
    - Detecta exe pelo tasklist se tiver nome
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


def abrir_menu_principal_skip_entrada() -> bool:
    """
    Abre o menu principal com --skip-entrada (sem duplicar se já estiver rodando).
    """
    menu_path = localizar_menu_principal()

    log_menu_run("=== abrir_menu_principal_skip_entrada() ===")
    log_menu_run(f"APP_DIR={APP_DIR}")
    log_menu_run(f"BASE_DIR={BASE_DIR}")
    log_menu_run(f"sys.executable={sys.executable}")
    log_menu_run(f"argv0={_this_script_abspath()}")
    log_menu_run(f"menu_path={menu_path}")

    if not menu_path:
        log_app("MENU: não encontrado.")
        return False

    if menu_ja_rodando(menu_path):
        log_app("MENU: já rodando (não duplicar).")
        return True

    try:
        cwd = os.path.dirname(menu_path) or APP_DIR
        is_exe = menu_path.lower().endswith(".exe")
        is_py = menu_path.lower().endswith(".py")

        if is_exe:
            cmd = [menu_path, "--skip-entrada"]
        elif is_py:
            if os.name == "nt":
                cmd = _pick_python_launcher_windows(
                ) + [menu_path, "--skip-entrada"]
            else:
                cmd = [sys.executable, menu_path, "--skip-entrada"]
        else:
            cmd = [menu_path, "--skip-entrada"]

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
# SESSÃO / USUÁRIO / PERMISSÕES (IGUAL tela_info_produto)
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
        for k, v in obj.items():
            kll = str(k).strip().lower()

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
            log_app("ACESSO: tabela programas/programa não encontrada.")
            return None

        up_table = "usuario_programa"
        if not _table_exists(cur, "Ekenox", up_table):
            log_app("ACESSO: tabela usuario_programa não encontrada.")
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
            log_app(
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
            log_app(
                f"ACESSO: sem registro em usuario_programa. user_id={user_id} programa={programa_codigo}")
            return 0

        nivel_int = int(row[0] or 0)
        permitido_val = row[1] if (up_perm_col and len(row) > 1) else True
        permitido = _bool_from_db(permitido_val)

        log_app(
            f"ACESSO: user_id={user_id} programa={programa_codigo} "
            f"nivel={nivel_int} permitido_col={up_perm_col} permitido_val={permitido_val!r} permitido={permitido}"
        )

        # ✅ BLOQUEIA se "permitido" existir e for falso
        if up_perm_col and (not permitido):
            return 0

        return nivel_int

    except Exception as e:
        log_app(f"ACESSO: erro obter_nivel_programa: {type(e).__name__}: {e}")
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
        log_app(f"RESOLVE: user_id via argumento/env = {user_id}")
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
    # só considera "veio do menu" se o Menu avisar explicitamente
    if "--standalone" in sys.argv:
        return False
    if "--from-menu" in sys.argv:
        return True
    env = (os.getenv("EKENOX_FROM_MENU") or "").strip().lower()
    return env in {"1", "true", "yes", "sim", "s"}


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
# UI / MODEL
# ============================================================

DEFAULT_GEOMETRY = "1200x700"
APP_TITLE = "Tela de Produtos"


@dataclass
class Produto:
    produtoId: Optional[str]
    nomeProduto: str
    sku: Optional[str]
    preco: Decimal
    custo: Decimal
    tipo: Optional[str]
    formato: Optional[str]
    descricaoCurta: Optional[str]
    idProdutoPai: Optional[str]
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

    def proximo_produto_id(self) -> int:
        sql = r'''
            SELECT COALESCE(
                MAX(NULLIF(regexp_replace(p."produtoId"::text, '\D', '', 'g'), '')::bigint),
                0
            ) + 1
            FROM "Ekenox"."produtos" AS p
        '''
        if not self.db.conectar():
            raise RuntimeError(
                f"Falha ao conectar no banco: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql)
            return int(self.db.cursor.fetchone()[0])
        finally:
            self.db.desconectar()

    def inserir(self, p: Produto) -> str:
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
            self.db.commit()
            return str(new_id)
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def atualizar(self, p: Produto) -> None:
        if not p.produtoId:
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

    def excluir(self, produto_id: str) -> None:
        sql = 'DELETE FROM "Ekenox"."produtos" WHERE "produtoId" = %s'

        if not self.db.conectar():
            raise RuntimeError(
                f"Falha ao conectar no banco: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(produto_id),))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()


def _clean_text(v: object) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def _to_text_or_none(v: object) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def _to_decimal(v: object) -> Decimal:
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
        raise ValueError(f"Valor numérico inválido: {v!r}")


class ProdutosService:
    def __init__(self, repo: ProdutosRepo) -> None:
        self.repo = repo

    def listar(self, termo: str | None = None) -> list[Produto]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo=termo)

    def proximo_codigo(self) -> int:
        return self.repo.proximo_produto_id()

    def salvar_from_form(self, form: dict) -> str | None:
        produtoId = _to_text_or_none(form.get("produtoId"))
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
            idProdutoPai=_to_text_or_none(form.get("idProdutoPai")),
            descImetro=_clean_text(form.get("descImetro")),
        )

        if p.preco < 0 or p.custo < 0:
            raise ValueError("preço/custo não podem ser negativos.")

        if p.produtoId is None:
            return self.repo.inserir(p)

        self.repo.atualizar(p)
        return None

    def excluir(self, produto_id: str) -> None:
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
    def __init__(self, master: tk.Misc, service: ProdutosService, acesso: SessaoAcesso, *, from_menu: bool):
        super().__init__(master)
        self.service = service
        self.acesso = acesso
        self.from_menu = bool(from_menu)

        self.vars: dict[str, tk.StringVar] = {
            k: tk.StringVar() for k, _ in CAMPOS}
        self.var_filtro = tk.StringVar()

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
        self.rowconfigure(3, weight=1)

        # TOPBAR (igual InfoProduto)
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

        # BUSCA / AÇÕES
        top = ttk.Frame(self)
        top.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (nome/sku/código):").grid(row=0,
                                                              column=0, sticky="w")
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
        form = ttk.LabelFrame(self, text="Produto")
        form.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
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

        # LISTA
        lst_frame = ttk.Frame(self)
        lst_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
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
        self.entries[key] = ent
        return ent

    def _aplicar_permissoes(self) -> None:
        n = int(self.acesso.nivel or 0)

        if n <= 0:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="disabled")
            for k, ent in self.entries.items():
                ent.configure(state="readonly")
            return

        if n == 1:
            self.btn_novo.configure(state="disabled")
            self.btn_salvar.configure(state="disabled")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            for k, ent in self.entries.items():
                ent.configure(state="readonly")
            return

        if n == 2:
            self.btn_novo.configure(state="normal")
            self.btn_salvar.configure(state="normal")
            self.btn_excluir.configure(state="disabled")
            self.btn_limpar.configure(state="normal")
            for k, ent in self.entries.items():
                # produtoId sempre readonly
                ent.configure(state="readonly" if k ==
                              "produtoId" else "normal")
            return

        # n >= 3
        self.btn_novo.configure(state="normal")
        self.btn_salvar.configure(state="normal")
        self.btn_excluir.configure(state="normal")
        self.btn_limpar.configure(state="normal")
        for k, ent in self.entries.items():
            ent.configure(state="readonly" if k == "produtoId" else "normal")

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
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para criar (somente leitura).")
            return

        self.limpar_form()
        self.vars["preco"].set("0")
        self.vars["custo"].set("0")

        try:
            prox = self.service.proximo_codigo()
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao gerar o próximo Código:\n{e}")
            return

        self.vars["produtoId"].set(str(prox))

    def limpar_form(self) -> None:
        for k in self.vars:
            self.vars[k].set("")
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        if not self._can_edit():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para salvar (somente leitura).")
            return

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
        if not self._can_delete():
            messagebox.showwarning(
                "Acesso", "Você não tem permissão para excluir (somente admin).")
            return

        produto_id_str = self.vars["produtoId"].get().strip()
        if not produto_id_str:
            messagebox.showwarning(
                "Atenção", "Selecione um produto para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir o produto Código {produto_id_str}?"):
            return

        try:
            self.service.excluir(produto_id_str)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Produto excluído.")
        self.limpar_form()
        self.atualizar_lista()

    def _voltar_ou_fechar(self) -> None:
        # ✅ se o menu já está rodando, não abre outro
        try:
            if menu_ja_rodando():
                self.winfo_toplevel().destroy()
                return
        except Exception:
            self.winfo_toplevel().destroy()
            return

        # se não está rodando e não veio do menu, abre menu
        if not self.from_menu:
            try:
                abrir_menu_principal_skip_entrada()
            finally:
                self.winfo_toplevel().destroy()
            return

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


def main():
    log_app("=== START tela_produtos ===")
    log_app(f"APP_DIR={APP_DIR}")
    log_app(f"BASE_DIR={BASE_DIR}")
    log_app(f"sys.executable={sys.executable}")
    log_app(f"argv0={_this_script_abspath()}")
    log_app(f"ARGV={sys.argv!r}")

    cfg = env_override(load_config())

    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--from-menu", action="store_true")
    ap.add_argument("--standalone", action="store_true")
    ap.add_argument("--reopen-menu-on-exit", action="store_true")

    ap.add_argument("--user-id", "--usuario-id", "--uid", dest="user_id")
    ap.add_argument("--session-file", dest="session_file")
    ns, _ = ap.parse_known_args()

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
            f"Erro:\n{type(e).__name__}: {e}"
        )
        return

    acesso = _build_access(cfg, ns)
    log_app(
        f"ACESSO_FINAL: nivel={acesso.nivel} origem={acesso.origem} usuario_id={acesso.usuario_id} nome={acesso.usuario_nome!r}")

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

        # se abrir standalone e quiser voltar menu
        if reopen_menu_on_exit and (not from_menu):
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
    repo = ProdutosRepo(db)
    service = ProdutosService(repo)

    tela = TelaProdutos(root, service, acesso, from_menu=from_menu)
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
    main()
