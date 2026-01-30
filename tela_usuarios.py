from __future__ import annotations

import argparse
import binascii
import hashlib
import json
import os
import secrets
import string
import sys
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple

import psycopg2


# ============================================================
# PATHS / LOG
# ============================================================

BASE_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
APP_DIR = BASE_DIR
os.makedirs(BASE_DIR, exist_ok=True)

LOG_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "tela_usuarios.log")
CONFIG_FILE = os.path.join(BASE_DIR, "config_op.json")


def log(msg: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


# ============================================================
# UTILITÁRIOS (TEXTO / EMAIL)
# ============================================================

def to_text(v: Any) -> str:
    """Converte bytes/memoryview para str (melhor esforço)."""
    if v is None:
        return ""
    if isinstance(v, memoryview):
        v = v.tobytes()
    if isinstance(v, (bytes, bytearray)):
        try:
            return bytes(v).decode("utf-8", errors="replace")
        except Exception:
            return ""
    return str(v)


def _is_printable(s: str) -> bool:
    s = s or ""
    return all((ch in string.printable) and (ch not in "\x0b\x0c") for ch in s)


def looks_like_email(s: str) -> bool:
    s = (s or "").strip()
    return ("@" in s) and (" " not in s) and (len(s) >= 5)


def bytes_to_hex(v: Any) -> str:
    """Converte bytea/memoryview para hex legível."""
    if v is None:
        return ""
    if isinstance(v, memoryview):
        v = v.tobytes()
    if isinstance(v, (bytes, bytearray)):
        return binascii.hexlify(bytes(v)).decode("ascii")
    return str(v)


def safe_text_from_any(v: Any) -> str:
    """Tenta obter texto legível de memoryview/bytes/str (sem retornar lixo)."""
    if v is None:
        return ""
    if isinstance(v, memoryview):
        v = v.tobytes()
    if isinstance(v, (bytes, bytearray)):
        try:
            s = bytes(v).decode("utf-8", errors="strict")
            return s if _is_printable(s) else ""
        except Exception:
            return ""
    s = str(v)
    return s if _is_printable(s) else ""


def email_from_columns(email_enc: Any, email_hash: Any) -> Tuple[str, str]:
    """
    Decide o e-mail para EXIBIÇÃO e para EDIÇÃO.

    1) Se email_enc for texto legível e parecer e-mail -> usa email_enc
    2) Senão, se email_hash for texto legível e parecer e-mail -> usa email_hash
       (isso resolve seu caso "e-mail está no hash")
    3) Senão, exibe [hash:xxxx] (curto) e não permite edição
    """
    enc_txt = safe_text_from_any(email_enc).strip()
    hash_txt = safe_text_from_any(email_hash).strip()

    if enc_txt and looks_like_email(enc_txt):
        return enc_txt, enc_txt

    if hash_txt and looks_like_email(hash_txt):
        return hash_txt, hash_txt

    h = bytes_to_hex(email_hash) or bytes_to_hex(email_enc)
    h = (h or "")[:12]
    return (f"[hash:{h}]" if h else "[não disponível]"), ""


def _norm_email(email: str) -> str:
    return (email or "").strip().lower()


# ============================================================
# CONFIG / DB
# ============================================================

@dataclass
class AppConfig:
    db_host: str = "10.0.0.154"
    db_database: str = "postgresekenox"
    db_user: str = "postgresekenox"
    db_password: str = "Ekenox5426"
    db_port: int = 55432


def load_config() -> AppConfig:
    if not os.path.exists(CONFIG_FILE):
        return AppConfig()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return AppConfig(**data)
    except Exception as e:
        log(f"load_config error: {type(e).__name__}: {e}")
        return AppConfig()


def db_connect(cfg: AppConfig):
    return psycopg2.connect(
        host=cfg.db_host,
        database=cfg.db_database,
        user=cfg.db_user,
        password=cfg.db_password,
        port=int(cfg.db_port),
        connect_timeout=5,
    )


def try_connect_db(cfg: AppConfig) -> Tuple[bool, str]:
    try:
        conn = db_connect(cfg)
        conn.close()
        return True, ""
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


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


def apply_window_icon(win: tk.Misc) -> None:
    try:
        ico = find_icon_path()
        if ico:
            try:
                win.iconbitmap(ico)
                return
            except Exception as e:
                log(f"iconbitmap failed: {ico} | {type(e).__name__}: {e}")
    except Exception as e:
        log(f"apply_window_icon fatal: {type(e).__name__}: {e}")


# ============================================================
# DETECÇÃO DE TIPOS (email_hash / email_enc)
# ============================================================

_EMAIL_HASH_IS_BYTEA: Optional[bool] = None
_EMAIL_ENC_IS_BYTEA: Optional[bool] = None


def _column_is_bytea(cfg: AppConfig, schema: str, table: str, column: str) -> bool:
    """
    Detecta se coluna é BYTEA (tenta information_schema; não depende de existir linha).
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT data_type
                  FROM information_schema.columns
                 WHERE table_schema=%s
                   AND table_name=%s
                   AND column_name=%s
                 LIMIT 1
                """,
                (schema, table, column),
            )
            r = cur.fetchone()
            if r and r[0]:
                return str(r[0]).lower() == "bytea"
    except Exception as e:
        log(f"_column_is_bytea error {schema}.{table}.{column}: {type(e).__name__}: {e}")
    finally:
        conn.close()
    return False


def is_email_hash_bytea(cfg: AppConfig) -> bool:
    global _EMAIL_HASH_IS_BYTEA
    if _EMAIL_HASH_IS_BYTEA is None:
        _EMAIL_HASH_IS_BYTEA = _column_is_bytea(
            cfg, "Ekenox", "usuarios", "email_hash")
    return bool(_EMAIL_HASH_IS_BYTEA)


def is_email_enc_bytea(cfg: AppConfig) -> bool:
    global _EMAIL_ENC_IS_BYTEA
    if _EMAIL_ENC_IS_BYTEA is None:
        _EMAIL_ENC_IS_BYTEA = _column_is_bytea(
            cfg, "Ekenox", "usuarios", "email_enc")
    return bool(_EMAIL_ENC_IS_BYTEA)


def email_hash_value(cfg: AppConfig, email: str) -> Any:
    """
    Retorna o hash compatível com o tipo da coluna:
      - BYTEA: digest (32 bytes)
      - TEXT: hexdigest (64 chars)
    """
    e = _norm_email(email)
    if is_email_hash_bytea(cfg):
        return hashlib.sha256(e.encode("utf-8")).digest()
    return hashlib.sha256(e.encode("utf-8")).hexdigest()


def email_hash_candidates(email: str) -> Tuple[str, bytes]:
    """Para busca: tenta TEXT (hex) e BYTEA (digest)."""
    e = _norm_email(email)
    return hashlib.sha256(e.encode("utf-8")).hexdigest(), hashlib.sha256(e.encode("utf-8")).digest()


# ============================================================
# PASSWORD
# ============================================================

def hash_password(password: str, salt: str) -> str:
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120_000,
        dklen=32,
    )
    return dk.hex()


def make_password_record(password: str) -> str:
    salt = secrets.token_hex(16)
    ph = hash_password(password, salt)
    return f"{salt}${ph}"


# ============================================================
# DB: USUÁRIOS
# ============================================================

def fetch_users(cfg: AppConfig, filtro: str = "") -> List[dict]:
    """
    - Lista usuários
    - Mostra email de forma legível (se estiver em email_enc OU em email_hash)
    - Se vier bytea/cripto/lixo, mostra [hash:xxxx]
    """
    filtro = (filtro or "").strip()
    where = ""
    params: List[Any] = []

    # filtro por nome e, se possível, por email (somente colunas text)
    if filtro:
        like = f"%{filtro.lower()}%"
        clauses = ['LOWER(COALESCE(u."nome",\'\')) LIKE %s']
        params.append(like)

        # se email_enc NÃO for bytea, dá pra filtrar também
        if not is_email_enc_bytea(cfg):
            clauses.append('LOWER(COALESCE(u."email_enc",\'\')) LIKE %s')
            params.append(like)

        # se email_hash NÃO for bytea, pode ser texto (às vezes o e-mail está aqui)
        if not is_email_hash_bytea(cfg):
            clauses.append(
                'LOWER(COALESCE(u."email_hash"::text,\'\')) LIKE %s')
            params.append(like)

        where = "WHERE " + " OR ".join(clauses)

    sql = f"""
        SELECT u."usuarioId",
               COALESCE(u."nome",'') AS nome,
               u."email_enc" AS email_enc,
               u."email_hash" AS email_hash,
               COALESCE(u."ativo", true) AS ativo
          FROM "Ekenox"."usuarios" u
          {where}
         ORDER BY u."usuarioId" DESC
    """

    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, tuple(params))
            rows = cur.fetchall()

        out: List[dict] = []
        for r in rows:
            email_disp, _ = email_from_columns(r[2], r[3])
            out.append(
                {
                    "usuarioId": int(r[0]),
                    "nome": to_text(r[1]),
                    "email_disp": email_disp,
                    "ativo": bool(r[4]),
                }
            )
        return out
    finally:
        conn.close()


def fetch_user_by_id(cfg: AppConfig, usuario_id: int) -> dict:
    """
    Busca usuário e devolve:
      - email_disp: pra mostrar
      - email_edit: pra editar (só se parecer e-mail)
    """
    sql = """
        SELECT u."usuarioId",
               COALESCE(u."nome",'') AS nome,
               u."email_enc" AS email_enc,
               u."email_hash" AS email_hash,
               COALESCE(u."ativo", true) AS ativo
          FROM "Ekenox"."usuarios" u
         WHERE u."usuarioId" = %s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
        if not r:
            raise RuntimeError("Usuário não encontrado.")

        email_disp, email_edit = email_from_columns(r[2], r[3])
        return {
            "usuarioId": int(r[0]),
            "nome": to_text(r[1]),
            "email_disp": email_disp,
            "email_edit": email_edit,
            "ativo": bool(r[4]),
        }
    finally:
        conn.close()


def insert_user(cfg: AppConfig, nome: str, email: str, ativo: bool) -> int:
    """
    Inserção CORRIGIDA:
      - email_enc: grava o e-mail como texto (ou bytes utf-8 se coluna for bytea)
      - email_hash: grava o hash sha256 compatível com o tipo da coluna
    """
    nome = (nome or "").strip()
    email_n = _norm_email(email)
    if not nome:
        raise ValueError("Nome obrigatório.")
    if not email_n:
        raise ValueError("E-mail obrigatório.")

    # email_enc pode ser TEXT ou BYTEA no seu banco
    email_enc_val: Any = email_n.encode(
        "utf-8") if is_email_enc_bytea(cfg) else email_n
    eh = email_hash_value(cfg, email_n)

    sql = """
        INSERT INTO "Ekenox"."usuarios"
            ("email_hash","email_enc","senha_hash","nome","ativo","criado_em","atualizado_em")
        VALUES
            (%s, %s, %s, %s, %s, NOW(), NOW())
        RETURNING "usuarioId"
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (eh, email_enc_val, "", nome, bool(ativo)))
            uid = cur.fetchone()[0]
        conn.commit()
        return int(uid)
    finally:
        conn.close()


def update_user(cfg: AppConfig, usuario_id: int, nome: str, email: str, ativo: bool) -> None:
    """
    Update CORRIGIDO:
      - Se antes o e-mail estava "no hash", ao salvar ele MIGRA:
        email_enc passa a ter e-mail (texto) e email_hash vira sha256.
    """
    nome = (nome or "").strip()
    email_n = _norm_email(email)
    if not nome:
        raise ValueError("Nome obrigatório.")
    if not email_n:
        raise ValueError("E-mail obrigatório.")

    email_enc_val: Any = email_n.encode(
        "utf-8") if is_email_enc_bytea(cfg) else email_n
    eh = email_hash_value(cfg, email_n)

    sql = """
        UPDATE "Ekenox"."usuarios"
           SET "nome" = %s,
               "email_enc" = %s,
               "email_hash" = %s,
               "ativo" = %s,
               "atualizado_em" = NOW()
         WHERE "usuarioId" = %s
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (nome, email_enc_val, eh,
                        bool(ativo), int(usuario_id)))
            if cur.rowcount != 1:
                raise RuntimeError("Usuário não encontrado para atualizar.")
        conn.commit()
    finally:
        conn.close()


def set_user_password(cfg: AppConfig, usuario_id: int, new_password: str) -> None:
    if len(new_password or "") < 4:
        raise ValueError("Senha deve ter pelo menos 4 caracteres.")
    rec = make_password_record(new_password)
    sql = """
        UPDATE "Ekenox"."usuarios"
           SET "senha_hash" = %s,
               "atualizado_em" = NOW()
         WHERE "usuarioId" = %s
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (rec, int(usuario_id)))
            if cur.rowcount != 1:
                raise RuntimeError(
                    "Usuário não encontrado para definir senha.")
        conn.commit()
    finally:
        conn.close()


def fetch_user_nome(cfg: AppConfig, usuario_id: int) -> str:
    sql = """
        SELECT COALESCE(u."nome",'')
          FROM "Ekenox"."usuarios" u
         WHERE u."usuarioId" = %s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
        return to_text(r[0]).strip() if r else ""
    finally:
        conn.close()


# ============================================================
# DB: PROGRAMAS / PERMISSÕES
# ============================================================

def fetch_programs(cfg: AppConfig) -> List[dict]:
    sql = """
        SELECT pr."programaId", COALESCE(pr."codigo",''), COALESCE(pr."nome",'')
          FROM "Ekenox"."programas" pr
         ORDER BY pr."nome"
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql)
            rows = cur.fetchall()
        out = []
        for r in rows:
            out.append(
                {
                    "programaId": int(r[0]),
                    "codigo": to_text(r[1]).strip(),
                    "nome": to_text(r[2]).strip(),
                }
            )
        return out
    finally:
        conn.close()


def fetch_user_program_levels(cfg: AppConfig, usuario_id: int) -> Dict[int, int]:
    sql = """
        SELECT up."programaId", COALESCE(up."nivel", 0)
          FROM "Ekenox"."usuario_programa" up
         WHERE up."usuarioId" = %s
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id),))
            rows = cur.fetchall()
        return {int(r[0]): int(r[1]) for r in rows}
    finally:
        conn.close()


def save_user_permissions(cfg: AppConfig, usuario_id: int, perms: Dict[int, int]) -> None:
    """
    Regras:
      - nivel 0 => NÃO grava (fica sem linha na usuario_programa)
      - nivel 1/2/3 => grava
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                'DELETE FROM "Ekenox"."usuario_programa" WHERE "usuarioId" = %s', (int(usuario_id),))
            for programa_id, nivel in perms.items():
                if int(nivel) <= 0:
                    continue
                cur.execute(
                    'INSERT INTO "Ekenox"."usuario_programa" ("usuarioId","programaId","nivel") VALUES (%s,%s,%s)',
                    (int(usuario_id), int(programa_id), int(nivel)),
                )
        conn.commit()
    finally:
        conn.close()


# ============================================================
# CONTROLE DE ACESSO (DA TELA)
# ============================================================

THIS_PROGRAMA_TERMO = "Usuários"

NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Admin",
}


def _parse_cli_user(cfg: AppConfig) -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--usuario-id", "--uid", dest="usuario_id", type=int)
    parser.add_argument("--email", dest="email", type=str)
    args, _ = parser.parse_known_args(sys.argv[1:])

    if args.usuario_id:
        return int(args.usuario_id)

    if args.email:
        email_n = _norm_email(args.email)
        hexh, byh = email_hash_candidates(email_n)

        # Monta WHERE de acordo com os tipos reais das colunas
        clauses: List[str] = []
        params: List[Any] = []

        # tenta bater por e-mail em email_enc, se for TEXT
        if not is_email_enc_bytea(cfg):
            clauses.append('LOWER(COALESCE(u."email_enc",\'\')) = %s')
            params.append(email_n)

        # tenta bater por e-mail em email_hash, se for TEXT (caso seu "email está no hash")
        if not is_email_hash_bytea(cfg):
            clauses.append('LOWER(COALESCE(u."email_hash"::text,\'\')) = %s')
            params.append(email_n)

        # tenta bater por hash
        if is_email_hash_bytea(cfg):
            clauses.append('u."email_hash" = %s')
            params.append(psycopg2.Binary(byh))
        else:
            clauses.append('u."email_hash" = %s')
            params.append(hexh)

        if not clauses:
            raise ValueError(
                "Não foi possível montar busca de usuário (tipos de coluna inesperados).")

        sql = f"""
            SELECT u."usuarioId"
              FROM "Ekenox"."usuarios" u
             WHERE {" OR ".join(clauses)}
             LIMIT 1
        """

        conn = db_connect(cfg)
        try:
            with conn.cursor() as cur:
                cur.execute(sql, tuple(params))
                r = cur.fetchone()
            if not r:
                raise ValueError(
                    "Usuário não encontrado para o e-mail informado.")
            return int(r[0])
        finally:
            conn.close()

    raise ValueError(
        "Usuário não informado. Use --usuario-id <id> ou --email <email>.")


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


# ============================================================
# UI
# ============================================================

class TelaUsuariosApp(tk.Tk):
    def __init__(self, cfg: AppConfig, *, usuario_logado_id: int, acesso_nivel: int) -> None:
        super().__init__()
        self.cfg = cfg
        self.usuario_logado_id = int(usuario_logado_id)
        self.acesso_nivel = int(acesso_nivel)
        self.usuario_logado_nome = fetch_user_nome(cfg, self.usuario_logado_id)

        self.title("Ekenox - Usuários e Programas")

        # TELA: abre menor (não maximizada)
        self.geometry("1250x680")
        self.minsize(1200, 650)

        apply_window_icon(self)
        self.protocol("WM_DELETE_WINDOW", self._close)

        ok, err = try_connect_db(self.cfg)
        if not ok:
            messagebox.showerror(
                "Banco de Dados", f"Falha ao conectar:\n{err}")

        # aquece caches
        try:
            is_email_hash_bytea(self.cfg)
            is_email_enc_bytea(self.cfg)
        except Exception:
            pass

        self.programs = fetch_programs(self.cfg)

        self.selected_user_id: Optional[int] = None
        self.program_rows: Dict[str, dict] = {}

        self._btn_novo: Optional[ttk.Button] = None
        self._btn_salvar_usuario: Optional[ttk.Button] = None
        self._btn_senha: Optional[ttk.Button] = None
        self._btn_toggle: Optional[ttk.Button] = None
        self._btn_salvar_perms: Optional[ttk.Button] = None

        self._ent_nome: Optional[ttk.Entry] = None
        self._ent_email: Optional[ttk.Entry] = None

        self._build()
        self._load_users()
        self._apply_access_rules()

    def _close(self) -> None:
        self.destroy()

    def _apply_access_rules(self) -> None:
        if self._ent_nome:
            self._ent_nome.state(
                ["readonly"] if self.acesso_nivel < 2 else ["!readonly"])
        if self._ent_email:
            self._ent_email.state(
                ["readonly"] if self.acesso_nivel < 2 else ["!readonly"])

        if self._btn_novo:
            self._btn_novo.state(
                ["disabled"] if self.acesso_nivel < 2 else ["!disabled"])
        if self._btn_salvar_usuario:
            self._btn_salvar_usuario.state(
                ["disabled"] if self.acesso_nivel < 2 else ["!disabled"])
        if self._btn_senha:
            self._btn_senha.state(
                ["disabled"] if self.acesso_nivel < 2 else ["!disabled"])
        if self._btn_toggle:
            self._btn_toggle.state(
                ["disabled"] if self.acesso_nivel < 2 else ["!disabled"])
        if self._btn_salvar_perms:
            self._btn_salvar_perms.state(
                ["disabled"] if self.acesso_nivel < 3 else ["!disabled"])

        nome = self.usuario_logado_nome or f"ID {self.usuario_logado_id}"
        nivel_txt = NIVEL_LABEL.get(self.acesso_nivel, str(self.acesso_nivel))
        self.lbl_access.config(
            text=f"Logado: {nome} | Nível: {nivel_txt}",
            foreground=("green" if self.acesso_nivel >= 3 else (
                "orange" if self.acesso_nivel == 2 else "gray")),
        )

    def _build(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        topbar = ttk.Frame(root)
        topbar.pack(fill="x", pady=(0, 8))

        self.lbl_access = ttk.Label(
            topbar, text="", foreground="gray", font=("Segoe UI", 9, "bold"))
        self.lbl_access.pack(side="left", anchor="w")

        ttk.Button(topbar, text="Voltar ao Menu",
                   command=self._close).pack(side="right")

        main = ttk.Frame(root)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main, width=360)
        right = ttk.Frame(main)

        left.pack(side="left", fill="y", expand=False, padx=(0, 10))
        left.pack_propagate(False)   # impede o left de crescer pelo conteúdo
        right.pack(side="left", fill="both", expand=True)

        # ================= LEFT (LISTA) =================
        ttk.Label(left, text="Usuários", font=(
            "Segoe UI", 11, "bold")).pack(anchor="w")

        row_search = ttk.Frame(left)
        row_search.pack(fill="x", pady=(6, 8))

        self.var_search = tk.StringVar()
        ent = ttk.Entry(row_search, textvariable=self.var_search, width=32)
        ent.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(row_search, text="Buscar",
                   command=self._load_users).pack(side="left")
        ttk.Button(row_search, text="Limpar", command=self._clear_search).pack(
            side="left", padx=(6, 0))
        ent.bind("<Return>", lambda e: self._load_users())

        # --- Área da lista + botões ao lado ---
        users_area = ttk.Frame(left)
        users_area.pack(fill="both", expand=True)

        users_area.rowconfigure(0, weight=1)
        users_area.columnconfigure(0, weight=1)

        cols = ("id", "nome", "email", "ativo")
        self.tree_users = ttk.Treeview(
            users_area, columns=cols, show="headings", height=28
        )
        self.tree_users.heading("id", text="ID")
        self.tree_users.heading("nome", text="Nome")
        self.tree_users.heading("email", text="E-mail")
        self.tree_users.heading("ativo", text="Ativo")

        # Ajuste de colunas (deixe mais compacto para caber com botões ao lado)
        self.tree_users.column("id", width=50, anchor="center", stretch=False)
        self.tree_users.column("nome", width=140, anchor="w", stretch=True)
        self.tree_users.column("email", width=170, anchor="w", stretch=True)
        self.tree_users.column(
            "ativo", width=55, anchor="center", stretch=False)

        sb = ttk.Scrollbar(users_area, orient="vertical",
                           command=self.tree_users.yview)
        self.tree_users.configure(yscrollcommand=sb.set)

        # Coluna de botões (lado direito da lista)
        btn_col_users = ttk.Frame(users_area)
        # (se quiser “colar” no topo, sticky="n"; se quiser ocupar tudo, sticky="ns")
        btn_col_users.grid(row=0, column=2, sticky="n", padx=(10, 0))

        self._btn_novo = ttk.Button(
            btn_col_users, text="Novo usuário", command=self._new_user)
        self._btn_novo.pack(fill="x", pady=(0, 6))

        ttk.Button(btn_col_users, text="Recarregar",
                   command=self._load_users).pack(fill="x")

        # Posicionamento Tree + Scroll (grid)
        self.tree_users.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        self.tree_users.bind("<<TreeviewSelect>>", self._on_select_user)

        # ================= RIGHT (FORM) =================
        top = ttk.LabelFrame(right, text="Cadastro do usuário", padding=10)
        top.pack(fill="x")

        frm = ttk.Frame(top)
        frm.pack(fill="x")

        self.var_nome = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_ativo = tk.BooleanVar(value=True)

        ttk.Label(frm, text="Nome:").grid(row=0, column=0, sticky="w")
        self._ent_nome = ttk.Entry(
            frm, textvariable=self.var_nome, width=40)  # menor
        self._ent_nome.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        ttk.Label(frm, text="E-mail:").grid(row=1,
                                            column=0, sticky="w", pady=(8, 0))
        self._ent_email = ttk.Entry(frm, textvariable=self.var_email, width=55)
        self._ent_email.grid(row=1, column=1, sticky="ew",
                             padx=(6, 0), pady=(8, 0))

        chk = ttk.Checkbutton(frm, text="Ativo", variable=self.var_ativo)
        chk.grid(row=2, column=1, sticky="w", pady=(10, 0))

        frm.columnconfigure(1, weight=1)

        row_btns = ttk.Frame(top)
        row_btns.pack(fill="x", pady=(10, 0))

        self._btn_salvar_usuario = ttk.Button(
            row_btns, text="Salvar usuário", command=self._save_user)
        self._btn_salvar_usuario.pack(side="left")

        self._btn_senha = ttk.Button(
            row_btns, text="Definir/Resetar senha", command=self._open_set_password)
        self._btn_senha.pack(side="left", padx=(8, 0))

        self._btn_toggle = ttk.Button(
            row_btns, text="Ativar/Desativar", command=self._toggle_active)
        self._btn_toggle.pack(side="left", padx=(8, 0))

        self.lbl_info = ttk.Label(top, text="", foreground="gray")
        self.lbl_info.pack(anchor="w", pady=(8, 0))

        # ================= PROGRAMAS =================
        bottom = ttk.LabelFrame(
            right, text="Programas do usuário (duplo clique para editar nível)", padding=10)
        bottom.pack(fill="both", expand=True, pady=(10, 0))
        cols_p = ("permitido", "codigo", "nome", "nivel")

        # container interno para organizar Tree + botões (lado a lado)
        prog_area = ttk.Frame(bottom)
        prog_area.pack(fill="both", expand=True)

        self.tree_prog = ttk.Treeview(
            prog_area, columns=cols_p, show="headings")
        self.tree_prog.heading("permitido", text="Permitido")
        self.tree_prog.heading("codigo", text="Código")
        self.tree_prog.heading("nome", text="Nome")
        self.tree_prog.heading("nivel", text="Nível")

        self.tree_prog.column("permitido", width=85,
                              anchor="center", stretch=False)
        self.tree_prog.column("codigo", width=130, anchor="w", stretch=False)
        self.tree_prog.column("nome", width=560, anchor="w",
                              stretch=True)   # AMARELO maior
        self.tree_prog.column("nivel", width=110,
                              anchor="center", stretch=False)  # LARANJA menor

        sbp = ttk.Scrollbar(prog_area, orient="vertical",
                            command=self.tree_prog.yview)
        self.tree_prog.configure(yscrollcommand=sbp.set)

        # coluna de botões (direita) - um abaixo do outro
        btn_col = ttk.Frame(prog_area)
        btn_col.grid(row=0, column=2, sticky="n", padx=(10, 0))

        self._btn_salvar_perms = ttk.Button(
            btn_col, text="Salvar permissões", command=self._save_permissions)
        self._btn_salvar_perms.pack(fill="x", pady=(0, 6))

        ttk.Button(btn_col, text="Marcar todos (nível 1)", command=lambda: self._set_all_level(1)).pack(
            fill="x", pady=(0, 6)
        )
        ttk.Button(btn_col, text="Desmarcar todos", command=lambda: self._set_all_level(0)).pack(
            fill="x"
        )

        # grid do Tree + Scroll
        self.tree_prog.grid(row=0, column=0, sticky="nsew")
        sbp.grid(row=0, column=1, sticky="ns")

        prog_area.columnconfigure(0, weight=1)  # Tree expande
        prog_area.rowconfigure(0, weight=1)

        self.tree_prog.bind("<Double-1>", self._on_double_click_program)

        self._render_programs_for_user(None)

        # ===== Ajuste dinâmico de colunas (encaixa melhor em qualquer tamanho) =====
        def _resize_cols(_evt=None):
            try:
                w = self.tree_users.winfo_width()
                if w > 50:
                    self.tree_users.column("id", width=50)
                    self.tree_users.column("ativo", width=55)
                    rem = max(220, w - 50 - 55 - 30)
                    self.tree_users.column("nome", width=int(rem * 0.40))
                    self.tree_users.column("email", width=int(rem * 0.60))

                wp = self.tree_prog.winfo_width()
                if wp > 50:
                    w_permit = 85
                    w_codigo = 130
                    w_nivel = 110   # menor
                    self.tree_prog.column("permitido", width=w_permit)
                    self.tree_prog.column("codigo", width=w_codigo)
                    self.tree_prog.column("nivel", width=w_nivel)

                    remp = max(280, wp - w_permit - w_codigo - w_nivel - 30)
                    # AMARELO ganha o resto
                    self.tree_prog.column("nome", width=remp)

            except Exception:
                pass

        self.bind("<Configure>", _resize_cols)

    # ============================================================
    # USERS
    # ============================================================

    def _clear_search(self) -> None:
        self.var_search.set("")
        self._load_users()

    def _load_users(self) -> None:
        filtro = self.var_search.get()
        try:
            users = fetch_users(self.cfg, filtro=filtro)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao buscar usuários:\n{e}")
            return

        for i in self.tree_users.get_children():
            self.tree_users.delete(i)

        for u in users:
            self.tree_users.insert(
                "",
                "end",
                values=(u["usuarioId"], u["nome"], u["email_disp"],
                        "Sim" if u["ativo"] else "Não"),
            )

    def _new_user(self) -> None:
        if self.acesso_nivel < 2:
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return
        self.selected_user_id = None
        self.var_nome.set("")
        self.var_email.set("")
        self.var_ativo.set(True)
        self.lbl_info.config(text="Novo usuário (ainda não salvo).")
        self._render_programs_for_user(None)

    def _on_select_user(self, event=None) -> None:
        sel = self.tree_users.selection()
        if not sel:
            return
        vals = self.tree_users.item(sel[0], "values")
        if not vals:
            return

        usuario_id = int(vals[0])
        self.selected_user_id = usuario_id

        try:
            u = fetch_user_by_id(self.cfg, usuario_id)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar usuário:\n{e}")
            return

        self.var_nome.set(u["nome"])
        # se e-mail estiver no hash, vem aqui
        self.var_email.set(u["email_edit"])
        self.var_ativo.set(bool(u["ativo"]))
        self.lbl_info.config(
            text=f"Editando usuário ID {usuario_id} | E-mail: {u['email_disp']}")
        self._render_programs_for_user(usuario_id)

    def _save_user(self) -> None:
        if self.acesso_nivel < 2:
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return

        nome = self.var_nome.get()
        email = self.var_email.get()
        ativo = bool(self.var_ativo.get())

        try:
            if self.selected_user_id is None:
                uid = insert_user(self.cfg, nome, email, ativo)
                self.selected_user_id = uid
                self.lbl_info.config(text=f"Usuário criado com ID {uid}.")
                self._load_users()
                self._render_programs_for_user(uid)
            else:
                update_user(self.cfg, self.selected_user_id,
                            nome, email, ativo)
                self.lbl_info.config(
                    text=f"Usuário ID {self.selected_user_id} atualizado.")
                self._load_users()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar usuário:\n{e}")

    def _toggle_active(self) -> None:
        if self.acesso_nivel < 2:
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return
        if self.selected_user_id is None:
            messagebox.showwarning(
                "Usuário", "Selecione um usuário para ativar/desativar.")
            return
        ativo_atual = bool(self.var_ativo.get())
        self.var_ativo.set(not ativo_atual)
        self._save_user()

    # ============================================================
    # PASSWORD
    # ============================================================

    def _open_set_password(self) -> None:
        if self.acesso_nivel < 2:
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return
        if self.selected_user_id is None:
            messagebox.showwarning(
                "Senha", "Selecione um usuário para definir/resetar a senha.")
            return

        win = tk.Toplevel(self)
        win.title("Definir/Resetar senha")
        win.resizable(False, False)
        apply_window_icon(win)
        win.transient(self)
        win.grab_set()
        win.geometry("420x220")

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=f"Usuário ID: {self.selected_user_id}", font=(
            "Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, text="Nova senha:", font=(
            "Segoe UI", 10)).pack(anchor="w", pady=(10, 0))

        v1 = tk.StringVar()
        v2 = tk.StringVar()

        e1 = ttk.Entry(frm, textvariable=v1, show="*")
        e2 = ttk.Entry(frm, textvariable=v2, show="*")
        e1.pack(fill="x", pady=(6, 6))
        ttk.Label(frm, text="Confirmar senha:", font=(
            "Segoe UI", 10)).pack(anchor="w")
        e2.pack(fill="x", pady=(6, 10))

        def salvar():
            s1 = v1.get()
            s2 = v2.get()
            if len(s1) < 4:
                messagebox.showwarning("Senha", "Mínimo 4 caracteres.")
                return
            if s1 != s2:
                messagebox.showwarning("Senha", "As senhas não conferem.")
                return
            try:
                set_user_password(self.cfg, int(self.selected_user_id), s1)
                messagebox.showinfo("OK", "Senha definida com sucesso.")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao definir senha:\n{e}")

        row = ttk.Frame(frm)
        row.pack(fill="x")
        ttk.Button(row, text="Salvar", command=salvar).pack(
            side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(row, text="Cancelar", command=win.destroy).pack(
            side="left", expand=True, fill="x")

        e1.focus_set()
        win.bind("<Return>", lambda e: salvar())
        win.bind("<Escape>", lambda e: win.destroy())

    # ============================================================
    # PROGRAMS / PERMS
    # ============================================================

    def _render_programs_for_user(self, usuario_id: Optional[int]) -> None:
        for i in self.tree_prog.get_children():
            self.tree_prog.delete(i)
        self.program_rows.clear()

        levels: Dict[int, int] = {}
        if usuario_id is not None:
            try:
                levels = fetch_user_program_levels(self.cfg, usuario_id)
            except Exception as e:
                messagebox.showerror(
                    "Erro", f"Falha ao buscar permissões:\n{e}")
                levels = {}

        for p in self.programs:
            pid = int(p["programaId"])
            nivel = int(levels.get(pid, 0))
            allowed = nivel > 0

            iid = self.tree_prog.insert(
                "",
                "end",
                values=("Sim" if allowed else "Não",
                        p["codigo"], p["nome"], NIVEL_LABEL.get(nivel, str(nivel))),
            )
            self.program_rows[iid] = {
                "programaId": pid,
                "codigo": p["codigo"],
                "nome": p["nome"],
                "nivel": nivel,
                "allowed": allowed,
            }

    def _on_double_click_program(self, event) -> None:
        if self.acesso_nivel < 3:
            return
        iid = self.tree_prog.identify_row(event.y)
        col = self.tree_prog.identify_column(event.x)
        if not iid or iid not in self.program_rows:
            return
        if col in ("#1", "#4"):
            self._open_level_editor(iid)

    def _open_level_editor(self, iid: str) -> None:
        d = self.program_rows[iid]

        win = tk.Toplevel(self)
        win.title("Definir nível de acesso")
        win.resizable(False, False)
        apply_window_icon(win)
        win.transient(self)
        win.grab_set()
        win.geometry("380x220")

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=d["nome"], font=(
            "Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, text=f'Código: {d["codigo"]}', foreground="gray").pack(
            anchor="w", pady=(0, 10))

        ttk.Label(frm, text="Selecione o nível de acesso:").pack(anchor="w")

        var = tk.StringVar(value=str(int(d.get("nivel", 0))))
        values = ["0 - Sem acesso", "1 - Leitura", "2 - Edição", "3 - Admin"]
        cbo = ttk.Combobox(frm, textvariable=var,
                           values=values, state="readonly", width=25)
        cbo.current(int(d.get("nivel", 0)))
        cbo.pack(anchor="w", pady=(6, 10))

        def salvar():
            try:
                txt = cbo.get()
                nv = int(txt.split("-")[0].strip())
                if nv not in (0, 1, 2, 3):
                    raise ValueError()
            except Exception:
                messagebox.showerror(
                    "Valor inválido", "Selecione um nível válido.")
                return

            d["nivel"] = nv
            d["allowed"] = nv > 0
            self._refresh_program_row(iid)
            win.destroy()

        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(12, 0))
        ttk.Button(row, text="OK", command=salvar).pack(
            side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(row, text="Cancelar", command=win.destroy).pack(
            side="left", expand=True, fill="x")

        cbo.focus_set()
        win.bind("<Return>", lambda e: salvar())
        win.bind("<Escape>", lambda e: win.destroy())

    def _refresh_program_row(self, iid: str) -> None:
        d = self.program_rows[iid]
        nivel = int(d.get("nivel", 0))
        self.tree_prog.item(
            iid,
            values=("Sim" if nivel > 0 else "Não",
                    d["codigo"], d["nome"], NIVEL_LABEL.get(nivel, str(nivel))),
        )

    def _set_all_level(self, nivel: int) -> None:
        if self.acesso_nivel < 3:
            messagebox.showwarning(
                "Acesso", "Somente Admin (nível 3) pode alterar permissões.")
            return
        if nivel not in (0, 1, 2, 3):
            return

        for iid, d in self.program_rows.items():
            d["nivel"] = int(nivel)
            d["allowed"] = int(nivel) > 0
            self._refresh_program_row(iid)

    def _save_permissions(self) -> None:
        if self.acesso_nivel < 3:
            messagebox.showwarning(
                "Acesso", "Somente Admin (nível 3) pode salvar permissões.")
            return
        if self.selected_user_id is None:
            messagebox.showwarning(
                "Permissões", "Selecione um usuário para salvar permissões.")
            return

        perms: Dict[int, int] = {}
        for d in self.program_rows.values():
            nv = int(d.get("nivel", 0))
            if nv > 0:
                perms[int(d["programaId"])] = nv

        try:
            save_user_permissions(self.cfg, int(self.selected_user_id), perms)
            messagebox.showinfo("OK", "Permissões salvas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar permissões:\n{e}")


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    cfg = load_config()

    try:
        usuario_id = _parse_cli_user(cfg)
    except Exception as e:
        _deny_and_exit(str(e))

    nivel, aviso = get_access_level_for_this_screen(cfg, usuario_id)
    if nivel <= 0:
        _deny_and_exit(aviso)

    app = TelaUsuariosApp(
        cfg, usuario_logado_id=usuario_id, acesso_nivel=nivel)

    if aviso:
        app.after(200, lambda: messagebox.showwarning(
            "Aviso", aviso, parent=app))

    app.mainloop()
