from __future__ import annotations

import argparse
import hashlib
import json
import os
import secrets
import sys
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Dict, List, Optional, Tuple, Union

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
# ÍCONE (favicon)
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
    """
    Aplica ícone .ico no Windows; se falhar, tenta .png.
    Mantém referência do PhotoImage para não ser coletado.
    """
    try:
        ico = find_icon_path()
        if ico:
            try:
                win.iconbitmap(ico)
                return
            except Exception as e:
                log(f"iconbitmap failed: {ico} | {type(e).__name__}: {e}")

        png_candidates = [
            os.path.join(BASE_DIR, "imagens", "favicon.png"),
            os.path.join(BASE_DIR, "favicon.png"),
            os.path.join(APP_DIR, "imagens", "favicon.png"),
            os.path.join(APP_DIR, "favicon.png"),
        ]
        png = next((p for p in png_candidates if os.path.isfile(p)), None)
        if png:
            try:
                img = tk.PhotoImage(file=png)
                win.iconphoto(True, img)
                win._icon_img = img  # type: ignore[attr-defined]
            except Exception as e:
                log(f"iconphoto failed: {png} | {type(e).__name__}: {e}")
    except Exception as e:
        log(f"apply_window_icon fatal: {type(e).__name__}: {e}")


# ============================================================
# SEGURANÇA (hash email + PBKDF2 senha) - COMPATÍVEL COM MENU
# ============================================================

def _norm_email(email: str) -> str:
    return (email or "").strip().lower()


def email_hash(email: str) -> str:
    e = _norm_email(email)
    return hashlib.sha256(e.encode("utf-8")).hexdigest()


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
# EMAIL_ENC: suporte a TEXT ou BYTEA (detecta no banco)
# ============================================================

EMAIL_PEPPER = (os.getenv("EKENOX_EMAIL_PEPPER")
                or "EKENOX-EMAIL-PEPPER").encode("utf-8")
_EMAIL_ENC_IS_BYTEA_CACHE: Optional[bool] = None


def _xor_bytes(data: bytes, key: bytes) -> bytes:
    out = bytearray(len(data))
    for i, b in enumerate(data):
        out[i] = b ^ key[i % len(key)]
    return bytes(out)


def _detect_email_enc_is_bytea(cfg: AppConfig) -> bool:
    global _EMAIL_ENC_IS_BYTEA_CACHE
    if _EMAIL_ENC_IS_BYTEA_CACHE is not None:
        return _EMAIL_ENC_IS_BYTEA_CACHE

    sql = """
        SELECT data_type
          FROM information_schema.columns
         WHERE table_schema = 'Ekenox'
           AND table_name = 'usuarios'
           AND column_name = 'email_enc'
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql)
        r = cur.fetchone()
        dt = (r[0] or "").strip().lower() if r else ""
        _EMAIL_ENC_IS_BYTEA_CACHE = (dt == "bytea")
        return _EMAIL_ENC_IS_BYTEA_CACHE
    except Exception as e:
        # se não conseguir detectar, assume TEXT (não quebra inserção)
        log(f"detect email_enc type failed: {type(e).__name__}: {e}")
        _EMAIL_ENC_IS_BYTEA_CACHE = False
        return False
    finally:
        conn.close()


def email_enc_to_db(cfg: AppConfig, email: str) -> Union[str, bytes]:
    """
    Retorna valor para salvar em email_enc:
    - se coluna for BYTEA -> XOR bytes
    - se for TEXT -> e-mail normalizado
    """
    e = _norm_email(email)
    if _detect_email_enc_is_bytea(cfg):
        raw = e.encode("utf-8")
        key = hashlib.sha256(EMAIL_PEPPER).digest()
        return _xor_bytes(raw, key)
    return e


def email_enc_from_db(cfg: AppConfig, v: Union[None, str, bytes, memoryview]) -> str:
    """
    Converte valor vindo do banco para string de e-mail exibível.
    """
    if v is None:
        return ""
    if not _detect_email_enc_is_bytea(cfg):
        return str(v or "")

    # BYTEA -> decrypt XOR
    try:
        if isinstance(v, memoryview):
            raw = v.tobytes()
        elif isinstance(v, bytes):
            raw = v
        else:
            # se vier como string (raro), apenas não quebra
            return ""
        key = hashlib.sha256(EMAIL_PEPPER).digest()
        dec = _xor_bytes(raw, key)
        return dec.decode("utf-8", errors="ignore")
    except Exception:
        return ""


# ============================================================
# DB: USUÁRIOS / PROGRAMAS / PERMISSÕES
# ============================================================

# Programa protegido: o usuário logado precisa ter nível 3 nele
# Ajuste para o "codigo" correto no banco se você já tiver definido.
THIS_PROGRAMA_CODIGOS_CANDIDATOS = [
    "MODULO_ACESSO",
    "USUARIOS_ACESSO",
    "USUARIOS_ACES",  # Adicione o código que está no banco
    "USUARIOS",
    "ACESSOS",
]


def _fetch_programa_id_por_codigo(cfg: AppConfig, codigo: str) -> Optional[int]:
    sql = """
        SELECT pr."programaId"
          FROM "Ekenox"."programas" pr
         WHERE UPPER(COALESCE(pr."codigo",'')) = UPPER(%s)
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, (codigo.strip(),))
        r = cur.fetchone()
        return int(r[0]) if r else None
    finally:
        conn.close()


def _resolve_programa_id_protegido(cfg: AppConfig) -> Optional[Tuple[int, str]]:
    for cod in THIS_PROGRAMA_CODIGOS_CANDIDATOS:
        pid = _fetch_programa_id_por_codigo(cfg, cod)
        if pid is not None:
            return pid, cod
    return None


def _user_esta_ativo(cfg: AppConfig, usuario_id: int) -> bool:
    sql = 'SELECT COALESCE(u."ativo", true) FROM "Ekenox"."usuarios" u WHERE u."usuarioId"=%s LIMIT 1'
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, (int(usuario_id),))
        r = cur.fetchone()
        return bool(r[0]) if r else False
    finally:
        conn.close()


def _fetch_user_nivel(cfg: AppConfig, usuario_id: int, programa_id: int) -> Optional[int]:
    sql = """
        SELECT up."nivel"
          FROM "Ekenox"."usuario_programa" up
         WHERE up."usuarioId" = %s
           AND up."programaId" = %s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, (int(usuario_id), int(programa_id)))
        r = cur.fetchone()
        return int(r[0]) if r and r[0] is not None else None
    finally:
        conn.close()


def _check_acesso_admin(cfg: AppConfig, usuario_id: int) -> Tuple[bool, str]:
    if not _user_esta_ativo(cfg, usuario_id):
        return False, "Usuário inativo ou não encontrado."

    resolved = _resolve_programa_id_protegido(cfg)
    if not resolved:
        return False, (
            "Não encontrei o programa protegido na tabela Ekenox.programas.\n\n"
            "Crie um registro em Ekenox.programas com código igual a um destes:\n"
            f"{', '.join(THIS_PROGRAMA_CODIGOS_CANDIDATOS)}\n\n"
            "e dê nível 3 para o usuário admin."
        )

    programa_id, cod = resolved
    nivel = _fetch_user_nivel(cfg, usuario_id, programa_id)
    if nivel != 3:
        return False, f"Acesso negado. Necessário nível 3 (Admin) em '{cod}'. Seu nível: {nivel or 0}."
    return True, ""


def _fetch_usuario_id_by_email(cfg: AppConfig, email: str) -> int:
    eh = email_hash(email)
    sql = """
        SELECT u."usuarioId"
          FROM "Ekenox"."usuarios" u
         WHERE u."email_hash" = %s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, (eh,))
        row = cur.fetchone()
        if not row:
            raise ValueError("Usuário não encontrado para o e-mail informado.")
        return int(row[0])
    finally:
        conn.close()


def _parse_cli_user(cfg: AppConfig) -> int:
    """
    Espera receber o usuário logado via CLI:
      --usuario-id 123
    Opcional:
      --email usuario@dominio.com
    """
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--usuario-id", "--user-id",
                        "--uid", dest="usuario_id", type=int)
    parser.add_argument("--email", dest="email", type=str)
    args, _ = parser.parse_known_args(sys.argv[1:])

    if args.usuario_id:
        return int(args.usuario_id)

    if args.email:
        e = _norm_email(args.email)
        if not e:
            raise ValueError("E-mail vazio.")
        return _fetch_usuario_id_by_email(cfg, e)

    raise ValueError(
        "Usuário não informado. Abra este programa com --usuario-id <id> (recomendado).")


def _deny_and_exit(msg: str) -> None:
    r = tk.Tk()
    try:
        r.withdraw()
        apply_window_icon(r)
        messagebox.showerror("Acesso negado", msg, parent=r)
    finally:
        try:
            r.destroy()
        except Exception:
            pass
    raise SystemExit(1)


# ---------------- CRUD ----------------

def fetch_users(cfg: AppConfig, filtro: str = "") -> List[dict]:
    """
    Busca robusta:
      - se filtro for número -> usuarioId = filtro
      - se contiver @ -> email_hash = sha256(email)
      - sempre busca por nome ILIKE e email_hash ILIKE
    """
    termo = (filtro or "").strip()
    like = f"%{termo}%" if termo else None

    termo_id: Optional[int] = int(termo) if termo.isdigit() else None
    termo_email_hash: Optional[str] = email_hash(
        termo) if ("@" in termo) else None

    sql = """
        SELECT u."usuarioId",
               COALESCE(u."nome",'') AS nome,
               u."email_enc" AS email_enc,
               COALESCE(u."email_hash",'') AS email_hash,
               COALESCE(u."ativo", true) AS ativo
          FROM "Ekenox"."usuarios" u
         WHERE (%s IS NULL)
            OR (COALESCE(u."nome",'') ILIKE %s)
            OR (COALESCE(u."email_hash",'') ILIKE %s)
            OR (%s IS NOT NULL AND u."email_hash" = %s)
            OR (%s IS NOT NULL AND u."usuarioId" = %s)
         ORDER BY u."usuarioId" DESC
    """
    params = (
        (termo or None),
        like, like,
        termo_email_hash, termo_email_hash,
        termo_id, termo_id,
    )

    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, params)
        rows = cur.fetchall()

        out = []
        for r in rows:
            uid = int(r[0])
            nome = r[1] or ""
            email_enc = r[2]
            email = email_enc_from_db(cfg, email_enc)
            out.append({
                "usuarioId": uid,
                "nome": nome,
                "email": email,
                "email_hash": r[3] or "",
                "ativo": bool(r[4]),
            })
        return out
    finally:
        conn.close()


def insert_user(cfg: AppConfig, nome: str, email: str, ativo: bool) -> int:
    nome = (nome or "").strip()
    email_n = _norm_email(email)
    if not nome:
        raise ValueError("Nome obrigatório.")
    if not email_n:
        raise ValueError("E-mail obrigatório.")

    eh = email_hash(email_n)
    email_enc_val = email_enc_to_db(cfg, email_n)

    sql = """
        INSERT INTO "Ekenox"."usuarios"
            ("email_hash","email_enc","senha_hash","nome","ativo","criado_em","atualizado_em")
        VALUES
            (%s, %s, %s, %s, %s, NOW(), NOW())
        RETURNING "usuarioId"
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, (eh, email_enc_val, "", nome, bool(ativo)))
        uid = cur.fetchone()[0]
        conn.commit()
        return int(uid)
    finally:
        conn.close()


def update_user(cfg: AppConfig, usuario_id: int, nome: str, email: str, ativo: bool) -> None:
    nome = (nome or "").strip()
    email_n = _norm_email(email)
    if not nome:
        raise ValueError("Nome obrigatório.")
    if not email_n:
        raise ValueError("E-mail obrigatório.")

    eh = email_hash(email_n)
    email_enc_val = email_enc_to_db(cfg, email_n)

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
        cur = conn.cursor()
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
        cur = conn.cursor()
        cur.execute(sql, (rec, int(usuario_id)))
        if cur.rowcount != 1:
            raise RuntimeError("Usuário não encontrado para definir senha.")
        conn.commit()
    finally:
        conn.close()


def fetch_programs(cfg: AppConfig) -> List[dict]:
    sql = """
        SELECT pr."programaId", COALESCE(pr."codigo",''), COALESCE(pr."nome",'')
          FROM "Ekenox"."programas" pr
         ORDER BY pr."programaId"
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        out = []
        for r in rows:
            out.append({
                "programaId": int(r[0]),
                "codigo": (r[1] or "").strip(),
                "nome": (r[2] or "").strip(),
            })
        return out
    finally:
        conn.close()


def fetch_user_program_levels(cfg: AppConfig, usuario_id: int) -> Dict[int, int]:
    sql = """
        SELECT up."programaId", COALESCE(up."nivel", 1)
          FROM "Ekenox"."usuario_programa" up
         WHERE up."usuarioId" = %s
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, (int(usuario_id),))
        rows = cur.fetchall()
        return {int(r[0]): int(r[1]) for r in rows}
    finally:
        conn.close()


def save_user_permissions(cfg: AppConfig, usuario_id: int, perms: Dict[int, int]) -> None:
    """
    perms: {programaId: nivel}
    Estratégia: delete + insert em transação.
    """
    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(
            'DELETE FROM "Ekenox"."usuario_programa" WHERE "usuarioId" = %s',
            (int(usuario_id),)
        )
        for programa_id, nivel in perms.items():
            cur.execute(
                'INSERT INTO "Ekenox"."usuario_programa" ("usuarioId","programaId","nivel") VALUES (%s,%s,%s)',
                (int(usuario_id), int(programa_id), int(nivel)),
            )
        conn.commit()
    finally:
        conn.close()


# ============================================================
# UI
# ============================================================

NIVEL_LABEL = {
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Admin",
}


class TelaUsuariosApp(tk.Tk):
    def __init__(self, cfg: AppConfig) -> None:
        super().__init__()
        self.cfg = cfg

        self.title("Usuários e Programas - Ekenox")
        self.geometry("1100x650")
        self.minsize(1000, 600)
        apply_window_icon(self)

        ok, err = try_connect_db(self.cfg)
        if not ok:
            messagebox.showerror(
                "Banco de Dados", f"Falha ao conectar:\n{err}")
            # deixa abrir mesmo assim para ver tela, mas tudo vai falhar

        try:
            self.programs = fetch_programs(self.cfg)
        except Exception as e:
            self.programs = []
            messagebox.showerror("Erro", f"Falha ao listar programas:\n{e}")

        self.selected_user_id: Optional[int] = None
        self.program_rows: Dict[str, dict] = {}  # iid -> data

        self._build()
        self._load_users()

    # ---------------- BUILD ----------------

    def _build(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        # layout: esquerda users / direita detalhes
        left = ttk.Frame(root)
        right = ttk.Frame(root)
        left.pack(side="left", fill="y", padx=(0, 10))
        right.pack(side="left", fill="both", expand=True)

        # ---------- LEFT: USERS ----------
        ttk.Label(left, text="Usuários", font=(
            "Segoe UI", 11, "bold")).pack(anchor="w")

        row_search = ttk.Frame(left)
        row_search.pack(fill="x", pady=(6, 8))

        self.var_search = tk.StringVar()
        ent = ttk.Entry(row_search, textvariable=self.var_search, width=26)
        ent.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(row_search, text="Buscar",
                   command=self._load_users).pack(side="left")
        ttk.Button(row_search, text="Limpar", command=self._clear_search).pack(
            side="left", padx=(6, 0))
        ent.bind("<Return>", lambda e: self._load_users())

        cols = ("id", "nome", "email", "ativo")
        self.tree_users = ttk.Treeview(
            left, columns=cols, show="headings", height=26)
        self.tree_users.heading("id", text="ID")
        self.tree_users.heading("nome", text="Nome")
        self.tree_users.heading("email", text="E-mail")
        self.tree_users.heading("ativo", text="Ativo")

        self.tree_users.column("id", width=60, anchor="center")
        self.tree_users.column("nome", width=180, anchor="w")
        self.tree_users.column("email", width=240, anchor="w")
        self.tree_users.column("ativo", width=60, anchor="center")

        sb = ttk.Scrollbar(left, orient="vertical",
                           command=self.tree_users.yview)
        self.tree_users.configure(yscrollcommand=sb.set)
        self.tree_users.pack(side="left", fill="y")
        sb.pack(side="left", fill="y")

        self.tree_users.bind("<<TreeviewSelect>>", self._on_select_user)

        btns_left = ttk.Frame(left)
        btns_left.pack(fill="x", pady=(10, 0))
        ttk.Button(btns_left, text="Novo usuário",
                   command=self._new_user).pack(fill="x")
        ttk.Button(btns_left, text="Recarregar",
                   command=self._load_users).pack(fill="x", pady=(6, 0))

        # ---------- RIGHT: USER FORM ----------
        top = ttk.LabelFrame(right, text="Cadastro do usuário", padding=10)
        top.pack(fill="x")

        frm = ttk.Frame(top)
        frm.pack(fill="x")

        self.var_nome = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_ativo = tk.BooleanVar(value=True)

        ttk.Label(frm, text="Nome:").grid(row=0, column=0, sticky="w")
        self.ent_nome = ttk.Entry(frm, textvariable=self.var_nome, width=55)
        self.ent_nome.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        ttk.Label(frm, text="E-mail:").grid(row=1,
                                            column=0, sticky="w", pady=(8, 0))
        self.ent_email = ttk.Entry(frm, textvariable=self.var_email, width=55)
        self.ent_email.grid(row=1, column=1, sticky="ew",
                            padx=(6, 0), pady=(8, 0))

        chk = ttk.Checkbutton(frm, text="Ativo", variable=self.var_ativo)
        chk.grid(row=2, column=1, sticky="w", pady=(10, 0))

        frm.columnconfigure(1, weight=1)

        row_btns = ttk.Frame(top)
        row_btns.pack(fill="x", pady=(10, 0))

        ttk.Button(row_btns, text="Salvar usuário",
                   command=self._save_user).pack(side="left")
        ttk.Button(row_btns, text="Definir/Resetar senha",
                   command=self._open_set_password).pack(side="left", padx=(8, 0))
        ttk.Button(row_btns, text="Ativar/Desativar",
                   command=self._toggle_active).pack(side="left", padx=(8, 0))

        self.lbl_info = ttk.Label(top, text="", foreground="gray")
        self.lbl_info.pack(anchor="w", pady=(8, 0))

        # ---------- RIGHT: PROGRAMS ----------
        bottom = ttk.LabelFrame(
            right, text="Programas do usuário (duplo clique para editar)", padding=10)
        bottom.pack(fill="both", expand=True, pady=(10, 0))

        cols_p = ("permitido", "codigo", "nome", "nivel")
        self.tree_prog = ttk.Treeview(bottom, columns=cols_p, show="headings")
        self.tree_prog.heading("permitido", text="Permitido")
        self.tree_prog.heading("codigo", text="Código")
        self.tree_prog.heading("nome", text="Nome")
        self.tree_prog.heading("nivel", text="Nível")

        self.tree_prog.column("permitido", width=90, anchor="center")
        self.tree_prog.column("codigo", width=140, anchor="w")
        self.tree_prog.column("nome", width=420, anchor="w")
        self.tree_prog.column("nivel", width=120, anchor="center")

        sbp = ttk.Scrollbar(bottom, orient="vertical",
                            command=self.tree_prog.yview)
        self.tree_prog.configure(yscrollcommand=sbp.set)

        self.tree_prog.pack(side="left", fill="both", expand=True)
        sbp.pack(side="left", fill="y")

        self.tree_prog.bind("<Double-1>", self._on_double_click_program)

        row_perm = ttk.Frame(bottom)
        row_perm.pack(fill="x", pady=(10, 0))
        ttk.Button(row_perm, text="Salvar permissões",
                   command=self._save_permissions).pack(side="left")
        ttk.Button(row_perm, text="Marcar todos (permitir)",
                   command=lambda: self._set_all_permit(True)).pack(side="left", padx=(8, 0))
        ttk.Button(row_perm, text="Desmarcar todos", command=lambda: self._set_all_permit(
            False)).pack(side="left", padx=(8, 0))

        # atalhos
        self.ent_nome.bind("<Return>", lambda e: self.ent_email.focus_set())
        self.ent_email.bind("<Return>", lambda e: self._save_user())

        # carregar programas (vazio até selecionar user)
        self._render_programs_for_user(None)

    # ---------------- USERS ----------------

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
                "", "end",
                values=(u["usuarioId"], u["nome"], u["email"],
                        "Sim" if u["ativo"] else "Não")
            )

    def _new_user(self) -> None:
        self.selected_user_id = None
        self.var_nome.set("")
        self.var_email.set("")
        self.var_ativo.set(True)
        self.lbl_info.config(text="Novo usuário (ainda não salvo).")
        self._render_programs_for_user(None)
        self.ent_nome.focus_set()

    def _on_select_user(self, event=None) -> None:
        sel = self.tree_users.selection()
        if not sel:
            return
        vals = self.tree_users.item(sel[0], "values")
        if not vals:
            return
        usuario_id = int(vals[0])

        self.selected_user_id = usuario_id
        self.var_nome.set(str(vals[1]))
        self.var_email.set(str(vals[2]))
        self.var_ativo.set(str(vals[3]).strip().lower() == "sim")
        self.lbl_info.config(text=f"Editando usuário ID {usuario_id}")

        self._render_programs_for_user(usuario_id)

    def _save_user(self) -> None:
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
        if self.selected_user_id is None:
            messagebox.showwarning(
                "Usuário", "Selecione um usuário para ativar/desativar.")
            return
        self.var_ativo.set(not bool(self.var_ativo.get()))
        self._save_user()

    # ---------------- PASSWORD ----------------

    def _open_set_password(self) -> None:
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

        win.update_idletasks()
        x = (win.winfo_screenwidth() // 2) - (win.winfo_width() // 2)
        y = (win.winfo_screenheight() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

    # ---------------- PROGRAMS / PERMS ----------------

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
            allowed = pid in levels
            nivel = int(levels.get(pid, 1))
            iid = self.tree_prog.insert(
                "", "end",
                values=("Sim" if allowed else "Não",
                        p["codigo"], p["nome"], NIVEL_LABEL.get(nivel, str(nivel)))
            )
            self.program_rows[iid] = {
                "programaId": pid,
                "codigo": p["codigo"],
                "nome": p["nome"],
                "allowed": allowed,
                "nivel": nivel,
            }

    def _on_double_click_program(self, event) -> None:
        iid = self.tree_prog.identify_row(event.y)
        col = self.tree_prog.identify_column(
            event.x)  # "#1" permitido, "#4" nivel
        if not iid or iid not in self.program_rows:
            return

        data = self.program_rows[iid]

        if col == "#1":
            data["allowed"] = not data["allowed"]
        elif col == "#4":
            data["nivel"] = 1 if data["nivel"] >= 3 else (data["nivel"] + 1)
            data["allowed"] = True
        else:
            return

        self._refresh_program_row(iid)

    def _refresh_program_row(self, iid: str) -> None:
        d = self.program_rows[iid]
        self.tree_prog.item(
            iid,
            values=(
                "Sim" if d["allowed"] else "Não",
                d["codigo"],
                d["nome"],
                NIVEL_LABEL.get(int(d["nivel"]), str(d["nivel"])),
            ),
        )

    def _set_all_permit(self, allow: bool) -> None:
        """
        Corrige bug: seus botões chamavam isso, mas o método não existia.
        """
        for iid, d in self.program_rows.items():
            d["allowed"] = bool(allow)
            if d["allowed"]:
                if int(d.get("nivel") or 1) not in (1, 2, 3):
                    d["nivel"] = 1
            self._refresh_program_row(iid)

    def _save_permissions(self) -> None:
        if self.selected_user_id is None:
            messagebox.showwarning(
                "Permissões", "Selecione um usuário para salvar permissões.")
            return

        perms: Dict[int, int] = {}
        for d in self.program_rows.values():
            if d["allowed"]:
                nivel = int(d["nivel"]) if int(d["nivel"]) in (1, 2, 3) else 1
                perms[int(d["programaId"])] = nivel

        try:
            save_user_permissions(self.cfg, int(self.selected_user_id), perms)
            messagebox.showinfo("OK", "Permissões salvas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar permissões:\n{e}")


def main() -> None:
    cfg = load_config()

    # --- validações iniciais
    ok, err = try_connect_db(cfg)
    if not ok:
        # ainda mostra popup, mas não deixa seguir
        _deny_and_exit(f"Falha ao conectar no banco:\n{err}")

    # --- valida acesso admin (nível 3)
    try:
        usuario_id = _parse_cli_user(cfg)
    except Exception as e:
        _deny_and_exit(str(e))

    ok, msg = _check_acesso_admin(cfg, usuario_id)
    if not ok:
        _deny_and_exit(msg)

    # --- abre o app
    app = TelaUsuariosApp(cfg)
    app.mainloop()


if __name__ == "__main__":
    main()
