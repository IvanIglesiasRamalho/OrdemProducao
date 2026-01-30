from __future__ import annotations

"""
tela_usuarios_acesso.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2

Funções:
- Cadastrar usuário (nome, email, ativo) em "Ekenox"."usuarios"
- Definir/Resetar senha (PBKDF2-HMAC-SHA256) no formato salt$hash (compatível com menu_principal)
- Definir nível por programa em "Ekenox"."usuario_programa" (0 a 3)
- Lista de programas vem de "Ekenox"."programas"

CORREÇÕES IMPORTANTES:
- NÃO usa ILIKE em email_enc (email_enc é bytea)
- Busca por termo:
    * se for número -> busca por usuarioId
    * se contiver @ -> calcula email_hash e busca por igualdade
    * sempre busca por nome ILIKE e email_hash ILIKE
- Favicon: aplica no root via apply_window_icon()
"""

import base64
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

def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
BASE_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
os.makedirs(BASE_DIR, exist_ok=True)

LOG_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "tela_usuarios_acesso.log")


def log(msg: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


# ============================================================
# CONFIG
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
    except Exception as e:
        log(f"load_config error: {type(e).__name__}: {e}")
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
# FAVICON
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
            img = tk.PhotoImage(file=png)
            win.iconphoto(True, img)
            win._icon_img = img  # type: ignore[attr-defined]
    except Exception as e:
        log(f"apply_window_icon fatal: {type(e).__name__}: {e}")


# ============================================================
# EMAIL HASH + EMAIL_ENC (bytea)
# ============================================================

# chave para "obfuscar" email_enc (NÃO é criptografia forte; só para não ficar legível no banco)
APP_PEPPER = (os.getenv("EKENOX_EMAIL_PEPPER")
              or "EKENOX-EMAIL-PEPPER").encode("utf-8")


def normalize_email(email: str) -> str:
    return (email or "").strip().lower()


def email_hash(email: str) -> str:
    # COMPATÍVEL com menu_principal.py: sha256(email_normalizado)
    e = normalize_email(email).encode("utf-8")
    return hashlib.sha256(e).hexdigest()


def _xor_bytes(data: bytes, key: bytes) -> bytes:
    out = bytearray(len(data))
    for i, b in enumerate(data):
        out[i] = b ^ key[i % len(key)]
    return bytes(out)


def email_encrypt_to_bytea(email: str) -> bytes:
    e = normalize_email(email).encode("utf-8")
    key = hashlib.sha256(APP_PEPPER).digest()
    return _xor_bytes(e, key)  # bytes -> ideal para coluna BYTEA


def email_decrypt_from_bytea(v: Union[None, bytes, memoryview, str]) -> str:
    try:
        if v is None:
            return ""
        if isinstance(v, memoryview):
            raw = v.tobytes()
        elif isinstance(v, bytes):
            raw = v
        elif isinstance(v, str):
            # caso antigo (se alguém salvou base64 texto no campo)
            try:
                raw = base64.urlsafe_b64decode(v.encode("ascii"))
            except Exception:
                return ""
        else:
            return ""

        key = hashlib.sha256(APP_PEPPER).digest()
        dec = _xor_bytes(raw, key)
        return dec.decode("utf-8", errors="ignore")
    except Exception:
        return ""


# ============================================================
# SENHA (PBKDF2 salt$hash) - COMPATÍVEL COM menu_principal.py
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
# DB
# ============================================================

class Database:
    def __init__(self, cfg: AppConfig) -> None:
        self.cfg = cfg

    def connect(self):
        return psycopg2.connect(
            host=self.cfg.db_host,
            database=self.cfg.db_database,
            user=self.cfg.db_user,
            password=self.cfg.db_password,
            port=int(self.cfg.db_port),
            connect_timeout=5,
        )


USUARIOS_TABLE = '"Ekenox"."usuarios"'
PROGRAMAS_TABLE = '"Ekenox"."programas"'
USUARIO_PROGRAMA_TABLE = '"Ekenox"."usuario_programa"'


@dataclass
class Usuario:
    usuarioId: Optional[int]
    nome: str
    ativo: bool
    email_hash: str
    email_enc: Optional[Union[bytes, memoryview]]  # bytea
    senha_hash: Optional[str]


@dataclass
class Programa:
    programaId: int
    codigo: str
    nome: str


class Repo:
    def __init__(self, db: Database) -> None:
        self.db = db

    # ---------- programas ----------
    def listar_programas(self) -> List[Programa]:
        sql = f'SELECT "programaId","codigo","nome" FROM {PROGRAMAS_TABLE} ORDER BY "programaId"'
        with self.db.connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql)
                rows = cur.fetchall()
                return [Programa(int(r[0]), str(r[1] or ""), str(r[2] or "")) for r in rows]

    # ---------- usuarios ----------
    def listar_usuarios(self, termo: Optional[str] = None, limit: int = 500) -> List[Usuario]:
        termo = (termo or "").strip()
        like = f"%{termo}%" if termo else None

        termo_id: Optional[int] = None
        if termo.isdigit():
            termo_id = int(termo)

        termo_email_hash: Optional[str] = None
        if "@" in termo:
            termo_email_hash = email_hash(termo)

        # NOTE: NÃO filtrar email_enc (bytea) com ILIKE
        sql = f"""
            SELECT "usuarioId","nome","ativo","email_hash","email_enc","senha_hash"
              FROM {USUARIOS_TABLE}
             WHERE (%s IS NULL)
                OR (COALESCE("nome",'') ILIKE %s)
                OR (COALESCE("email_hash",'') ILIKE %s)
                OR (%s IS NOT NULL AND "email_hash" = %s)
                OR (%s IS NOT NULL AND "usuarioId" = %s)
             ORDER BY "nome"
             LIMIT %s
        """
        params = (
            (termo or None),
            like, like,
            termo_email_hash, termo_email_hash,
            termo_id, termo_id,
            int(limit),
        )

        with self.db.connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, params)
                rows = cur.fetchall()
                out: List[Usuario] = []
                for r in rows:
                    out.append(
                        Usuario(
                            usuarioId=int(r[0]) if r[0] is not None else None,
                            nome=str(r[1] or ""),
                            ativo=bool(r[2]),
                            email_hash=str(r[3] or ""),
                            # bytea -> memoryview
                            email_enc=(r[4] if r[4] is not None else None),
                            senha_hash=(str(r[5]) if r[5]
                                        is not None else None),
                        )
                    )
                return out

    def obter_usuario_por_id(self, usuario_id: int) -> Optional[Usuario]:
        sql = f"""
            SELECT "usuarioId","nome","ativo","email_hash","email_enc","senha_hash"
              FROM {USUARIOS_TABLE}
             WHERE "usuarioId" = %s
             LIMIT 1
        """
        with self.db.connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (int(usuario_id),))
                r = cur.fetchone()
                if not r:
                    return None
                return Usuario(
                    usuarioId=int(r[0]),
                    nome=str(r[1] or ""),
                    ativo=bool(r[2]),
                    email_hash=str(r[3] or ""),
                    email_enc=(r[4] if r[4] is not None else None),
                    senha_hash=(str(r[5]) if r[5] is not None else None),
                )

    def salvar_usuario(self, u: Usuario) -> int:
        if not u.nome.strip():
            raise ValueError("Nome é obrigatório.")
        if not u.email_hash.strip():
            raise ValueError("email_hash inválido.")

        # duplicidade de email_hash (quando alterando)
        sql_dup = f'SELECT 1 FROM {USUARIOS_TABLE} WHERE "email_hash" = %s AND COALESCE("usuarioId",0) <> %s'

        sql_ins = f"""
            INSERT INTO {USUARIOS_TABLE}
                ("nome","ativo","email_hash","email_enc","senha_hash","criado_em","atualizado_em")
            VALUES
                (%s,%s,%s,%s,%s,NOW(),NOW())
            RETURNING "usuarioId"
        """

        sql_upd = f"""
            UPDATE {USUARIOS_TABLE}
               SET "nome"=%s,
                   "ativo"=%s,
                   "email_hash"=%s,
                   "email_enc"=%s,
                   "atualizado_em"=NOW()
             WHERE "usuarioId"=%s
        """

        with self.db.connect() as conn:
            with conn.cursor() as cur:
                uid = int(u.usuarioId or 0)

                cur.execute(sql_dup, (u.email_hash, uid))
                if cur.fetchone():
                    raise ValueError(
                        "Já existe outro usuário com este e-mail (hash duplicado).")

                if u.usuarioId is None:
                    # senha_hash em branco no cadastro inicial
                    cur.execute(sql_ins, (u.nome, bool(u.ativo),
                                u.email_hash, u.email_enc, u.senha_hash or ""))
                    new_id = cur.fetchone()[0]
                    conn.commit()
                    return int(new_id)

                cur.execute(sql_upd, (u.nome, bool(u.ativo),
                            u.email_hash, u.email_enc, int(u.usuarioId)))
                if cur.rowcount != 1:
                    raise RuntimeError(
                        "Usuário não encontrado para atualizar.")
                conn.commit()
                return int(u.usuarioId)

    def atualizar_senha(self, usuario_id: int, senha_hash_nova: str) -> None:
        sql = f'UPDATE {USUARIOS_TABLE} SET "senha_hash"=%s, "atualizado_em"=NOW() WHERE "usuarioId"=%s'
        with self.db.connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (senha_hash_nova, int(usuario_id)))
                if cur.rowcount != 1:
                    raise RuntimeError(
                        "Usuário não encontrado para atualizar senha.")
                conn.commit()

    # ---------- acessos ----------
    def listar_acessos(self, usuario_id: int) -> Dict[int, int]:
        sql = f'SELECT "programaId","nivel" FROM {USUARIO_PROGRAMA_TABLE} WHERE "usuarioId"=%s'
        with self.db.connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql, (int(usuario_id),))
                rows = cur.fetchall()
                return {int(r[0]): int(r[1]) for r in rows}

    def salvar_acessos(self, usuario_id: int, niveis_por_programa: Dict[int, int]) -> None:
        # robusto: DELETE + INSERT (não depende de constraint unique)
        sql_del_all = f'DELETE FROM {USUARIO_PROGRAMA_TABLE} WHERE "usuarioId"=%s'
        sql_ins = f'INSERT INTO {USUARIO_PROGRAMA_TABLE} ("usuarioId","programaId","nivel") VALUES (%s,%s,%s)'

        with self.db.connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql_del_all, (int(usuario_id),))
                for pid, nivel in niveis_por_programa.items():
                    n = int(nivel)
                    if n <= 0:
                        continue
                    cur.execute(sql_ins, (int(usuario_id), int(pid), n))
                conn.commit()


# ============================================================
# UI AUX
# ============================================================

class PasswordDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc, title: str = "Definir senha"):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result: Optional[str] = None

        apply_window_icon(self)

        self.var1 = tk.StringVar()
        self.var2 = tk.StringVar()

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Nova senha:").grid(row=0, column=0, sticky="w")
        e1 = ttk.Entry(frm, textvariable=self.var1, show="*")
        e1.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(frm, text="Confirmar:").grid(
            row=1, column=0, sticky="w", pady=(8, 0))
        e2 = ttk.Entry(frm, textvariable=self.var2, show="*")
        e2.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))

        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancelar",
                   command=self._cancel).pack(side="right")
        ttk.Button(btns, text="OK", command=self._ok).pack(
            side="right", padx=(0, 8))

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())

        self.transient(master)
        self.grab_set()
        e1.focus_set()

    def _ok(self) -> None:
        s1 = self.var1.get()
        s2 = self.var2.get()
        if not s1:
            messagebox.showwarning("Senha", "Informe a nova senha.")
            return
        if len(s1) < 4:
            messagebox.showwarning(
                "Senha", "A senha deve ter pelo menos 4 caracteres.")
            return
        if s1 != s2:
            messagebox.showerror("Senha", "As senhas não conferem.")
            return
        self.result = s1
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


# ============================================================
# UI PRINCIPAL
# ============================================================

class TelaUsuariosAcesso(ttk.Frame):
    def __init__(self, master: tk.Misc, repo: Repo):
        super().__init__(master)
        self.repo = repo

        self.var_busca = tk.StringVar()
        self.var_usuario_id = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_ativo = tk.BooleanVar(value=True)

        self.programas: List[Programa] = []
        self._build_ui()
        self._load_programas()
        self._refresh_users()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=(10, 10, 10, 6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(
            top, text="Buscar usuário (nome / email / hash / ID):").grid(row=0, column=0, sticky="w")
        ent = ttk.Entry(top, textvariable=self.var_busca)
        ent.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ent.bind("<Return>", lambda e: self._refresh_users())

        ttk.Button(top, text="Atualizar",
                   command=self._refresh_users).grid(row=0, column=2)
        ttk.Button(top, text="Novo", command=self._novo).grid(
            row=0, column=3, padx=(6, 0))
        ttk.Button(top, text="Salvar", command=self._salvar).grid(
            row=0, column=4, padx=(6, 0))

        body = ttk.Panedwindow(self, orient="horizontal")
        body.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # esquerda: lista usuários
        left = ttk.Frame(body, padding=10)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        ttk.Label(left, text="Usuários").grid(row=0, column=0, sticky="w")
        self.lst_users = ttk.Treeview(left, columns=(
            "id", "nome", "ativo"), show="headings", selectmode="browse")
        self.lst_users.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.lst_users.heading("id", text="ID")
        self.lst_users.heading("nome", text="Nome")
        self.lst_users.heading("ativo", text="Ativo")
        self.lst_users.column("id", width=70, anchor="e", stretch=False)
        self.lst_users.column("nome", width=240, anchor="w", stretch=True)
        self.lst_users.column(
            "ativo", width=70, anchor="center", stretch=False)
        self.lst_users.bind("<<TreeviewSelect>>", self._on_select_user)

        vsb = ttk.Scrollbar(left, orient="vertical",
                            command=self.lst_users.yview)
        self.lst_users.configure(yscrollcommand=vsb.set)
        vsb.grid(row=1, column=1, sticky="ns", pady=(6, 0))

        body.add(left, weight=1)

        # direita: detalhes + acessos
        right = ttk.Frame(body, padding=10)
        right.columnconfigure(1, weight=1)
        right.rowconfigure(3, weight=1)

        ttk.Label(right, text="Cadastro").grid(
            row=0, column=0, columnspan=2, sticky="w")

        frm = ttk.LabelFrame(right, text="Usuário", padding=10)
        frm.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 10))
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="ID:").grid(row=0, column=0, sticky="w")
        self.e_id = ttk.Entry(frm, textvariable=self.var_usuario_id,
                              state="readonly", width=12)
        self.e_id.grid(row=0, column=1, sticky="w")

        ttk.Label(frm, text="Nome:").grid(
            row=1, column=0, sticky="w", pady=(8, 0))
        self.e_nome = ttk.Entry(frm, textvariable=self.var_nome)
        self.e_nome.grid(row=1, column=1, sticky="ew", pady=(8, 0))

        ttk.Label(frm, text="E-mail:").grid(row=2,
                                            column=0, sticky="w", pady=(8, 0))
        self.e_email = ttk.Entry(frm, textvariable=self.var_email)
        self.e_email.grid(row=2, column=1, sticky="ew", pady=(8, 0))

        chk = ttk.Checkbutton(frm, text="Ativo", variable=self.var_ativo)
        chk.grid(row=3, column=1, sticky="w", pady=(8, 0))

        btn_row = ttk.Frame(frm)
        btn_row.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btn_row, text="Definir/Resetar Senha...",
                   command=self._definir_senha).pack(side="right")

        ttk.Label(right, text="Acessos por programa (nível 0 a 3):").grid(
            row=2, column=0, columnspan=2, sticky="w")

        self.tree_acc = ttk.Treeview(
            right,
            columns=("programaId", "codigo", "programa", "nivel"),
            show="headings",
            selectmode="browse",
        )
        self.tree_acc.grid(row=3, column=0, columnspan=2,
                           sticky="nsew", pady=(6, 0))
        self.tree_acc.heading("programaId", text="ID")
        self.tree_acc.heading("codigo", text="Código")
        self.tree_acc.heading("programa", text="Programa")
        self.tree_acc.heading("nivel", text="Nível")
        self.tree_acc.column("programaId", width=60, anchor="e", stretch=False)
        self.tree_acc.column("codigo", width=120, anchor="w", stretch=False)
        self.tree_acc.column("programa", width=320, anchor="w", stretch=True)
        self.tree_acc.column("nivel", width=70, anchor="center", stretch=False)

        vsb2 = ttk.Scrollbar(right, orient="vertical",
                             command=self.tree_acc.yview)
        self.tree_acc.configure(yscrollcommand=vsb2.set)
        vsb2.grid(row=3, column=2, sticky="ns", pady=(6, 0))

        # editor simples de nível
        self.var_nivel = tk.StringVar(value="0")
        edit = ttk.Frame(right)
        edit.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Label(edit, text="Nível selecionado:").pack(side="left")
        cb = ttk.Combobox(edit, textvariable=self.var_nivel, values=[
                          "0", "1", "2", "3"], width=5, state="readonly")
        cb.pack(side="left", padx=(8, 8))
        ttk.Button(edit, text="Aplicar ao selecionado",
                   command=self._aplicar_nivel).pack(side="left")

        body.add(right, weight=3)

        self.e_nome.bind("<Return>", lambda e: self.e_email.focus_set())
        self.e_email.bind("<Return>", lambda e: self._salvar())

    def _load_programas(self) -> None:
        try:
            self.programas = self.repo.listar_programas()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar programas:\n{e}")
            self.programas = []
            return

        self._render_access_tree({})

    def _render_access_tree(self, niveis: Dict[int, int]) -> None:
        for it in self.tree_acc.get_children():
            self.tree_acc.delete(it)

        for p in self.programas:
            nivel = int(niveis.get(p.programaId, 0))
            self.tree_acc.insert("", "end", values=(
                p.programaId, p.codigo, p.nome, nivel))

    def _refresh_users(self) -> None:
        termo = self.var_busca.get().strip() or None
        for it in self.lst_users.get_children():
            self.lst_users.delete(it)

        try:
            usuarios = self.repo.listar_usuarios(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar usuários:\n{e}")
            return

        for u in usuarios:
            ativo = "Sim" if u.ativo else "Não"
            self.lst_users.insert("", "end", values=(
                u.usuarioId or "", u.nome, ativo))

    def _on_select_user(self, _evt=None) -> None:
        sel = self.lst_users.selection()
        if not sel:
            return
        vals = self.lst_users.item(sel[0], "values")
        uid = int(vals[0])

        u = self.repo.obter_usuario_por_id(uid)
        if not u:
            return

        self.var_usuario_id.set(str(u.usuarioId or ""))
        self.var_nome.set(u.nome)
        self.var_email.set(email_decrypt_from_bytea(
            u.email_enc) if u.email_enc else "")
        self.var_ativo.set(bool(u.ativo))

        acc = self.repo.listar_acessos(uid)
        self._render_access_tree(acc)

    def _novo(self) -> None:
        # limpa seleção da lista (opcional, mas ajuda)
        try:
            self.lst_users.selection_remove(self.lst_users.selection())
        except Exception:
            pass

        # ✅ deixa claro que ainda não existe ID
        self.var_usuario_id.set("(novo)")
        self.var_nome.set("")
        self.var_email.set("")
        self.var_ativo.set(True)
        self._render_access_tree({})

        try:
            self.e_nome.focus_set()  # ✅ vai direto para o nome
        except Exception:
            pass

    def _definir_senha(self) -> None:
        uid_s = self.var_usuario_id.get().strip()
        if not uid_s:
            messagebox.showwarning(
                "Senha", "Salve o usuário antes de definir senha.")
            return
        uid = int(uid_s)

        dlg = PasswordDialog(self.winfo_toplevel(), "Definir/Resetar senha")
        self.wait_window(dlg)
        if not dlg.result:
            return

        try:
            # ✅ compatível com menu_principal
            sh = make_password_record(dlg.result)
            self.repo.atualizar_senha(uid, sh)
            messagebox.showinfo("OK", "Senha atualizada com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao atualizar senha:\n{e}")

    def _aplicar_nivel(self) -> None:
        sel = self.tree_acc.selection()
        if not sel:
            messagebox.showwarning("Acesso", "Selecione um programa na lista.")
            return

        try:
            n = int(self.var_nivel.get().strip())
        except ValueError:
            n = 0

        if n < 0 or n > 3:
            messagebox.showwarning("Acesso", "Nível deve ser 0 a 3.")
            return

        vals = list(self.tree_acc.item(sel[0], "values"))
        vals[3] = n
        self.tree_acc.item(sel[0], values=vals)

    def _salvar(self) -> None:
        nome = self.var_nome.get().strip()
        email = self.var_email.get().strip()
        ativo = bool(self.var_ativo.get())

        if not nome:
            messagebox.showwarning("Validação", "Nome é obrigatório.")
            return
        if not email or "@" not in email:
            messagebox.showwarning("Validação", "Informe um e-mail válido.")
            return

        uid_s = self.var_usuario_id.get().strip()
        if not uid_s or uid_s.lower() == "(novo)":
            uid = None
        else:
            uid = int(uid_s)

        u = Usuario(
            usuarioId=uid,
            nome=nome,
            ativo=ativo,
            email_hash=email_hash(email),
            email_enc=email_encrypt_to_bytea(email),  # ✅ bytea
            senha_hash=None,  # só altera na tela de senha
        )

        try:
            new_id = self.repo.salvar_usuario(u)
            self.var_usuario_id.set(str(new_id))
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar usuário:\n{e}")
            return

        # salvar acessos
        try:
            niveis: Dict[int, int] = {}
            for it in self.tree_acc.get_children():
                pid, _cod, _nome, nivel = self.tree_acc.item(it, "values")
                niveis[int(pid)] = int(nivel)
            self.repo.salvar_acessos(int(self.var_usuario_id.get()), niveis)
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Usuário salvo, mas falhou ao salvar acessos:\n{e}")
            return

        messagebox.showinfo("OK", "Usuário e acessos salvos.")
        self._refresh_users()


def main() -> None:
    cfg = env_override(load_config())

    root = tk.Tk()
    root.title("Usuários e Acessos - Ekenox")
    root.geometry("1100x650")
    apply_window_icon(root)  # ✅ favicon corrigido

    # tema
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    repo = Repo(Database(cfg))
    tela = TelaUsuariosAcesso(root, repo)
    tela.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
