from __future__ import annotations

import hashlib
import json
import os
import secrets
import subprocess
import sys
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Optional, Tuple

import psycopg2


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
APP_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
os.makedirs(BASE_DIR, exist_ok=True)

LOG_PATH = os.path.join(BASE_DIR, "logs")
os.makedirs(LOG_PATH, exist_ok=True)
LOG_FILE = os.path.join(LOG_PATH, "reset_senha.log")

CONFIG_FILE = os.path.join(BASE_DIR, "config_op.json")


def log(msg: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


@dataclass
class AppConfig:
    db_host: str = "10.0.0.154"
    db_database: str = "postgresekenox"
    db_user: str = "postgresekenox"
    db_password: str = "Ekenox5426"
    db_port: int = 55432


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
# CONFIG LOAD / DB CONNECT
# ============================================================

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
# HASH / SENHA
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
# DB: BUSCA / UPDATE
# ============================================================

def fetch_user_by_email(cfg: AppConfig, email: str) -> Optional[dict]:
    eh = email_hash(email)
    sql = """
        SELECT u."usuarioId", u."nome", u."ativo"
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
            return None
        return {"usuarioId": int(row[0]), "nome": row[1] or "", "ativo": bool(row[2])}
    finally:
        conn.close()


def set_user_password(cfg: AppConfig, usuario_id: int, new_password: str) -> None:
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
            raise RuntimeError(
                "Nenhuma linha foi atualizada (usuarioId não encontrado).")
        conn.commit()
    finally:
        conn.close()


# ============================================================
# VOLTAR AO MENU PRINCIPAL
# ============================================================

def pick_python_gui_cmd() -> list[str]:
    if os.name != "nt":
        return [sys.executable]
    for cmd in (["pyw", "-3.12"], ["pyw"], ["pythonw"], ["python"]):
        try:
            subprocess.run(cmd + ["-c", "print('ok')"],
                           capture_output=True, text=True, timeout=2)
            return cmd
        except Exception:
            continue
    return ["python"]


def build_run_cmd(script_path: str, extra_args: Optional[list[str]] = None) -> list[str]:
    extra_args = extra_args or []
    if script_path.lower().endswith(".exe"):
        return [script_path] + extra_args
    if os.name == "nt" and script_path.lower().endswith(".py"):
        return pick_python_gui_cmd() + [script_path] + extra_args
    return [sys.executable, script_path] + extra_args


def find_menu_principal() -> Optional[str]:
    candidates = [
        os.path.join(APP_DIR, "menu_principal.py"),
        os.path.join(BASE_DIR, "menu_principal.py"),
        os.path.join(APP_DIR, "menu_principal.exe"),
        os.path.join(BASE_DIR, "menu_principal.exe"),
    ]
    return next((p for p in candidates if os.path.isfile(p)), None)


def voltar_menu_principal() -> None:
    mp = find_menu_principal()
    if not mp:
        messagebox.showwarning(
            "Menu Principal",
            "Não encontrei menu_principal.py / menu_principal.exe para retornar.\n"
            "Verifique se ele está em APP_DIR ou BASE_DIR."
        )
        return
    try:
        cmd = build_run_cmd(mp, extra_args=["--skip-entrada"])
        subprocess.Popen(cmd, cwd=os.path.dirname(mp) or APP_DIR)
    except Exception as e:
        log(f"Falha ao voltar ao menu_principal: {type(e).__name__}: {e}")
        messagebox.showerror(
            "Menu Principal", f"Falha ao abrir menu_principal:\n{e}")


# ============================================================
# UI
# ============================================================

class ResetSenhaApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        apply_window_icon(self)
        self.cfg = load_config()

        self.title("Ekenox - Reset de Senha (Admin)")
        self.resizable(False, False)
        self.geometry("460x360")

        ok, err = try_connect_db(self.cfg)
        self.db_ok = ok
        self.db_err = err

        self.admin_code_expected = "Ekenox-Admin"

        self.protocol("WM_DELETE_WINDOW", self._exit_to_menu)

        self._build()

    def _build(self) -> None:
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Reset de senha (sem senha atual)",
                  font=("Segoe UI", 12, "bold")).pack(anchor="w")

        status_txt = "Conectado ao banco" if self.db_ok else f"ERRO BD: {self.db_err}"
        ttk.Label(frm, text=status_txt,
                  foreground=("green" if self.db_ok else "red")).pack(anchor="w", pady=(4, 12))

        ttk.Label(frm, text="Código administrador:",
                  font=("Segoe UI", 10)).pack(anchor="w")
        self.var_admin = tk.StringVar()
        self.ent_admin = ttk.Entry(frm, textvariable=self.var_admin, show="*")
        self.ent_admin.pack(fill="x", pady=(4, 12))

        ttk.Label(frm, text="E-mail do usuário:",
                  font=("Segoe UI", 10)).pack(anchor="w")
        self.var_email = tk.StringVar()
        self.ent_email = ttk.Entry(frm, textvariable=self.var_email)
        self.ent_email.pack(fill="x", pady=(4, 12))

        ttk.Label(frm, text="Nova senha:", font=(
            "Segoe UI", 10)).pack(anchor="w")
        self.var_p1 = tk.StringVar()
        self.ent_p1 = ttk.Entry(frm, textvariable=self.var_p1, show="*")
        self.ent_p1.pack(fill="x", pady=(4, 10))

        ttk.Label(frm, text="Confirmar nova senha:",
                  font=("Segoe UI", 10)).pack(anchor="w")
        self.var_p2 = tk.StringVar()
        self.ent_p2 = ttk.Entry(frm, textvariable=self.var_p2, show="*")
        self.ent_p2.pack(fill="x", pady=(4, 14))

        row = ttk.Frame(frm)
        row.pack(fill="x")

        self.btn_reset = ttk.Button(
            row, text="Alterar senha", command=self._do_reset)
        self.btn_reset.pack(side="left", expand=True, fill="x", padx=(0, 8))

        ttk.Button(row, text="Voltar ao Menu", command=self._exit_to_menu).pack(
            side="left", expand=True, fill="x"
        )

        self.bind("<Return>", lambda e: self._do_reset())
        self.bind("<Escape>", lambda e: self._exit_to_menu())
        self.ent_admin.focus_set()

        if not self.db_ok:
            self.btn_reset.state(["disabled"])

    def _exit_to_menu(self) -> None:
        try:
            self.destroy()
        finally:
            voltar_menu_principal()

    def _do_reset(self) -> None:
        if not self.db_ok:
            messagebox.showerror("Banco", "Sem conexão com o banco.")
            return

        admin = (self.var_admin.get() or "").strip()
        email = _norm_email(self.var_email.get())
        p1 = self.var_p1.get()
        p2 = self.var_p2.get()

        if admin != self.admin_code_expected:
            messagebox.showerror("Permissão", "Código administrador inválido.")
            self.ent_admin.focus_set()
            return

        if not email:
            messagebox.showwarning("Entrada", "Informe o e-mail do usuário.")
            self.ent_email.focus_set()
            return

        if len(p1) < 4:
            messagebox.showwarning(
                "Senha", "A senha deve ter pelo menos 4 caracteres.")
            self.ent_p1.focus_set()
            return

        if p1 != p2:
            messagebox.showwarning("Senha", "As senhas não conferem.")
            self.ent_p2.focus_set()
            return

        try:
            u = fetch_user_by_email(self.cfg, email)
            if not u:
                messagebox.showerror("Usuário", "Usuário não encontrado.")
                return
            if not u.get("ativo", True):
                messagebox.showerror("Usuário", "Usuário inativo.")
                return

            set_user_password(self.cfg, int(u["usuarioId"]), p1)

            self.var_p1.set("")
            self.var_p2.set("")
            messagebox.showinfo(
                "OK", f"Senha alterada com sucesso.\n\nUsuário: {u.get('nome') or u['usuarioId']}")

        except Exception as e:
            log(f"reset error: {type(e).__name__}: {e}")
            messagebox.showerror("Erro", f"Falha ao alterar senha:\n{e}")


if __name__ == "__main__":
    app = ResetSenhaApp()
    app.mainloop()
