from __future__ import annotations

"""
menu_principal.py
Python 3.12+ | Tkinter | Postgres 16+ | psycopg2

OBJETIVO:
- Mostrar a tela "Ekenox - Entrada" (login: e-mail + senha) SOMENTE na primeira abertura normal.
- Ao reabrir o menu voltando de outro programa (ex: tela_produtos.py), NÃO mostrar a entrada.

MECANISMOS:
1) CLI:
   python menu_principal.py --skip-entrada
   -> abre direto no menu (sem entrada)

2) Flag de sessão (arquivo):
   Ao logar com sucesso, cria BASE_DIR\\.session_skip_entrada
   Enquanto esse arquivo existir, não mostra Entrada.

Quando o usuário "Sair" (confirmando), a flag é removida.

NOTA (Windows):
- Não usar withdraw() no root para controlar Toplevel (pode "sumir" janela).
- Root começa com alpha=0.0 (invisível, mas existente); depois volta alpha=1.0 ao mostrar menu.
"""

import hashlib
import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple

import psycopg2


# ============================================================
# PATHS
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


def log(msg: str) -> None:
    _log_write("menu_principal.log", msg)


# ============================================================
# FLAG DE SESSÃO (pula entrada após login OK)
# ============================================================

SESSION_FLAG = os.path.join(BASE_DIR, ".session_skip_entrada")


def set_session_skip_entrada() -> None:
    try:
        with open(SESSION_FLAG, "w", encoding="utf-8") as f:
            f.write("1")
    except Exception:
        pass


def clear_session_skip_entrada() -> None:
    try:
        if os.path.exists(SESSION_FLAG):
            os.remove(SESSION_FLAG)
    except Exception:
        pass


def has_session_skip_entrada() -> bool:
    try:
        return os.path.exists(SESSION_FLAG)
    except Exception:
        return False


# ============================================================
# ÍCONE / FAVICON
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
    Tenta aplicar .ico (Windows) e, se falhar, tenta .png.
    Mantém referência do PhotoImage no objeto para não ser coletado.
    """
    try:
        ico = find_icon_path()
        if ico:
            try:
                # no Windows, isso é o que costuma funcionar melhor
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


def load_png_icon(rel_path: str, max_side: int = 48) -> Optional[tk.PhotoImage]:
    candidates = [
        os.path.join(BASE_DIR, rel_path),
        os.path.join(APP_DIR, rel_path),
    ]
    path = next((p for p in candidates if os.path.isfile(p)), None)
    if not path:
        return None
    try:
        img = tk.PhotoImage(file=path)
        w, h = img.width(), img.height()
        maior = max(w, h)
        if maior > max_side:
            fator = max(1, int(maior / max_side))
            img = img.subsample(fator, fator)
        return img
    except Exception:
        return None


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
# SEGURANÇA: email_hash + senha_hash (salt$hash PBKDF2)
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


def verify_password(password: str, stored: str) -> bool:
    try:
        if not stored or "$" not in stored:
            return False
        salt, ph = stored.split("$", 1)
        return hash_password(password, salt) == ph
    except Exception:
        return False


def fetch_user_by_email(cfg: AppConfig, email: str) -> Optional[dict]:
    eh = email_hash(email)
    sql = """
        SELECT u."usuarioId", u."email_hash", u."senha_hash", u."nome", u."ativo"
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
        return {
            "usuarioId": int(row[0]),
            "email_hash": row[1] or "",
            "senha_hash": row[2] or "",
            "nome": row[3] or "",
            "ativo": bool(row[4]),
        }
    finally:
        conn.close()


# ============================================================
# PERMISSÃO (NÍVEL POR PROGRAMA) — exige NÍVEL 3
# ============================================================

def fetch_user_max_nivel_for_program_keys(
    cfg: AppConfig,
    usuario_id: int,
    program_keys: List[str],
) -> int:
    """
    Retorna o MAIOR nível encontrado para o usuário em qualquer programa que
    combine com uma das chaves (por codigo/nome).
    Se não encontrar, retorna 0.
    """
    keys = [k.strip() for k in (program_keys or []) if (k or "").strip()]
    if not keys:
        return 0

    # Monta filtros OR: lower(codigo)=lower(%s) OR lower(nome)=lower(%s) OR nome ILIKE %s ...
    clauses: List[str] = []
    params: List[Any] = [int(usuario_id)]

    for k in keys:
        clauses.append('LOWER(COALESCE(p."codigo", \'\')) = LOWER(%s)')
        params.append(k)
        clauses.append('LOWER(COALESCE(p."nome", \'\')) = LOWER(%s)')
        params.append(k)
        clauses.append('COALESCE(p."nome", \'\') ILIKE %s')
        params.append(f"%{k}%")

    where_program = " OR ".join(f"({c})" for c in clauses)

    sql = f"""
        SELECT COALESCE(MAX(COALESCE(up."nivel",0)), 0) AS nivel
          FROM "Ekenox"."usuario_programa" up
          JOIN "Ekenox"."programas" p ON p."programaId" = up."programaId"
         WHERE up."usuarioId" = %s
           AND ({where_program})
    """

    conn = db_connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(sql, tuple(params))
        row = cur.fetchone()
        return int(row[0] or 0) if row else 0
    finally:
        conn.close()


# ============================================================
# PYTHON GUI CMD (para abrir scripts .py sem console)
# ============================================================

def pick_python_gui_cmd() -> List[str]:
    if os.name != "nt":
        return [sys.executable]

    candidates = [
        ["pyw", "-3.12"],
        ["pyw"],
        ["pythonw"],
        ["python"],
    ]
    for cmd in candidates:
        try:
            subprocess.run(cmd + ["-c", "print('ok')"],
                           capture_output=True, text=True, timeout=2)
            return cmd
        except Exception:
            continue
    return ["python"]


def build_run_cmd(script_path: str, extra_args: Optional[List[str]] = None) -> List[str]:
    extra_args = extra_args or []

    if script_path.lower().endswith(".exe"):
        return [script_path] + extra_args

    if os.name == "nt" and script_path.lower().endswith(".py"):
        return pick_python_gui_cmd() + [script_path] + extra_args

    return [sys.executable, script_path] + extra_args


# ============================================================
# SPLASH
# ============================================================

class Splash(tk.Toplevel):
    def __init__(self, master: tk.Tk, titulo: str = "Inicializando"):
        super().__init__(master)
        self.title(titulo)
        self.geometry("430x170")
        self.resizable(False, False)
        self.transient(master)
        apply_window_icon(self)

        self.lbl = ttk.Label(self, text="Iniciando...",
                             font=("Segoe UI", 11, "bold"))
        self.lbl.pack(pady=(18, 8))

        self.pb = ttk.Progressbar(self, mode="indeterminate")
        self.pb.pack(fill="x", padx=18, pady=8)

        self.info = ttk.Label(self, text="", foreground="gray")
        self.info.pack(pady=(8, 0))

        self.pb.start(10)
        self._center()

        try:
            self.lift()
            self.attributes("-topmost", True)
            self.after(250, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

    def set_text(self, t: str) -> None:
        self.lbl.config(text=t)
        self.update_idletasks()

    def set_info(self, t: str) -> None:
        self.info.config(text=t)
        self.update_idletasks()

    def _center(self) -> None:
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")


# ============================================================
# PROGRAMAS (DIVIDIDO EM: Cadastro / Movimento)
# ============================================================

PROGRAMAS: List[Dict[str, Any]] = [
    # ---- CADASTRO ----
    {"menu": "Cadastro", "nome": "Produtos", "icone": "imagens/menu_produtos.png",
     "arquivo": os.path.join(APP_DIR, "tela_produtos.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Informação Produto", "icone": "imagens/menu_infoproduto.png",
     "arquivo": os.path.join(APP_DIR, "tela_info_produto.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Arranjo", "icone": "imagens/menu_arranjo.png",
     "arquivo": os.path.join(APP_DIR, "tela_arranjo.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Categoria", "icone": "imagens/menu_categoria.png",
     "arquivo": os.path.join(APP_DIR, "tela_categoria.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Depósito", "icone": "imagens/menu_deposito.png",
     "arquivo": os.path.join(APP_DIR, "tela_deposito.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Estoque", "icone": "imagens/menu_estoque.png",
     "arquivo": os.path.join(APP_DIR, "tela_estoque.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Estrutura", "icone": "imagens/menu_estrutura.png",
     "arquivo": os.path.join(APP_DIR, "tela_estrutura.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Fornecedor", "icone": "imagens/menu_fornecedor.png",
     "arquivo": os.path.join(APP_DIR, "tela_fornecedor.py"), "tipo": "py"},
    {"menu": "Cadastro", "nome": "Situação", "icone": "imagens/menu_situacao.png",
     "arquivo": os.path.join(APP_DIR, "tela_situacao.py"), "tipo": "py"},

    # ---- MOVIMENTO ----
    {"menu": "Movimento", "nome": "Ordem de Produção", "icone": "imagens/menu_op.png",
     "arquivo": os.path.join(APP_DIR, "Ordem_Producao.py"), "tipo": "py"},
]


# ============================================================
# APP
# ============================================================

class MenuPrincipalApp(tk.Tk):
    def __init__(self, *, show_entrada: bool) -> None:
        super().__init__()

        self.cfg = load_config()
        self.connected = False
        self.db_err = ""

        self.show_entrada = bool(show_entrada) and (
            not has_session_skip_entrada())

        self.title("Menu Principal - Ekenox")
        self.geometry("1100x650")
        self.minsize(1100, 650)
        apply_window_icon(self)

        self._closing = False
        self._ui_built = False
        self._icons: Dict[str, tk.PhotoImage] = {}

        self.user: Optional[dict] = None

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # ✅ não esconder Toplevel no Windows
        try:
            self.deiconify()
            self.attributes("-alpha", 0.0)
        except Exception:
            self.withdraw()

        self.after(80, self._startup)

    # ---------------- permissões ----------------

    def _require_nivel3(self, program_keys: List[str]) -> bool:
        if not self.user or not self.user.get("usuarioId"):
            messagebox.showerror("Permissão", "Faça login novamente.")
            return False

        try:
            uid = int(self.user["usuarioId"])
        except Exception:
            messagebox.showerror(
                "Permissão", "Usuário inválido. Faça login novamente.")
            return False

        try:
            nivel = fetch_user_max_nivel_for_program_keys(
                self.cfg, uid, program_keys)
        except Exception as e:
            log(f"Falha ao validar permissão: {type(e).__name__}: {e}")
            messagebox.showerror(
                "Permissão", f"Falha ao validar permissão:\n{e}")
            return False

        if nivel < 3:
            messagebox.showerror(
                "Permissão",
                "Acesso negado.\n\nSomente usuários com NÍVEL 3 (Admin) podem abrir este programa."
            )
            return False

        return True

    # ---------------- abrir programas do sistema ----------------

    def _open_reset_senha(self) -> None:
        """Abre o programa reset_senha.py (ou exe) — protegido por NÍVEL 3."""
        # (opcional) protege reset senha também
        if not self._require_nivel3(["RESET_SENHA", "Reset", "Senha", "Admin"]):
            return

        candidates = [
            os.path.join(APP_DIR, "reset_senha.py"),
            os.path.join(APP_DIR, "reset_senha.exe"),
            os.path.join(BASE_DIR, "reset_senha.py"),
            os.path.join(BASE_DIR, "reset_senha.exe"),
        ]
        arquivo = next((p for p in candidates if os.path.isfile(p)), None)

        if not arquivo:
            messagebox.showerror(
                "Reset de Senha",
                "Não encontrei o arquivo reset_senha.py / reset_senha.exe.\n\n"
                "Coloque o arquivo na mesma pasta do menu_principal.py (APP_DIR) ou em BASE_DIR."
            )
            return

        try:
            cwd = os.path.dirname(arquivo) or APP_DIR
            if arquivo.lower().endswith(".exe"):
                subprocess.Popen([arquivo], cwd=cwd)
            else:
                cmd = build_run_cmd(arquivo)
                subprocess.Popen(cmd, cwd=cwd)

            self.after(150, self.destroy)

        except Exception as e:
            log(f"Falha ao abrir reset_senha: {type(e).__name__}: {e}")
            messagebox.showerror("Reset de Senha", f"Falha ao abrir:\n{e}")

    def _open_usuarios_acesso(self) -> None:
        """Abre tela_usuarios_acesso.py/.exe — protegido por NÍVEL 3."""
        # Ajuste aqui caso o programa no banco tenha outro codigo/nome
        PROGRAM_KEYS_USUARIOS_ACESSO = [
            "USUARIOS_ACESSO",
            "USUARIOS",
            "ACESSO",
            "Usuários e Acessos",
            "Usuarios e Acessos",
        ]

        if not self._require_nivel3(PROGRAM_KEYS_USUARIOS_ACESSO):
            return

        candidates = [
            os.path.join(APP_DIR, "tela_usuarios_acesso.py"),
            os.path.join(APP_DIR, "tela_usuarios_acesso.exe"),
            os.path.join(BASE_DIR, "tela_usuarios_acesso.py"),
            os.path.join(BASE_DIR, "tela_usuarios_acesso.exe"),
        ]
        arquivo = next((p for p in candidates if os.path.isfile(p)), None)

        if not arquivo:
            messagebox.showerror(
                "Usuários e Acessos",
                "Não encontrei o arquivo tela_usuarios_acesso.py / tela_usuarios_acesso.exe.\n\n"
                "Coloque o arquivo na mesma pasta do menu_principal.py (APP_DIR) ou em BASE_DIR."
            )
            return

        # Se quiser passar o usuário logado para a tela (opcional):
        extra_args: List[str] = []
        try:
            if self.user and self.user.get("usuarioId"):
                extra_args = ["--usuario-id", str(int(self.user["usuarioId"]))]
        except Exception:
            extra_args = []

        try:
            cwd = os.path.dirname(arquivo) or APP_DIR
            if arquivo.lower().endswith(".exe"):
                subprocess.Popen([arquivo] + extra_args, cwd=cwd)
            else:
                cmd = build_run_cmd(arquivo, extra_args=extra_args)
                subprocess.Popen(cmd, cwd=cwd)

            self.after(150, self.destroy)

        except Exception as e:
            log(
                f"Falha ao abrir tela_usuarios_acesso: {type(e).__name__}: {e}")
            messagebox.showerror("Usuários e Acessos", f"Falha ao abrir:\n{e}")

    # ---------------- STARTUP ----------------

    def _startup(self) -> None:
        splash = Splash(self, "Inicializando")
        splash.set_text("Conectando ao banco...")

        def worker() -> None:
            ok, err = try_connect_db(self.cfg)

            def finish() -> None:
                self.connected = bool(ok)
                self.db_err = str(err)

                try:
                    splash.set_info("Conectado" if ok else (
                        self.db_err[:70] + ("..." if len(self.db_err) > 70 else "")))
                    splash.set_text("Carregando menu...")
                    self._build_ui_once()
                finally:
                    try:
                        splash.destroy()
                    except Exception:
                        pass

                if self.show_entrada:
                    self._show_login_dialog()
                else:
                    self._show_menu()

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _build_ui_once(self) -> None:
        if self._ui_built:
            return
        self._ui_built = True
        self._build_ui()

    def _show_menu(self) -> None:
        if self._closing:
            return

        try:
            self.attributes("-alpha", 1.0)
        except Exception:
            pass

        self.deiconify()
        self.lift()
        try:
            self.focus_force()
        except Exception:
            pass

    # ---------------- LOGIN ----------------

    def _center_window(self, win: tk.Toplevel) -> None:
        win.update_idletasks()
        w, h = win.winfo_width(), win.winfo_height()
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{x}+{y}")

    def _bring_to_front(self, win: tk.Toplevel) -> None:
        try:
            win.update_idletasks()
            win.lift()
            win.attributes("-topmost", True)
            win.after(250, lambda: win.attributes("-topmost", False))
            win.focus_force()
        except Exception:
            pass

    def _show_login_dialog(self) -> None:
        if self._closing:
            return

        if has_session_skip_entrada():
            self._show_menu()
            return

        tela = tk.Toplevel(self)
        apply_window_icon(tela)
        tela.title("Ekenox - Entrada")
        tela.resizable(False, False)
        tela.configure(bg="#121212")

        # ✅ Ajustado para caber tudo
        tela.geometry("420x500")

        tela.transient(self)
        tela.grab_set()

        def do_cancel() -> None:
            try:
                tela.grab_release()
            except Exception:
                pass
            tela.destroy()
            self.on_closing(force=True)

        tela.protocol("WM_DELETE_WINDOW", do_cancel)

        frame = tk.Frame(tela, bg="#121212", padx=16, pady=12)
        frame.pack(fill="both", expand=True)

        candidatos = [
            os.path.join(BASE_DIR, "imagens", "avatar_ekenox.png"),
            os.path.join(BASE_DIR, "avatar_ekenox.png"),
            os.path.join(BASE_DIR, "imagens", "Ekenox.png"),
            os.path.join(BASE_DIR, "Ekenox.png"),
        ]
        avatar_path = next((p for p in candidatos if os.path.isfile(p)), None)
        avatar_img = None
        if avatar_path:
            try:
                img = tk.PhotoImage(file=avatar_path)
                w, h = img.width(), img.height()
                maior = max(w, h)
                max_logo = 160  # ✅ menor para caber
                if maior > max_logo:
                    fator = max(1, int(maior / max_logo))
                    img = img.subsample(fator, fator)
                avatar_img = img
            except Exception:
                avatar_img = None

        if avatar_img:
            lbl_img = tk.Label(frame, image=avatar_img, bg="#121212")
            lbl_img.image = avatar_img
            lbl_img.pack(pady=(0, 8))
        else:
            tk.Label(frame, text="EKENOX", bg="#121212", fg="#ff9f1a",
                     font=("Segoe UI", 18, "bold")).pack(pady=(0, 8))

        tk.Label(frame, text="Sistema - Menu Principal", bg="#121212",
                 fg="#ffffff", font=("Segoe UI", 11, "bold")).pack()

        status = "Conectado ao banco" if self.connected else f"ERRO BD: {self.db_err}"
        tk.Label(frame, text=status, bg="#121212",
                 fg=("#34d399" if self.connected else "#f87171"),
                 font=("Segoe UI", 9, "bold")).pack(pady=(6, 12))

        form = tk.Frame(frame, bg="#121212")
        form.pack(fill="x")

        var_email = tk.StringVar()
        var_senha = tk.StringVar()

        tk.Label(form, text="E-mail:", bg="#121212", fg="#e5e7eb",
                 font=("Segoe UI", 10)).pack(anchor="w")
        ent_email = tk.Entry(
            form, textvariable=var_email,
            bg="#0b0b0b", fg="#ffffff", insertbackground="#ffffff",
            relief="flat", highlightthickness=1,
            highlightbackground="#2a2a2a", highlightcolor="#ff9f1a",
            font=("Segoe UI", 11),
        )
        ent_email.pack(fill="x", pady=(6, 10))

        tk.Label(form, text="Senha:", bg="#121212", fg="#e5e7eb",
                 font=("Segoe UI", 10)).pack(anchor="w")
        ent_senha = tk.Entry(
            form, textvariable=var_senha, show="*",
            bg="#0b0b0b", fg="#ffffff", insertbackground="#ffffff",
            relief="flat", highlightthickness=1,
            highlightbackground="#2a2a2a", highlightcolor="#ff9f1a",
            font=("Segoe UI", 11),
        )
        ent_senha.pack(fill="x", pady=(6, 8))

        def do_login(event=None) -> None:
            if not self.connected:
                messagebox.showerror(
                    "Banco de dados", "Sem conexão com o banco.\n\nVerifique rede/configuração.")
                return

            email = _norm_email(var_email.get())
            senha = var_senha.get()

            if not email:
                messagebox.showwarning("Entrada", "Informe o e-mail.")
                ent_email.focus_set()
                return
            if not senha:
                messagebox.showwarning("Entrada", "Informe a senha.")
                ent_senha.focus_set()
                return

            try:
                u = fetch_user_by_email(self.cfg, email)
                if not u:
                    messagebox.showerror("Entrada", "Usuário não encontrado.")
                    return
                if not u.get("ativo", True):
                    messagebox.showerror("Entrada", "Usuário inativo.")
                    return

                stored = (u.get("senha_hash") or "").strip()
                if not verify_password(senha, stored):
                    messagebox.showerror("Entrada", "Senha incorreta.")
                    return

                self.user = u
                set_session_skip_entrada()

                try:
                    tela.grab_release()
                except Exception:
                    pass
                tela.destroy()

                self._show_menu()

            except Exception as e:
                log(f"Falha no login: {type(e).__name__}: {e}")
                messagebox.showerror("Erro", f"Falha no login:\n{e}")

        botoes = tk.Frame(frame, bg="#121212")
        botoes.pack(fill="x", pady=(12, 0))

        btn_entrar = ttk.Button(botoes, text="Entrar", command=do_login)
        btn_entrar.pack(side="left", expand=True, fill="x", padx=(0, 8))
        ttk.Button(botoes, text="Cancelar", command=do_cancel).pack(
            side="left", expand=True, fill="x"
        )

        if not self.connected:
            try:
                btn_entrar.state(["disabled"])
            except Exception:
                btn_entrar.config(state="disabled")

        tela.bind("<Return>", do_login)
        tela.bind("<Escape>", lambda e: do_cancel())

        ent_email.focus_set()
        self._center_window(tela)
        self._bring_to_front(tela)
        self.wait_window(tela)

    # ---------------- UI (Cadastro / Movimento / Sistema) ----------------

    def _build_ui(self) -> None:
        # ----- MENUBAR -----
        menubar = tk.Menu(self)

        m_cadastro = tk.Menu(menubar, tearoff=0)
        m_movimento = tk.Menu(menubar, tearoff=0)
        m_sistema = tk.Menu(menubar, tearoff=0)

        cad_items = [p for p in PROGRAMAS if (
            p.get("menu") or "").strip().lower() == "cadastro"]
        mov_items = [p for p in PROGRAMAS if (
            p.get("menu") or "").strip().lower() == "movimento"]

        for p in cad_items:
            m_cadastro.add_command(
                label=p["nome"], command=lambda pp=p: self.open_program(pp))
        for p in mov_items:
            m_movimento.add_command(
                label=p["nome"], command=lambda pp=p: self.open_program(pp))

        # ✅ Sistema
        m_sistema.add_command(label="Usuários e Acessos",
                              command=self._open_usuarios_acesso)
        m_sistema.add_command(label="Reset de Senha",
                              command=self._open_reset_senha)
        m_sistema.add_separator()
        m_sistema.add_command(label="Sair", command=self.on_closing)

        menubar.add_cascade(label="Cadastro", menu=m_cadastro)
        menubar.add_cascade(label="Movimento", menu=m_movimento)
        menubar.add_cascade(label="Sistema", menu=m_sistema)

        self.config(menu=menubar)

        # ----- CORPO -----
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        status_frame = ttk.Frame(main)
        status_frame.pack(fill="x")

        status_txt = "Conectado ao banco de dados" if self.connected else "Erro ao conectar ao banco"
        status_fg = "green" if self.connected else "red"
        ttk.Label(status_frame, text=status_txt, foreground=status_fg,
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        user_txt = ""
        if self.user:
            user_txt = f"Usuário: {self.user.get('nome') or self.user.get('usuarioId')}"
        ttk.Label(status_frame, text=user_txt,
                  foreground="gray").pack(side="right")

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=12)

        # Frames separados
        grp_cad = ttk.LabelFrame(main, text="Cadastro", padding=10)
        grp_mov = ttk.LabelFrame(main, text="Movimento", padding=10)
        grp_cad.pack(fill="x", pady=(0, 10))
        grp_mov.pack(fill="x", pady=(0, 10))

        for cc in range(5):
            grp_cad.columnconfigure(cc, weight=1)
            grp_mov.columnconfigure(cc, weight=1)

        def add_buttons(group: ttk.LabelFrame, items: List[Dict[str, Any]]) -> None:
            cols = 5
            r = 0
            c = 0
            for p in items:
                img = load_png_icon(p.get("icone", ""), max_side=32)
                if img:
                    self._icons[p["nome"]] = img

                btn = ttk.Button(
                    group,
                    text=p["nome"],
                    image=(img if img else None),
                    compound="left",
                    command=lambda pp=p: self.open_program(pp),
                    width=22,
                )
                btn.grid(row=r, column=c, sticky="ew", padx=6, pady=6, ipady=6)

                c += 1
                if c >= cols:
                    c = 0
                    r += 1

        add_buttons(grp_cad, cad_items)
        add_buttons(grp_mov, mov_items)

    # ---------------- abrir programas ----------------

    def open_program(self, programa: Dict[str, Any]) -> None:
        nome = programa.get("nome", "Programa")
        arquivo = programa.get("arquivo", "")
        tipo = (programa.get("tipo") or "").lower()

        if not arquivo:
            messagebox.showerror("Abrir", f"Arquivo não definido para {nome}.")
            return

        if not os.path.isfile(arquivo):
            messagebox.showerror("Abrir", f"Não encontrei:\n{arquivo}")
            return

        try:
            cwd = os.path.dirname(arquivo) or APP_DIR

            if tipo == "exe":
                if os.name == "nt":
                    os.startfile(arquivo)  # type: ignore[attr-defined]
                else:
                    subprocess.Popen([arquivo], cwd=cwd)
            else:
                cmd = build_run_cmd(arquivo)
                subprocess.Popen(cmd, cwd=cwd)

            self.after(150, self.destroy)

        except Exception as e:
            messagebox.showerror("Abrir", f"Falha ao abrir {nome}:\n{e}")

    # ---------------- fechar ----------------

    def on_closing(self, force: bool = False) -> None:
        if self._closing:
            return
        try:
            if force or messagebox.askokcancel("Sair", "Deseja realmente sair?"):
                self._closing = True
                clear_session_skip_entrada()
                self.destroy()
        except Exception:
            try:
                clear_session_skip_entrada()
                self.destroy()
            except Exception:
                pass


# ============================================================
# CLI
# ============================================================

def has_arg(flag: str) -> bool:
    flag = flag.lower().strip()
    return any(a.lower().strip() == flag for a in sys.argv[1:])


if __name__ == "__main__":
    skip = has_arg("--skip-entrada") or has_arg("--no-entrada")
    show_entrada = not skip

    log("=== START menu_principal ===")
    log(f"APP_DIR={APP_DIR}")
    log(f"BASE_DIR={BASE_DIR}")
    log(f"sys.executable={sys.executable}")
    log(f"argv={sys.argv}")
    log(f"show_entrada(cli)={show_entrada}")
    log(f"has_session_skip_entrada={has_session_skip_entrada()}")

    app = MenuPrincipalApp(show_entrada=show_entrada)
    app.mainloop()
