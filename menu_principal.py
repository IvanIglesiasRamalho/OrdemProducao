from __future__ import annotations

"""
menu_principal.py
Python 3.12+ | Tkinter

OBJETIVO:
- A tela "Ekenox - Entrada" deve aparecer SOMENTE na primeira abertura normal do sistema.
- Quando o menu for reaberto ao voltar de outro programa (ex: tela_produtos.py),
  NÃO deve aparecer a tela de entrada.

MECANISMOS:
1) CLI:
   python menu_principal.py --skip-entrada
   -> abre direto no menu (sem entrada)

2) Flag de sessão (arquivo):
   Ao clicar "Entrar", cria BASE_DIR\\.session_skip_entrada
   Enquanto esse arquivo existir, não mostra Entrada.

Quando o usuário "Sair" (confirmando), a flag é removida.
"""

import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple


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
# FLAG DE SESSÃO (pula entrada após "Entrar")
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
            win._icon_img = img  # type: ignore[attr-defined]
    except Exception:
        pass


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


def try_connect_db(cfg: AppConfig) -> Tuple[bool, str]:
    try:
        import psycopg2  # type: ignore
        conn = psycopg2.connect(
            host=cfg.db_host,
            database=cfg.db_database,
            user=cfg.db_user,
            password=cfg.db_password,
            port=int(cfg.db_port),
            connect_timeout=5,
        )
        conn.close()
        return True, ""
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


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

    # Se for exe, roda direto
    if script_path.lower().endswith(".exe"):
        return [script_path] + extra_args

    # Se for .py no Windows, preferir pyw/pythonw
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
# PROGRAMAS (AJUSTE OS NOMES/ARQUIVOS)
# ============================================================

PROGRAMAS: List[Dict[str, Any]] = [
    {"nome": "Ordem de Produção", "icone": "imagens/menu_op.png",
        "arquivo": os.path.join(APP_DIR, "Ordem_Producao.py"), "tipo": "py"},
    {"nome": "Produtos", "icone": "imagens/menu_produtos.png",
        "arquivo": os.path.join(APP_DIR, "tela_produtos.py"), "tipo": "py"},
    {"nome": "Informação Produto", "icone": "imagens/menu_infoproduto.png",
        "arquivo": os.path.join(APP_DIR, "tela_info_produto.py"), "tipo": "py"},
    {"nome": "Arranjo", "icone": "imagens/menu_arranjo.png",
        "arquivo": os.path.join(APP_DIR, "tela_arranjo.py"), "tipo": "py"},
    {"nome": "Categoria", "icone": "imagens/menu_categoria.png",
        "arquivo": os.path.join(APP_DIR, "tela_categoria.py"), "tipo": "py"},
    {"nome": "Depósito", "icone": "imagens/menu_deposito.png",
        "arquivo": os.path.join(APP_DIR, "tela_deposito.py"), "tipo": "py"},
    {"nome": "Estoque", "icone": "imagens/menu_estoque.png",
        "arquivo": os.path.join(APP_DIR, "tela_estoque.py"), "tipo": "py"},
    {"nome": "Estrutura", "icone": "imagens/menu_estrutura.png",
        "arquivo": os.path.join(APP_DIR, "tela_estrutura.py"), "tipo": "py"},
    {"nome": "Fornecedor", "icone": "imagens/menu_fornecedor.png",
        "arquivo": os.path.join(APP_DIR, "tela_fornecedor.py"), "tipo": "py"},
    {"nome": "Situação", "icone": "imagens/menu_situacao.png",
        "arquivo": os.path.join(APP_DIR, "tela_situacao.py"), "tipo": "py"},
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

        # regra final: se existe flag de sessão, nunca mostrar entrada
        self.show_entrada = bool(show_entrada) and (
            not has_session_skip_entrada())

        self.title("Menu Principal - Ekenox")
        self.geometry("1100x650")
        self.minsize(1100, 650)
        apply_window_icon(self)

        self._closing = False
        self._ui_built = False

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.withdraw()
        self._icons: Dict[str, tk.PhotoImage] = {}

        self.after(80, self._startup)

    def _startup(self) -> None:
        splash = Splash(self, "Inicializando")

        def worker() -> None:
            splash.set_text("Conectando ao banco...")
            ok, err = try_connect_db(self.cfg)
            self.connected = ok
            self.db_err = err

            splash.set_info("Conectado" if ok else (
                err[:70] + ("..." if len(err) > 70 else "")))
            splash.set_text("Carregando menu...")

            self.after(0, self._build_ui_once)

            splash.set_text("Pronto")
            self.after(250, splash.destroy)

            if self.show_entrada:
                self.after(300, self._show_entrada)
            else:
                self.after(300, self._show_menu_direct)

        threading.Thread(target=worker, daemon=True).start()

    def _build_ui_once(self) -> None:
        if self._ui_built:
            return
        self._ui_built = True
        self._build_ui()

    def _show_menu_direct(self) -> None:
        if self._closing:
            return
        self.deiconify()
        self.lift()
        try:
            self.focus_force()
        except Exception:
            pass

    # ---------------- Entrada ----------------

    def _load_avatar(self, path: str, max_side: int = 260) -> Optional[tk.PhotoImage]:
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

    def _show_entrada(self) -> None:
        if self._closing:
            return

        # segurança extra: se por algum motivo a flag apareceu, não abre entrada
        if has_session_skip_entrada():
            self._show_menu_direct()
            return

        tela = tk.Toplevel(self)
        apply_window_icon(tela)
        tela.title("Ekenox - Entrada")
        tela.resizable(False, False)
        tela.configure(bg="#121212")
        tela.protocol("WM_DELETE_WINDOW", self.on_closing)

        frame = tk.Frame(tela, bg="#121212", padx=30, pady=25)
        frame.pack(fill="both", expand=True)

        candidatos = [
            os.path.join(BASE_DIR, "imagens", "avatar_ekenox.png"),
            os.path.join(BASE_DIR, "avatar_ekenox.png"),
            os.path.join(BASE_DIR, "imagens", "Ekenox.png"),
            os.path.join(BASE_DIR, "Ekenox.png"),
        ]
        avatar_path = next((p for p in candidatos if os.path.isfile(p)), None)

        avatar_img = self._load_avatar(avatar_path) if avatar_path else None
        if avatar_img:
            lbl_img = tk.Label(frame, image=avatar_img, bg="#121212")
            lbl_img.image = avatar_img
            lbl_img.pack(pady=(0, 15))
        else:
            tk.Label(frame, text="(Avatar não encontrado)", bg="#121212",
                     fg="#aaaaaa", font=("Segoe UI", 10)).pack(pady=(0, 15))

        tk.Label(frame, text="Sistema - Menu Principal", bg="#121212",
                 fg="#ffffff", font=("Segoe UI", 14, "bold")).pack()
        tk.Label(frame, text="Ekenox", bg="#121212", fg="#ff9f1a",
                 font=("Segoe UI", 18, "bold")).pack(pady=(2, 18))

        status = "Conectado ao banco" if self.connected else f"ERRO BD: {self.db_err}"
        tk.Label(frame, text=status, bg="#121212",
                 fg=("#34d399" if self.connected else "#f87171"),
                 font=("Segoe UI", 10, "bold")).pack(pady=(0, 14))

        botoes = tk.Frame(frame, bg="#121212")
        botoes.pack(fill="x")

        def entrar(event=None) -> None:
            # marca sessão para nunca mais mostrar entrada ao reabrir
            set_session_skip_entrada()
            try:
                tela.destroy()
            except Exception:
                pass
            self._show_menu_direct()

        ttk.Button(botoes, text="Entrar", command=entrar).pack(
            side="left", expand=True, fill="x", padx=(0, 8))
        ttk.Button(botoes, text="Sair", command=self.on_closing).pack(
            side="left", expand=True, fill="x")

        tela.bind("<Return>", entrar)
        tela.bind("<Escape>", lambda e: self.on_closing())

        tela.update_idletasks()
        w, h = tela.winfo_width(), tela.winfo_height()
        x = (tela.winfo_screenwidth() // 2) - (w // 2)
        y = (tela.winfo_screenheight() // 2) - (h // 2)
        tela.geometry(f"+{x}+{y}")

    # ---------------- UI ----------------

    def _build_ui(self) -> None:
        menubar = tk.Menu(self)

        m_programas = tk.Menu(menubar, tearoff=0)
        for p in PROGRAMAS:
            m_programas.add_command(
                label=p["nome"], command=lambda pp=p: self.open_program(pp))
        menubar.add_cascade(label="Programas", menu=m_programas)

        m_sistema = tk.Menu(menubar, tearoff=0)
        m_sistema.add_command(label="Sair", command=self.on_closing)
        menubar.add_cascade(label="Sistema", menu=m_sistema)

        self.config(menu=menubar)

        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        status_frame = ttk.Frame(main)
        status_frame.pack(fill="x")

        status_txt = "Conectado ao banco de dados" if self.connected else "Erro ao conectar ao banco"
        status_fg = "green" if self.connected else "red"
        ttk.Label(status_frame, text=status_txt, foreground=status_fg,
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        ttk.Separator(main, orient="horizontal").pack(fill="x", pady=12)

        grid = ttk.Frame(main)
        grid.pack(fill="both", expand=True)

        cols = 5
        for cc in range(cols):
            grid.columnconfigure(cc, weight=1)

        r = 0
        c = 0

        for p in PROGRAMAS:
            img = load_png_icon(p.get("icone", ""), max_side=32)
            if img:
                self._icons[p["nome"]] = img

            btn = ttk.Button(
                grid,
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

            # fecha menu atual; o filho deve reabrir menu com --skip-entrada
            self.after(150, self.destroy)

        except Exception as e:
            messagebox.showerror("Abrir", f"Falha ao abrir {nome}:\n{e}")

    # ---------------- fechar ----------------

    def on_closing(self) -> None:
        if self._closing:
            return
        try:
            if messagebox.askokcancel("Sair", "Deseja realmente sair?"):
                self._closing = True
                # limpando a flag => próxima abertura normal volta a mostrar Entrada
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
