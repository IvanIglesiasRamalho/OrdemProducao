from __future__ import annotations

"""
tela_situacao.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Tabela situacao (conforme imagem):
  id (PK), nome, idHerdado (opcional)

Recursos:
- CRUD completo (inserir/atualizar/excluir/listar/buscar)
- Popup para escolher "Situação Herdada" (idHerdado) na própria tabela
- Config: BASE_DIR\\config_op.json + override DB_*
- Log em BASE_DIR\\logs\\tela_situacao.log
- Ícone favicon.ico/png (se existir)
- Ao fechar: reabre menu_principal.py com --skip-entrada

OBS:
- O programa detecta automaticamente se a tabela é "situacao" ou "Ekenox.situacao"
"""

import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass
from typing import Optional, Any, List, Tuple

import psycopg2


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
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def log_situacao(msg: str) -> None:
    _log_write("tela_situacao.log", msg)


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
            win._icon_img = img  # mantém referência
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


# ============================================================
# AUTO-DETECT TABELA (schema ou não)
# ============================================================

SITUACAO_TABLES = [
    '"Ekenox"."situacao"',
    '"situacao"',
]


def _table_exists(cfg: AppConfig, table_name: str) -> bool:
    conn = None
    try:
        conn = psycopg2.connect(
            host=cfg.db_host,
            database=cfg.db_database,
            user=cfg.db_user,
            password=cfg.db_password,
            port=int(cfg.db_port),
            connect_timeout=5,
        )
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
        ok = _table_exists(cfg, t)
        log_situacao(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


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
# MENU PRINCIPAL (ao fechar)
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
    pastas = [APP_DIR, BASE_DIR]
    for pasta in pastas:
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
        log_situacao(
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
        log_situacao(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_situacao(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# MODEL
# ============================================================

@dataclass
class Situacao:
    id: int
    nome: str
    idHerdado: Optional[int]


def _clean_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


# ============================================================
# REPOSITORY
# ============================================================

class SituacaoRepo:
    def __init__(self, db: Database, table: str) -> None:
        self.db = db
        self.table = table

    def proximo_id_preview(self) -> int:
        sql = f'SELECT COALESCE(MAX("id"), 0) + 1 FROM {self.table};'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql)
            return int(self.db.cursor.fetchone()[0])
        finally:
            self.db.desconectar()

    def listar(self, termo: Optional[str] = None, limit: int = 1200) -> List[Situacao]:
        like = f"%{termo}%" if termo else None

        def ilike_text(expr_sql: str) -> str:
            return f"COALESCE(CAST({expr_sql} AS TEXT),'') ILIKE %s"

        sql = f"""
            SELECT s."id", s."nome", s."idHerdado"
            FROM {self.table} AS s
            WHERE (%s IS NULL)
               OR {ilike_text('s.\"id\"')}
               OR {ilike_text('s.\"nome\"')}
               OR {ilike_text('s.\"idHerdado\"')}
            ORDER BY s."id" DESC
            LIMIT %s
        """
        params = (termo, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            out: List[Situacao] = []
            for r in rows:
                out.append(Situacao(
                    id=int(r[0]),
                    nome=str(r[1] or ""),
                    idHerdado=(int(r[2]) if r[2] is not None else None),
                ))
            return out
        finally:
            self.db.desconectar()

    def buscar_por_id(self, sid: int) -> Optional[Situacao]:
        sql = f"""
            SELECT s."id", s."nome", s."idHerdado"
            FROM {self.table} AS s
            WHERE s."id" = %s
            LIMIT 1
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (sid,))
            r = self.db.cursor.fetchone()
            if not r:
                return None
            return Situacao(id=int(r[0]), nome=str(r[1] or ""), idHerdado=(int(r[2]) if r[2] is not None else None))
        finally:
            self.db.desconectar()

    def existe_id(self, sid: int) -> bool:
        sql = f'SELECT 1 FROM {self.table} WHERE "id"=%s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (sid,))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, nome: str, idHerdado: Optional[int]) -> int:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None

            # trava a tabela durante a geração do próximo id e o insert
            self.db.cursor.execute(
                f'LOCK TABLE {self.table} IN EXCLUSIVE MODE;')
            self.db.cursor.execute(
                f'SELECT COALESCE(MAX("id"), 0) + 1 FROM {self.table};')
            new_id = int(self.db.cursor.fetchone()[0])

            sql = f"""
                INSERT INTO {self.table} ("id","nome","idHerdado")
                VALUES (%s,%s,%s)
                RETURNING "id"
            """
            self.db.cursor.execute(sql, (new_id, nome, idHerdado))
            rid = int(self.db.cursor.fetchone()[0])
            self.db.commit()
            return rid
        finally:
            self.db.desconectar()

    def atualizar(self, sid: int, nome: str, idHerdado: Optional[int]) -> None:
        sql = f"""
            UPDATE {self.table}
               SET "nome"=%s,
                   "idHerdado"=%s
             WHERE "id"=%s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (nome, idHerdado, sid))
            self.db.commit()
        finally:
            self.db.desconectar()

    def excluir(self, sid: int) -> None:
        sql = f'DELETE FROM {self.table} WHERE "id"=%s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (sid,))
            self.db.commit()
        finally:
            self.db.desconectar()

    def buscar_para_popup(self, termo: Optional[str], limit: int = 600) -> List[Tuple[int, str, Optional[int]]]:
        like = f"%{termo}%" if termo else None

        def ilike_text(expr_sql: str) -> str:
            return f"COALESCE(CAST({expr_sql} AS TEXT),'') ILIKE %s"

        sql = f"""
            SELECT s."id", s."nome", s."idHerdado"
            FROM {self.table} AS s
            WHERE (%s IS NULL)
               OR {ilike_text('s.\"id\"')}
               OR {ilike_text('s.\"nome\"')}
            ORDER BY s."nome"
            LIMIT %s
        """
        params = (termo, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()
            out: List[Tuple[int, str, Optional[int]]] = []
            for r in rows:
                out.append((int(r[0]), str(r[1] or ""),
                           (int(r[2]) if r[2] is not None else None)))
            return out
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class SituacaoService:
    def __init__(self, repo: SituacaoRepo) -> None:
        self.repo = repo

    def listar(self, termo: Optional[str]) -> List[Situacao]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def validar(self, sid: Optional[int], nome: str, idHerdado: Optional[int]) -> None:
        nome = (nome or "").strip()
        if not nome:
            raise ValueError("Nome é obrigatório.")

        if idHerdado is not None:
            if idHerdado <= 0:
                raise ValueError("idHerdado inválido.")
            if sid is not None and idHerdado == sid:
                raise ValueError("idHerdado não pode ser igual ao próprio id.")
            if not self.repo.existe_id(idHerdado):
                raise ValueError(f"idHerdado ({idHerdado}) não existe.")

    def salvar(self, id_original: Optional[int], nome: str, idHerdado: Optional[int]) -> Tuple[str, int]:
        nome = (nome or "").strip()
        self.validar(id_original, nome, idHerdado)

        if id_original is None:
            new_id = self.repo.inserir(nome, idHerdado)
            return ("inserida", new_id)

        if not self.repo.existe_id(id_original):
            new_id = self.repo.inserir(nome, idHerdado)
            return ("inserida", new_id)

        self.repo.atualizar(id_original, nome, idHerdado)
        return ("atualizada", id_original)

    def excluir(self, sid: int) -> None:
        self.repo.excluir(sid)

    def popup_buscar(self, termo: Optional[str]) -> List[Tuple[int, str, Optional[int]]]:
        termo = (termo or "").strip() or None
        return self.repo.buscar_para_popup(termo)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1000x650"
APP_TITLE = "Tela de Situação"

TREE_COLS = ["id", "nome", "idHerdado"]


class SituacaoPicker(tk.Toplevel):
    """
    Popup para escolher uma Situação (para preencher idHerdado).
    Retorna (id, nome).
    """

    def __init__(self, master: tk.Misc, service: SituacaoService, on_pick):
        super().__init__(master)
        self.service = service
        self.on_pick = on_pick

        self.title("Buscar Situação (Herdado)")
        self.geometry("720x450")
        self.minsize(650, 380)
        apply_window_icon(self)

        self.var_busca = tk.StringVar()

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (ID/Nome):").grid(row=0,
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
            "id", "nome", "herd"), show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.heading("id", text="ID")
        self.tree.heading("nome", text="Nome")
        self.tree.heading("herd", text="Herdado")

        self.tree.column("id", width=90, anchor="e", stretch=False)
        self.tree.column("nome", width=460, anchor="w", stretch=True)
        self.tree.column("herd", width=90, anchor="e", stretch=False)

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
            rows = self.service.popup_buscar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao buscar situações:\n{e}")
            return

        for sid, nome, herd in rows:
            self.tree.insert("", "end", values=(
                sid, nome, "" if herd is None else herd))

    def _pick(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selecionar", "Selecione uma situação.")
            return
        sid, nome, _herd = self.tree.item(sel[0], "values")
        try:
            self.on_pick(int(sid), str(nome))
        finally:
            self.destroy()


class TelaSituacao(ttk.Frame):
    def __init__(self, master: tk.Misc, service: SituacaoService):
        super().__init__(master)
        self.service = service

        self.var_filtro = tk.StringVar()

        self.var_id = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_herd_id = tk.StringVar()
        self.var_herd_nome = tk.StringVar()

        self._id_original: Optional[int] = None

        self._build_ui()
        self.atualizar_lista()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # Topo
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (ID/Nome/Herdado):").grid(row=0,
                                                              column=0, sticky="w")
        ent_busca = ttk.Entry(top, textvariable=self.var_filtro)
        ent_busca.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent_busca.bind("<Return>", lambda e: self.atualizar_lista())

        ttk.Button(top, text="Atualizar", command=self.atualizar_lista).grid(
            row=0, column=2, padx=(0, 6))
        ttk.Button(top, text="Novo", command=self.novo).grid(
            row=0, column=3, padx=(0, 6))
        ttk.Button(top, text="Salvar", command=self.salvar).grid(
            row=0, column=4, padx=(0, 6))
        ttk.Button(top, text="Excluir", command=self.excluir).grid(
            row=0, column=5, padx=(0, 6))
        ttk.Button(top, text="Limpar", command=self.limpar_form).grid(
            row=0, column=6)

        # Form
        form = ttk.LabelFrame(self, text="Situação")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(10):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="ID:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_id, state="readonly", width=10).grid(
            row=0, column=1, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Nome:").grid(
            row=0, column=2, sticky="w", padx=(10, 6), pady=6)
        ent_nome = ttk.Entry(form, textvariable=self.var_nome)
        ent_nome.grid(row=0, column=3, sticky="ew",
                      padx=(0, 10), pady=6, columnspan=5)

        ttk.Label(form, text="Herdado (ID):").grid(
            row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        ent_herd = ttk.Entry(form, textvariable=self.var_herd_id, width=12)
        ent_herd.grid(row=1, column=1, sticky="w", padx=(0, 10), pady=6)

        ttk.Button(form, text="Buscar Herdado...", command=self.buscar_herdado_popup).grid(
            row=1, column=2, sticky="w", padx=(0, 6), pady=6
        )
        ttk.Button(form, text="Remover Herdado", command=self.remover_herdado).grid(
            row=1, column=3, sticky="w", padx=(0, 6), pady=6
        )

        ttk.Label(form, text="Herdado (Nome):").grid(
            row=1, column=4, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_herd_nome, state="readonly").grid(
            row=1, column=5, sticky="ew", padx=(0, 10), pady=6, columnspan=4
        )

        # Lista
        lst = ttk.Frame(self)
        lst.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
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

        self.tree.heading("id", text="ID")
        self.tree.heading("nome", text="Nome")
        self.tree.heading("idHerdado", text="Herdado")

        self.tree.column("id", width=90, anchor="e", stretch=False)
        self.tree.column("nome", width=700, anchor="w", stretch=True)
        self.tree.column("idHerdado", width=90, anchor="e", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # atalhos
        ent_nome.bind("<Return>", lambda e: self.salvar())
        ent_herd.bind(
            "<Return>", lambda e: self._preencher_nome_herdado_por_id())

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar situações:\n{e}")
            return

        for s in rows:
            self.tree.insert("", "end", values=(
                s.id, s.nome, "" if s.idHerdado is None else s.idHerdado))

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        sid, nome, herd = self.tree.item(sel[0], "values")

        self._id_original = int(sid)
        self.var_id.set(str(sid))
        self.var_nome.set(str(nome or ""))

        herd_str = (str(herd).strip() if herd is not None else "")
        if herd_str == "" or herd_str == "0":
            self.var_herd_id.set("")
            self.var_herd_nome.set("")
        else:
            self.var_herd_id.set(herd_str)
            self._preencher_nome_herdado_por_id()

    def novo(self) -> None:
        self.limpar_form()
        try:
            nid = self.service.repo.proximo_id_preview()
            self.var_id.set(str(nid))   # <-- aparece no campo ID
        except Exception:
            self.var_id.set("")

    def limpar_form(self) -> None:
        self._id_original = None
        self.var_id.set("")
        self.var_nome.set("")
        self.var_herd_id.set("")
        self.var_herd_nome.set("")
        self.tree.selection_remove(self.tree.selection())

    def remover_herdado(self) -> None:
        self.var_herd_id.set("")
        self.var_herd_nome.set("")

    def _preencher_nome_herdado_por_id(self) -> None:
        raw = (self.var_herd_id.get() or "").strip()
        if not raw:
            self.var_herd_nome.set("")
            return
        try:
            hid = int(raw)
        except ValueError:
            self.var_herd_nome.set("")
            return

        try:
            sit = self.service.repo.buscar_por_id(hid)
        except Exception:
            self.var_herd_nome.set("")
            return

        if not sit:
            self.var_herd_nome.set("(não encontrado)")
        else:
            self.var_herd_nome.set(sit.nome)

    def buscar_herdado_popup(self) -> None:
        def on_pick(sid: int, nome: str) -> None:
            self.var_herd_id.set(str(sid))
            self.var_herd_nome.set(nome)

        SituacaoPicker(self.winfo_toplevel(), self.service, on_pick)

    def salvar(self) -> None:
        nome = self.var_nome.get()
        herd_raw = (self.var_herd_id.get() or "").strip()

        herd: Optional[int]
        if herd_raw == "" or herd_raw.lower() in {"none", "null"}:
            herd = None
        else:
            try:
                herd = int(herd_raw)
            except ValueError:
                messagebox.showerror(
                    "Validação", "Herdado (ID) deve ser número ou vazio.")
                return

        try:
            status, sid = self.service.salvar(self._id_original, nome, herd)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        messagebox.showinfo("OK", f"Situação {status} com sucesso.\nID: {sid}")
        self.atualizar_lista()
        self._id_original = sid
        self.var_id.set(str(sid))
        self._preencher_nome_herdado_por_id()

    def excluir(self) -> None:
        if self._id_original is None:
            messagebox.showwarning(
                "Atenção", "Selecione uma situação para excluir.")
            return

        sid = self._id_original
        if not messagebox.askyesno("Confirmar", f"Excluir Situação ID {sid}?"):
            return

        try:
            self.service.excluir(sid)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Situação excluída.")
        self.limpar_form()
        self.atualizar_lista()


# ============================================================
# STARTUP
# ============================================================

def test_connection_or_die(cfg: AppConfig) -> None:
    conn = None
    try:
        conn = psycopg2.connect(
            host=cfg.db_host,
            database=cfg.db_database,
            user=cfg.db_user,
            password=cfg.db_password,
            port=int(cfg.db_port),
            connect_timeout=5,
        )
        cur = conn.cursor()
        cur.execute("SELECT 1")
        cur.fetchone()
        cur.close()
        conn.close()
    except Exception as e:
        try:
            if conn:
                conn.close()
        except Exception:
            pass
        raise RuntimeError(f"{type(e).__name__}: {e}")


def main() -> None:
    cfg = env_override(load_config())

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
        root.destroy()
        return

    situacao_table = detectar_tabela(cfg, SITUACAO_TABLES, '"situacao"')

    db = Database(cfg)
    repo = SituacaoRepo(db, table=situacao_table)
    service = SituacaoService(repo)

    tela = TelaSituacao(root, service)
    tela.pack(fill="both", expand=True)

    closing = {"done": False}

    def open_menu_then_close():
        if closing["done"]:
            return
        closing["done"] = True
        try:
            abrir_menu_principal_skip_entrada()
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def on_close():
        try:
            root.withdraw()
            root.update_idletasks()
        except Exception:
            pass
        root.after(50, open_menu_then_close)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_situacao(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_situacao:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
