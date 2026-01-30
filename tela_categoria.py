from __future__ import annotations

import argparse
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


def log_categoria(msg: str) -> None:
    _log_write("tela_categoria.log", msg)


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


# ============================================================
# DB
# ============================================================

def db_connect(cfg: AppConfig):
    return psycopg2.connect(
        host=cfg.db_host,
        database=cfg.db_database,
        user=cfg.db_user,
        password=cfg.db_password,
        port=int(cfg.db_port),
        connect_timeout=5,
    )


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
# MENU PRINCIPAL (mantido, mas NÃO abrimos menu no fechar)
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
    # Mantido por compatibilidade, mas não usado no fechamento.
    menu_path = localizar_menu_principal()
    if not menu_path:
        log_categoria(
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
        log_categoria(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_categoria(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# AUTO-DETECT TABELA
# ============================================================

CATEGORIA_TABLES = [
    '"Ekenox"."categoria"',
    '"categoria"',
    '"Ekenox"."categorias"',
    '"categorias"',
]


def _table_exists(cfg: AppConfig, table_name: str) -> bool:
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
        ok = _table_exists(cfg, t)
        log_categoria(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


# ============================================================
# HELPERS SQL
# ============================================================

def _qident(name: str) -> str:
    return '"' + str(name).replace('"', '""') + '"'


def _is_numeric_pg_type(typname: str) -> bool:
    t = (typname or "").lower()
    return t in {"int2", "int4", "int8", "numeric", "float4", "float8"}


# ============================================================
# CONTROLE DE ACESSO (PERMISSÕES)
# ============================================================

# ajuste se no banco estiver "Categorias", etc.
THIS_PROGRAMA_TERMO = "Categoria"

NIVEL_LABEL = {
    0: "0 - Sem acesso",
    1: "1 - Leitura",
    2: "2 - Edição",
    3: "3 - Edição",
}


def _parse_cli_user() -> Optional[int]:
    """
    Aceita:
      --usuario-id <id> (ou --uid <id>)
    Se não vier, retorna None e a tela abre em NÍVEL 1 (Leitura).
    """
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--usuario-id", "--uid", dest="usuario_id", type=int)
    args, _ = parser.parse_known_args(sys.argv[1:])
    return int(args.usuario_id) if args.usuario_id else None


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


def fetch_user_nome(cfg: AppConfig, usuario_id: int) -> str:
    sql = """
        SELECT COALESCE(u."nome",'')
          FROM "Ekenox"."usuarios" u
         WHERE u."usuarioId"=%s
         LIMIT 1
    """
    conn = db_connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (int(usuario_id),))
            r = cur.fetchone()
        return str(r[0] or "").strip() if r else ""
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
    """
    Regras:
      - Usuário inativo => nega (0)
      - Se programa não encontrado => abre N1 com aviso
      - Se permissão não cadastrada (nivel<=0) => abre N1 com aviso
      - Nivel 2 e 3 => edição (mesma regra)
    """
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

    # Nível 3 = edição também (igual nível 2)
    if nivel == 3:
        nivel = 3

    return nivel, ""


# ============================================================
# MODEL
# ============================================================

@dataclass
class Categoria:
    codigo: str
    nome: str
    pai: str  # pode ser ""


# ============================================================
# REPOSITORY
# ============================================================

class CategoriaRepo:
    def __init__(self, db: Database, categoria_table: str) -> None:
        self.db = db
        self.categoria_table = categoria_table

        self.pk_col = self._find_primary_key_column() or self._find_existing_column(
            ["codigo", "categoriaId", "idCategoria", "id", "Codigo"]
        ) or "codigo"

        self.nome_col = self._find_existing_column(
            ["nomeCategoria", "nome", "descricao", "descr", "NomeCategoria"]
        ) or "nomeCategoria"

        self.pai_col = self._find_existing_column(
            ["pai", "paiId", "fkPai", "categoriaPai", "idPai", "Pai"]
        ) or "pai"

        self.pk_typname = self._col_typname(self.categoria_table, self.pk_col)
        self.pk_is_numeric = _is_numeric_pg_type(self.pk_typname)

    def _list_columns(self) -> List[str]:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT a.attname
                  FROM pg_attribute a
                 WHERE a.attrelid = %s::regclass
                   AND a.attnum > 0
                   AND NOT a.attisdropped
                 ORDER BY a.attnum
                """,
                (self.categoria_table,),
            )
            return [str(r[0]) for r in self.db.cursor.fetchall()]
        finally:
            self.db.desconectar()

    def _find_existing_column(self, candidates: List[str]) -> Optional[str]:
        try:
            cols = {c.lower(): c for c in self._list_columns()}
        except Exception:
            cols = {}
        for cand in candidates:
            key = str(cand).lower()
            if key in cols:
                return cols[key]
        return None

    def _find_primary_key_column(self) -> Optional[str]:
        if not self.db.conectar():
            return None
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                """
                SELECT a.attname
                  FROM pg_index i
                  JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)
                 WHERE i.indrelid = %s::regclass
                   AND i.indisprimary
                 ORDER BY a.attnum
                 LIMIT 1
                """,
                (self.categoria_table,),
            )
            r = self.db.cursor.fetchone()
            return str(r[0]) if r and r[0] else None
        except Exception:
            return None
        finally:
            self.db.desconectar()

    def _col_typname(self, table_reg: str, col: str) -> str:
        if not self.db.conectar():
            return ""
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
        except Exception:
            return ""
        finally:
            self.db.desconectar()

    def proximo_codigo(self) -> int:
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            pk = _qident(self.pk_col)

            if self.pk_is_numeric:
                sql = f"SELECT COALESCE(MAX({pk})::bigint, 0) + 1 FROM {self.categoria_table}"
                self.db.cursor.execute(sql)
            else:
                sql = f"""
                    SELECT COALESCE(
                        MAX((NULLIF(regexp_replace({pk}::text,'[^0-9]','','g'),'') )::bigint),
                        0
                    ) + 1
                    FROM {self.categoria_table}
                """
                self.db.cursor.execute(sql)

            r = self.db.cursor.fetchone()
            return int(r[0] or 1)
        finally:
            self.db.desconectar()

    def listar(self, termo: Optional[str] = None, limit: int = 2000) -> List[Categoria]:
        termo = (termo or "").strip()
        like = f"%{termo}%" if termo else None

        pk = _qident(self.pk_col)
        nm = _qident(self.nome_col)
        pai = _qident(self.pai_col)

        sql = f"""
            SELECT
                {pk}::text AS codigo,
                COALESCE({nm}::text,'') AS nome,
                COALESCE({pai}::text,'') AS pai
            FROM {self.categoria_table}
            WHERE (%s IS NULL)
               OR ({pk}::text ILIKE %s)
               OR (COALESCE({nm}::text,'') ILIKE %s)
               OR (COALESCE({pai}::text,'') ILIKE %s)
            ORDER BY {pk}::text
            LIMIT %s
        """
        params = (termo or None, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            out: List[Categoria] = []
            for (codigo, nome, pai_val) in self.db.cursor.fetchall():
                out.append(Categoria(
                    codigo=str(codigo or ""),
                    nome=str(nome or ""),
                    pai=str(pai_val or ""),
                ))
            return out
        finally:
            self.db.desconectar()

    def buscar_nome_por_codigo(self, codigo_txt: str) -> str:
        pk = _qident(self.pk_col)
        nm = _qident(self.nome_col)
        sql = f"SELECT COALESCE({nm}::text,'') FROM {self.categoria_table} WHERE {pk}::text = %s LIMIT 1"
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(codigo_txt),))
            r = self.db.cursor.fetchone()
            return str(r[0] or "") if r else ""
        finally:
            self.db.desconectar()

    def existe(self, codigo_txt: str) -> bool:
        pk = _qident(self.pk_col)
        sql = f"SELECT 1 FROM {self.categoria_table} WHERE {pk}::text = %s"
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(codigo_txt),))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, codigo_int: int, nome: str, pai_txt: str) -> str:
        pk = _qident(self.pk_col)
        nm = _qident(self.nome_col)
        pai = _qident(self.pai_col)

        codigo_param = codigo_int if self.pk_is_numeric else str(codigo_int)
        pai_param = (pai_txt or "").strip() or None

        sql = f"""
            INSERT INTO {self.categoria_table} ({pk},{nm},{pai})
            VALUES (%s,%s,%s)
            RETURNING {pk}::text
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (codigo_param, nome, pai_param))
            new_code = self.db.cursor.fetchone()[0]
            self.db.commit()
            return str(new_code)
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def atualizar(self, codigo_txt: str, nome: str, pai_txt: str) -> None:
        pk = _qident(self.pk_col)
        nm = _qident(self.nome_col)
        pai = _qident(self.pai_col)

        pai_param = (pai_txt or "").strip() or None

        sql = f"""
            UPDATE {self.categoria_table}
               SET {nm} = %s,
                   {pai} = %s
             WHERE {pk}::text = %s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (nome, pai_param, str(codigo_txt)))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def excluir(self, codigo_txt: str) -> None:
        pk = _qident(self.pk_col)
        sql = f"DELETE FROM {self.categoria_table} WHERE {pk}::text = %s"
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (str(codigo_txt),))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class CategoriaService:
    def __init__(self, repo: CategoriaRepo) -> None:
        self.repo = repo

    def listar(self, termo: Optional[str]) -> List[Categoria]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def proximo_codigo(self) -> int:
        return self.repo.proximo_codigo()

    def nome_por_codigo(self, codigo_txt: str) -> str:
        codigo_txt = (codigo_txt or "").strip()
        if not codigo_txt:
            return ""
        return self.repo.buscar_nome_por_codigo(codigo_txt)

    def salvar(self, codigo_txt: str, nome: str, pai_txt: str) -> Tuple[str, str]:
        nome = (nome or "").strip()
        if not nome:
            raise ValueError("Nome Categoria é obrigatório.")

        codigo_txt = (codigo_txt or "").strip()
        pai_txt = (pai_txt or "").strip()

        if not codigo_txt:
            codigo_int = self.repo.proximo_codigo()
            codigo_salvo = self.repo.inserir(codigo_int, nome, pai_txt)
            return ("inserido", codigo_salvo)

        try:
            int(codigo_txt)
        except ValueError:
            raise ValueError(
                "Código inválido (não numérico). Use o botão NOVO para gerar automaticamente.")

        if self.repo.existe(codigo_txt):
            self.repo.atualizar(codigo_txt, nome, pai_txt)
            return ("atualizado", codigo_txt)

        codigo_salvo = self.repo.inserir(int(codigo_txt), nome, pai_txt)
        return ("inserido", codigo_salvo)

    def excluir(self, codigo_txt: str) -> None:
        self.repo.excluir(codigo_txt)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1100x650"
APP_TITLE = "Tela de Categoria"
TREE_COLS = ["codigo", "nome", "pai"]


class SeletorPai(tk.Toplevel):
    def __init__(self, master: tk.Misc, service: CategoriaService, on_pick):
        super().__init__(master)
        self.title("Selecionar Pai")
        self.geometry("780x420")
        self.service = service
        self.on_pick = on_pick

        self.var_busca = tk.StringVar()

        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (Código/Nome):").grid(row=0,
                                                          column=0, sticky="w")
        ent = ttk.Entry(top, textvariable=self.var_busca)
        ent.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent.bind("<Return>", lambda e: self._load())

        ttk.Button(top, text="Buscar", command=self._load).grid(
            row=0, column=2)

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            frm, columns=["codigo", "nome"], show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.heading("codigo", text="Código")
        self.tree.heading("nome", text="Nome Categoria")
        self.tree.column("codigo", width=120, anchor="e", stretch=False)
        self.tree.column("nome", width=560, anchor="w", stretch=True)

        self.tree.bind("<Double-1>", lambda e: self._pick())

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Selecionar",
                   command=self._pick).pack(side="right")

        self._load()
        ent.focus_set()
        self.transient(master)
        self.grab_set()

    def _load(self):
        termo = self.var_busca.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao listar categorias:\n{e}", parent=self)
            return

        for r in rows:
            self.tree.insert("", "end", values=(r.codigo, r.nome))

    def _pick(self):
        sel = self.tree.selection()
        if not sel:
            return
        codigo, nome = self.tree.item(sel[0], "values")
        self.on_pick(str(codigo), str(nome))
        self.destroy()


class TelaCategoria(ttk.Frame):
    def __init__(
        self,
        master: tk.Misc,
        service: CategoriaService,
        *,
        usuario_logado_id: int,
        acesso_nivel: int,
        usuario_logado_nome: str = "",
    ):
        super().__init__(master)
        self.service = service

        self.usuario_logado_id = int(usuario_logado_id)
        self.acesso_nivel = int(acesso_nivel)
        self.usuario_logado_nome = (usuario_logado_nome or "").strip()

        self.var_filtro = tk.StringVar()
        self.var_codigo = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_pai_id = tk.StringVar()
        self.var_pai_nome = tk.StringVar()

        self._lbl_access: Optional[ttk.Label] = None
        self._ent_codigo: Optional[ttk.Entry] = None
        self._ent_nome: Optional[ttk.Entry] = None
        self._ent_pai_id: Optional[ttk.Entry] = None

        self._btn_novo: Optional[ttk.Button] = None
        self._btn_salvar: Optional[ttk.Button] = None
        self._btn_excluir: Optional[ttk.Button] = None
        self._btn_buscar_pai: Optional[ttk.Button] = None
        self._btn_remover_pai: Optional[ttk.Button] = None

        self._build_ui()
        self.atualizar_lista()
        self._apply_access_rules()

    def _can_edit(self) -> bool:
        # Regra pedida: 1 leitura / 2 edição / 3 edição também
        return self.acesso_nivel >= 2

    def _apply_access_rules(self) -> None:
        can_edit = self._can_edit()

        if self._lbl_access:
            nome = self.usuario_logado_nome or (
                f"ID {self.usuario_logado_id}" if self.usuario_logado_id else "Não informado")
            nivel_txt = NIVEL_LABEL.get(
                self.acesso_nivel, str(self.acesso_nivel))
            self._lbl_access.config(
                text=f"Logado: {nome} | Nível: {nivel_txt}",
                foreground=("green" if self.acesso_nivel >= 2 else "gray"),
            )

        # Campos: readonly no nível 1
        if self._ent_codigo:
            self._ent_codigo.configure(
                state=("normal" if can_edit else "readonly"))
        if self._ent_nome:
            self._ent_nome.configure(
                state=("normal" if can_edit else "readonly"))
        if self._ent_pai_id:
            self._ent_pai_id.configure(
                state=("normal" if can_edit else "readonly"))

        # Botões que alteram dados
        for btn in (self._btn_novo, self._btn_salvar, self._btn_excluir, self._btn_buscar_pai, self._btn_remover_pai):
            if btn:
                btn.state(["!disabled"] if can_edit else ["disabled"])

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        # Barra de status / permissões
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        bar.columnconfigure(0, weight=1)

        self._lbl_access = ttk.Label(
            bar, text="", foreground="gray", font=("Segoe UI", 9, "bold"))
        self._lbl_access.grid(row=0, column=0, sticky="w")

        # Topo (busca + botões)
        top = ttk.Frame(self)
        top.grid(row=1, column=0, sticky="ew", padx=10, pady=(6, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Buscar (Código/Nome/Pai):").grid(row=0,
                                                              column=0, sticky="w")
        ent_busca = ttk.Entry(top, textvariable=self.var_filtro)
        ent_busca.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ent_busca.bind("<Return>", lambda e: self.atualizar_lista())

        ttk.Button(top, text="Atualizar", command=self.atualizar_lista).grid(
            row=0, column=2, padx=(0, 6))

        self._btn_novo = ttk.Button(top, text="Novo", command=self.novo)
        self._btn_novo.grid(row=0, column=3, padx=(0, 6))

        self._btn_salvar = ttk.Button(top, text="Salvar", command=self.salvar)
        self._btn_salvar.grid(row=0, column=4, padx=(0, 6))

        self._btn_excluir = ttk.Button(
            top, text="Excluir", command=self.excluir)
        self._btn_excluir.grid(row=0, column=5, padx=(0, 6))

        ttk.Button(top, text="Limpar", command=self.limpar_form).grid(
            row=0, column=6)

        # Form
        form = ttk.LabelFrame(self, text="Categoria")
        form.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(12):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="Código:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        self._ent_codigo = ttk.Entry(
            form, textvariable=self.var_codigo, width=14)
        self._ent_codigo.grid(row=0, column=1, sticky="w",
                              padx=(0, 10), pady=6)

        ttk.Label(form, text="Nome Categoria:").grid(
            row=0, column=3, sticky="w", padx=(10, 6), pady=6)
        self._ent_nome = ttk.Entry(form, textvariable=self.var_nome)
        self._ent_nome.grid(row=0, column=4, sticky="ew",
                            padx=(0, 10), pady=6, columnspan=8)

        ttk.Label(form, text="Pai (ID):").grid(
            row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        self._ent_pai_id = ttk.Entry(
            form, textvariable=self.var_pai_id, width=14)
        self._ent_pai_id.grid(row=1, column=1, sticky="w",
                              padx=(0, 10), pady=6)

        self._btn_buscar_pai = ttk.Button(
            form, text="Buscar Pai...", command=self.buscar_pai)
        self._btn_buscar_pai.grid(
            row=1, column=2, sticky="w", padx=(0, 10), pady=6)

        self._btn_remover_pai = ttk.Button(
            form, text="Remover Pai", command=self.remover_pai)
        self._btn_remover_pai.grid(
            row=1, column=3, sticky="w", padx=(0, 10), pady=6)

        ttk.Label(form, text="Pai (Nome):").grid(
            row=1, column=4, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_pai_nome, state="readonly").grid(
            row=1, column=5, sticky="ew", padx=(0, 10), pady=6, columnspan=7
        )

        # Lista
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

        self.tree.heading("codigo", text="Código")
        self.tree.heading("nome", text="Nome Categoria")
        self.tree.heading("pai", text="Pai")

        self.tree.column("codigo", width=130, anchor="e", stretch=False)
        self.tree.column("nome", width=700, anchor="w", stretch=True)
        self.tree.column("pai", width=140, anchor="e", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar categoria:\n{e}")
            return

        for c in rows:
            self.tree.insert("", "end", values=(c.codigo, c.nome, c.pai))

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        codigo, nome, pai = self.tree.item(sel[0], "values")
        self.var_codigo.set(str(codigo or ""))
        self.var_nome.set(str(nome or ""))
        self.var_pai_id.set(str(pai or ""))
        self._atualizar_pai_nome()

    def _atualizar_pai_nome(self):
        pid = (self.var_pai_id.get() or "").strip()
        if not pid:
            self.var_pai_nome.set("")
            return
        try:
            self.var_pai_nome.set(self.service.nome_por_codigo(pid) or "")
        except Exception:
            self.var_pai_nome.set("")

    def buscar_pai(self):
        if not self._can_edit():
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return

        def picked(cod, nome):
            self.var_pai_id.set(cod)
            self.var_pai_nome.set(nome)

        SeletorPai(self.winfo_toplevel(), self.service, picked)

    def remover_pai(self):
        if not self._can_edit():
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return
        self.var_pai_id.set("")
        self.var_pai_nome.set("")

    def limpar_form(self) -> None:
        self.var_codigo.set("")
        self.var_nome.set("")
        self.var_pai_id.set("")
        self.var_pai_nome.set("")
        self.tree.selection_remove(self.tree.selection())

    def novo(self) -> None:
        if not self._can_edit():
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return

        self.limpar_form()
        try:
            prox = self.service.proximo_codigo()
            self.var_codigo.set(str(prox))
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao gerar próximo código:\n{e}")

    def salvar(self) -> None:
        if not self._can_edit():
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return

        try:
            status, codigo = self.service.salvar(
                self.var_codigo.get(),
                self.var_nome.get(),
                self.var_pai_id.get(),
            )
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        self.var_codigo.set(str(codigo))
        self._atualizar_pai_nome()
        messagebox.showinfo(
            "OK", f"Categoria {status} com sucesso.\nCódigo: {codigo}")
        self.atualizar_lista()

    def excluir(self) -> None:
        if not self._can_edit():
            messagebox.showwarning("Acesso", "Seu nível é somente leitura.")
            return

        codigo = (self.var_codigo.get() or "").strip()
        if not codigo:
            messagebox.showwarning(
                "Atenção", "Selecione um registro para excluir.")
            return

        if not messagebox.askyesno("Confirmar", f"Excluir Categoria {codigo}?"):
            return

        try:
            self.service.excluir(codigo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Registro excluído.")
        self.limpar_form()
        self.atualizar_lista()


# ============================================================
# STARTUP
# ============================================================

def test_connection_or_die(cfg: AppConfig) -> None:
    conn = None
    try:
        conn = db_connect(cfg)
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

    # Descobre usuário e permissão (se não vier, abre em leitura)
    cli_user_id = _parse_cli_user()

    if cli_user_id is None:
        usuario_id = 0
        nivel = 1
        aviso = (
            "Atenção: usuário não informado ao abrir a tela.\n\n"
            "Abrindo em NÍVEL 1 (Leitura).\n"
            "Para respeitar permissões reais, chame com --usuario-id <id>."
        )
        usuario_nome = ""
    else:
        usuario_id = int(cli_user_id)
        nivel, aviso = get_access_level_for_this_screen(cfg, usuario_id)
        if nivel <= 0:
            _deny_and_exit(aviso)
        usuario_nome = fetch_user_nome(cfg, usuario_id)

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

    categoria_table = detectar_tabela(cfg, CATEGORIA_TABLES, '"categoria"')

    db = Database(cfg)
    repo = CategoriaRepo(db, categoria_table=categoria_table)
    service = CategoriaService(repo)

    tela = TelaCategoria(
        root,
        service,
        usuario_logado_id=usuario_id,
        acesso_nivel=nivel,
        usuario_logado_nome=usuario_nome,
    )
    tela.pack(fill="both", expand=True)

    if aviso:
        root.after(200, lambda: messagebox.showwarning(
            "Aviso", aviso, parent=root))

    # Fechamento: NÃO abre menu novo. Apenas fecha esta tela.
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_categoria(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_categoria:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
