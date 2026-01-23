from __future__ import annotations

"""
tela_fornecedor.py
Python 3.12+ | Postgres 16+ | Tkinter + psycopg2 (sem pool)

Tabela fornecedor:
  nome, codigo, situacao, numeroDocumentaca(o), telefone, celular, idFornecedor

Recursos:
- CRUD completo
- Busca (ILIKE) convertendo campos para TEXT (evita erro double precision)
- Detecta automaticamente a coluna do documento:
    "numeroDocumentacao" OU "numeroDocumentaca"
- Botão "Novo": mostra o próximo código via nextval() (com quote correto de schema/table/seq)
- Se não houver sequence vinculada ao idFornecedor, tenta localizar; se não achar, tenta criar e vincular
- Ao fechar: reabre menu_principal.py com --skip-entrada
"""

import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass
from typing import Optional, Any, List, Dict, Tuple

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


def log_fornecedor(msg: str) -> None:
    _log_write("tela_fornecedor.log", msg)


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

FORNECEDOR_TABLES = [
    '"Ekenox"."fornecedor"',
    '"fornecedor"',
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
        log_fornecedor(f"TABELA {'OK' if ok else 'FAIL'}: {t}")
        if ok:
            return t
    return fallback


# ============================================================
# DETECT COLUNA (numeroDocumentacao vs numeroDocumentaca)
# ============================================================

def _col_exists(cfg: AppConfig, table_name: str, col_name: str) -> bool:
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
        cur.execute(f'SELECT {col_name} FROM {table_name} LIMIT 1')
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


def detectar_coluna(cfg: AppConfig, table_name: str, candidates: List[str], fallback: str) -> str:
    for c in candidates:
        ok = _col_exists(cfg, table_name, c)
        log_fornecedor(f"COLUNA {'OK' if ok else 'FAIL'}: {table_name}.{c}")
        if ok:
            return c
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
        log_fornecedor(
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
        log_fornecedor(f"MENU: iniciado -> {cmd}")

    except Exception as e:
        log_fornecedor(f"MENU: erro ao abrir: {type(e).__name__}: {e}")


# ============================================================
# MODEL
# ============================================================

@dataclass
class Fornecedor:
    idFornecedor: int
    nome: str
    codigo: Optional[str]
    situacao: Optional[str]
    numeroDocumento: Optional[str]   # nome “neutro” no app
    telefone: Optional[str]
    celular: Optional[str]


def _clean_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def _qident(name: str) -> str:
    """Quote seguro para identificadores (schema/tabela/seq)."""
    return '"' + str(name).replace('"', '""') + '"'


# ============================================================
# REPOSITORY
# ============================================================

class FornecedorRepo:
    def __init__(self, db: Database, fornecedor_table: str, col_doc: str) -> None:
        self.db = db
        # ex: '"Ekenox"."fornecedor"' ou '"fornecedor"'
        self.table = fornecedor_table
        # ex: '"numeroDocumentaca"' ou '"numeroDocumentacao"'
        self.col_doc = col_doc
        self.pk_col = "idFornecedor"

    def _table_regclass_text(self) -> str:
        """
        Retorna o texto exato para casts/regclass: mantém aspas.
        Ex: '"Ekenox"."fornecedor"' ou '"fornecedor"'
        """
        return (self.table or "").strip()

    def _get_schema_and_table_real(self) -> Tuple[str, str]:
        """
        Descobre schema e nome real da tabela via OID (respeita case).
        """
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
                (self._table_regclass_text(),),
            )
            r = self.db.cursor.fetchone()
            if not r:
                raise RuntimeError(
                    f"Não consegui resolver a tabela via regclass: {self._table_regclass_text()}")
            return str(r[0]), str(r[1])
        finally:
            self.db.desconectar()

    def _find_sequence_via_pg_get_serial_sequence(self) -> Optional[str]:
        """
        Tenta achar sequence ligada por SERIAL/IDENTITY/DEFAULT.
        Retorna texto pronto para regclass (pode vir com aspas).
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                "SELECT pg_get_serial_sequence(%s, %s)",
                (self._table_regclass_text(), self.pk_col),
            )
            r = self.db.cursor.fetchone()
            if r and r[0]:
                return str(r[0])  # pode vir como: public.seq ou "Ekenox"."seq"
            return None
        finally:
            self.db.desconectar()

    def _find_sequence_owned_by_column(self) -> Optional[str]:
        """
        Tenta achar sequence 'owned by' a coluna (pg_depend).
        Retorna texto com quote_ident(schema).quote_ident(seq)
        """
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
                (self._table_regclass_text(), self.pk_col),
            )
            r = self.db.cursor.fetchone()
            if r and r[0]:
                return str(r[0])
            return None
        finally:
            self.db.desconectar()

    def _ensure_sequence_and_default(self) -> str:
        """
        Garante que exista uma sequence e que a coluna tenha DEFAULT nextval(seq).
        Retorna o nome qualificado (com aspas quando necessário), pronto para %s::regclass.
        """
        # 1) tenta métodos padrão
        seq = self._find_sequence_via_pg_get_serial_sequence()
        if seq:
            return seq

        seq2 = self._find_sequence_owned_by_column()
        if seq2:
            return seq2

        # 2) tenta criar e vincular (se tiver permissão)
        schema, table = self._get_schema_and_table_real()
        seq_name = f"{table}_{self.pk_col}_seq"

        seq_qual = f"{_qident(schema)}.{_qident(seq_name)}"
        table_qual = f"{_qident(schema)}.{_qident(table)}"
        col_qual = _qident(self.pk_col)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None

            # cria sequence se não existir
            self.db.cursor.execute(f"CREATE SEQUENCE IF NOT EXISTS {seq_qual}")

            # define owned by (boa prática)
            try:
                self.db.cursor.execute(
                    f"ALTER SEQUENCE {seq_qual} OWNED BY {table_qual}.{col_qual}")
            except Exception:
                # se não tiver permissão, segue (mas pode impactar drop automático)
                self.db.rollback()

            # seta default nextval na coluna
            try:
                self.db.cursor.execute(
                    f"ALTER TABLE {table_qual} ALTER COLUMN {col_qual} SET DEFAULT nextval(%s::regclass)",
                    (seq_qual,),
                )
            except Exception:
                self.db.rollback()
                raise

            # ajusta o valor da sequence para MAX(idFornecedor)+1 (sem "pular" no primeiro nextval)
            self.db.cursor.execute(
                f"""
                SELECT setval(
                    %s::regclass,
                    GREATEST((SELECT COALESCE(MAX({col_qual}), 0) FROM {table_qual}) + 1, 1),
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

    def proximo_id_nextval(self) -> int:
        """
        Reserva o próximo ID com nextval(). (Isso consome o número — comportamento esperado do nextval.)
        """
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

    def listar(self, termo: Optional[str] = None, limit: int = 1500) -> List[Fornecedor]:
        like = f"%{termo}%" if termo else None

        def ilike_text(expr_sql: str) -> str:
            return f"COALESCE(CAST({expr_sql} AS TEXT),'') ILIKE %s"

        sql = f"""
            SELECT
                f."idFornecedor",
                f."nome",
                f."codigo",
                f."situacao",
                f.{self.col_doc} AS doc,
                f."telefone",
                f."celular"
            FROM {self.table} AS f
            WHERE (%s IS NULL)
               OR {ilike_text('f.\"idFornecedor\"')}
               OR {ilike_text('f.\"nome\"')}
               OR {ilike_text('f.\"codigo\"')}
               OR {ilike_text('f.\"situacao\"')}
               OR {ilike_text(f'f.{self.col_doc}')}
               OR {ilike_text('f.\"telefone\"')}
               OR {ilike_text('f.\"celular\"')}
            ORDER BY f."idFornecedor" DESC
            LIMIT %s
        """

        params = (termo, like, like, like, like, like, like, like, limit)

        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")

        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, params)
            rows = self.db.cursor.fetchall()

            out: List[Fornecedor] = []
            for r in rows:
                out.append(Fornecedor(
                    idFornecedor=int(r[0]),
                    nome=str(r[1] or ""),
                    codigo=_clean_text(r[2]),
                    situacao=_clean_text(r[3]),
                    numeroDocumento=_clean_text(r[4]),
                    telefone=_clean_text(r[5]),
                    celular=_clean_text(r[6]),
                ))
            return out
        finally:
            self.db.desconectar()

    def existe_id(self, fornecedor_id: int) -> bool:
        sql = f'SELECT 1 FROM {self.table} WHERE "idFornecedor"=%s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fornecedor_id,))
            return self.db.cursor.fetchone() is not None
        finally:
            self.db.desconectar()

    def inserir(self, fornecedor_id: int, nome: str, codigo: Optional[str], situacao: Optional[str],
                doc: Optional[str], telefone: Optional[str], celular: Optional[str]) -> int:
        sql = f"""
            INSERT INTO {self.table}
              ("idFornecedor","nome","codigo","situacao",{self.col_doc},"telefone","celular")
            VALUES (%s,%s,%s,%s,%s,%s,%s)
            RETURNING "idFornecedor"
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                sql, (fornecedor_id, nome, codigo, situacao, doc, telefone, celular))
            new_id = self.db.cursor.fetchone()[0]
            self.db.commit()
            return int(new_id)
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def atualizar(self, fornecedor_id: int, nome: str, codigo: Optional[str], situacao: Optional[str],
                  doc: Optional[str], telefone: Optional[str], celular: Optional[str]) -> None:
        sql = f"""
            UPDATE {self.table}
               SET "nome"=%s,
                   "codigo"=%s,
                   "situacao"=%s,
                   {self.col_doc}=%s,
                   "telefone"=%s,
                   "celular"=%s
             WHERE "idFornecedor"=%s
        """
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(
                sql, (nome, codigo, situacao, doc, telefone, celular, fornecedor_id))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()

    def excluir(self, fornecedor_id: int) -> None:
        sql = f'DELETE FROM {self.table} WHERE "idFornecedor"=%s'
        if not self.db.conectar():
            raise RuntimeError(f"Falha ao conectar: {self.db.ultimo_erro}")
        try:
            assert self.db.cursor is not None
            self.db.cursor.execute(sql, (fornecedor_id,))
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise
        finally:
            self.db.desconectar()


# ============================================================
# SERVICE
# ============================================================

class FornecedorService:
    def __init__(self, repo: FornecedorRepo) -> None:
        self.repo = repo

    def proximo_id_nextval(self) -> int:
        return self.repo.proximo_id_nextval()

    def listar(self, termo: Optional[str]) -> List[Fornecedor]:
        termo = (termo or "").strip() or None
        return self.repo.listar(termo)

    def validar(self, nome: str) -> None:
        nome = (nome or "").strip()
        if not nome:
            raise ValueError("Nome é obrigatório.")

    def salvar(self, id_original: Optional[int], form: Dict[str, Any]) -> Tuple[str, int]:
        nome = (form.get("nome") or "").strip()
        codigo = _clean_text(form.get("codigo"))
        situacao = _clean_text(form.get("situacao"))
        doc = _clean_text(form.get("doc"))
        telefone = _clean_text(form.get("telefone"))
        celular = _clean_text(form.get("celular"))

        self.validar(nome)

        # id vindo do botão "Novo" (reservado via nextval)
        id_txt = (form.get("idFornecedor") or "").strip()
        fornecedor_id_form: Optional[int] = None
        if id_txt.isdigit():
            fornecedor_id_form = int(id_txt)

        if id_original is None:
            # se o usuário não clicou "Novo", reserva agora via nextval
            if fornecedor_id_form is None:
                fornecedor_id_form = self.repo.proximo_id_nextval()
            new_id = self.repo.inserir(
                fornecedor_id_form, nome, codigo, situacao, doc, telefone, celular)
            return ("inserido", new_id)

        if not self.repo.existe_id(id_original):
            # registro sumiu; insere novo (mantém id_original se quiser, mas aqui respeita o fluxo padrão)
            if fornecedor_id_form is None:
                fornecedor_id_form = self.repo.proximo_id_nextval()
            new_id = self.repo.inserir(
                fornecedor_id_form, nome, codigo, situacao, doc, telefone, celular)
            return ("inserido", new_id)

        self.repo.atualizar(id_original, nome, codigo,
                            situacao, doc, telefone, celular)
        return ("atualizado", id_original)

    def excluir(self, fornecedor_id: int) -> None:
        self.repo.excluir(fornecedor_id)


# ============================================================
# UI
# ============================================================

DEFAULT_GEOMETRY = "1200x700"
APP_TITLE = "Tela de Fornecedor"

TREE_COLS = ["idFornecedor", "nome", "codigo",
             "situacao", "doc", "telefone", "celular"]


class TelaFornecedor(ttk.Frame):
    def __init__(self, master: tk.Misc, service: FornecedorService):
        super().__init__(master)
        self.service = service

        self.var_filtro = tk.StringVar()

        self.var_id = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_codigo = tk.StringVar()
        self.var_situacao = tk.StringVar()
        self.var_doc = tk.StringVar()
        self.var_tel = tk.StringVar()
        self.var_cel = tk.StringVar()

        self._id_original: Optional[int] = None

        self._build_ui()
        self.atualizar_lista()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(
            top, text="Buscar (ID/Nome/Código/Doc/Fone):").grid(row=0, column=0, sticky="w")
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

        form = ttk.LabelFrame(self, text="Fornecedor")
        form.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        for c in range(12):
            form.columnconfigure(c, weight=1)

        ttk.Label(form, text="Código:").grid(
            row=0, column=0, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_id, state="readonly", width=12).grid(
            row=0, column=1, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Nome:").grid(
            row=0, column=2, sticky="w", padx=(10, 6), pady=6)
        ent_nome = ttk.Entry(form, textvariable=self.var_nome)
        ent_nome.grid(row=0, column=3, sticky="ew",
                      padx=(0, 10), pady=6, columnspan=5)

        ttk.Label(form, text="Código:").grid(
            row=0, column=8, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_codigo, width=18).grid(
            row=0, column=9, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Situação:").grid(
            row=1, column=0, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_situacao, width=22).grid(
            row=1, column=1, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Documento:").grid(
            row=1, column=2, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_doc, width=26).grid(
            row=1, column=3, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Telefone:").grid(
            row=1, column=5, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_tel, width=18).grid(
            row=1, column=6, sticky="w", padx=(0, 10), pady=6
        )

        ttk.Label(form, text="Celular:").grid(
            row=1, column=8, sticky="w", padx=(10, 6), pady=6)
        ttk.Entry(form, textvariable=self.var_cel, width=18).grid(
            row=1, column=9, sticky="w", padx=(0, 10), pady=6
        )

        ent_nome.bind("<Return>", lambda e: self.salvar())

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

        headings = {
            "idFornecedor": "Código",
            "nome": "Nome",
            "codigo": "Código Origem",
            "situacao": "Situação",
            "doc": "Documento",
            "telefone": "Telefone",
            "celular": "Celular",
        }
        for col in TREE_COLS:
            self.tree.heading(col, text=headings.get(col, col))

        self.tree.column("idFornecedor", width=110, anchor="e", stretch=False)
        self.tree.column("nome", width=360, anchor="w", stretch=True)
        self.tree.column("codigo", width=120, anchor="w", stretch=False)
        self.tree.column("situacao", width=120, anchor="w", stretch=False)
        self.tree.column("doc", width=170, anchor="w", stretch=False)
        self.tree.column("telefone", width=140, anchor="w", stretch=False)
        self.tree.column("celular", width=140, anchor="w", stretch=False)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def atualizar_lista(self) -> None:
        termo = self.var_filtro.get().strip() or None
        for it in self.tree.get_children():
            self.tree.delete(it)

        try:
            rows = self.service.listar(termo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao listar fornecedores:\n{e}")
            return

        for f in rows:
            self.tree.insert("", "end", values=(
                f.idFornecedor,
                f.nome,
                f.codigo or "",
                f.situacao or "",
                f.numeroDocumento or "",
                f.telefone or "",
                f.celular or "",
            ))

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")

        self._id_original = int(vals[0])
        self.var_id.set(str(vals[0] or ""))
        self.var_nome.set(str(vals[1] or ""))
        self.var_codigo.set(str(vals[2] or ""))
        self.var_situacao.set(str(vals[3] or ""))
        self.var_doc.set(str(vals[4] or ""))
        self.var_tel.set(str(vals[5] or ""))
        self.var_cel.set(str(vals[6] or ""))

    def novo(self) -> None:
        self.limpar_form()
        try:
            next_id = self.service.proximo_id_nextval()
            self.var_id.set(str(next_id))
            # IMPORTANT: como nextval "reserva" o número, guardamos como "original"
            # para que o salvar trate como INSERÇÃO (id_original None) mas com id preenchido.
            # Mantemos _id_original = None para o service inserir, não atualizar.
            self._id_original = None
        except Exception as e:
            messagebox.showerror(
                "Erro",
                "Falha ao obter próximo código (nextval).\n"
                "Obs: se não existir sequence vinculada, o programa tenta criar/vincular.\n"
                "Se falhar, pode ser falta de permissão no banco.\n\n"
                f"Detalhe:\n{e}"
            )

    def limpar_form(self) -> None:
        self._id_original = None
        self.var_id.set("")
        self.var_nome.set("")
        self.var_codigo.set("")
        self.var_situacao.set("")
        self.var_doc.set("")
        self.var_tel.set("")
        self.var_cel.set("")
        self.tree.selection_remove(self.tree.selection())

    def salvar(self) -> None:
        form = {
            "idFornecedor": self.var_id.get(),   # essencial p/ inserir com o id reservado
            "nome": self.var_nome.get(),
            "codigo": self.var_codigo.get(),
            "situacao": self.var_situacao.get(),
            "doc": self.var_doc.get(),
            "telefone": self.var_tel.get(),
            "celular": self.var_cel.get(),
        }

        try:
            status, new_id = self.service.salvar(self._id_original, form)
        except Exception as e:
            messagebox.showerror("Validação/Erro", str(e))
            return

        messagebox.showinfo(
            "OK", f"Fornecedor {status} com sucesso.\nCódigo: {new_id}")
        self._id_original = new_id
        self.var_id.set(str(new_id))
        self.atualizar_lista()

    def excluir(self) -> None:
        if self._id_original is None:
            messagebox.showwarning(
                "Atenção", "Selecione um fornecedor para excluir.")
            return

        fid = self._id_original
        if not messagebox.askyesno("Confirmar", f"Excluir Fornecedor Código {fid}?"):
            return

        try:
            self.service.excluir(fid)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao excluir:\n{e}")
            return

        messagebox.showinfo("OK", "Fornecedor excluído.")
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

    fornecedor_table = detectar_tabela(cfg, FORNECEDOR_TABLES, '"fornecedor"')

    # Detecta a coluna “documento” (variação no nome)
    col_doc = detectar_coluna(
        cfg,
        fornecedor_table,
        candidates=['"numeroDocumentacao"', '"numeroDocumentaca"'],
        fallback='"numeroDocumentaca"',
    )

    db = Database(cfg)
    repo = FornecedorRepo(
        db, fornecedor_table=fornecedor_table, col_doc=col_doc)
    service = FornecedorService(repo)

    tela = TelaFornecedor(root, service)
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
        log_fornecedor(f"FATAL: {type(e).__name__}: {e}")
        try:
            messagebox.showerror(
                "Erro", f"Falha ao iniciar tela_fornecedor:\n{type(e).__name__}: {e}")
        except Exception:
            pass
        raise
