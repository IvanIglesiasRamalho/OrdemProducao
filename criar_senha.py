from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from getpass import getpass
from typing import Optional, Dict, Any

import psycopg2


# ========= AJUSTE SE PRECISAR =========
BASE_DIR = r"C:\Users\User\Desktop\Pyton\OrdemProducao"
USUARIOS_TABLE = '"Ekenox"."usuarios"'


# ========= CONFIG =========
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
    return AppConfig(db_host=host, db_port=port, db_database=dbname, db_user=user, db_password=password)


class Database:
    def __init__(self, cfg: AppConfig) -> None:
        self.cfg = cfg
        self.conn = None
        self.cur = None

    def __enter__(self):
        self.conn = psycopg2.connect(
            host=self.cfg.db_host,
            database=self.cfg.db_database,
            user=self.cfg.db_user,
            password=self.cfg.db_password,
            port=int(self.cfg.db_port),
            connect_timeout=5,
        )
        self.cur = self.conn.cursor()
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.cur:
                self.cur.close()
        except Exception:
            pass
        try:
            if self.conn:
                if exc is None:
                    self.conn.commit()
                else:
                    self.conn.rollback()
                self.conn.close()
        except Exception:
            pass
        self.cur = None
        self.conn = None


def _get_column_types(db: Database) -> Dict[str, str]:
    """
    Descobre o tipo das colunas email_hash / email_enc / senha_hash
    para gerar SQL compatível (text vs bytea).
    """
    sql = """
        SELECT column_name, data_type
        FROM information_schema.columns
        WHERE table_schema = 'Ekenox'
          AND table_name = 'usuarios'
          AND column_name IN ('email_hash','email_enc','senha_hash','nome','ativo');
    """
    db.cur.execute(sql)
    return {r[0]: r[1] for r in db.cur.fetchall()}


def _ensure_pgcrypto(db: Database) -> None:
    db.cur.execute('CREATE EXTENSION IF NOT EXISTS pgcrypto')


def criar_usuario(email: str, senha: str, nome: Optional[str], ativo: bool, secret: str) -> None:
    email_norm = (email or "").strip().lower()
    if not email_norm:
        raise ValueError("Email vazio.")
    if not senha:
        raise ValueError("Senha vazia.")

    cfg = env_override(load_config())

    with Database(cfg) as db:
        _ensure_pgcrypto(db)

        col_types = _get_column_types(db)
        email_hash_type = col_types.get("email_hash", "text")
        email_enc_type = col_types.get("email_enc", "bytea")
        senha_hash_type = col_types.get("senha_hash", "text")

        # --- Monta expressões conforme o tipo real da coluna ---
        # email_hash:
        # - se for text: encode(digest(...),'hex') -> text
        # - se for bytea: digest(...) -> bytea
        if email_hash_type == "bytea":
            email_hash_expr = "digest(%s, 'sha256')"
            email_hash_param = email_norm.encode("utf-8")
        else:
            email_hash_expr = "encode(digest(%s, 'sha256'), 'hex')"
            email_hash_param = email_norm

        # email_enc:
        # - se for bytea: pgp_sym_encrypt(%s, %s) -> bytea
        # - se for text: encode(pgp_sym_encrypt(...), 'base64') -> text
        if email_enc_type == "bytea":
            email_enc_expr = "pgp_sym_encrypt(%s, %s)"
        else:
            email_enc_expr = "encode(pgp_sym_encrypt(%s, %s), 'base64')"

        # senha_hash:
        # geralmente é text (crypt retorna text)
        # se a coluna for bytea (raro), convertemos para bytea via convert_to
        if senha_hash_type == "bytea":
            senha_hash_expr = "convert_to(crypt(%s, gen_salt('bf', 12)), 'UTF8')"
        else:
            senha_hash_expr = "crypt(%s, gen_salt('bf', 12))"

        # Verifica se já existe (por email_hash)
        sql_exists = f"""
            SELECT "usuarioId"
            FROM {USUARIOS_TABLE}
            WHERE "email_hash" = ({email_hash_expr})
            LIMIT 1
        """
        db.cur.execute(sql_exists, (email_hash_param,))
        row = db.cur.fetchone()

        if row:
            # Atualiza senha / nome / ativo
            usuario_id = int(row[0])
            sql_upd = f"""
                UPDATE {USUARIOS_TABLE}
                   SET "senha_hash" = ({senha_hash_expr}),
                       "nome" = %s,
                       "ativo" = %s,
                       "atualizado_em" = NOW()
                 WHERE "usuarioId" = %s
            """
            db.cur.execute(sql_upd, (senha, nome, ativo, usuario_id))
            print(
                f"Usuário já existia. Senha atualizada. usuarioId={usuario_id}")
            return

        # Insere novo
        sql_ins = f"""
            INSERT INTO {USUARIOS_TABLE}
                ("email_hash","email_enc","senha_hash","nome","ativo","criado_em","atualizado_em")
            VALUES
                ( ({email_hash_expr}),
                  ({email_enc_expr}),
                  ({senha_hash_expr}),
                  %s, %s, NOW(), NOW()
                )
            RETURNING "usuarioId"
        """

        params = (
            email_hash_param,
            email_norm, secret,
            senha,
            nome,
            ativo
        )
        db.cur.execute(sql_ins, params)
        new_id = db.cur.fetchone()[0]
        print(f"OK. Usuário criado com sucesso. usuarioId={new_id}")


def main() -> int:
    os.makedirs(BASE_DIR, exist_ok=True)

    try:
        email = input("Email: ").strip()
        if not email:
            print("Cancelado: email vazio.")
            return 1

        senha = getpass("Senha: ").strip()
        if not senha:
            print("Cancelado: senha vazia.")
            return 1

        nome = input("Nome (opcional): ").strip() or None

        # Chave de criptografia do email (pgp_sym_encrypt).
        # IMPORTANTE: use sempre a mesma (se mudar, não consegue descriptografar depois).
        # Sugestão: colocar em variável de ambiente EMAIL_SECRET.
        secret = (os.getenv("EMAIL_SECRET") or "").strip()
        if not secret:
            # fallback simples: cria/usa um arquivo local
            secret_file = os.path.join(BASE_DIR, ".email_secret")
            if os.path.exists(secret_file):
                secret = open(secret_file, "r",
                              encoding="utf-8").read().strip()
            else:
                # se quiser, você pode trocar isso por uma string fixa sua
                secret = "EKENOX_EMAIL_SECRET_2026"
                with open(secret_file, "w", encoding="utf-8") as f:
                    f.write(secret)

        criar_usuario(email=email, senha=senha, nome=nome,
                      ativo=True, secret=secret)
        return 0

    except KeyboardInterrupt:
        print("\nCancelado pelo usuário.")
        return 130
    except Exception as e:
        print(f"\nErro: {type(e).__name__}: {e}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
