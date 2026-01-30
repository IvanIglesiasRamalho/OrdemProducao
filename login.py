import os
from typing import Dict, Optional, Tuple
import psycopg2
from security import email_hash, verify_password

EMAIL_KEY = os.getenv("APP_EMAIL_KEY", "")


def auth_login(cfg, email: str, senha: str) -> Tuple[bool, str, Optional[int], Dict[str, int]]:
    """
    Retorna:
      ok, msg, usuarioId, permissoes {codigo_programa: nivel}
    """
    ehash = email_hash(email)

    conn = psycopg2.connect(
        host=cfg.db_host, database=cfg.db_database, user=cfg.db_user,
        password=cfg.db_password, port=int(cfg.db_port), connect_timeout=5
    )
    try:
        cur = conn.cursor()

        cur.execute("""
            SELECT "usuarioId", "senha_hash", "ativo"
            FROM "Ekenox"."usuarios"
            WHERE "email_hash" = %s
            LIMIT 1
        """, (ehash,))
        row = cur.fetchone()
        if not row:
            return False, "Usuário ou senha inválidos.", None, {}

        usuario_id, senha_hash_db, ativo = row
        if not ativo:
            return False, "Usuário inativo.", None, {}

        if not verify_password(senha, senha_hash_db):
            return False, "Usuário ou senha inválidos.", None, {}

        cur.execute("""
            SELECT pr."codigo", up."nivel"
            FROM "Ekenox"."usuario_programa" up
            JOIN "Ekenox"."programas" pr ON pr."programaId" = up."programaId"
            WHERE up."usuarioId" = %s
        """, (usuario_id,))
        perms = {codigo: int(nivel) for codigo, nivel in cur.fetchall()}

        return True, "OK", int(usuario_id), perms

    finally:
        conn.close()
