# estoque_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class EstoqueCRUDMixin:
    """
    CRUD para "Ekenox"."estoque"
    Colunas:
      - fkProduto (PK/Chave lógica)
      - saldoFisico
      - saldoVirtual
    """

    ESTOQUE_COLS = ("fkProduto", "saldoFisico", "saldoVirtual")

    def estoque_get(self, fk_produto: int) -> Optional[Dict[str, Any]]:
        """READ: busca 1 registro por fkProduto."""
        try:
            pid = int(fk_produto)
            sql = """
                SELECT
                    e."fkProduto",
                    e."saldoFisico",
                    e."saldoVirtual"
                FROM "Ekenox"."estoque" e
                WHERE e."fkProduto"::bigint = %s
                LIMIT 1;
            """
            self._q(sql, (pid,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.ESTOQUE_COLS, r))
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def estoque_list(self, limit: int = 500, offset: int = 0) -> List[Dict[str, Any]]:
        """READ: lista paginada."""
        try:
            sql = """
                SELECT
                    e."fkProduto",
                    e."saldoFisico",
                    e."saldoVirtual"
                FROM "Ekenox"."estoque" e
                ORDER BY e."fkProduto"::bigint
                LIMIT %s OFFSET %s;
            """
            self._q(sql, (int(limit), int(offset)))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.ESTOQUE_COLS, row)) for row in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ CREATE com parâmetros explícitos
    def estoque_create(
        self,
        fk_produto: int,
        saldo_fisico: float | int = 0,
        saldo_virtual: float | int = 0,
    ) -> Tuple[bool, str]:
        """CREATE: insere 1 linha por fkProduto."""
        try:
            pid = int(fk_produto)

            if self.estoque_get(pid):
                return False, f"Já existe estoque para fkProduto={pid}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."estoque"
                    ("fkProduto", "saldoFisico", "saldoVirtual")
                VALUES
                    (%s, %s, %s);
            """
            self._q(sql, (pid, saldo_fisico, saldo_virtual))
            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir estoque: {type(e).__name__}: {e}"

    def estoque_update(self, fk_produto: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        UPDATE: patch por fkProduto.
        Campos aceitos: saldoFisico, saldoVirtual
        """
        try:
            pid = int(fk_produto)

            allowed = {"saldoFisico", "saldoVirtual"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar (saldoFisico/saldoVirtual)."

            set_sql = ", ".join([f'e."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [pid]

            sql = f"""
                UPDATE "Ekenox"."estoque" e
                   SET {set_sql}
                 WHERE e."fkProduto"::bigint = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para fkProduto={pid}."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar estoque: {type(e).__name__}: {e}"

    def estoque_upsert(
        self,
        fk_produto: int,
        saldo_fisico: float | int | None = None,
        saldo_virtual: float | int | None = None,
    ) -> Tuple[bool, str]:
        """
        Se existe fkProduto -> update (somente campos != None)
        Senão -> create (defaults 0 para campos None)
        """
        pid = int(fk_produto)

        if self.estoque_get(pid):
            patch: Dict[str, Any] = {}
            if saldo_fisico is not None:
                patch["saldoFisico"] = saldo_fisico
            if saldo_virtual is not None:
                patch["saldoVirtual"] = saldo_virtual

            if not patch:
                return True, ""

            return self.estoque_update(pid, patch)

        return self.estoque_create(
            fk_produto=pid,
            saldo_fisico=saldo_fisico if saldo_fisico is not None else 0,
            saldo_virtual=saldo_virtual if saldo_virtual is not None else 0,
        )

    def estoque_delete(self, fk_produto: int) -> Tuple[bool, str]:
        """DELETE: exclui por fkProduto."""
        try:
            pid = int(fk_produto)
            sql = """
                DELETE FROM "Ekenox"."estoque"
                WHERE "fkProduto"::bigint = %s;
            """
            self._q(sql, (pid,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (fkProduto={pid})."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir estoque: {type(e).__name__}: {e}"

    # ------------------------------------------------------------
    # Helpers úteis (opcional, mas ajuda muito no app)
    # ------------------------------------------------------------

    def estoque_get_saldos(self, fk_produto: int) -> Tuple[float, float]:
        """Retorna (saldoFisico, saldoVirtual). Se não existe, retorna (0, 0)."""
        r = self.estoque_get(fk_produto)
        if not r:
            return 0.0, 0.0
        sf = float(r.get("saldoFisico") or 0.0)
        sv = float(r.get("saldoVirtual") or 0.0)
        return sf, sv

    def estoque_set_saldos(self, fk_produto: int, saldo_fisico: float, saldo_virtual: float) -> Tuple[bool, str]:
        """Define saldos (upsert)."""
        return self.estoque_upsert(fk_produto, saldo_fisico=saldo_fisico, saldo_virtual=saldo_virtual)
