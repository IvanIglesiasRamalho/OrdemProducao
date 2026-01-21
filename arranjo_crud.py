# arranjo_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class ArranjoCRUDMixin:
    ARRANJO_COLS = ("sku", "nomeproduto", "quantidade", "chapa", "material")

    def arranjo_get(self, sku: str) -> Optional[Dict[str, Any]]:
        try:
            sku = (sku or "").strip()
            if not sku:
                return None

            sql = """
                SELECT
                    a."sku",
                    a."nomeproduto",
                    a."quantidade",
                    a."chapa",
                    a."material"
                FROM "Ekenox"."arranjo" a
                WHERE a."sku" = %s
                LIMIT 1;
            """
            self._q(sql, (sku,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.ARRANJO_COLS, r))
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def arranjo_list(self, limit: int = 200, offset: int = 0, sku_like: str = "") -> List[Dict[str, Any]]:
        try:
            sku_like = (sku_like or "").strip()

            if sku_like:
                sql = """
                    SELECT
                        a."sku",
                        a."nomeproduto",
                        a."quantidade",
                        a."chapa",
                        a."material"
                    FROM "Ekenox"."arranjo" a
                    WHERE a."sku" ILIKE %s
                    ORDER BY a."sku"
                    LIMIT %s OFFSET %s;
                """
                params = (f"%{sku_like}%", int(limit), int(offset))
            else:
                sql = """
                    SELECT
                        a."sku",
                        a."nomeproduto",
                        a."quantidade",
                        a."chapa",
                        a."material"
                    FROM "Ekenox"."arranjo" a
                    ORDER BY a."sku"
                    LIMIT %s OFFSET %s;
                """
                params = (int(limit), int(offset))

            self._q(sql, params)
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.ARRANJO_COLS, row)) for row in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ CREATE com parâmetros
    def arranjo_create(
        self,
        sku: str,
        nomeproduto: str | None = None,
        quantidade: float | int | None = None,
        chapa: str | None = None,
        material: str | None = None,
    ) -> Tuple[bool, str]:
        """
        CREATE: Insere novo arranjo.
        Chave assumida: sku
        """
        try:
            sku = (sku or "").strip()
            if not sku:
                return False, "sku é obrigatório para inserir arranjo."

            if self.arranjo_get(sku):
                return False, f"Já existe arranjo para sku={sku}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."arranjo"
                    ("sku", "nomeproduto", "quantidade", "chapa", "material")
                VALUES
                    (%s, %s, %s, %s, %s);
            """
            self._q(sql, (sku, nomeproduto, quantidade, chapa, material))
            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir arranjo: {type(e).__name__}: {e}"

    def arranjo_update(self, sku: str, data: Dict[str, Any]) -> Tuple[bool, str]:
        try:
            sku = (sku or "").strip()
            if not sku:
                return False, "sku inválido."

            allowed = set(self.ARRANJO_COLS) - {"sku"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar."

            set_sql = ", ".join([f'a."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [sku]

            sql = f"""
                UPDATE "Ekenox"."arranjo" a
                   SET {set_sql}
                 WHERE a."sku" = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para sku={sku}."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar arranjo: {type(e).__name__}: {e}"

    # ✅ UPSERT com parâmetros (chama create/update)
    def arranjo_upsert(
        self,
        sku: str,
        nomeproduto: str | None = None,
        quantidade: float | int | None = None,
        chapa: str | None = None,
        material: str | None = None,
    ) -> Tuple[bool, str]:
        """
        UPSERT: se existe sku -> update; senão -> create.
        (no update, só altera campos que vierem != None)
        """
        sku = (sku or "").strip()
        if not sku:
            return False, "sku é obrigatório no UPSERT."

        if self.arranjo_get(sku):
            patch: Dict[str, Any] = {}
            if nomeproduto is not None:
                patch["nomeproduto"] = nomeproduto
            if quantidade is not None:
                patch["quantidade"] = quantidade
            if chapa is not None:
                patch["chapa"] = chapa
            if material is not None:
                patch["material"] = material

            if not patch:
                return True, ""  # nada pra alterar

            return self.arranjo_update(sku, patch)

        return self.arranjo_create(
            sku=sku,
            nomeproduto=nomeproduto,
            quantidade=quantidade,
            chapa=chapa,
            material=material,
        )

    def arranjo_delete(self, sku: str) -> Tuple[bool, str]:
        try:
            sku = (sku or "").strip()
            if not sku:
                return False, "sku inválido."

            sql = """
                DELETE FROM "Ekenox"."arranjo"
                WHERE "sku" = %s;
            """
            self._q(sql, (sku,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (sku={sku})."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir arranjo: {type(e).__name__}: {e}"
