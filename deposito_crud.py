# deposito_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class DepositoCRUDMixin:
    """
    CRUD para "Ekenox"."deposito"
    Colunas:
      - id (PK)
      - descricao
      - situacao
      - padrao
      - desconsiderarsaldo
    """

    DEPOSITO_COLS = ("id", "descricao", "situacao",
                     "padrao", "desconsiderarsaldo")

    def deposito_get(self, deposito_id: int) -> Optional[Dict[str, Any]]:
        try:
            did = int(deposito_id)
            sql = """
                SELECT
                    d."id",
                    d."descricao",
                    d."situacao",
                    d."padrao",
                    d."desconsiderarsaldo"
                FROM "Ekenox"."deposito" d
                WHERE d."id" = %s
                LIMIT 1;
            """
            self._q(sql, (did,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.DEPOSITO_COLS, r))

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def deposito_list(
        self,
        limit: int = 300,
        offset: int = 0,
        descricao_like: str = "",
        situacao: int | None = None,
        padrao: int | None = None,
        desconsiderarsaldo: int | None = None,
    ) -> List[Dict[str, Any]]:
        """
        Lista depósitos com filtros opcionais.
        - descricao_like: ILIKE na descricao
        - situacao/padrao/desconsiderarsaldo: filtra igualdade (se não for None)
        """
        try:
            limit = int(limit)
            offset = int(offset)
            descricao_like = (descricao_like or "").strip()

            where_parts: List[str] = []
            params: List[Any] = []

            if descricao_like:
                where_parts.append('d."descricao" ILIKE %s')
                params.append(f"%{descricao_like}%")

            if situacao is not None:
                where_parts.append('d."situacao" = %s')
                params.append(int(situacao))

            if padrao is not None:
                where_parts.append('d."padrao" = %s')
                params.append(int(padrao))

            if desconsiderarsaldo is not None:
                where_parts.append('d."desconsiderarsaldo" = %s')
                params.append(int(desconsiderarsaldo))

            where_sql = ""
            if where_parts:
                where_sql = "WHERE " + " AND ".join(where_parts)

            sql = f"""
                SELECT
                    d."id",
                    d."descricao",
                    d."situacao",
                    d."padrao",
                    d."desconsiderarsaldo"
                FROM "Ekenox"."deposito" d
                {where_sql}
                ORDER BY d."descricao"
                LIMIT %s OFFSET %s;
            """
            params.extend([limit, offset])

            self._q(sql, tuple(params))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.DEPOSITO_COLS, row)) for row in rows]

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ CREATE com parâmetros explícitos
    def deposito_create(
        self,
        deposito_id: int,
        descricao: str,
        situacao: int = 1,
        padrao: int = 0,
        desconsiderarsaldo: int = 0,
    ) -> Tuple[bool, str]:
        """
        Insere depósito.
        - deposito_id e descricao obrigatórios.
        - situacao/padrao/desconsiderarsaldo default 1/0/0.
        """
        try:
            did = int(deposito_id)
            desc = (descricao or "").strip()
            if not desc:
                return False, "descricao é obrigatório."

            if self.deposito_get(did):
                return False, f"Já existe depósito id={did}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."deposito"
                    ("id", "descricao", "situacao", "padrao", "desconsiderarsaldo")
                VALUES
                    (%s, %s, %s, %s, %s);
            """
            self._q(sql, (did, desc, int(situacao), int(
                padrao), int(desconsiderarsaldo)))
            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir depósito: {type(e).__name__}: {e}"

    def deposito_update(self, deposito_id: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Atualiza campos (patch) do depósito.
        Campos aceitos em data: descricao, situacao, padrao, desconsiderarsaldo
        """
        try:
            did = int(deposito_id)

            allowed = {"descricao", "situacao", "padrao", "desconsiderarsaldo"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar (descricao/situacao/padrao/desconsiderarsaldo)."

            # normaliza descricao se vier
            if "descricao" in campos:
                data["descricao"] = (data.get("descricao") or "").strip()

            set_sql = ", ".join([f'd."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [did]

            sql = f"""
                UPDATE "Ekenox"."deposito" d
                   SET {set_sql}
                 WHERE d."id" = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para id={did}."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar depósito: {type(e).__name__}: {e}"

    def deposito_upsert(
        self,
        deposito_id: int,
        descricao: str | None = None,
        situacao: int | None = None,
        padrao: int | None = None,
        desconsiderarsaldo: int | None = None,
    ) -> Tuple[bool, str]:
        """
        Se existe id -> update (apenas campos não-None)
        Senão -> create (descricao obrigatória para criar)
        """
        did = int(deposito_id)
        existe = self.deposito_get(did)

        if existe:
            patch: Dict[str, Any] = {}
            if descricao is not None:
                patch["descricao"] = (descricao or "").strip()
            if situacao is not None:
                patch["situacao"] = int(situacao)
            if padrao is not None:
                patch["padrao"] = int(padrao)
            if desconsiderarsaldo is not None:
                patch["desconsiderarsaldo"] = int(desconsiderarsaldo)

            if not patch:
                return True, ""

            return self.deposito_update(did, patch)

        # create exige descricao
        if descricao is None or not (descricao or "").strip():
            return False, "Para criar (depósito novo), informe descricao."

        return self.deposito_create(
            deposito_id=did,
            descricao=descricao,
            situacao=int(situacao) if situacao is not None else 1,
            padrao=int(padrao) if padrao is not None else 0,
            desconsiderarsaldo=int(
                desconsiderarsaldo) if desconsiderarsaldo is not None else 0,
        )

    def deposito_delete(self, deposito_id: int) -> Tuple[bool, str]:
        try:
            did = int(deposito_id)
            sql = """
                DELETE FROM "Ekenox"."deposito"
                WHERE "id" = %s;
            """
            self._q(sql, (did,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (id={did})."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir depósito: {type(e).__name__}: {e}"
