# situacao_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class SituacaoCRUDMixin:
    """
    CRUD para "Ekenox"."situacao"
    Colunas:
      - id (PK)
      - nome
      - idHerdado (referência para situacao.id, pode ser NULL)
    """

    SITUACAO_COLS = ("id", "nome", "idHerdado")

    # ----------------------------
    # READ
    # ----------------------------

    def situacao_get(self, situacao_id: int) -> Optional[Dict[str, Any]]:
        """Busca 1 situação por id."""
        try:
            sid = int(situacao_id)
            sql = """
                SELECT
                    s."id",
                    s."nome",
                    s."idHerdado"
                FROM "Ekenox"."situacao" s
                WHERE s."id" = %s
                LIMIT 1;
            """
            self._q(sql, (sid,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.SITUACAO_COLS, r))
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def situacao_get_com_herdado(self, situacao_id: int) -> Optional[Dict[str, Any]]:
        """
        Busca 1 situação e traz o nome da situação herdada (self-join).
        Retorno:
          {
            id, nome, idHerdado,
            herdado_nome
          }
        """
        try:
            sid = int(situacao_id)
            sql = """
                SELECT
                    s."id",
                    s."nome",
                    s."idHerdado",
                    sh."nome" AS herdado_nome
                FROM "Ekenox"."situacao" s
                LEFT JOIN "Ekenox"."situacao" sh
                       ON sh."id" = s."idHerdado"
                WHERE s."id" = %s
                LIMIT 1;
            """
            self._q(sql, (sid,))
            r = self.cursor.fetchone()
            if not r:
                return None

            return {
                "id": r[0],
                "nome": r[1],
                "idHerdado": r[2],
                "herdado_nome": r[3],
            }
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def situacao_list(
        self,
        limit: int = 300,
        offset: int = 0,
        nome_like: str = "",
        id_herdado: int | None = None,
    ) -> List[Dict[str, Any]]:
        """Lista situações com filtros opcionais."""
        try:
            limit = int(limit)
            offset = int(offset)
            nome_like = (nome_like or "").strip()

            where_parts: List[str] = []
            params: List[Any] = []

            if nome_like:
                where_parts.append('s."nome" ILIKE %s')
                params.append(f"%{nome_like}%")

            if id_herdado is not None:
                where_parts.append('s."idHerdado" = %s')
                params.append(int(id_herdado))

            where_sql = ""
            if where_parts:
                where_sql = "WHERE " + " AND ".join(where_parts)

            sql = f"""
                SELECT
                    s."id",
                    s."nome",
                    s."idHerdado"
                FROM "Ekenox"."situacao" s
                {where_sql}
                ORDER BY s."nome"
                LIMIT %s OFFSET %s;
            """
            params.extend([limit, offset])

            self._q(sql, tuple(params))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.SITUACAO_COLS, row)) for row in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    def situacao_list_com_herdado(
        self,
        limit: int = 300,
        offset: int = 0,
        nome_like: str = "",
    ) -> List[Dict[str, Any]]:
        """
        Lista situações já com 'herdado_nome' (bom para grid/relatório).
        """
        try:
            limit = int(limit)
            offset = int(offset)
            nome_like = (nome_like or "").strip()

            if nome_like:
                sql = """
                    SELECT
                        s."id",
                        s."nome",
                        s."idHerdado",
                        sh."nome" AS herdado_nome
                    FROM "Ekenox"."situacao" s
                    LEFT JOIN "Ekenox"."situacao" sh
                           ON sh."id" = s."idHerdado"
                    WHERE s."nome" ILIKE %s
                    ORDER BY s."nome"
                    LIMIT %s OFFSET %s;
                """
                params = (f"%{nome_like}%", limit, offset)
            else:
                sql = """
                    SELECT
                        s."id",
                        s."nome",
                        s."idHerdado",
                        sh."nome" AS herdado_nome
                    FROM "Ekenox"."situacao" s
                    LEFT JOIN "Ekenox"."situacao" sh
                           ON sh."id" = s."idHerdado"
                    ORDER BY s."nome"
                    LIMIT %s OFFSET %s;
                """
                params = (limit, offset)

            self._q(sql, params)
            rows = self.cursor.fetchall() or []
            out = []
            for r in rows:
                out.append({
                    "id": r[0],
                    "nome": r[1],
                    "idHerdado": r[2],
                    "herdado_nome": r[3],
                })
            return out
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ Auxiliar: lista leve para combo
    def aux_situacao_list_para_combo(self, somente_ativas: bool = False) -> List[Dict[str, Any]]:
        """
        Retorna lista leve: [{"id":..., "nome":...}, ...]
        'somente_ativas' fica aqui para você reaproveitar caso depois exista alguma regra
        (como tabela de status ativa/inativa). Hoje não filtra porque a tabela não tem campo.
        """
        try:
            sql = """
                SELECT s."id", s."nome"
                FROM "Ekenox"."situacao" s
                ORDER BY s."nome";
            """
            self._q(sql, ())
            rows = self.cursor.fetchall() or []
            return [{"id": r[0], "nome": r[1]} for r in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ----------------------------
    # CREATE/UPDATE/UPSERT/DELETE
    # ----------------------------

    def situacao_create(
        self,
        situacao_id: int,
        nome: str,
        id_herdado: int | None = None,
    ) -> Tuple[bool, str]:
        """CREATE com parâmetros explícitos."""
        try:
            sid = int(situacao_id)
            nome = (nome or "").strip()
            if not nome:
                return False, "nome é obrigatório."

            if self.situacao_get(sid):
                return False, f"Já existe situação id={sid}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."situacao"
                    ("id", "nome", "idHerdado")
                VALUES
                    (%s, %s, %s);
            """
            self._q(sql, (sid, nome, int(id_herdado)
                    if id_herdado is not None else None))
            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir situação: {type(e).__name__}: {e}"

    def situacao_update(self, situacao_id: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        UPDATE patch por id.
        Campos aceitos: nome, idHerdado
        """
        try:
            sid = int(situacao_id)

            allowed = {"nome", "idHerdado"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar (nome/idHerdado)."

            if "nome" in campos:
                data["nome"] = (data.get("nome") or "").strip()

            set_sql = ", ".join([f's."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [sid]

            sql = f"""
                UPDATE "Ekenox"."situacao" s
                   SET {set_sql}
                 WHERE s."id" = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para situação id={sid}."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar situação: {type(e).__name__}: {e}"

    def situacao_upsert(
        self,
        situacao_id: int,
        nome: str | None = None,
        id_herdado: int | None = None,
    ) -> Tuple[bool, str]:
        """
        UPSERT com parâmetros:
          - se existe -> update só campos != None
          - se não existe -> create (nome obrigatório)
        """
        sid = int(situacao_id)
        existe = self.situacao_get(sid)

        if existe:
            patch: Dict[str, Any] = {}
            if nome is not None:
                patch["nome"] = (nome or "").strip()
            if id_herdado is not None:
                patch["idHerdado"] = int(id_herdado)

            if not patch:
                return True, ""

            return self.situacao_update(sid, patch)

        if nome is None or not (nome or "").strip():
            return False, "Para criar situação nova, informe nome."

        return self.situacao_create(situacao_id=sid, nome=nome, id_herdado=id_herdado)

    def situacao_delete(self, situacao_id: int) -> Tuple[bool, str]:
        """DELETE por id."""
        try:
            sid = int(situacao_id)
            sql = """
                DELETE FROM "Ekenox"."situacao"
                WHERE "id" = %s;
            """
            self._q(sql, (sid,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (id={sid})."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir situação: {type(e).__name__}: {e}"
