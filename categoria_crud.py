# categoria_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class CategoriaCRUDMixin:
    """
    CRUD para "Ekenox"."categoria"
    Colunas:
      - categoriaId (PK)
      - nomeCategoria
      - categoriaPai (FK/relacionamento para categoriaId, pode ser NULL)
    """

    CATEGORIA_COLS = ("categoriaId", "nomeCategoria", "categoriaPai")

    def categoria_get(self, categoria_id: int) -> Optional[Dict[str, Any]]:
        try:
            cid = int(categoria_id)

            sql = """
                SELECT
                    c."categoriaId",
                    c."nomeCategoria",
                    c."categoriaPai"
                FROM "Ekenox"."categoria" c
                WHERE c."categoriaId" = %s
                LIMIT 1;
            """
            self._q(sql, (cid,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.CATEGORIA_COLS, r))

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def categoria_list(
        self,
        limit: int = 200,
        offset: int = 0,
        nome_like: str = "",
        categoria_pai: int | None = None,
    ) -> List[Dict[str, Any]]:
        """
        Lista categorias com filtros opcionais:
          - nome_like: busca por nomeCategoria (ILIKE)
          - categoria_pai: filtra por categoriaPai (inclusive NULL se passar None? -> não filtra)
        """
        try:
            limit = int(limit)
            offset = int(offset)
            nome_like = (nome_like or "").strip()

            where_parts = []
            params: List[Any] = []

            if nome_like:
                where_parts.append('c."nomeCategoria" ILIKE %s')
                params.append(f"%{nome_like}%")

            if categoria_pai is not None:
                where_parts.append('c."categoriaPai" = %s')
                params.append(int(categoria_pai))

            where_sql = ""
            if where_parts:
                where_sql = "WHERE " + " AND ".join(where_parts)

            sql = f"""
                SELECT
                    c."categoriaId",
                    c."nomeCategoria",
                    c."categoriaPai"
                FROM "Ekenox"."categoria" c
                {where_sql}
                ORDER BY c."nomeCategoria"
                LIMIT %s OFFSET %s;
            """
            params.extend([limit, offset])

            self._q(sql, tuple(params))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.CATEGORIA_COLS, row)) for row in rows]

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ CREATE com parâmetros explícitos
    def categoria_create(
        self,
        categoria_id: int,
        nome_categoria: str,
        categoria_pai: int | None = None,
    ) -> Tuple[bool, str]:
        """
        Insere uma categoria.
        - categoria_id e nome_categoria são obrigatórios.
        - categoria_pai pode ser NULL.
        """
        try:
            cid = int(categoria_id)
            nome = (nome_categoria or "").strip()
            if not nome:
                return False, "nome_categoria é obrigatório."

            if self.categoria_get(cid):
                return False, f"Já existe categoriaId={cid}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."categoria"
                    ("categoriaId", "nomeCategoria", "categoriaPai")
                VALUES
                    (%s, %s, %s);
            """
            self._q(sql, (cid, nome, int(categoria_pai)
                    if categoria_pai is not None else None))
            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir categoria: {type(e).__name__}: {e}"

    # Update em estilo "patch" (dict) para manter consistente com o resto do seu projeto
    def categoria_update(self, categoria_id: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Atualiza campos (patch) da categoria.
        Campos aceitos em data: nomeCategoria, categoriaPai
        """
        try:
            cid = int(categoria_id)

            allowed = {"nomeCategoria", "categoriaPai"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar (use nomeCategoria/categoriaPai)."

            set_sql = ", ".join([f'c."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [cid]

            sql = f"""
                UPDATE "Ekenox"."categoria" c
                   SET {set_sql}
                 WHERE c."categoriaId" = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para categoriaId={cid}."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar categoria: {type(e).__name__}: {e}"

    # ✅ UPSERT com parâmetros explícitos
    def categoria_upsert(
        self,
        categoria_id: int,
        nome_categoria: str | None = None,
        categoria_pai: int | None = None,
    ) -> Tuple[bool, str]:
        """
        Se existe categoriaId -> update (apenas campos não-None)
        Senão -> create (nome_categoria obrigatório para criar)
        """
        cid = int(categoria_id)
        existe = self.categoria_get(cid)

        if existe:
            patch: Dict[str, Any] = {}
            if nome_categoria is not None:
                patch["nomeCategoria"] = (nome_categoria or "").strip()
            if categoria_pai is not None:
                patch["categoriaPai"] = int(categoria_pai)

            if not patch:
                return True, ""

            return self.categoria_update(cid, patch)

        # create exige nome_categoria
        if nome_categoria is None or not (nome_categoria or "").strip():
            return False, "Para criar (categoria nova), informe nome_categoria."

        return self.categoria_create(
            categoria_id=cid,
            nome_categoria=nome_categoria,
            categoria_pai=categoria_pai,
        )

    def categoria_delete(self, categoria_id: int) -> Tuple[bool, str]:
        try:
            cid = int(categoria_id)

            sql = """
                DELETE FROM "Ekenox"."categoria"
                WHERE "categoriaId" = %s;
            """
            self._q(sql, (cid,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (categoriaId={cid})."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir categoria: {type(e).__name__}: {e}"
