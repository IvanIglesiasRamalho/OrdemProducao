# info_produto_crud.py
from __future__ import annotations

from typing import Optional, Dict, Any, List, Tuple


class InfoProdutoCRUDMixin:
    """
    Mixin CRUD para a tabela "Ekenox"."infoProduto".

    Requer que a classe que herdar tenha:
      - self._q(sql: str, params: tuple)
      - self.conn (psycopg2 connection)
      - self.cursor (psycopg2 cursor)
    """

    INFO_PRODUTO_COLS = (
        "estoqueMinimo",
        "estoqueMaximo",
        "estoqueLocalizacao",
        "unidade",
        "pesoLiquido",
        "pesoBruto",
        "volumes",
        "itensPorCaixa",
        "gtin",
        "tipoProducao",
        "marca",
        "precoCompra",
        "largura",
        "altura",
        "profundidade",
        "unidadeMedida",
        "fkFornecedor",
        "fkCategoria",
        "prazoEntrega",
        "fkProduto",
    )

    def info_produto_get(self, fk_produto: int) -> Optional[Dict[str, Any]]:
        """READ: Busca 1 registro por fkProduto."""
        try:
            sql = """
                SELECT
                    i."estoqueMinimo",
                    i."estoqueMaximo",
                    i."estoqueLocalizacao",
                    i."unidade",
                    i."pesoLiquido",
                    i."pesoBruto",
                    i."volumes",
                    i."itensPorCaixa",
                    i."gtin",
                    i."tipoProducao",
                    i."marca",
                    i."precoCompra",
                    i."largura",
                    i."altura",
                    i."profundidade",
                    i."unidadeMedida",
                    i."fkFornecedor",
                    i."fkCategoria",
                    i."prazoEntrega",
                    i."fkProduto"
                FROM "Ekenox"."infoProduto" i
                WHERE i."fkProduto"::bigint = %s
                LIMIT 1;
            """
            self._q(sql, (int(fk_produto),))
            row = self.cursor.fetchone()
            if not row:
                return None

            keys = list(self.INFO_PRODUTO_COLS[:-1]) + ["fkProduto"]
            return dict(zip(keys, row))
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def info_produto_list(self, limit: int = 200, offset: int = 0) -> List[Dict[str, Any]]:
        """READ: Lista registros (paginado)."""
        try:
            sql = """
                SELECT
                    i."estoqueMinimo",
                    i."estoqueMaximo",
                    i."estoqueLocalizacao",
                    i."unidade",
                    i."pesoLiquido",
                    i."pesoBruto",
                    i."volumes",
                    i."itensPorCaixa",
                    i."gtin",
                    i."tipoProducao",
                    i."marca",
                    i."precoCompra",
                    i."largura",
                    i."altura",
                    i."profundidade",
                    i."unidadeMedida",
                    i."fkFornecedor",
                    i."fkCategoria",
                    i."prazoEntrega",
                    i."fkProduto"
                FROM "Ekenox"."infoProduto" i
                ORDER BY i."fkProduto"::bigint
                LIMIT %s OFFSET %s;
            """
            self._q(sql, (int(limit), int(offset)))
            rows = self.cursor.fetchall() or []
            keys = list(self.INFO_PRODUTO_COLS[:-1]) + ["fkProduto"]
            return [dict(zip(keys, r)) for r in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    def info_produto_create(self, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        CREATE: Insere novo registro.
        Espera pelo menos: data["fkProduto"].
        """
        try:
            fk_produto = int(data.get("fkProduto") or 0)
            if fk_produto <= 0:
                return False, "fkProduto é obrigatório para inserir infoProduto."

            # evita duplicar (se fkProduto for único)
            if self.info_produto_get(fk_produto):
                return False, f"Já existe infoProduto para fkProduto={fk_produto}. Use UPDATE."

            cols_insert = [
                "estoqueMinimo", "estoqueMaximo", "estoqueLocalizacao", "unidade",
                "pesoLiquido", "pesoBruto", "volumes", "itensPorCaixa", "gtin",
                "tipoProducao", "marca", "precoCompra", "largura", "altura",
                "profundidade", "unidadeMedida", "fkFornecedor", "fkCategoria",
                "prazoEntrega", "fkProduto"
            ]

            placeholders = ", ".join(["%s"] * len(cols_insert))
            col_sql = ", ".join([f'"{c}"' for c in cols_insert])

            values = [data.get(c) for c in cols_insert]
            values[-1] = fk_produto  # garante int

            sql = f'''
                INSERT INTO "Ekenox"."infoProduto" ({col_sql})
                VALUES ({placeholders});
            '''
            self._q(sql, tuple(values))
            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir infoProduto: {type(e).__name__}: {e}"

    def info_produto_update(self, fk_produto: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        UPDATE: Atualiza por fkProduto.
        Atualiza somente campos presentes em 'data' (exceto fkProduto).
        """
        try:
            fk_produto = int(fk_produto)
            if fk_produto <= 0:
                return False, "fkProduto inválido."

            allowed = set(self.INFO_PRODUTO_COLS) - {"fkProduto"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar."

            set_sql = ", ".join([f'i."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [fk_produto]

            sql = f"""
                UPDATE "Ekenox"."infoProduto" i
                   SET {set_sql}
                 WHERE i."fkProduto"::bigint = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para fkProduto={fk_produto}."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar infoProduto: {type(e).__name__}: {e}"

    def info_produto_upsert(self, data: Dict[str, Any]) -> Tuple[bool, str]:
        """UPSERT: se existe -> UPDATE, se não -> CREATE."""
        fk_produto = int(data.get("fkProduto") or 0)
        if fk_produto <= 0:
            return False, "fkProduto é obrigatório no UPSERT."

        if self.info_produto_get(fk_produto):
            data2 = dict(data)
            data2.pop("fkProduto", None)
            return self.info_produto_update(fk_produto, data2)

        return self.info_produto_create(data)

    def info_produto_delete(self, fk_produto: int) -> Tuple[bool, str]:
        """DELETE: Exclui por fkProduto."""
        try:
            fk_produto = int(fk_produto)
            sql = """
                DELETE FROM "Ekenox"."infoProduto"
                WHERE "fkProduto"::bigint = %s;
            """
            self._q(sql, (fk_produto,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (fkProduto={fk_produto})."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir infoProduto: {type(e).__name__}: {e}"
