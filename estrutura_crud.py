# estrutura_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class EstruturaCRUDMixin:
    """
    CRUD para "Ekenox"."estrutura"
    Colunas:
      - fkproduto
      - componente
      - quantidade
      - dados

    Chave lógica assumida: (fkproduto, componente)
    """

    ESTRUTURA_COLS = ("fkproduto", "componente", "quantidade", "dados")

    def estrutura_get(self, fkproduto: int, componente: int) -> Optional[Dict[str, Any]]:
        """READ: busca 1 item da estrutura por (fkproduto, componente)."""
        try:
            fk = int(fkproduto)
            comp = int(componente)

            sql = """
                SELECT
                    e."fkproduto",
                    e."componente",
                    e."quantidade",
                    e."dados"
                FROM "Ekenox"."estrutura" e
                WHERE e."fkproduto"::bigint = %s
                  AND e."componente"::bigint = %s
                LIMIT 1;
            """
            self._q(sql, (fk, comp))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.ESTRUTURA_COLS, r))

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def estrutura_list_do_produto(self, fkproduto: int) -> List[Dict[str, Any]]:
        """READ: lista toda a estrutura (BOM) de um produto."""
        try:
            fk = int(fkproduto)
            sql = """
                SELECT
                    e."fkproduto",
                    e."componente",
                    e."quantidade",
                    e."dados"
                FROM "Ekenox"."estrutura" e
                WHERE e."fkproduto"::bigint = %s
                ORDER BY e."componente"::bigint;
            """
            self._q(sql, (fk,))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.ESTRUTURA_COLS, row)) for row in rows]

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    def estrutura_list(
        self,
        limit: int = 500,
        offset: int = 0,
        fkproduto: int | None = None,
        componente: int | None = None,
    ) -> List[Dict[str, Any]]:
        """READ: lista paginada com filtros opcionais."""
        try:
            limit = int(limit)
            offset = int(offset)

            where_parts: List[str] = []
            params: List[Any] = []

            if fkproduto is not None:
                where_parts.append('e."fkproduto"::bigint = %s')
                params.append(int(fkproduto))

            if componente is not None:
                where_parts.append('e."componente"::bigint = %s')
                params.append(int(componente))

            where_sql = ""
            if where_parts:
                where_sql = "WHERE " + " AND ".join(where_parts)

            sql = f"""
                SELECT
                    e."fkproduto",
                    e."componente",
                    e."quantidade",
                    e."dados"
                FROM "Ekenox"."estrutura" e
                {where_sql}
                ORDER BY e."fkproduto"::bigint, e."componente"::bigint
                LIMIT %s OFFSET %s;
            """
            params.extend([limit, offset])

            self._q(sql, tuple(params))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.ESTRUTURA_COLS, row)) for row in rows]

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ CREATE com parâmetros explícitos
    def estrutura_create(
        self,
        fkproduto: int,
        componente: int,
        quantidade: float | int,
        dados: Any = None,
    ) -> Tuple[bool, str]:
        """CREATE: insere 1 item na estrutura."""
        try:
            fk = int(fkproduto)
            comp = int(componente)

            # evita duplicar
            if self.estrutura_get(fk, comp):
                return False, f"Já existe estrutura para fkproduto={fk} e componente={comp}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."estrutura"
                    ("fkproduto", "componente", "quantidade", "dados")
                VALUES
                    (%s, %s, %s, %s);
            """
            self._q(sql, (fk, comp, quantidade, dados))
            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir estrutura: {type(e).__name__}: {e}"

    def estrutura_update(
        self,
        fkproduto: int,
        componente: int,
        data: Dict[str, Any],
    ) -> Tuple[bool, str]:
        """
        UPDATE: patch por (fkproduto, componente).
        Campos aceitos: quantidade, dados
        """
        try:
            fk = int(fkproduto)
            comp = int(componente)

            allowed = {"quantidade", "dados"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar (quantidade/dados)."

            set_sql = ", ".join([f'e."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [fk, comp]

            sql = f"""
                UPDATE "Ekenox"."estrutura" e
                   SET {set_sql}
                 WHERE e."fkproduto"::bigint = %s
                   AND e."componente"::bigint = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado (fkproduto={fk}, componente={comp})."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar estrutura: {type(e).__name__}: {e}"

    def estrutura_upsert(
        self,
        fkproduto: int,
        componente: int,
        quantidade: float | int | None = None,
        dados: Any = None,
        atualizar_dados_quando_none: bool = False,
    ) -> Tuple[bool, str]:
        """
        UPSERT por (fkproduto, componente).

        - Se existir: atualiza apenas campos informados.
        - Se não existir: cria (quantidade é obrigatória para criar).

        Param extra:
          atualizar_dados_quando_none:
            - False (padrão): se dados=None, NÃO atualiza 'dados'
            - True: atualiza 'dados' para NULL
        """
        fk = int(fkproduto)
        comp = int(componente)

        existe = self.estrutura_get(fk, comp)
        if existe:
            patch: Dict[str, Any] = {}

            if quantidade is not None:
                patch["quantidade"] = quantidade

            if atualizar_dados_quando_none:
                patch["dados"] = dados
            else:
                if dados is not None:
                    patch["dados"] = dados

            if not patch:
                return True, ""

            return self.estrutura_update(fk, comp, patch)

        # criar exige quantidade
        if quantidade is None:
            return False, "Para criar item novo, informe quantidade."

        return self.estrutura_create(fk, comp, quantidade, dados)

    def estrutura_delete(self, fkproduto: int, componente: int) -> Tuple[bool, str]:
        """DELETE: exclui 1 item por (fkproduto, componente)."""
        try:
            fk = int(fkproduto)
            comp = int(componente)

            sql = """
                DELETE FROM "Ekenox"."estrutura"
                WHERE "fkproduto"::bigint = %s
                  AND "componente"::bigint = %s;
            """
            self._q(sql, (fk, comp))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (fkproduto={fk}, componente={comp})."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir estrutura: {type(e).__name__}: {e}"

    def estrutura_delete_all_do_produto(self, fkproduto: int) -> Tuple[bool, str]:
        """DELETE: exclui toda a estrutura de um produto (útil para reimportar BOM)."""
        try:
            fk = int(fkproduto)
            sql = """
                DELETE FROM "Ekenox"."estrutura"
                WHERE "fkproduto"::bigint = %s;
            """
            self._q(sql, (fk,))
            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir estrutura do produto: {type(e).__name__}: {e}"
