# fornecedor_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class FornecedorCRUDMixin:
    """
    CRUD para "Ekenox"."fornecedor"
    Colunas:
      - idFornecedor (PK)
      - codigo
      - nome
      - situacao
      - numeroDocumentacao
      - telefone
      - celular
    """

    FORNECEDOR_COLS = (
        "idFornecedor",
        "codigo",
        "nome",
        "situacao",
        "numeroDocumentacao",
        "telefone",
        "celular",
    )

    def fornecedor_get(self, id_fornecedor: int) -> Optional[Dict[str, Any]]:
        """READ: busca 1 fornecedor por idFornecedor."""
        try:
            fid = int(id_fornecedor)
            sql = """
                SELECT
                    f."idFornecedor",
                    f."codigo",
                    f."nome",
                    f."situacao",
                    f."numeroDocumentacao",
                    f."telefone",
                    f."celular"
                FROM "Ekenox"."fornecedor" f
                WHERE f."idFornecedor" = %s
                LIMIT 1;
            """
            self._q(sql, (fid,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.FORNECEDOR_COLS, r))
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def fornecedor_list(
        self,
        limit: int = 300,
        offset: int = 0,
        nome_like: str = "",
        codigo_like: str = "",
        situacao: int | None = None,
    ) -> List[Dict[str, Any]]:
        """
        Lista fornecedores com filtros opcionais:
          - nome_like: filtra por nome (ILIKE)
          - codigo_like: filtra por codigo (ILIKE)
          - situacao: filtra igualdade (se não for None)
        """
        try:
            limit = int(limit)
            offset = int(offset)
            nome_like = (nome_like or "").strip()
            codigo_like = (codigo_like or "").strip()

            where_parts: List[str] = []
            params: List[Any] = []

            if nome_like:
                where_parts.append('f."nome" ILIKE %s')
                params.append(f"%{nome_like}%")

            if codigo_like:
                where_parts.append('f."codigo" ILIKE %s')
                params.append(f"%{codigo_like}%")

            if situacao is not None:
                where_parts.append('f."situacao" = %s')
                params.append(int(situacao))

            where_sql = ""
            if where_parts:
                where_sql = "WHERE " + " AND ".join(where_parts)

            sql = f"""
                SELECT
                    f."idFornecedor",
                    f."codigo",
                    f."nome",
                    f."situacao",
                    f."numeroDocumentacao",
                    f."telefone",
                    f."celular"
                FROM "Ekenox"."fornecedor" f
                {where_sql}
                ORDER BY f."nome"
                LIMIT %s OFFSET %s;
            """
            params.extend([limit, offset])

            self._q(sql, tuple(params))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.FORNECEDOR_COLS, row)) for row in rows]

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ✅ CREATE com parâmetros explícitos
    def fornecedor_create(
        self,
        id_fornecedor: int,
        nome: str,
        codigo: str | None = None,
        situacao: int = 1,
        numero_documentacao: str | None = None,
        telefone: str | None = None,
        celular: str | None = None,
    ) -> Tuple[bool, str]:
        """
        Insere fornecedor.
        - id_fornecedor e nome obrigatórios.
        """
        try:
            fid = int(id_fornecedor)
            nome = (nome or "").strip()
            if not nome:
                return False, "nome é obrigatório."

            if self.fornecedor_get(fid):
                return False, f"Já existe fornecedor idFornecedor={fid}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."fornecedor"
                    ("idFornecedor", "codigo", "nome", "situacao", "numeroDocumentacao", "telefone", "celular")
                VALUES
                    (%s, %s, %s, %s, %s, %s, %s);
            """
            self._q(sql, (fid, codigo, nome, int(situacao),
                    numero_documentacao, telefone, celular))
            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir fornecedor: {type(e).__name__}: {e}"

    def fornecedor_update(self, id_fornecedor: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Atualiza campos (patch) do fornecedor.
        Campos aceitos em data: codigo, nome, situacao, numeroDocumentacao, telefone, celular
        """
        try:
            fid = int(id_fornecedor)

            allowed = {"codigo", "nome", "situacao",
                       "numeroDocumentacao", "telefone", "celular"}
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar."

            if "nome" in campos:
                data["nome"] = (data.get("nome") or "").strip()

            set_sql = ", ".join([f'f."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [fid]

            sql = f"""
                UPDATE "Ekenox"."fornecedor" f
                   SET {set_sql}
                 WHERE f."idFornecedor" = %s;
            """
            self._q(sql, tuple(values))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para idFornecedor={fid}."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar fornecedor: {type(e).__name__}: {e}"

    # ✅ UPSERT com parâmetros explícitos
    def fornecedor_upsert(
        self,
        id_fornecedor: int,
        nome: str | None = None,
        codigo: str | None = None,
        situacao: int | None = None,
        numero_documentacao: str | None = None,
        telefone: str | None = None,
        celular: str | None = None,
    ) -> Tuple[bool, str]:
        """
        Se existe idFornecedor -> update (apenas campos não-None)
        Senão -> create (nome obrigatório para criar)
        """
        fid = int(id_fornecedor)
        existe = self.fornecedor_get(fid)

        if existe:
            patch: Dict[str, Any] = {}
            if codigo is not None:
                patch["codigo"] = codigo
            if nome is not None:
                patch["nome"] = (nome or "").strip()
            if situacao is not None:
                patch["situacao"] = int(situacao)
            if numero_documentacao is not None:
                patch["numeroDocumentacao"] = numero_documentacao
            if telefone is not None:
                patch["telefone"] = telefone
            if celular is not None:
                patch["celular"] = celular

            if not patch:
                return True, ""

            return self.fornecedor_update(fid, patch)

        if nome is None or not (nome or "").strip():
            return False, "Para criar (fornecedor novo), informe nome."

        return self.fornecedor_create(
            id_fornecedor=fid,
            nome=nome,
            codigo=codigo,
            situacao=int(situacao) if situacao is not None else 1,
            numero_documentacao=numero_documentacao,
            telefone=telefone,
            celular=celular,
        )

    def fornecedor_delete(self, id_fornecedor: int) -> Tuple[bool, str]:
        """DELETE: exclui por idFornecedor."""
        try:
            fid = int(id_fornecedor)
            sql = """
                DELETE FROM "Ekenox"."fornecedor"
                WHERE "idFornecedor" = %s;
            """
            self._q(sql, (fid,))

            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (idFornecedor={fid})."

            self.conn.commit()
            return True, ""

        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir fornecedor: {type(e).__name__}: {e}"
