# produtos_crud.py
from __future__ import annotations
from typing import Optional, Dict, Any, List, Tuple


class ProdutosCRUDMixin:
    """
    CRUD para "Ekenox"."produtos"

    Colunas (conforme print):
      - produtoId (PK)
      - nomeProduto
      - sku
      - preco
      - custo
      - tipo
      - formato
      - descricaoCurta
      - idProdutoPai
      - descImetro

    Requer:
      - self._q(sql, params)
      - self.conn
      - self.cursor

    OBS: Auxiliares mais comuns no seu cenário:
      - "Ekenox"."infoProduto" (fkCategoria, fkFornecedor, estoqueMinimo/Maximo etc.)
      - "Ekenox"."categoria"
      - "Ekenox"."fornecedor"
      - "Ekenox"."arranjo"
      - "Ekenox"."deposito"
    """

    PRODUTOS_COLS = (
        "produtoId",
        "nomeProduto",
        "sku",
        "preco",
        "custo",
        "tipo",
        "formato",
        "descricaoCurta",
        "idProdutoPai",
        "descImetro",
    )

    # ----------------------------
    # PRODUTOS - READ
    # ----------------------------

    def produto_get(self, produto_id: int) -> Optional[Dict[str, Any]]:
        try:
            pid = int(produto_id)
            sql = """
                SELECT
                    p."produtoId",
                    p."nomeProduto",
                    p."sku",
                    p."preco",
                    p."custo",
                    p."tipo",
                    p."formato",
                    p."descricaoCurta",
                    p."idProdutoPai",
                    p."descImetro"
                FROM "Ekenox"."produtos" p
                WHERE p."produtoId"::bigint = %s
                LIMIT 1;
            """
            self._q(sql, (pid,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return dict(zip(self.PRODUTOS_COLS, r))
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def produto_list(
        self,
        limit: int = 300,
        offset: int = 0,
        nome_like: str = "",
        sku_like: str = "",
        tipo: str = "",
    ) -> List[Dict[str, Any]]:
        """
        Lista produtos com filtros opcionais:
          - nome_like: ILIKE nomeProduto
          - sku_like: ILIKE sku
          - tipo: igualdade (se informado)
        """
        try:
            limit = int(limit)
            offset = int(offset)
            nome_like = (nome_like or "").strip()
            sku_like = (sku_like or "").strip()
            tipo = (tipo or "").strip()

            where_parts: List[str] = []
            params: List[Any] = []

            if nome_like:
                where_parts.append('p."nomeProduto" ILIKE %s')
                params.append(f"%{nome_like}%")

            if sku_like:
                where_parts.append('p."sku" ILIKE %s')
                params.append(f"%{sku_like}%")

            if tipo:
                where_parts.append('p."tipo" = %s')
                params.append(tipo)

            where_sql = ""
            if where_parts:
                where_sql = "WHERE " + " AND ".join(where_parts)

            sql = f"""
                SELECT
                    p."produtoId",
                    p."nomeProduto",
                    p."sku",
                    p."preco",
                    p."custo",
                    p."tipo",
                    p."formato",
                    p."descricaoCurta",
                    p."idProdutoPai",
                    p."descImetro"
                FROM "Ekenox"."produtos" p
                {where_sql}
                ORDER BY p."nomeProduto"
                LIMIT %s OFFSET %s;
            """
            params.extend([limit, offset])

            self._q(sql, tuple(params))
            rows = self.cursor.fetchall() or []
            return [dict(zip(self.PRODUTOS_COLS, row)) for row in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    # ----------------------------
    # PRODUTOS - CREATE/UPDATE/UPSERT/DELETE
    # ----------------------------

    def produto_create(
        self,
        produto_id: int,
        nome_produto: str,
        sku: str = "",
        preco: float | int | None = None,
        custo: float | int | None = None,
        tipo: str | None = None,
        formato: str | None = None,
        descricao_curta: str | None = None,
        id_produto_pai: int | None = None,
        desc_imetro: str | None = None,
    ) -> Tuple[bool, str]:
        """
        CREATE com parâmetros explícitos.
        """
        try:
            pid = int(produto_id)
            nome = (nome_produto or "").strip()
            if not nome:
                return False, "nome_produto é obrigatório."

            if self.produto_get(pid):
                return False, f"Já existe produtoId={pid}. Use UPDATE."

            sql = """
                INSERT INTO "Ekenox"."produtos"
                    ("produtoId","nomeProduto","sku","preco","custo","tipo","formato","descricaoCurta","idProdutoPai","descImetro")
                VALUES
                    (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);
            """
            self._q(sql, (
                pid,
                nome,
                (sku or "").strip(),
                preco,
                custo,
                tipo,
                formato,
                descricao_curta,
                int(id_produto_pai) if id_produto_pai is not None else None,
                desc_imetro,
            ))
            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao inserir produto: {type(e).__name__}: {e}"

    def produto_update(self, produto_id: int, data: Dict[str, Any]) -> Tuple[bool, str]:
        """
        UPDATE patch por produtoId.
        Campos aceitos: nomeProduto, sku, preco, custo, tipo, formato, descricaoCurta, idProdutoPai, descImetro
        """
        try:
            pid = int(produto_id)

            allowed = {
                "nomeProduto", "sku", "preco", "custo", "tipo", "formato",
                "descricaoCurta", "idProdutoPai", "descImetro"
            }
            campos = [k for k in data.keys() if k in allowed]
            if not campos:
                return False, "Nenhum campo válido para atualizar."

            if "nomeProduto" in campos:
                data["nomeProduto"] = (data.get("nomeProduto") or "").strip()
            if "sku" in campos:
                data["sku"] = (data.get("sku") or "").strip()

            set_sql = ", ".join([f'p."{c}" = %s' for c in campos])
            values = [data.get(c) for c in campos] + [pid]

            sql = f"""
                UPDATE "Ekenox"."produtos" p
                   SET {set_sql}
                 WHERE p."produtoId"::bigint = %s;
            """
            self._q(sql, tuple(values))
            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro encontrado para produtoId={pid}."

            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao atualizar produto: {type(e).__name__}: {e}"

    def produto_upsert(
        self,
        produto_id: int,
        nome_produto: str | None = None,
        sku: str | None = None,
        preco: float | int | None = None,
        custo: float | int | None = None,
        tipo: str | None = None,
        formato: str | None = None,
        descricao_curta: str | None = None,
        id_produto_pai: int | None = None,
        desc_imetro: str | None = None,
    ) -> Tuple[bool, str]:
        """
        UPSERT com parâmetros explícitos:
          - se existe -> update só campos != None
          - se não existe -> create (nome_produto obrigatório)
        """
        pid = int(produto_id)
        existe = self.produto_get(pid)

        if existe:
            patch: Dict[str, Any] = {}
            if nome_produto is not None:
                patch["nomeProduto"] = (nome_produto or "").strip()
            if sku is not None:
                patch["sku"] = (sku or "").strip()
            if preco is not None:
                patch["preco"] = preco
            if custo is not None:
                patch["custo"] = custo
            if tipo is not None:
                patch["tipo"] = tipo
            if formato is not None:
                patch["formato"] = formato
            if descricao_curta is not None:
                patch["descricaoCurta"] = descricao_curta
            if id_produto_pai is not None:
                patch["idProdutoPai"] = int(id_produto_pai)
            if desc_imetro is not None:
                patch["descImetro"] = desc_imetro

            if not patch:
                return True, ""

            return self.produto_update(pid, patch)

        if nome_produto is None or not (nome_produto or "").strip():
            return False, "Para criar produto novo, informe nome_produto."

        return self.produto_create(
            produto_id=pid,
            nome_produto=nome_produto,
            sku=sku or "",
            preco=preco,
            custo=custo,
            tipo=tipo,
            formato=formato,
            descricao_curta=descricao_curta,
            id_produto_pai=id_produto_pai,
            desc_imetro=desc_imetro,
        )

    def produto_delete(self, produto_id: int) -> Tuple[bool, str]:
        try:
            pid = int(produto_id)
            sql = """
                DELETE FROM "Ekenox"."produtos"
                WHERE "produtoId"::bigint = %s;
            """
            self._q(sql, (pid,))
            if (self.cursor.rowcount or 0) == 0:
                self.conn.rollback()
                return False, f"Nenhum registro para excluir (produtoId={pid})."
            self.conn.commit()
            return True, ""
        except Exception as e:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return False, f"Erro ao excluir produto: {type(e).__name__}: {e}"

    # ----------------------------
    # AUXILIARES (para carregar combos/listas)
    # ----------------------------

    def aux_categoria_list(self, limit: int = 500, offset: int = 0, nome_like: str = "") -> List[Dict[str, Any]]:
        """Carrega categorias para combos."""
        try:
            nome_like = (nome_like or "").strip()
            if nome_like:
                sql = """
                    SELECT c."categoriaId", c."nomeCategoria", c."categoriaPai"
                    FROM "Ekenox"."categoria" c
                    WHERE c."nomeCategoria" ILIKE %s
                    ORDER BY c."nomeCategoria"
                    LIMIT %s OFFSET %s;
                """
                params = (f"%{nome_like}%", int(limit), int(offset))
            else:
                sql = """
                    SELECT c."categoriaId", c."nomeCategoria", c."categoriaPai"
                    FROM "Ekenox"."categoria" c
                    ORDER BY c."nomeCategoria"
                    LIMIT %s OFFSET %s;
                """
                params = (int(limit), int(offset))

            self._q(sql, params)
            rows = self.cursor.fetchall() or []
            return [{"categoriaId": r[0], "nomeCategoria": r[1], "categoriaPai": r[2]} for r in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    def aux_fornecedor_list(self, limit: int = 500, offset: int = 0, nome_like: str = "") -> List[Dict[str, Any]]:
        """Carrega fornecedores para combos."""
        try:
            nome_like = (nome_like or "").strip()
            if nome_like:
                sql = """
                    SELECT f."idFornecedor", f."nome", f."codigo"
                    FROM "Ekenox"."fornecedor" f
                    WHERE f."nome" ILIKE %s
                    ORDER BY f."nome"
                    LIMIT %s OFFSET %s;
                """
                params = (f"%{nome_like}%", int(limit), int(offset))
            else:
                sql = """
                    SELECT f."idFornecedor", f."nome", f."codigo"
                    FROM "Ekenox"."fornecedor" f
                    ORDER BY f."nome"
                    LIMIT %s OFFSET %s;
                """
                params = (int(limit), int(offset))

            self._q(sql, params)
            rows = self.cursor.fetchall() or []
            return [{"idFornecedor": r[0], "nome": r[1], "codigo": r[2]} for r in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    def aux_deposito_list(self, limit: int = 500, offset: int = 0, descricao_like: str = "") -> List[Dict[str, Any]]:
        """Carrega depósitos para combos."""
        try:
            descricao_like = (descricao_like or "").strip()
            if descricao_like:
                sql = """
                    SELECT d."id", d."descricao", d."situacao", d."padrao", d."desconsiderarsaldo"
                    FROM "Ekenox"."deposito" d
                    WHERE d."descricao" ILIKE %s
                    ORDER BY d."descricao"
                    LIMIT %s OFFSET %s;
                """
                params = (f"%{descricao_like}%", int(limit), int(offset))
            else:
                sql = """
                    SELECT d."id", d."descricao", d."situacao", d."padrao", d."desconsiderarsaldo"
                    FROM "Ekenox"."deposito" d
                    ORDER BY d."descricao"
                    LIMIT %s OFFSET %s;
                """
                params = (int(limit), int(offset))

            self._q(sql, params)
            rows = self.cursor.fetchall() or []
            return [{"id": r[0], "descricao": r[1], "situacao": r[2], "padrao": r[3], "desconsiderarsaldo": r[4]} for r in rows]
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return []

    def aux_arranjo_get(self, sku: str) -> Optional[Dict[str, Any]]:
        """Busca arranjo por SKU (útil no fluxo do arranjo)."""
        try:
            sku = (sku or "").strip()
            if not sku:
                return None
            sql = """
                SELECT a."sku", a."nomeproduto", a."quantidade", a."chapa", a."material"
                FROM "Ekenox"."arranjo" a
                WHERE a."sku" = %s
                LIMIT 1;
            """
            self._q(sql, (sku,))
            r = self.cursor.fetchone()
            if not r:
                return None
            return {"sku": r[0], "nomeproduto": r[1], "quantidade": r[2], "chapa": r[3], "material": r[4]}
        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None

    def produto_get_com_auxiliares(self, produto_id: int) -> Optional[Dict[str, Any]]:
        """
        Retorna o produto e (se existirem) informações auxiliares:
          - infoProduto (estoqueMinimo, estoqueMaximo, fkCategoria, fkFornecedor, etc.)
          - categoria (nomeCategoria)
          - fornecedor (nome)
          - arranjo (quantidade por sku)
        """
        try:
            pid = int(produto_id)

            sql = """
                SELECT
                    p."produtoId",
                    p."nomeProduto",
                    p."sku",
                    p."preco",
                    p."custo",
                    p."tipo",
                    p."formato",
                    p."descricaoCurta",
                    p."idProdutoPai",
                    p."descImetro",

                    ip."estoqueMinimo",
                    ip."estoqueMaximo",
                    ip."fkCategoria",
                    ip."fkFornecedor",
                    ip."prazoEntrega",

                    c."nomeCategoria",
                    f."nome" as fornecedor_nome,

                    a."quantidade" as arranjo_quantidade,
                    a."chapa" as arranjo_chapa,
                    a."material" as arranjo_material
                FROM "Ekenox"."produtos" p
                LEFT JOIN "Ekenox"."infoProduto" ip
                       ON ip."fkProduto"::bigint = p."produtoId"::bigint
                LEFT JOIN "Ekenox"."categoria" c
                       ON c."categoriaId" = ip."fkCategoria"
                LEFT JOIN "Ekenox"."fornecedor" f
                       ON f."idFornecedor" = ip."fkFornecedor"
                LEFT JOIN "Ekenox"."arranjo" a
                       ON a."sku" = p."sku"
                WHERE p."produtoId"::bigint = %s
                LIMIT 1;
            """
            self._q(sql, (pid,))
            r = self.cursor.fetchone()
            if not r:
                return None

            # produto base
            produto = {
                "produtoId": r[0],
                "nomeProduto": r[1],
                "sku": r[2],
                "preco": r[3],
                "custo": r[4],
                "tipo": r[5],
                "formato": r[6],
                "descricaoCurta": r[7],
                "idProdutoPai": r[8],
                "descImetro": r[9],
            }

            # auxiliares
            info_produto = {
                "estoqueMinimo": r[10],
                "estoqueMaximo": r[11],
                "fkCategoria": r[12],
                "fkFornecedor": r[13],
                "prazoEntrega": r[14],
            }

            produto["infoProduto"] = info_produto
            produto["categoria_nome"] = r[15]
            produto["fornecedor_nome"] = r[16]
            produto["arranjo"] = {
                "quantidade": r[17],
                "chapa": r[18],
                "material": r[19],
            }

            return produto

        except Exception:
            if getattr(self, "conn", None):
                self.conn.rollback()
            return None
