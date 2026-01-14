# Select Antiga
            SELECT
                e.{parent}      AS produto_pai,
                e.{child}       AS componente_id,
                e.{qty}         AS qtd_por,
                p."nomeProduto" AS componente_nome,
                p."sku"         AS componente_sku
            FROM "Ekenox"."estrutura" e
            JOIN "Ekenox"."produtos" p
              ON p."produtoId" = e.{child}
            WHERE e.{parent}: : bigint = %s
            ORDER BY p."nomeProduto"
        """