from __future__ import annotations

from typing import Optional, Tuple, Any
import psycopg2
from psycopg2.extensions import connection as PGConn
from psycopg2.extensions import cursor as PGCursor


# ---------------------------------------------------------------------
# IMPORTA OS MIXINS CRUD (ajuste os nomes conforme seus arquivos)
# ---------------------------------------------------------------------
from info_produto_crud import InfoProdutoCRUDMixin
from arranjo_crud import ArranjoCRUDMixin
from categoria_crud import CategoriaCRUDMixin
from deposito_crud import DepositoCRUDMixin
from estoque_crud import EstoqueCRUDMixin
from estrutura_crud import EstruturaCRUDMixin
from fornecedor_crud import FornecedorCRUDMixin
from produtos_crud import ProdutosCRUDMixin
from situacao_crud import SituacaoCRUDMixin


class SistemaOrdemProducao(
    InfoProdutoCRUDMixin,
    ArranjoCRUDMixin,
    CategoriaCRUDMixin,
    DepositoCRUDMixin,
    EstoqueCRUDMixin,
    EstruturaCRUDMixin,
    FornecedorCRUDMixin,
    ProdutosCRUDMixin,
    SituacaoCRUDMixin,
):
    """
    Loader principal do sistema:
    - gerencia conexão (conn/cursor)
    - fornece _q() para os mixins
    - agrega todos os CRUDs em um único objeto: self.sistema
    """

    def __init__(self, cfg):
        self.cfg = cfg
        self.conn: Optional[PGConn] = None
        self.cursor: Optional[PGCursor] = None
        self.ultimo_erro: Optional[str] = None

    # -----------------------------
    # Conexão
    # -----------------------------
    def conectar(self) -> Tuple[bool, str]:
        """
        Retorna (ok, err). Não levanta exception para não derrubar a UI.
        """
        self.ultimo_erro = None
        try:
            self.conn = psycopg2.connect(
                host=self.cfg.db_host,
                database=self.cfg.db_database,
                user=self.cfg.db_user,
                password=self.cfg.db_password,
                port=int(self.cfg.db_port),
                connect_timeout=5,
            )
            self.cursor = self.conn.cursor()
            return True, ""
        except Exception as e:
            self.ultimo_erro = f"{type(e).__name__}: {e}"
            return False, self.ultimo_erro

    def desconectar(self) -> None:
        try:
            if self.cursor:
                self.cursor.close()
        except Exception:
            pass
        try:
            if self.conn:
                self.conn.close()
        except Exception:
            pass
        self.cursor = None
        self.conn = None

    # -----------------------------
    # Executor padrão para os Mixins
    # -----------------------------
    def _q(self, sql: str, params: tuple = ()) -> None:
        """
        Execute helper (usado pelos CRUDMixins).
        """
        if not self.cursor:
            raise RuntimeError("Sem cursor (não conectado).")
        self.cursor.execute(sql, params)

    def commit(self) -> None:
        if self.conn:
            self.conn.commit()

    def rollback(self) -> None:
        if self.conn:
            self.conn.rollback()
