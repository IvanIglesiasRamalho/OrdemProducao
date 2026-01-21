# sistema_cruds.py
from __future__ import annotations

from typing import Optional, Any, Tuple

import psycopg2
import psycopg2.extras

# ---- Mixins CRUD (seus arquivos) ----
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
    Classe agregadora: reúne todos os CRUDs (mixins).

    Pré-requisitos atendidos para os mixins:
      - self.conn
      - self.cursor
      - self._q(sql, params)
    """

    def __init__(self, cfg: Any):
        self.cfg = cfg
        self.conn: Optional[psycopg2.extensions.connection] = None
        self.cursor: Optional[psycopg2.extensions.cursor] = None
        self._ultimo_erro_bd: Optional[str] = None

    # ----------------------------
    # Conexão / Execução
    # ----------------------------

    def conectar(self) -> Tuple[bool, str]:
        """
        Ajuste aqui para como seu cfg guarda as configs.
        Exemplo esperado:
          cfg.DB_HOST, cfg.DB_PORT, cfg.DB_NAME, cfg.DB_USER, cfg.DB_PASSWORD
        """
        try:
            if self.conn:
                return True, ""

            self.conn = psycopg2.connect(
                host=getattr(self.cfg, "DB_HOST", "localhost"),
                port=int(getattr(self.cfg, "DB_PORT", 5432)),
                dbname=getattr(self.cfg, "DB_NAME", "postgres"),
                user=getattr(self.cfg, "DB_USER", "postgres"),
                password=getattr(self.cfg, "DB_PASSWORD", ""),
            )
            # cursor padrão (pode trocar por DictCursor se quiser)
            self.cursor = self.conn.cursor()
            return True, ""
        except Exception as e:
            self._ultimo_erro_bd = f"{type(e).__name__}: {e}"
            return False, self._ultimo_erro_bd

    def desconectar(self) -> None:
        try:
            if self.cursor:
                try:
                    self.cursor.close()
                except Exception:
                    pass
            self.cursor = None

            if self.conn:
                try:
                    self.conn.close()
                except Exception:
                    pass
            self.conn = None
        except Exception:
            self.conn = None
            self.cursor = None

    def _q(self, sql: str, params: tuple = ()) -> None:
        """
        Executor único para os mixins.
        Mantém o último erro e garante que cursor exista.
        """
        if not self.conn or not self.cursor:
            raise RuntimeError(
                "Sem conexão com o banco. Chame conectar() antes.")

        try:
            self.cursor.execute(sql, params)
            self._ultimo_erro_bd = None
        except Exception as e:
            self._ultimo_erro_bd = f"{type(e).__name__}: {e}"
            raise

    # ----------------------------
    # Dispatcher opcional (um "hub" para chamar por nome)
    # ----------------------------

    def call(self, nome: str, *args, **kwargs):
        """
        Permite chamar qualquer método do sistema pelo nome:
          sistema.call("arranjo_get", "VIX9444N")
          sistema.call("produto_get_com_auxiliares", 123)
        """
        fn = getattr(self, nome, None)
        if not callable(fn):
            raise AttributeError(
                f"Método '{nome}' não existe em SistemaOrdemProducao.")
        return fn(*args, **kwargs)
