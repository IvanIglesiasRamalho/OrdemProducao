# launcher.py
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
import importlib


# ------------------------------------------------------------
# Ajuste aqui: nome do módulo onde está sua OP
# Ex.: se seu arquivo é Ordem_Producao.py, então MODULE_OP = "Ordem_Producao"
# ------------------------------------------------------------
MODULE_OP = "Ordem_Producao"

# ------------------------------------------------------------
# Ajuste aqui: módulo do sistema agregador (o loader que você criou)
# Ex.: se você salvou como sistema_loader.py, então MODULE_SISTEMA = "sistema_loader"
# ------------------------------------------------------------
MODULE_SISTEMA = "sistema_loader"


def _safe_import(module_name: str):
    try:
        return importlib.import_module(module_name)
    except Exception as e:
        messagebox.showerror(
            "Erro", f"Não foi possível importar {module_name}:\n{type(e).__name__}: {e}")
        return None


class LauncherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ekenox - Menu Principal")
        self.geometry("1200x700")
        self.minsize(1100, 650)

        # ---- carrega cfg + sistema ----
        self.cfg = self._load_cfg()
        self.sistema = self._create_and_connect_sistema(self.cfg)

        # ---- menu ----
        self._build_menu()

        # ---- área principal (tela inicial: Ordem de Produção) ----
        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self._open_ordem_producao_embedded()

    # -------------------------
    # Config / Sistema
    # -------------------------
    def _load_cfg(self):
        mod_op = _safe_import(MODULE_OP)
        if not mod_op:
            raise SystemExit

        # tenta usar load_config do seu projeto (se existir)
        load_config = getattr(mod_op, "load_config", None)
        if callable(load_config):
            return load_config()

        # se não existir, devolve None e o sistema_loader deve lidar
        return None

    def _create_and_connect_sistema(self, cfg):
        mod_sis = _safe_import(MODULE_SISTEMA)
        if not mod_sis:
            raise SystemExit

        SistemaOrdemProducao = getattr(mod_sis, "SistemaOrdemProducao", None)
        if SistemaOrdemProducao is None:
            messagebox.showerror(
                "Erro", f"{MODULE_SISTEMA}.py não tem a classe SistemaOrdemProducao.")
            raise SystemExit

        sistema = SistemaOrdemProducao(cfg)
        ok, err = sistema.conectar()
        if not ok:
            messagebox.showerror("Banco", f"Falha ao conectar:\n{err}")
        return sistema

    # -------------------------
    # Menu
    # -------------------------
    def _build_menu(self):
        menubar = tk.Menu(self)

        m_programas = tk.Menu(menubar, tearoff=0)
        m_programas.add_command(label="Ordem de Produção",
                                command=self._show_ordem_producao)
        m_programas.add_separator()
        m_programas.add_command(label="Produtos", command=lambda: self._open_program(
            "tela_produtos", "ProdutosWindow"))
        m_programas.add_command(label="InfoProduto", command=lambda: self._open_program(
            "tela_info_produto", "InfoProdutoWindow"))
        m_programas.add_command(label="Arranjo", command=lambda: self._open_program(
            "tela_arranjo", "ArranjoWindow"))
        m_programas.add_command(label="Categoria", command=lambda: self._open_program(
            "tela_categoria", "CategoriaWindow"))
        m_programas.add_command(label="Fornecedor", command=lambda: self._open_program(
            "tela_fornecedor", "FornecedorWindow"))
        m_programas.add_command(label="Depósito", command=lambda: self._open_program(
            "tela_deposito", "DepositoWindow"))
        m_programas.add_command(label="Estoque", command=lambda: self._open_program(
            "tela_estoque", "EstoqueWindow"))
        m_programas.add_command(label="Estrutura", command=lambda: self._open_program(
            "tela_estrutura", "EstruturaWindow"))
        m_programas.add_command(label="Situação", command=lambda: self._open_program(
            "tela_situacao", "SituacaoWindow"))
        menubar.add_cascade(label="Programas", menu=m_programas)

        m_sistema = tk.Menu(menubar, tearoff=0)
        m_sistema.add_command(label="Reconectar Banco",
                              command=self._reconnect)
        m_sistema.add_separator()
        m_sistema.add_command(label="Sair", command=self._on_exit)
        menubar.add_cascade(label="Sistema", menu=m_sistema)

        self.config(menu=menubar)

    # -------------------------
    # Tela principal: OP
    # -------------------------
    def _clear_container(self):
        for w in self.container.winfo_children():
            w.destroy()

    def _open_ordem_producao_embedded(self):
        """
        Tenta embutir a tela de OP como Frame.
        Se não existir OrdemProducaoFrame, abre em janela separada.
        """
        self._clear_container()

        mod_op = _safe_import(MODULE_OP)
        if not mod_op:
            return

        # 1) Preferencial: Frame embutível
        OrdemProducaoFrame = getattr(mod_op, "OrdemProducaoFrame", None)
        if OrdemProducaoFrame is not None:
            frame = OrdemProducaoFrame(self.container, self.cfg, self.sistema)
            frame.pack(fill="both", expand=True)
            return

        # 2) Fallback: abre sua janela OP atual como Toplevel
        OrdemProducaoApp = getattr(mod_op, "OrdemProducaoApp", None)
        if OrdemProducaoApp is not None:
            info = ttk.Label(
                self.container,
                text=(
                    "Sua OP ainda está como tk.Tk (janela própria).\n"
                    "Abrindo em janela separada. Se você criar OrdemProducaoFrame, eu embuto aqui."
                ),
                justify="left"
            )
            info.pack(padx=20, pady=20, anchor="w")

            ttk.Button(self.container, text="Abrir Ordem de Produção", command=self._open_op_toplevel).pack(
                padx=20, pady=10, anchor="w"
            )
            return

        messagebox.showerror(
            "Erro", f"No módulo {MODULE_OP} não encontrei OrdemProducaoFrame nem OrdemProducaoApp.")

    def _open_op_toplevel(self):
        mod_op = _safe_import(MODULE_OP)
        if not mod_op:
            return

        OrdemProducaoApp = getattr(mod_op, "OrdemProducaoApp", None)
        if OrdemProducaoApp is None:
            messagebox.showerror(
                "Erro", f"{MODULE_OP} não tem OrdemProducaoApp.")
            return

        # Cria a janela OP como Toplevel para não conflitar com o Tk principal do Launcher
        win = tk.Toplevel(self)
        win.title("Ordem de Produção - Ekenox")

        # Se sua OrdemProducaoApp for tk.Tk, ideal é criar uma versão Frame.
        # Aqui: tentamos instanciar uma classe "OrdemProducaoFrame" se existir.
        OrdemProducaoFrame = getattr(mod_op, "OrdemProducaoFrame", None)
        if OrdemProducaoFrame is not None:
            frame = OrdemProducaoFrame(win, self.cfg, self.sistema)
            frame.pack(fill="both", expand=True)
        else:
            messagebox.showwarning(
                "Ajuste recomendado",
                "Para embutir no launcher e abrir como Toplevel corretamente, crie OrdemProducaoFrame.\n"
                "Hoje só consigo abrir embutido se existir OrdemProducaoFrame."
            )
            win.destroy()

    def _show_ordem_producao(self):
        self._open_ordem_producao_embedded()

    # -------------------------
    # Abertura dos “programas”
    # -------------------------
    def _open_program(self, module_name: str, class_name: str):
        """
        Espera que cada tela exista num módulo separado:
          - tela_produtos.py -> class ProdutosWindow(tk.Toplevel)
          - tela_arranjo.py  -> class ArranjoWindow(tk.Toplevel)
        e que o construtor aceite: (master, cfg, sistema)
        """
        mod = _safe_import(module_name)
        if not mod:
            return

        cls = getattr(mod, class_name, None)
        if cls is None:
            messagebox.showerror(
                "Erro", f"No módulo {module_name}.py não existe a classe {class_name}.")
            return

        try:
            cls(self, self.cfg, self.sistema)  # abre como Toplevel
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Falha ao abrir {class_name}:\n{type(e).__name__}: {e}")

    # -------------------------
    # Sistema
    # -------------------------
    def _reconnect(self):
        try:
            self.sistema.desconectar()
        except Exception:
            pass

        ok, err = self.sistema.conectar()
        if ok:
            messagebox.showinfo("Banco", "Reconectado com sucesso.")
        else:
            messagebox.showerror("Banco", f"Falha ao reconectar:\n{err}")

    def _on_exit(self):
        try:
            self.sistema.desconectar()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = LauncherApp()
    app.mainloop()
