import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
import psycopg2

# ===================== CONFIG BANCO =====================
DB_CONFIG = {
    "host": "10.0.0.154",
    "database": "postgresekenox",
    "user": "postgresekenox",
    "password": "Ekenox5426",
    "port": 55432,
}


# ===================== ÍCONE (mesma pasta / imagens) =====================
def obter_caminho_icone():
    """
    Tenta localizar o favicon.ico:
      1) Na mesma pasta do .py / .exe
      2) Na subpasta 'imagens'
    Funciona tanto em execução normal quanto em .exe do PyInstaller.
    """
    if getattr(sys, "frozen", False):
        # executável (PyInstaller)
        base_dir = os.path.dirname(sys.executable)
    else:
        # script .py
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # base_dir = r'C:\Users\User\Desktop\Pyton'  # <<< CORRETO
    base_dir = r'Z:\Planilhas_OP'

    candidatos = [
        os.path.join(base_dir, "favicon.ico"),
        os.path.join(base_dir, "imagens", "favicon.ico"),
    ]

    for caminho in candidatos:
        if os.path.isfile(caminho):
            return caminho
    return None


# ===================== FUNÇÕES AUXILIARES =====================
def ultimo_caractere(texto: str):
    """Retorna o último caractere de uma string, ou None se vazio."""
    if not texto:
        return None
    return texto[-1]


def listar_produtos(janela_pai, entry_produto, entry_modelo, entry_serie):
    """
    Abre uma janela para o usuário selecionar um produto.
    Ao selecionar (duplo clique ou Enter), preenche:
      - campo 'Produto' com DESCIMETRO (Imetro)
      - campo 'Modelo' com SKU (removendo 'n'/'N' no final, se houver)
      - campo 'Número de Série (prefixo/base)' com o número do pedido
        de venda mais recente.
    """
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        sql = '''
            SELECT 
                p."produtoId"                        AS Produto,
                p."nomeProduto"                      AS Nome,
                p."sku"                              AS SKU,
                COALESCE(NULLIF(TRIM(p."descImetro"), ''), p."nomeProduto") AS Imetro,
                (
                    SELECT RIGHT(ped."numero"::text, 4) AS Numero
                    FROM "Ekenox"."itens"   AS i
                    JOIN "Ekenox"."pedidos" AS ped
                      ON ped."idPedido" = i."fkPedido"
                    WHERE i."fkProduto" = p."produtoId"
                    ORDER BY ped."data" DESC
                    LIMIT 1
                ) AS numero_pedido
            FROM "Ekenox"."produtos" AS p
            LEFT JOIN "Ekenox"."infoProduto" AS ip
              ON p."produtoId" = ip."fkProduto"
            WHERE (p."descImetro"  IS NOT NULL AND TRIM(p."descImetro")  <> '')
            ORDER BY p."nomeProduto" ASC;
        '''
        cur.execute(sql)
        produtos = cur.fetchall()
        cur.close()
        conn.close()

        if not produtos:
            messagebox.showinfo(
                "Produtos",
                "Nenhum produto encontrado com os filtros configurados.",
                parent=janela_pai,
            )
            return

    except Exception as e:
        messagebox.showerror(
            "Erro ao buscar produtos",
            f"Ocorreu um erro ao consultar o banco de dados:\n{e}",
            parent=janela_pai,
        )
        return

    # ----- Janela de seleção -----
    janela = tk.Toplevel(janela_pai)
    janela.title("Selecionar Produto")
    janela.geometry("900x400")
    janela.transient(janela_pai)
    janela.grab_set()

    frame = tk.Frame(janela, padx=10, pady=10)
    frame.pack(fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    cols = ("ID", "Nome", "SKU", "DescInmetro", "Pedido")
    tree = ttk.Treeview(
        frame,
        columns=cols,
        show="headings",
        yscrollcommand=scrollbar.set
    )

    tree.heading("ID", text="ID")
    tree.heading("Nome", text="Nome do Produto")
    tree.heading("SKU", text="SKU")
    tree.heading("DescInmetro", text="Desc. Inmetro")
    tree.heading("Pedido", text="Nº Pedido (último)")

    tree.column("ID", width=60, anchor="center")
    tree.column("Nome", width=300, anchor="w")
    tree.column("SKU", width=120, anchor="w")
    tree.column("DescInmetro", width=260, anchor="w")
    tree.column("Pedido", width=110, anchor="center")

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=tree.yview)

    # Preenche a lista com os produtos buscados
    for prod_id, nome, sku, desc_inm, num_ped in produtos:
        tree.insert(
            "",
            tk.END,
            values=(
                prod_id,
                nome or "",
                sku or "",
                desc_inm or "",
                num_ped or "",
            ),
        )

    def selecionar_produto(event=None):
        selecao = tree.selection()
        if not selecao:
            return

        valores = tree.item(selecao[0])["values"]

        prod_id = valores[0]           # não usamos, mas deixei para referência
        nome_produto = valores[1]      # idem
        sku_val = (valores[2] or "").strip()
        desc_inmetro = (valores[3] or "").strip()
        numero_pedido = valores[4]

        # Produto = Desc. Inmetro
        texto_produto = desc_inmetro

        # Trata SKU: remove N/n final, se houver
        if sku_val and ultimo_caractere(sku_val).upper() == "N":
            sku_val = sku_val[:-1]

        # Campo Produto
        entry_produto.delete(0, tk.END)
        entry_produto.insert(0, texto_produto)

        # Campo Modelo (SKU)
        entry_modelo.delete(0, tk.END)
        entry_modelo.insert(0, sku_val)

        # Campo Número de Série (prefixo/base) = nº do pedido, se houver
        if numero_pedido not in (None, ""):
            entry_serie.delete(0, tk.END)
            entry_serie.insert(0, str(numero_pedido).strip())

        janela.destroy()

    tree.bind("<Double-Button-1>", selecionar_produto)
    tree.bind("<Return>", selecionar_produto)

    # Centraliza a janela de consulta
    janela.update_idletasks()
    x = (janela.winfo_screenwidth() // 2) - (janela.winfo_width() // 2)
    y = (janela.winfo_screenheight() // 2) - (janela.winfo_height() // 2)
    janela.geometry(f"+{x}+{y}")


def gerar_etiquetas(janela_pai,
                    entry_empresa,
                    entry_endereco,
                    entry_bairro,
                    entry_cidade,
                    entry_estado,
                    entry_cep,
                    entry_telefone,
                    entry_email,
                    entry_produto,
                    entry_modelo,
                    entry_classe,
                    combo_tensao,
                    entry_potencia,
                    entry_temperatura,
                    entry_frequencia,
                    entry_serie,
                    entry_quantidade):
    """Gera o PDF etiquetas.pdf com base nos dados preenchidos na tela."""
    try:
        # Dados da empresa
        empresa = {
            "company_name": entry_empresa.get().strip(),
            "company_address": entry_endereco.get().strip(),
            "company_district": entry_bairro.get().strip(),
            "company_city": entry_cidade.get().strip(),
            "company_state": entry_estado.get().strip(),
            "company_cep": entry_cep.get().strip(),
            "company_phone": entry_telefone.get().strip(),
            "company_email": entry_email.get().strip(),
        }

        # Dados do produto
        produto = {
            "product_title": entry_produto.get().strip(),
            "product_model": entry_modelo.get().strip(),
            "product_classe": entry_classe.get().strip(),
            "voltage": combo_tensao.get().strip(),
            "power": entry_potencia.get().strip(),
            "temperature": entry_temperatura.get().strip(),
            "frequency": entry_frequencia.get().strip(),
        }

        if not produto["product_title"]:
            messagebox.showerror(
                "Erro",
                "O campo 'Produto' deve ser preenchido (selecione na lista).",
                parent=janela_pai,
            )
            return

        quantidade_str = entry_quantidade.get().strip()
        if not quantidade_str:
            messagebox.showerror(
                "Erro",
                "Informe a quantidade de etiquetas.",
                parent=janela_pai,
            )
            return

        try:
            quantidade = int(quantidade_str)
            if quantidade <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Erro",
                "A quantidade de etiquetas deve ser um número inteiro maior que zero.",
                parent=janela_pai,
            )
            return

        serie_base = entry_serie.get().strip()
        if not serie_base:
            messagebox.showerror(
                "Erro",
                "O campo 'Número de Série (prefixo/base)' deve ser preenchido!",
                parent=janela_pai,
            )
            return

        largura, altura = 100 * mm, 75 * mm
        c = canvas.Canvas("etiquetas.pdf", pagesize=(largura, altura))

        # Colunas fixas
        x_titulo = 10
        x_valor = 70
        espaco = 10

        for i in range(quantidade):
            serial = f"{serie_base}-{i+1:03d}"

            # Borda
            c.setLineWidth(1)
            c.rect(5, 5, largura - 10, altura - 10)

            # Cabeçalho (nome da empresa)
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(largura / 2, altura - 15,
                                empresa["company_name"])

            y = altura - 30

            # Dados da empresa
            campos_empresa = [
                ("Endereço:", empresa["company_address"]),
                ("Bairro:", empresa["company_district"]),
                ("Cidade:",
                 f"{empresa['company_city']} - {empresa['company_state']}"),
                ("CEP:", empresa["company_cep"]),
                ("Telefone:", empresa["company_phone"]),
                ("Email SAC:", empresa["company_email"]),
            ]

            for titulo, valor in campos_empresa:
                c.setFont("Helvetica-Bold", 9)
                c.drawString(x_titulo, y, titulo)
                c.setFont("Helvetica", 9)
                c.drawString(x_valor, y, valor)
                y -= espaco

            # Linha separadora
            c.line(x_titulo, y, largura - 10, y)
            y -= espaco

            # Dados do produto
            produto_campos = [
                ("Produto:", produto["product_title"]),
                ("Modelo:", produto["product_model"]),
                ("Classe:", produto["product_classe"]),
                ("Tensão:", produto["voltage"]),
                ("Potência:", produto["power"]),
                ("Temp:", produto["temperature"]),
                ("Freq:", produto["frequency"]),
            ]

            for titulo, valor in produto_campos:
                c.setFont("Helvetica-Bold", 9)
                c.drawString(x_titulo, y, titulo)
                c.setFont("Helvetica", 9)
                c.drawString(x_valor, y, valor)
                y -= espaco

            # Linha separadora antes do número de série
            c.line(x_titulo, y, largura - 10, y)
            y -= espaco * 2

            # Número de série
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(largura / 2, y, f"Nº Série: {serial}")

            c.showPage()

        c.save()
        messagebox.showinfo(
            "Sucesso",
            "PDF 'etiquetas.pdf' gerado com sucesso!",
            parent=janela_pai,
        )

    except Exception as e:
        messagebox.showerror("Erro", str(e), parent=janela_pai)


# ===================== CLASSE PRINCIPAL =====================
class EtiquetaApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Etiquetas EKENOX")
        self.geometry("680x720")

        # Ícone
        icon_path = obter_caminho_icone()
        if icon_path:
            try:
                self.iconbitmap(default=icon_path)
                print(f"✓ Ícone carregado: {icon_path}")
            except Exception as e:
                print(f"⚠ Não foi possível carregar o ícone: {e}")
        else:
            print("⚠ favicon.ico não encontrado; executando sem ícone.")

        # Atalho ESC para fechar
        self.bind("<Escape>", lambda e: self.destroy())

        # ---------- Montagem da interface ----------
        self._montar_interface()

    def _montar_interface(self):
        # --- Campos Empresa ---
        frame_empresa = tk.LabelFrame(
            self, text="Informações da Empresa", padx=10, pady=10
        )
        frame_empresa.pack(fill="both", padx=10, pady=5)

        tk.Label(frame_empresa, text="Nome da Empresa:").grid(
            row=0, column=0, sticky="e"
        )
        self.entry_empresa = tk.Entry(frame_empresa, width=50)
        self.entry_empresa.insert(0, "EKENOX DISTRIBUIDORA DE COZ. IND. LTDA")
        self.entry_empresa.grid(row=0, column=1, pady=2)

        tk.Label(frame_empresa, text="Endereço:").grid(
            row=1, column=0, sticky="e"
        )
        self.entry_endereco = tk.Entry(frame_empresa, width=50)
        self.entry_endereco.insert(0, "Rua: José de Ribamar Souza, 499")
        self.entry_endereco.grid(row=1, column=1, pady=2)

        tk.Label(frame_empresa, text="Bairro:").grid(
            row=2, column=0, sticky="e"
        )
        self.entry_bairro = tk.Entry(frame_empresa, width=50)
        self.entry_bairro.insert(0, "Pq. Industrial")
        self.entry_bairro.grid(row=2, column=1, pady=2)

        tk.Label(frame_empresa, text="Cidade:").grid(
            row=3, column=0, sticky="e"
        )
        self.entry_cidade = tk.Entry(frame_empresa, width=50)
        self.entry_cidade.insert(0, "Catanduva")
        self.entry_cidade.grid(row=3, column=1, pady=2)

        tk.Label(frame_empresa, text="Estado:").grid(
            row=4, column=0, sticky="e"
        )
        self.entry_estado = tk.Entry(frame_empresa, width=50)
        self.entry_estado.insert(0, "SP")
        self.entry_estado.grid(row=4, column=1, pady=2)

        tk.Label(frame_empresa, text="CEP:").grid(
            row=5, column=0, sticky="e"
        )
        self.entry_cep = tk.Entry(frame_empresa, width=50)
        self.entry_cep.insert(0, "15803-290")
        self.entry_cep.grid(row=5, column=1, pady=2)

        tk.Label(frame_empresa, text="Telefone:").grid(
            row=6, column=0, sticky="e"
        )
        self.entry_telefone = tk.Entry(frame_empresa, width=50)
        self.entry_telefone.insert(0, "(11)98740-3669")
        self.entry_telefone.grid(row=6, column=1, pady=2)

        tk.Label(frame_empresa, text="Email SAC:").grid(
            row=7, column=0, sticky="e"
        )
        self.entry_email = tk.Entry(frame_empresa, width=50)
        self.entry_email.insert(0, "sac@ekenox.com.br")
        self.entry_email.grid(row=7, column=1, pady=2)

        # --- Campos Produto ---
        frame_produto = tk.LabelFrame(
            self, text="Informações do Produto", padx=10, pady=10
        )
        frame_produto.pack(fill="both", padx=10, pady=5)

        tk.Label(frame_produto, text="Produto:").grid(
            row=0, column=0, sticky="e"
        )
        self.entry_produto = tk.Entry(frame_produto, width=45)
        self.entry_produto.insert(0, "BUFFET TÉRMICO")
        self.entry_produto.grid(row=0, column=1, pady=2, sticky="w")

        btn_buscar_prod = tk.Button(
            frame_produto,
            text="Selecionar...",
            command=lambda: listar_produtos(
                self,
                self.entry_produto,
                self.entry_modelo,
                self.entry_serie,
            ),
        )
        btn_buscar_prod.grid(row=0, column=2, padx=5, pady=2, sticky="w")

        tk.Label(frame_produto, text="Classe:").grid(
            row=1, column=0, sticky="e"
        )
        self.entry_classe = tk.Entry(frame_produto, width=50)
        self.entry_classe.insert(0, "IPX4")
        self.entry_classe.grid(
            row=1, column=1, columnspan=2, pady=2, sticky="w"
        )

        tk.Label(frame_produto, text="Modelo (SKU):").grid(
            row=2, column=0, sticky="e"
        )
        self.entry_modelo = tk.Entry(frame_produto, width=50)
        self.entry_modelo.insert(0, "VIX8368")
        self.entry_modelo.grid(
            row=2, column=1, columnspan=2, pady=2, sticky="w"
        )

        tk.Label(frame_produto, text="Tensão:").grid(
            row=3, column=0, sticky="e"
        )
        self.combo_tensao = ttk.Combobox(
            frame_produto,
            values=["127V", "220V"],
            state="readonly",
            width=47,
        )
        self.combo_tensao.grid(
            row=3, column=1, columnspan=2, pady=2, sticky="w"
        )
        self.combo_tensao.set("127V")

        tk.Label(frame_produto, text="Potência:").grid(
            row=4, column=0, sticky="e"
        )
        self.entry_potencia = ttk.Combobox(
            frame_produto,
            values=["1000W", "2000W"],
            state="readonly",
            width=47,
        )
        self.entry_potencia.grid(
            row=4, column=1, columnspan=2, pady=2, sticky="w"
        )
        self.entry_potencia.set("2000W")

        tk.Label(frame_produto, text="Temperatura:").grid(
            row=5, column=0, sticky="e"
        )
        self.entry_temperatura = tk.Entry(frame_produto, width=50)
        self.entry_temperatura.insert(0, "30°C a 120°C")
        self.entry_temperatura.grid(
            row=5, column=1, columnspan=2, pady=2, sticky="w"
        )

        tk.Label(frame_produto, text="Frequência:").grid(
            row=6, column=0, sticky="e"
        )
        self.entry_frequencia = tk.Entry(frame_produto, width=50)
        self.entry_frequencia.insert(0, "60Hz")
        self.entry_frequencia.grid(
            row=6, column=1, columnspan=2, pady=2, sticky="w"
        )

        tk.Label(
            frame_produto,
            text="Número de Série (prefixo/base):",
        ).grid(row=7, column=0, sticky="e")
        self.entry_serie = tk.Entry(frame_produto, width=50)
        self.entry_serie.insert(0, "EKX2024")
        self.entry_serie.grid(
            row=7, column=1, columnspan=2, pady=2, sticky="w"
        )

        tk.Label(
            frame_produto,
            text="Quantidade de etiquetas:",
        ).grid(row=8, column=0, sticky="e")
        self.entry_quantidade = tk.Entry(frame_produto, width=50)
        self.entry_quantidade.insert(0, "5")
        self.entry_quantidade.grid(
            row=8, column=1, columnspan=2, pady=2, sticky="w"
        )

        # Botões inferiores
        frame_botoes = tk.Frame(self, pady=10)
        frame_botoes.pack(fill="x")

        btn_gerar = tk.Button(
            frame_botoes,
            text="Gerar PDF",
            command=lambda: gerar_etiquetas(
                self,
                self.entry_empresa,
                self.entry_endereco,
                self.entry_bairro,
                self.entry_cidade,
                self.entry_estado,
                self.entry_cep,
                self.entry_telefone,
                self.entry_email,
                self.entry_produto,
                self.entry_modelo,
                self.entry_classe,
                self.combo_tensao,
                self.entry_potencia,
                self.entry_temperatura,
                self.entry_frequencia,
                self.entry_serie,
                self.entry_quantidade,
            ),
            bg="#2563eb",
            fg="white",
            font=("Arial", 12, "bold"),
            width=15,
        )
        btn_gerar.pack(side="left", padx=(40, 10))

        btn_fechar = tk.Button(
            frame_botoes,
            text="Fechar",
            command=self.destroy,
            bg="#ef4444",
            fg="white",
            font=("Arial", 12, "bold"),
            width=15,
        )
        btn_fechar.pack(side="left")

        # --------- Atalho F12: abrir CONSULTA de produtos ---------
        self.bind(
            "<F12>",
            lambda e: listar_produtos(
                self,
                self.entry_produto,
                self.entry_modelo,
                self.entry_serie,
            ),
        )


# ===================== MAIN =====================
def main():
    app = EtiquetaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
