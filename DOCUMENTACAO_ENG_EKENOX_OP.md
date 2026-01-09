# Documentação de Engenharia — Sistema de Ordem de Produção (Ekenox)

**Artefato:** `ordem_producao_windows_corrigido.py`  
**Stack:** Python + Tkinter + PostgreSQL (`psycopg2`) + OpenPyXL + Webhook n8n (opcional)  
**Plataforma alvo:** Windows (há trechos específicos para Windows; execução em outras plataformas é possível com limitações)

> Este documento descreve arquitetura, fluxos, integrações, decisões técnicas e recomendações para evolução/manutenção do código fornecido.

---

## 1. Objetivos do sistema

O sistema tem como objetivo permitir a criação e manutenção de **Ordens de Produção (OPs)** no banco PostgreSQL (schema `"Ekenox"`), além de:
- Listar produtos, depósitos e situações cadastradas.
- Inserir novas OPs com validações básicas.
- Listar OPs existentes e permitir exclusão.
- Finalizar OPs pendentes (sem data_fim) alterando data_fim e situação.
- (Opcional) Enviar eventos para um **webhook n8n**.
- (Opcional) Abrir aplicativo externo de etiquetas (`etiqueta.exe`).

---

## 2. Visão geral da arquitetura

### 2.1 Componentes principais
1. **UI (Tkinter)** — classe `OrdemProducaoApp`
   - Splash/entrada (`Toplevel`) e janela principal (`Tk`).
   - Formulário de criação de OP e ações de listagem/finalização.
   - Bindings de atalhos (`F2`, `F3`, `F4`, `F5`, `F6`, `F8`, `F10`, `F11`, `F12`).

2. **Acesso a dados / Regras** — classe `SistemaOrdemProducao`
   - Conexão e cursor `psycopg2`.
   - Validações de FK (produto e situação).
   - Listagens (produtos, depósitos, situações, ordens).
   - Escrita (inserção, exclusão, finalização).

3. **Geração de Excel (OpenPyXL)**
   - Geração de abas por fornecedor/pedido para planilha "Pedido de Compra".
   - Funções: `gerar_abas_fornecedor_pedido`, `gerar_planilha_excel`, `gerar_pedido_compra`.

4. **Integrações**
   - **n8n** via `requests.post()` em:
     - `SistemaOrdemProducao.finalizar_ordem_individual`
     - `OrdemProducaoApp.enviar_webhook_ordem`
   - **Executável externo** `etiqueta.exe` via `os.startfile()` (Windows).

### 2.2 Decisões de design já aplicadas
- **Apenas 1 instância de `Tk()`**: evita o erro clássico de Tkinter (`wm command` após destroy).
- `WM_DELETE_WINDOW` registrado no `__init__` para evitar referências inválidas.
- `withdraw()` antes do splash e `deiconify()` após “Entrar”.
- Tratamento de *merged cells* no Excel para evitar exceções ao preencher.

---

## 3. Estrutura de configuração

### 3.1 Caminhos e diretórios
- `BASE_DIR = r"Z:\Planilhas_OP"` (hardcoded)
- `APP_DIR` é derivado de `sys.frozen` (PyInstaller) ou do `__file__`

**Arquivos esperados em `BASE_DIR`:**
- `pedido-de-compra v2.xlsx` (modelo)
- `saida_pedido-de-compra v2.xlsx` (gerada/atualizada)
- `erro_app.log` (log de exceções)
- `etiqueta.exe` (opcional)
- `imagens/favicon.ico` (opcional)
- `imagens/avatar_ekenox.png` (opcional)

### 3.2 Webhook n8n
- `N8N_WEBHOOK_URL = "http://localhost:56789/webhook/ordem-producao"`
- Desabilitar deixando string vazia: `N8N_WEBHOOK_URL = ""`

### 3.3 Warnings OpenPyXL
Warnings específicos são ignorados para reduzir ruído em runtime.

---

## 4. Banco de dados (PostgreSQL)

### 4.1 Dependências de schema/tabelas
O código assume schema `"Ekenox"` e consulta/escreve nestas tabelas:
- `"produtos"`
- `"infoProduto"`
- `"situacao"`
- `"deposito"`
- `"ordem_producao"`
- `"sequenciadores"`

### 4.2 Inserção de OP (tabela `"ordem_producao"`)
A inserção exige (mínimo):
- `numero`
- `deposito_id_destino`
- `deposito_id_origem`
- `situacao_id`
- `fkprodutoid`
- `quantidade`
e opcionais:
- `responsavel`, datas, `valor`, `observacao`

**Validações prévias:**
- `validar_produto(fkprodutoid)`
- `validar_situacao(situacao_id)`

**Geração de ID:**
- Hoje é feita por `SELECT COALESCE(MAX(id), 0) + 1`, o que é **arriscado** sob concorrência.

✅ Recomendação: migrar `id` para `GENERATED AS IDENTITY` e remover geração manual no app.

### 4.3 Finalização de OP pendente
`finalizar_ordem_individual()` faz:
- `data_fim = today()`
- `situacao_id = 18162`
- Somente quando `data_fim IS NULL` ou `data_fim = '1970-01-01'`

✅ Recomendação: extrair `18162` para constante (`SITUACAO_FINALIZADA_ID`).

### 4.4 Listagens e performance
Listagens atuais carregam **tudo** (sem paginação), com `Treeview` preenchendo muitas linhas.
- Pode travar a UI com bases grandes.

✅ Recomendação: adicionar:
- filtros/busca (nome, sku, id)
- paginação (LIMIT/OFFSET) ou cursor incremental
- cache de resultados (produtos/situações/depósitos) por sessão

---

## 5. UI/UX — Fluxos e atalhos

### 5.1 Fluxo de inicialização
1. App instancia `OrdemProducaoApp()`
2. Configura janela + icon
3. `withdraw()`
4. Conecta ao banco e carrega totais
5. Monta widgets
6. Define número da ordem sugerido (`gerar_numero_ordem`)
7. Abre splash (`mostrar_tela_entrada`) com `after(50)`
8. Ao clicar **Entrar**, fecha splash e `deiconify()` a janela principal

### 5.2 Atalhos globais
- `F2` Lista Produtos
- `F3` Lista Situações
- `F4` Lista Depósitos (Origem)
- `F5` Totais
- `F6` Detalhes do cálculo de quantidade
- `F8` Lista Depósitos (Destino)
- `F10` Ordens existentes
- `F11` Finalizar pendentes
- `F12` Etiquetas (`etiqueta.exe`)

### 5.3 Formulário: validações
No `salvar_ordem()`:
- Campos obrigatórios checados
- Datas validadas (formato `DD/MM/AAAA` e coerência entre início/fim)
- Confirmação via `messagebox.askyesno`

---

## 6. Excel — Geração de Pedido de Compra

### 6.1 Modelo e estrutura
A função `gerar_abas_fornecedor_pedido()`:
- Carrega `CAMINHO_SAIDA` se existir, senão `CAMINHO_MODELO`
- Duplica a aba `Pedido de Compra`
- Preenche:
  - `D6` número pedido
  - `D8` data pedido
  - `D10` fornecedor
- Limpa linhas 16..42 (inclui) em colunas 2..9 (B..I)
- Escreve itens a partir da linha 16

### 6.2 Merge cells
- `_escrever_celula_segura` e `_set_cell_segura_rc` evitam escrever em `MergedCell` direto.

### 6.3 Ponto de atenção
- A função salva sempre em `CAMINHO_SAIDA`, acumulando abas (risco de poluição ao longo do tempo).

✅ Recomendações:
- Opção de gerar arquivo “por pedido” (timestamp/número) para auditoria.
- Limpar abas antigas automaticamente (se desejado).
- Validar concorrência (dois usuários gerando ao mesmo tempo no mesmo share).

---

## 7. Webhook n8n — Eventos e confiabilidade

### 7.1 Tipos de eventos
1. **Finalização**: payload simples `{ordem_id, data_fim}`
2. **Inserção**: payload completo em `enviar_webhook_ordem()`

### 7.2 Boas práticas recomendadas
- Introduzir **timeout + retries** com backoff controlado quando webhook for crítico.
- Logar falhas em arquivo (não só `print`).
- Autenticação (token) e HTTPS se sair do localhost.
- Versionar endpoint (`/v1/...`).

---

## 8. Logging e observabilidade

### 8.1 Log de exceções
`log_exception()` registra stacktrace em:
- `BASE_DIR/erro_app.log`

✅ Recomendações:
- Usar `logging` do Python com `RotatingFileHandler`
- Log estruturado (JSON) para correlacionar erros e eventos do webhook
- Capturar também status/erro de banco quando `self._ultimo_erro_bd` for setado

---

## 9. Problemas técnicos conhecidos e melhorias

### 9.1 Segurança (crítico)
- Credenciais do banco **hardcoded** no código.

✅ Migrar para:
- Variáveis de ambiente
- `.env` (não versionado) + `python-dotenv`
- arquivo `config.ini` em pasta protegida
- secret manager (ideal)

### 9.2 IDs gerados no app (concorrência)
- `gerar_id_ordem()` = `MAX(id) + 1` é vulnerável a corrida.

✅ Solução:
- `id` como `GENERATED AS IDENTITY`
- `INSERT ... RETURNING id`
- Remover campo `id` do insert e deixar o banco gerar

### 9.3 UI bloqueada em operações longas
- Listagens grandes e queries podem travar o Tkinter (single-thread UI).

✅ Soluções:
- Paginação e filtro
- Thread/async com cuidado (comunicar com UI via `after()`)

### 9.4 Cálculo de quantidade está simplificado
`atualizar_quantidade_producao` sugere `1.0` e define variáveis para F6.

✅ Evolução:
- Extrair cálculo para módulo separado com testes unitários.

---

## 10. Padrões recomendados para refatoração

### 10.1 Separação em camadas
- `ui/` (Tkinter)
- `db/` (repos/repository)
- `services/` (regras: cálculo, validações, webhook)
- `infra/` (config, logging, paths)

### 10.2 Injeção de configuração
- Carregar config em um objeto (dataclass) e passar para `OrdemProducaoApp` e `SistemaOrdemProducao`

### 10.3 Testes
- Unitário: parsing de datas, validações, cálculo de quantidade, geração de payload webhook
- Integração: queries em banco de staging; geração de Excel com template fixture

---

## 11. Checklist de engenharia (antes de release)

- [ ] Credenciais removidas do código (config externa).
- [ ] `id` no banco com auto-incremento (IDENTITY) e `INSERT ... RETURNING id`.
- [ ] Paginação/filtro para listas grandes.
- [ ] Log estruturado e rotação de logs.
- [ ] Webhook protegido (token + HTTPS quando necessário).
- [ ] Template do Excel validado (aba e ranges corretos).
- [ ] Teste de fechamento/abertura sem erros (inclusive splash).
- [ ] Teste de concorrência (2 usuários inserindo OP simultaneamente).

---

## 12. Apêndice: SQL recomendado para auto-incremento do ID

```sql
ALTER TABLE "Ekenox"."ordem_producao"
  ALTER COLUMN id ADD GENERATED BY DEFAULT AS IDENTITY;
```

E no app, usar `INSERT ... RETURNING id` (deixar o banco gerar o id).

---

## 13. Contatos (preencher)
- Responsável técnico:
- DBA/Infra:
- Dono do processo:
- Canal de suporte:
