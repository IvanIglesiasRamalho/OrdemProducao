# Guia de Manutenção — Sistema de Ordem de Produção (Ekenox)

> Documento voltado à equipe de engenharia/manutenção para operar, ajustar, evoluir e dar suporte ao aplicativo **ordem_producao_windows_corrigido.py** (Tkinter + PostgreSQL + OpenPyXL + Webhook n8n).

## 1) Escopo e responsabilidades

Este guia cobre:

- Operação e suporte do aplicativo (execução, troubleshooting, logs).
- Manutenção de configurações (diretórios, Excel, webhook, executáveis auxiliares).
- Manutenção de integrações (PostgreSQL, n8n, Excel).
- Alterações comuns de negócio (situação de finalização, filtros de categoria, regras de quantidade).
- Processo de build (dev e PyInstaller) e boas práticas de segurança.

Não cobre:
- Regras completas de estoque/planejamento (o cálculo de quantidade nesta versão está **simplificado**).
- Infra/DBA (backup, replicação, HA) em profundidade — apenas diretrizes mínimas.

---

## 2) Visão rápida do sistema

### Componentes
- **UI**: Tkinter (`OrdemProducaoApp`), com Splash (entrada) e janela principal.
- **Persistência**: PostgreSQL via `psycopg2` (`SistemaOrdemProducao`).
- **Excel**: geração de Pedido de Compra via `openpyxl`.
- **Integrações**:
  - Webhook n8n (opcional) ao salvar e finalizar OP.
  - Programa externo `etiqueta.exe` (atalho F12).

### Artefatos esperados em `BASE_DIR`
- `pedido-de-compra v2.xlsx` (modelo)
- `saida_pedido-de-compra v2.xlsx` (saída)
- `erro_app.log` (log de exceções)
- `etiqueta.exe` (opcional)
- `imagens/favicon.ico` e `imagens/avatar_ekenox.png` (opcionais)

---

## 3) Ambiente e dependências

### Requisitos
- Python 3.10+ (recomendado).
- Acesso à instância PostgreSQL do Ekenox (rede, credenciais e schema).
- Permissão de leitura/escrita no `BASE_DIR` (rede `Z:` no exemplo).

### Dependências Python
Recomendado `requirements.txt`:

```txt
psycopg2-binary==2.9.*
openpyxl==3.1.*
requests==2.32.*
```

> `tkinter` geralmente já vem com o Python do Windows.

---

## 4) Configurações e pontos de ajuste

### 4.1 Diretório operacional (BASE_DIR)
No código:
```py
BASE_DIR = r"Z:\Planilhas_OP"
```

**Responsabilidades do mantenedor:**
- Garantir que o caminho exista e é acessível na máquina do usuário.
- Em rede (Z:), validar mapeamento e permissões (especialmente em execução como `.exe`).

**Melhoria recomendada (futuro):**
- Tornar configurável via `.env` / `.ini` para reduzir hardcode.

---

### 4.2 Arquivos Excel (modelo e saída)
No código:
- `CAMINHO_MODELO = BASE_DIR/pedido-de-compra v2.xlsx`
- `CAMINHO_SAIDA  = BASE_DIR/saida_pedido-de-compra v2.xlsx`

**Checklist de manutenção do modelo:**
- Confirmar que existe uma aba **exatamente** chamada `Pedido de Compra`.
- Validar as células usadas:
  - `D6` (número do pedido)
  - `D8` (data)
  - `D10` (fornecedor)
  - Tabela: linhas **16..42**, colunas **B..I**

**Quando atualizar o layout do Excel:**
- Se o template mudar (linhas/colunas), atualizar:
  - limpeza de tabela (ranges)
  - mapeamento dos campos (D6/D8/D10, colunas B..I)
- Sempre testar com *merged cells* (o código já trata).

---

### 4.3 Webhook n8n
No código:
```py
N8N_WEBHOOK_URL = "http://localhost:56789/webhook/ordem-producao"  # ou ""
```

**Para desabilitar:**
```py
N8N_WEBHOOK_URL = ""
```

**Payloads importantes**
- Finalização (`finalizar_ordem_individual`):
```json
{ "ordem_id": 123, "data_fim": "YYYY-MM-DD" }
```
- Salvar OP (`enviar_webhook_ordem`): payload completo (id, numero, depósitos, datas, produto, etc.)

**Boas práticas recomendadas:**
- Colocar segredo/token no webhook (query param, header) e validar no n8n.
- Implementar retentativas (retry) com limite e backoff quando webhook for crítico.
- Versionar endpoint (ex: `/v1/ordem-producao`).

---

### 4.4 Integração com `etiqueta.exe`
- Caminho esperado:
```py
exe_path = os.path.join(BASE_DIR, "etiqueta.exe")
```

**Manutenção:**
- Garantir que o arquivo existe e é confiável (origem).
- Se mudar o nome/localização do executável, atualizar o caminho no método `abrir_programa_etiqueta`.

---

## 5) Banco de dados: manutenção operacional

### 5.1 Conexão
Parâmetros estão hardcoded no `__init__` do app (`OrdemProducaoApp`).

**Ações recomendadas:**
- **Remover senha hardcoded**: usar variáveis de ambiente ou arquivo local não versionado.
- Validar porta e host (firewall/VPN).

### 5.2 Tabelas utilizadas (mínimo)
O aplicativo pressupõe o schema `"Ekenox"` e usa (no mínimo):
- `produtos`
- `infoProduto`
- `situacao`
- `deposito`
- `ordem_producao`
- `sequenciadores` (há métodos, embora não esteja central no fluxo atual)

### 5.3 Queries críticas (pontos de falha)
- Listagens com `JOIN` em `produtos` e `situacao` para montar descrição legível.
- Inserção na tabela `"ordem_producao"` com validações prévias.
- Finalização: `UPDATE` em `"ordem_producao"` alterando `data_fim` e `situacao_id`.

### 5.4 Alteração da situação “finalizada”
No código, finalização define:
```sql
situacao_id = 18162
```

**Se o ID mudar:**
- Atualizar o valor `18162` no método `finalizar_ordem_individual`.
- (Opcional) guardar em constante no topo do arquivo (`SITUACAO_FINALIZADA_ID`).

---

## 6) Mudanças comuns (runbook de alterações)

### 6.1 Atualizar filtro de categorias excluídas (produtos disponíveis)
A query de produtos filtra categorias via `NOT IN (...)`.

**Quando alterar:**
- Se a regra de negócio mudar (categorias que devem aparecer/não aparecer).
- Se o cadastro de categorias for reorganizado.

**Como manter:**
- Documentar o motivo de cada categoria ID excluída (ideal: comentário ou arquivo separado).
- Validar impacto de performance (lista longa de `NOT IN` pode pesar).

Melhoria recomendada:
- Mover esses IDs para uma tabela de configuração no banco (ex: `categorias_excluidas_op`).

---

### 6.2 Reintroduzir cálculo completo de quantidade
Atualmente `atualizar_quantidade_producao()` é simplificado (qtd = 1).

**Como evoluir corretamente:**
- Extrair cálculo para função pura (`services/quantidade.py`) para permitir testes.
- Retornar também `variaveis_quantidade` para o modal “Detalhes (F6)”.

**Checklist antes de subir:**
- Testes com produtos inexistentes.
- Verificar performance (não fazer N queries por digitação).
- Prevenir travamento de UI (usar cache ou operações rápidas).

---

### 6.3 Ajustar validações de datas
Validações atuais:
- `prev_final >= prev_inicio`
- `data_fim >= data_inicio`

Se regras mudarem:
- Ajustar no método `salvar_ordem`.
- Manter mensagens claras ao usuário.

---

### 6.4 Ajustar campos obrigatórios
Obrigatórios atuais:
- Número da ordem, depósitos, situação, produto, quantidade.

Se mudar:
- Ajustar `salvar_ordem()` e o texto UI (labels com `*`).

---

## 7) Logs, diagnóstico e troubleshooting

### 7.1 Log de exceções
`log_exception()` grava stacktrace em:
- `BASE_DIR/erro_app.log`

**Como usar em suporte:**
- Pedir ao usuário o arquivo `erro_app.log`.
- Identificar o bloco mais recente (separado por `=====`).

### 7.2 Problemas comuns e resolução

#### “Erro ao conectar ao banco”
- Verificar host/porta.
- Verificar credenciais.
- Verificar acesso de rede.
- Verificar schema `"Ekenox"` e permissões do usuário.

#### “Aba de modelo 'Pedido de Compra' não encontrada”
- Abrir o Excel e confirmar o nome da aba.
- Confirmar se o arquivo usado é o correto (`CAMINHO_MODELO`).

#### Travamento ao abrir listas (Produtos/Depósitos/Situações)
- Possível volume alto (carregando muitos registros).
- Mitigação:
  - usar `LIMIT` e paginação (melhoria futura)
  - adicionar filtro/busca no Treeview

#### Webhook n8n falha
- Verificar se o n8n está de pé e endpoint correto.
- Verificar firewall/localhost/porta.
- Como mitigação:
  - manter webhook opcional (já é).
  - registrar falha em log estruturado (melhoria).

#### `etiqueta.exe` não encontrado
- Confirmar existência em `BASE_DIR`.
- Confirmar permissões de execução.

---

## 8) Processo de build e release (recomendado)

### 8.1 Execução em desenvolvimento
```bash
python ordem_producao_windows_corrigido.py
```

### 8.2 Empacotamento com PyInstaller (diretriz)
Pontos importantes:
- Incluir arquivos: imagens, Excel modelo e (opcional) `etiqueta.exe`.
- O código já detecta `sys.frozen` para `APP_DIR`.

Recomendação:
- Evitar depender de `Z:` dentro do executável; preferir pasta do usuário e copiar o template na primeira execução.

---

## 9) Segurança (obrigatório para manutenção)

### 9.1 Nunca versionar credenciais
Hoje há senha no código. Isso deve ser removido assim que possível.
- Usar `.env` local (não versionado) ou `config.ini` fora do repositório.
- Permissões no banco com princípio do menor privilégio.

### 9.2 Webhook
- Proteger com token (header) e validar no n8n.
- Se rodar fora do localhost, usar HTTPS.

### 9.3 Executáveis externos
- Controlar distribuição e integridade do `etiqueta.exe`.
- Evitar permitir caminho arbitrário vindo do usuário.

---

## 10) Melhores práticas de manutenção (roadmap técnico)

1. **Config externa**: `.env` / `.ini` e secrets fora do código.
2. **Logging**: trocar `print()` por `logging` e rotacionar arquivo.
3. **Arquitetura em camadas**: UI → service → repository.
4. **Testes**:
   - unitários (cálculo de quantidade, validações)
   - integração (queries críticas) em ambiente de staging
5. **Performance**:
   - paginação/filtro nas listas (Treeview)
   - cache de consultas (produtos/situações/depósitos)
6. **UX**:
   - campo de busca nas listas
   - bloquear botões durante operações longas

---

## 11) Checklist rápido pós-alteração

Antes de publicar uma versão nova:

- [ ] App abre e fecha sem erros (incluindo splash).
- [ ] Conexão com banco OK.
- [ ] Inserção de OP funciona.
- [ ] Lista de ordens abre.
- [ ] Exclusão funciona (onde permitido).
- [ ] Finalização pendentes funciona e atualiza data/situação.
- [ ] Webhook (se ativo) envia payload corretamente.
- [ ] Excel gera sem quebrar merges e com aba “Pedido de Compra” válida.
- [ ] `erro_app.log` registra erros inesperados (teste forçado).
- [ ] Atalho F12 abre `etiqueta.exe` (se aplicável).

---

## 12) Apêndice: onde mexer no código (mapa rápido)

- Config (paths/webhook): topo do arquivo (`BASE_DIR`, `CAMINHO_MODELO`, `N8N_WEBHOOK_URL`)
- Banco e SQL: classe `SistemaOrdemProducao`
- UI/fluxo: classe `OrdemProducaoApp`
- Cálculo de quantidade: `OrdemProducaoApp.atualizar_quantidade_producao`
- Finalização (situação final): `SistemaOrdemProducao.finalizar_ordem_individual`
- Excel (template/saída): `gerar_abas_fornecedor_pedido` e helpers

---

### Contatos e responsabilidades (preencher)
- Responsável técnico:
- DBA / Infra:
- Dono do processo (negócio):
- Canal de suporte:

