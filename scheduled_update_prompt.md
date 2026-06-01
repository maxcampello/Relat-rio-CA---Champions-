# SCHEDULED TASK — Atualização Diária da Planilha Meta Ads
**Horário:** 09:00 BRT (UTC-3) todos os dias  
**Conta Meta Ads:** `1597556981579641`  
**Google Sheets ID:** `1utbHyuOl9fO5Acf544-AgflDkO9vNsc8`  
**URL da planilha:** https://docs.google.com/spreadsheets/d/1utbHyuOl9fO5Acf544-AgflDkO9vNsc8/edit

---

## OBJETIVO

Atualizar automaticamente a planilha de controle de investimentos Meta Ads com os dados mais recentes da conta `1597556981579641`. A planilha tem duas abas:
- **Aba 1 — "Visão por Estrutura":** hierarquia Campanha → Conjunto → Anúncio com orçamento e status atual
- **Aba 2 — "Visão Temporal":** investimento por mês, semana e dia (dados acumulados históricos + ontem)

---

## PASSO 1 — COLETAR DADOS DA META ADS API

Execute as seguintes chamadas **em paralelo** usando a ferramenta `mcp__7f275e72-0aa3-48b5-9756-0af2125ef323__ads_get_ad_entities`:

### 1a. Campanhas (com spend total)
```
ad_account_id: "act_1597556981579641"
entity_type: "CAMPAIGN"
fields: ["id","name","status","daily_budget","lifetime_budget","budget_remaining","objective","spend_cap"]
date_preset: "maximum"
```

### 1b. Conjuntos de anúncios (com orçamento e status)
```
ad_account_id: "act_1597556981579641"
entity_type: "ADSET"
fields: ["id","name","status","campaign_id","daily_budget","lifetime_budget","budget_remaining","targeting"]
date_preset: "maximum"
```

### 1c. Anúncios (com status)
```
ad_account_id: "act_1597556981579641"
entity_type: "AD"
fields: ["id","name","status","adset_id","campaign_id","creative"]
date_preset: "maximum"
```

### 1d. Gasto diário (para Visão Temporal)
```
ad_account_id: "act_1597556981579641"
entity_type: "ACCOUNT"
fields: ["spend","date_start","date_stop","impressions","clicks"]
date_preset: "maximum"
time_increment: "1"
```

### 1e. Gasto mensal (para consolidação por mês)
```
ad_account_id: "act_1597556981579641"
entity_type: "ACCOUNT"
fields: ["spend","date_start","date_stop"]
date_preset: "maximum"
time_increment: "monthly"
```

---

## PASSO 2 — PROCESSAR E MAPEAR OS DADOS

### Mapeamento Canal/Produto
Aplique as seguintes regras de mapeamento com base no **nome do conjunto de anúncios** ou da **campanha**:

| Palavra-chave no nome | Canal/Produto |
|---|---|
| corporativo, eventos corporativos | Corporativo |
| aniversário, infantil | Infantil |
| FINAL, champions, ingressos, VÍDEOS | Final da Champions |
| copa, COPA | Copa |
| almoço executivo, ALMOÇO | Almoço Executivo |
| reconhecimento, alcance, awareness | Geral (Awareness) |
| vagas, recrutamento, emprego | Recrutamento |

### Tipo de Budget
- Se o orçamento (`daily_budget` ou `lifetime_budget`) está no **ad set** → `ABO`
- Se o orçamento está na **campanha** → `CBO`

### Formatação de Orçamento
- `daily_budget` → `"R$ XX,00/dia"` (Meta retorna em centavos — dividir por 100)
- `lifetime_budget` → `"R$ XXX,00 total"` (dividir por 100)
- Se não há orçamento definido no nível → herda do nível acima

### Status
- `ACTIVE` → `"Ativa"`
- `PAUSED` → `"Pausada"`
- `ARCHIVED` → `"Arquivada"`
- `DELETED` → ignorar (não incluir)

### Consolidação temporal
- **Por mês:** agrupar os dados diários por `YYYY-MM`, somar `spend`
- **Por semana:** agrupar os dados do mês atual por semana ISO (segunda a domingo), somar `spend`
- **Por dia:** todos os dias do mês atual em ordem cronológica

---

## PASSO 3 — GERAR O ARQUIVO XLSX ATUALIZADO

Execute o seguinte script Python via ferramenta `Bash` (instale openpyxl se necessário: `pip install openpyxl -q`):

O script deve:
1. Usar a estrutura de formatação do arquivo `gerar_planilha_meta.py` existente no repositório
2. Substituir os dados hardcoded pelas variáveis coletadas no Passo 1
3. Atualizar a data no cabeçalho para a data de hoje no formato `DD/MM/AAAA`
4. Salvar em `/tmp/controle_investimentos_meta.xlsx`

**Cores e formatação a manter:**
- Cabeçalho: `AZUL_ESCURO = "1B2A4A"`, sub-cabeçalho: `AZUL_MEDIO = "2E4A7A"`
- Status Ativa: fundo `D6EAD7`, texto `1A5C1A` bold
- Status Pausada: fundo `FAD7D7`, texto `8B1A1A`
- Cores por Canal:
  - Corporativo: `D6EAF8`
  - Infantil: `D5F5E3`
  - Final da Champions: `FEF9E7`
  - Almoço Executivo: `FDEDEC`
  - Copa: `EBF5FB`
  - Geral (Awareness): `F4ECF7`
  - Recrutamento: `FDFEFE`
- Freeze panes: Aba1=`A4`, Aba2=`A22`
- Larguras Aba1: `[12, 18, 52, 42, 28, 20, 12, 12]`
- Larguras Aba2: `[18, 10, 55, 22, 22, 10]`
- Formato moeda: `'"R$"#,##0.00'`

---

## PASSO 4 — ATUALIZAR O ARQUIVO NO GOOGLE DRIVE

Use a ferramenta `mcp__6369182a-5546-4bd5-b9ee-2bfdb337773e__create_file` para substituir o arquivo existente:

```
name: "controle_investimentos_meta.xlsx"
mime_type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
content: <conteúdo binário do arquivo gerado em /tmp/controle_investimentos_meta.xlsx>
```

> **Nota:** Se a ferramenta suportar sobrescrever por ID, use o ID `1utbHyuOl9fO5Acf544-AgflDkO9vNsc8` para substituir o arquivo existente em vez de criar um novo.

---

## PASSO 5 — VALIDAÇÃO E LOG

Após a atualização, confirme:
1. O arquivo foi gerado com sucesso em `/tmp/controle_investimentos_meta.xlsx`
2. O upload para o Google Drive foi concluído
3. A planilha está acessível em: https://docs.google.com/spreadsheets/d/1utbHyuOl9fO5Acf544-AgflDkO9vNsc8/edit

Registre no output:
- Data/hora da atualização
- Número de campanhas encontradas
- Número de conjuntos de anúncios
- Número de anúncios
- Total de gasto histórico (R$)
- Gasto do dia anterior (R$)

---

## CONTEXTO DE NEGÓCIO

**Empresa:** Reis Andrade & Co — espaço de eventos em arena  
**Conta Meta Ads:** `1597556981579641`  
**Início das campanhas:** 12/02/2026  
**Categorias de eventos:**
- Corporativo (eventos B2B, MICE)
- Infantil (festas de aniversário)
- Final da Champions (evento esportivo premium)
- Copa (campeonato de futebol)
- Almoço Executivo (produto gastronômico)
- Geral / Awareness (reconhecimento de marca)
- Recrutamento (vagas internas)

**Estrutura de campanhas:**
- Todas as campanhas usam modelo **ABO** (orçamento a nível de conjunto de anúncios)
- Objetivo predominante: **ENGAJAMENTO** (mensagens para WhatsApp)
- Funis: ToFU (reconhecimento) → BoFU (conversão/mensagem)

---

## TRATAMENTO DE ERROS

Se qualquer chamada à Meta Ads API falhar:
1. Tente novamente 2x com intervalo de 5 segundos
2. Se persistir, mantenha os dados da última atualização bem-sucedida na planilha
3. Adicione uma nota de rodapé na célula `A1` da Aba 2 indicando: `⚠️ Dados de [ÚLTIMA DATA BEM-SUCEDIDA] — falha na atualização de [DATA ATUAL]`

Se o upload para o Google Drive falhar:
1. Tente novamente 2x
2. Salve o arquivo em `/tmp/controle_investimentos_meta_YYYYMMDD.xlsx` como backup
3. Registre o erro no log de output

---

*Prompt gerado em: 01/06/2026 | Conta: 1597556981579641 | Planilha: 1utbHyuOl9fO5Acf544-AgflDkO9vNsc8*
