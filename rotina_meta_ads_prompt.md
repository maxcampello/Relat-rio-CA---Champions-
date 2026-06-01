# PROMPT — Rotina: Atualização Diária Meta Ads — Champions League

> Cole este texto no campo "Instruções" ao criar a nova rotina no Claude Cowork.
> Configurar: Runs daily at 9:00 BRT | Conectores: Meta MCP + Google Drive

---

## INSTRUÇÕES (cole no campo da rotina)

Você é um agente de automação de marketing digital. Sua tarefa é atualizar a planilha de controle de investimentos Meta Ads da conta **1597556981579641** com os dados do dia anterior.

---

## PASSO 1 — COLETAR DADOS DA META ADS (em paralelo)

Use o conector Meta MCP para buscar simultaneamente:

**1a. Estrutura de campanhas:**
- Entidade: CAMPAIGN → fields: id, name, status, daily_budget, lifetime_budget, objective
- Entidade: ADSET → fields: id, name, status, campaign_id, daily_budget, lifetime_budget
- Entidade: AD → fields: id, name, status, adset_id, campaign_id
- Conta: act_1597556981579641 | date_preset: maximum

**1b. Investimento por período:**
- Entidade: ACCOUNT | fields: spend, date_start, date_stop | date_preset: maximum | time_increment: monthly
- Entidade: ACCOUNT | fields: spend, date_start, date_stop, impressions, clicks | date_preset: maximum | time_increment: 1

---

## PASSO 2 — PROCESSAR OS DADOS

**Mapeamento Canal/Produto** (pelo nome do conjunto ou campanha):
- "corporativo" ou "eventos corporativos" → Corporativo
- "aniversário" ou "infantil" → Infantil
- "FINAL", "champions", "ingressos", "VÍDEOS" → Final da Champions
- "copa" ou "COPA" → Copa
- "almoço" ou "executivo" → Almoço Executivo
- "reconhecimento", "alcance", "awareness" → Geral (Awareness)
- "vagas", "recrutamento", "emprego" → Recrutamento

**Tipo de Budget:**
- Orçamento no ad set → ABO
- Orçamento na campanha → CBO

**Conversão de valores:** Meta retorna em centavos — dividir por 100 para obter R$

**Status:**
- ACTIVE → "Ativa" | PAUSED → "Pausada" | ARCHIVED → "Arquivada" | DELETED → ignorar

**Consolidação temporal:**
- Por mês: agrupar dados diários por YYYY-MM e somar spend
- Por semana: agrupar dias do mês atual por semana (seg–dom) e somar spend
- Por dia: todos os dias do mês atual em ordem cronológica com coluna de acumulado

---

## PASSO 3 — GERAR OS DOIS CSVs

### Arquivo 1 — Visão por Estrutura
Colunas: Plataforma, Canal/Produto, Nome da Campanha (Nível 1), Conjunto de Anúncios (Nível 2), Nome do Anúncio (Nível 3), Orçamento Configurado, Tipo de Budget, Status

- Uma linha por anúncio
- Ignorar anúncios com status DELETED
- Ordenar por: Canal/Produto → Campanha → Conjunto → Anúncio
- title: "Visão por Estrutura — Meta Ads | Champions League [DD/MM/YYYY]"
- contentMimeType: text/csv
- base64Content: base64 do CSV 1

### Arquivo 2 — Visão Temporal
Três seções separadas por linha em branco:

**Seção 1 — Por Mês** (colunas: Mês/Período, Gasto Real R$, Campanhas Ativas, Observações)
- Uma linha por mês desde fev/2026 até o mês atual
- Última linha: TOTAL INVESTIDO

**Seção 2 — Por Semana do mês atual** (colunas: Semana, Período, Gasto Real R$, % do Mês)
- Semana 1 a N do mês corrente
- Última linha: TOTAL DO MÊS

**Seção 3 — Por Dia do mês atual** (colunas: Data, Semana, Campanhas Ativas no Dia, Gasto R$, Acumulado Mês R$)
- Um linha por dia, do dia 1 até ontem
- Última linha: TOTAL MÊS ATÉ HOJE

- title: "Visão Temporal — Meta Ads | Champions League [DD/MM/YYYY]"
- contentMimeType: text/csv
- base64Content: base64 do CSV 2

Ambos os arquivos serão automaticamente convertidos para Google Sheets nativos.

---

## REFERÊNCIA DE CAMPANHAS

Campanha 01: [ENGAJAMENTO][BoFU][MENSAGEM] - 01
→ Canal: Corporativo + Infantil | ABO | Ativa (conjuntos 01) e Pausada (conjuntos 00, 02)

Campanha 02: [ENGAJAMENTO][BoFU][MENSAGEM] - 02
→ Canal: Corporativo + Infantil | ABO | Ativa (conjuntos 01, 03) e Pausada (conjuntos 00, 01, 02)

Campanha 03 FINAL: [ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL
→ Canal: Corporativo | ABO | Ativa (conjuntos 01) e Pausada (conjuntos 00)

Tráfego Perfil: [TRÁFEGO][ToFU][VISITAS AO PERFIL] - 12/02/2026
→ Canal: Almoço Executivo (ADS04 Ativa) + Corporativo (ADS01-03 Pausadas)

Copa: [ENGAJAMENTO][MENSAGEM] - COPA
→ Canal: Copa | ABO | Ativa (AD01, AD02, AD03)

Champions Vídeos: [ENGAJAMENTO][MENSAGEM][VÍDEOS] - FINAL
→ Canal: Final da Champions | ABO | Ativa (AD 01, AD 02)

---

## CRITÉRIOS DE SUCESSO
✅ Dois arquivos Google Sheets criados no Drive com a data de hoje no título
✅ Dados desde 12/02/2026 até ontem sem lacunas
✅ Nenhuma célula com #ERROR
✅ Totais mensais = soma exata dos dias
✅ Seção semanal cobre apenas o mês atual
✅ Valores monetários com 2 casas decimais em R$
✅ Coluna "Acumulado Mês" cresce progressivamente dia a dia
✅ Exibir no final os links dos dois arquivos criados no Google Drive
