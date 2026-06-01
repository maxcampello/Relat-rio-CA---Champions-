import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── Cores ──────────────────────────────────────────────────────────────────────
AZUL_ESCURO   = "1B2A4A"   # cabeçalho principal
AZUL_MEDIO    = "2E4A7A"   # sub-cabeçalho
CINZA_CLARO   = "F2F4F7"   # linhas alternadas
BRANCO        = "FFFFFF"
VERDE_STATUS  = "D6EAD7"
VERMELHO_STAT = "FAD7D7"
AMARELO_STAT  = "FFF3CD"
FONTE_BRANCA  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
FONTE_TITULO  = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
FONTE_NORMAL  = Font(name="Calibri", size=10)
FONTE_BOLD    = Font(name="Calibri", bold=True, size=10)
FONTE_DARK    = Font(name="Calibri", bold=True, color=AZUL_ESCURO, size=10)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="D0D5DD")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

BRL_FMT = '"R$"#,##0.00'

# ══════════════════════════════════════════════════════════════════════════════
#  ABA 1 — VISÃO POR ESTRUTURA
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Visão por Estrutura"

# Título
ws1.merge_cells("A1:H1")
c = ws1["A1"]
c.value = "META ADS — VISÃO POR ESTRUTURA DE CAMPANHAS"
c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
c.fill  = fill(AZUL_ESCURO)
c.alignment = center()
ws1.row_dimensions[1].height = 30

ws1.merge_cells("A2:H2")
c = ws1["A2"]
c.value = "Conta: 1597556981579641  |  Plataforma: Meta Ads  |  Atualizado em: 28/05/2026"
c.font  = Font(name="Calibri", italic=True, color="FFFFFF", size=10)
c.fill  = fill(AZUL_MEDIO)
c.alignment = center()
ws1.row_dimensions[2].height = 18

# Cabeçalhos
headers1 = [
    "Plataforma", "Canal / Produto", "Nome da Campanha (Nível 1)",
    "Conjunto de Anúncios (Nível 2)", "Nome do Anúncio (Nível 3)",
    "Orçamento Configurado", "Tipo de Budget", "Status"
]
for col, h in enumerate(headers1, 1):
    c = ws1.cell(row=3, column=col, value=h)
    c.font      = FONTE_BRANCA
    c.fill      = fill(AZUL_ESCURO)
    c.alignment = center()
    c.border    = thin_border()
ws1.row_dimensions[3].height = 28

ws1.freeze_panes = "A4"

# ── Dados ─────────────────────────────────────────────────────────────────────
# Estrutura: (plataforma, canal, campanha, conjunto, anuncio, orcamento, tipo, status)
dados_aba1 = [
    # ── [ENGAJAMENTO][BoFU][MENSAGEM] - 02 ─────────────────────────────────
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","00 - Eventos corporativos","ADS07 - Vídeo","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","01 - Eventos corporativos - 1KM","ADS07 - Vídeo","R$ 20,00/dia","ABO","Ativa"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","03 - Aniversário","ADS05 - Vídeo","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","00 - Aniversário","ADS05 - Vídeo","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","00 - Aniversário","ADS01 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","00 - Aniversário","ADS02 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","00 - Aniversário","ADS03 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","00 - Aniversário","ADS04 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","01 - Aniversário","ADS05 - Vídeo","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 02","02 - Aniversário","ADS05 - Vídeo","R$ 15,00/dia","ABO","Pausada"),
    # ── [ENGAJAMENTO][BoFU][MENSAGEM] - 01 ─────────────────────────────────
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","01 - Eventos corporativos","ADS01 - FEED/STORY","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","01 - Eventos corporativos","ADS03 - FEED/STORY","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","03 - Aniversário","ADS01 - FEED/STORY","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","03 - Aniversário","ADS02 - FEED/STORY","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","03 - Aniversário","ADS03 - FEED/STORY","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","03 - Aniversário","ADS04 - FEED/STORY","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","00 - Eventos corporativos","ADS01 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","00 - Eventos corporativos","ADS03 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","00 - Aniversário","ADS01 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","00 - Aniversário","ADS02 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","00 - Aniversário","ADS03 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","00 - Aniversário","ADS04 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","01 - Aniversário","ADS01 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","01 - Aniversário","ADS02 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","02 - Aniversário","ADS01 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Infantil","[ENGAJAMENTO][BoFU][MENSAGEM] - 01","02 - Aniversário","ADS02 - FEED/STORY","R$ 15,00/dia","ABO","Pausada"),
    # ── [ENGAJAMENTO][BoFU][MENSAGEM] - 03 ─────────────────────────────────
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","01 - Eventos corporativos","ADS03 - IMAGEM - FINAL","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","01 - Eventos corporativos","ADS04 - IMAGEM - FINAL","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","01 - Eventos corporativos","ADS05 - IMAGEM - FINAL","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","01 - Eventos corporativos","ADS06 - IMAGEM - FINAL","R$ 25,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","01 - Eventos corporativos","ADS01 - VÍDEO - SEMI","R$ 25,00/dia","ABO","Pausada"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","00 - Eventos corporativos","ADS01 - VÍDEO - SEMI","R$ 20,00/dia","ABO","Pausada"),
    ("Meta Ads","Corporativo","[ENGAJAMENTO][BoFU][MENSAGEM] - 03 - FINAL","00 - Eventos corporativos","ADS02 - VÍDEO - SEMI","R$ 20,00/dia","ABO","Pausada"),
    # ── [TRÁFEGO][ToFU][VISITAS AO PERFIL] ─────────────────────────────────
    ("Meta Ads","Almoço Executivo","[TRÁFEGO][ToFU][VISITAS AO PERFIL] - 12/02/2026","00 - Aberto","ADS04 - Almoço Executivo","R$ 10,00/dia","ABO","Ativa"),
    ("Meta Ads","Corporativo","[TRÁFEGO][ToFU][VISITAS AO PERFIL] - 12/02/2026","00 - Aberto","ADS01 - Vídeo","R$ 10,00/dia","ABO","Pausada"),
    ("Meta Ads","Corporativo","[TRÁFEGO][ToFU][VISITAS AO PERFIL] - 12/02/2026","00 - Aberto","ADS02 - Vídeo","R$ 10,00/dia","ABO","Pausada"),
    ("Meta Ads","Corporativo","[TRÁFEGO][ToFU][VISITAS AO PERFIL] - 12/02/2026","00 - Aberto","ADS03 - Vídeo","R$ 10,00/dia","ABO","Pausada"),
    # ── [TRÁFEGO][SITE][FINAL] ──────────────────────────────────────────────
    ("Meta Ads","Final da Champions","[TRÁFEGO][SITE][FINAL] - 07/05/2026","00 - Aberto","AD01 - VÍDEO","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[TRÁFEGO][SITE][FINAL] - 07/05/2026","00 - Aberto","AD02 - VÍDEO","R$ 15,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[TRÁFEGO][SITE][FINAL] - 07/05/2026","01 - Eventos corporativos","AD01 - VÍDEO","R$ 70,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[TRÁFEGO][SITE][FINAL] - 07/05/2026","01 - Eventos corporativos","AD02 - VÍDEO","R$ 70,00/dia","ABO","Pausada"),
    # ── [ENGAJAMENTO][VENDA DE INGRESSOS] ──────────────────────────────────
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Interesses AAA - [Número antigo]","ADS05 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Interesses AAA - [Número antigo]","ADS06 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Interesses AAA - [Número antigo]","ADS01 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Interesses AAA - [Número antigo]","ADS02 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Interesses AAA - [Número antigo]","ADS03 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Interesses AAA - [Número antigo]","ADS04 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Lookalike 3% + Interesses AAA","ADS01 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Lookalike 3% + Interesses AAA","ADS02 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Lookalike 3% + Interesses AAA","ADS03 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Lookalike 3% + Interesses AAA","ADS04 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Lookalike 3% + Interesses AAA","ADS05 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][VENDA DE INGRESSOS][MENSAGEM] - FINAL","00 - Lookalike 3% + Interesses AAA","ADS06 - FEED/STORY","R$ 40,00/dia","ABO","Pausada"),
    # ── [VENDAS][BoFU][SITE] - 22/04 ───────────────────────────────────────
    ("Meta Ads","Final da Champions","[VENDAS][BoFU][SITE] - 22/04/2026","00 - Publico Aberto","ADS01 - VÍDEO","R$ 50,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[VENDAS][BoFU][SITE] - 22/04/2026","00 - Publico Aberto","ADS02 - VÍDEO","R$ 50,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[VENDAS][BoFU][SITE] - 22/04/2026","00 - Publico Aberto","ADS03 - VÍDEO","R$ 50,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[VENDAS][BoFU][SITE] - 22/04/2026","00 - Publico Aberto","ADS04 - VÍDEO","R$ 50,00/dia","ABO","Pausada"),
    # ── [VENDAS][BoFU][SITE] - 04/05 ───────────────────────────────────────
    ("Meta Ads","Final da Champions","[VENDAS][BoFU][SITE] - 04/05/2026","00 - Publico Aberto","ADS03 - VÍDEO","R$ 50,00/dia","ABO","Pausada"),
    ("Meta Ads","Final da Champions","[VENDAS][BoFU][SITE] - 04/05/2026","00 - Publico Aberto","ADS04 - VÍDEO","R$ 50,00/dia","ABO","Pausada"),
    # ── [ENGAJAMENTO][MENSAGEM] - COPA ──────────────────────────────────────
    ("Meta Ads","Copa","[ENGAJAMENTO][MENSAGEM] - COPA","00 - Corporativo + Interesses AAA","AD01 - VÍDEO","R$ 40,00/dia","ABO","Ativa"),
    ("Meta Ads","Copa","[ENGAJAMENTO][MENSAGEM] - COPA","00 - Corporativo + Interesses AAA","AD02 - VÍDEO","R$ 40,00/dia","ABO","Ativa"),
    ("Meta Ads","Copa","[ENGAJAMENTO][MENSAGEM] - COPA","00 - Corporativo + Interesses AAA","AD03 - VÍDEO","R$ 40,00/dia","ABO","Ativa"),
    # ── [ENGAJAMENTO][MENSAGEM][VÍDEOS] - FINAL ─────────────────────────────
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][MENSAGEM][VÍDEOS] - FINAL","00 - Interesses AAA","AD 01","R$ 40,00/dia","ABO","Ativa"),
    ("Meta Ads","Final da Champions","[ENGAJAMENTO][MENSAGEM][VÍDEOS] - FINAL","00 - Interesses AAA","AD 02","R$ 40,00/dia","ABO","Ativa"),
    # ── [RECONHECIMENTO][ToFU][ALCANCE] ─────────────────────────────────────
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral — Vídeos novos","ADS01 - Vídeo_Arena Boutique","R$ 22,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral — Vídeos novos","ADS01 - Vídeo_Restaurante","R$ 22,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral — Vídeos novos","ADS01 - Vídeo_Happy Hour","R$ 22,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral — Vídeos novos","ADS01 - Vídeo_Jogos","R$ 22,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral","ADS01 - Vídeo_Restaurante","R$ 6,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral","ADS01 - Vídeo_Happy Hour","R$ 6,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral","ADS01 - Vídeo_Jogos","R$ 6,00/dia","ABO","Pausada"),
    ("Meta Ads","Geral (Awareness)","[RECONHECIMENTO][ToFU][ALCANCE] - 08/04/2026","00 - Geral","ADS01 - Vídeo_Arena Boutique","R$ 6,00/dia","ABO","Pausada"),
    # ── [RECONHECIMENTO][VAGAS DE EMPREGO] ──────────────────────────────────
    ("Meta Ads","Recrutamento","[RECONHECIMENTO][VAGAS DE EMPREGO] - 27/02/2026","00 - [Cozinheiro, Pizzaiolo, Confeiteiro...]","AD01 - FEED/STORY","R$ 6,00/dia","ABO","Pausada"),
]

CORES_CANAL = {
    "Corporativo":       "D6EAF8",
    "Infantil":          "D5F5E3",
    "Final da Champions":"FEF9E7",
    "Almoço Executivo":  "FDEDEC",
    "Copa":              "EBF5FB",
    "Geral (Awareness)": "F4ECF7",
    "Recrutamento":      "FDFEFE",
}

row = 4
prev_camp = None
for d in dados_aba1:
    plat, canal, camp, conj, anuncio, orcamento, tipo, status = d
    alt = (row % 2 == 0)
    bg  = CORES_CANAL.get(canal, BRANCO)

    for col, val in enumerate([plat, canal, camp, conj, anuncio, orcamento, tipo, status], 1):
        c = ws1.cell(row=row, column=col, value=val)
        c.font   = FONTE_NORMAL
        c.border = thin_border()
        c.alignment = left()
        if col in (1, 7, 8):
            c.alignment = center()
        # cor de fundo por canal
        c.fill = fill(bg) if not alt else fill(BRANCO)

    # coluna status colorida
    sc = ws1.cell(row=row, column=8)
    if status == "Ativa":
        sc.fill = fill("D6EAD7")
        sc.font = Font(name="Calibri", color="1A5C1A", bold=True, size=10)
    else:
        sc.fill = fill("FAD7D7")
        sc.font = Font(name="Calibri", color="8B1A1A", size=10)

    ws1.row_dimensions[row].height = 20
    row += 1

# Larguras colunas
col_widths1 = [12, 18, 52, 42, 28, 20, 12, 12]
for i, w in enumerate(col_widths1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w


# ══════════════════════════════════════════════════════════════════════════════
#  ABA 2 — VISÃO TEMPORAL
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Visão Temporal")

def titulo_secao(ws, row, col, texto, span, cor=AZUL_ESCURO):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+span-1)
    c = ws.cell(row=row, column=col, value=texto)
    c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    c.fill = fill(cor)
    c.alignment = center()
    ws.row_dimensions[row].height = 22
    return c

def header_row(ws, row, cols_vals, cor=AZUL_MEDIO):
    for col, val in enumerate(cols_vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill = fill(cor)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[row].height = 22

# ── Título aba 2 ──────────────────────────────────────────────────────────────
ws2.merge_cells("A1:F1")
c = ws2["A1"]
c.value = "META ADS — VISÃO TEMPORAL DE INVESTIMENTOS"
c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
c.fill  = fill(AZUL_ESCURO)
c.alignment = center()
ws2.row_dimensions[1].height = 30

ws2.merge_cells("A2:F2")
c = ws2["A2"]
c.value = "Conta: 1597556981579641  |  Início do projeto: 12/02/2026  |  Atualizado em: 28/05/2026"
c.font  = Font(name="Calibri", italic=True, color="FFFFFF", size=10)
c.fill  = fill(AZUL_MEDIO)
c.alignment = center()

# ─────────────────────────────────────────────────────────────────────────────
#  SEÇÃO 1 – INVESTIMENTO POR MÊS
# ─────────────────────────────────────────────────────────────────────────────
titulo_secao(ws2, 4, 1, "SEÇÃO 1 — INVESTIMENTO POR MÊS (Consolidado)", 6)
header_row(ws2, 5, ["Mês / Período","Meta Ads (Gasto Real)","Campanhas Ativas","Observações","",""])

dados_mes = [
    ("Fevereiro / 2026 (12–28 fev)", 3402.32, "Campanhas 01, 02, Tráfego ToFU, Vagas", "Início das campanhas em 12/02"),
    ("Março / 2026",                 5821.94, "Campanhas 02, Tráfego ToFU, Vagas",      "Mês de maior gasto — Campanha 02 intensificada"),
    ("Abril / 2026",                 5346.28, "Campanhas 01, 02, 03, Tráfego, Reconhecimento, Vendas", "Entrada das campanhas de Vendas e Reconhecimento"),
    ("Maio / 2026 (01–27 mai)",      5829.82, "Todas ativas + novas: Copa, Vídeos, Ingressos", "Campanhas de Champions e Copa ativadas"),
]

total_gasto = sum(d[1] for d in dados_mes)

for i, (mes, valor, camps, obs) in enumerate(dados_mes):
    r = 6 + i
    bg = CINZA_CLARO if i % 2 == 0 else BRANCO
    for col in range(1, 7):
        ws2.cell(row=r, column=col).fill  = fill(bg)
        ws2.cell(row=r, column=col).border = thin_border()
        ws2.cell(row=r, column=col).alignment = left()
        ws2.cell(row=r, column=col).font = FONTE_NORMAL
    ws2.cell(row=r, column=1).value = mes
    c = ws2.cell(row=r, column=2)
    c.value          = valor
    c.number_format  = BRL_FMT
    c.font           = FONTE_BOLD
    c.alignment      = center()
    ws2.cell(row=r, column=3).value = camps
    ws2.cell(row=r, column=4).value = obs
    ws2.row_dimensions[r].height = 20

# Total
rt = 10
ws2.merge_cells(f"A{rt}:A{rt}")
c = ws2.cell(row=rt, column=1, value="TOTAL INVESTIDO (meta ads)")
c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
c.fill  = fill(AZUL_ESCURO)
c.alignment = center()
c.border = thin_border()
c2 = ws2.cell(row=rt, column=2, value=total_gasto)
c2.number_format = BRL_FMT
c2.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
c2.fill  = fill(AZUL_ESCURO)
c2.alignment = center()
c2.border = thin_border()
for col in range(3, 7):
    c = ws2.cell(row=rt, column=col)
    c.fill   = fill(AZUL_ESCURO)
    c.border = thin_border()
ws2.row_dimensions[rt].height = 22

# ─────────────────────────────────────────────────────────────────────────────
#  SEÇÃO 2 – INVESTIMENTO POR SEMANA (Maio/2026)
# ─────────────────────────────────────────────────────────────────────────────
titulo_secao(ws2, 12, 1, "SEÇÃO 2 — INVESTIMENTO POR SEMANA — Maio/2026", 6, AZUL_MEDIO)
header_row(ws2, 13, ["Semana","Período","Meta Ads (Gasto Real)","% do Mês","Ritmo vs Budget",""])

dados_semana = [
    ("Semana 1", "01/05 – 07/05/2026", 1031.94),
    ("Semana 2", "08/05 – 14/05/2026", 1655.88),
    ("Semana 3", "15/05 – 21/05/2026", 1735.54),
    ("Semana 4", "22/05 – 27/05/2026", 1363.72),
]
total_maio_sem = sum(s[2] for s in dados_semana)

for i, (sem, periodo, valor) in enumerate(dados_semana):
    r = 14 + i
    bg = CINZA_CLARO if i % 2 == 0 else BRANCO
    for col in range(1, 7):
        ws2.cell(row=r, column=col).fill  = fill(bg)
        ws2.cell(row=r, column=col).border = thin_border()
        ws2.cell(row=r, column=col).alignment = center()
        ws2.cell(row=r, column=col).font = FONTE_NORMAL
    ws2.cell(row=r, column=1).value = sem
    ws2.cell(row=r, column=2).value = periodo
    c = ws2.cell(row=r, column=3)
    c.value = valor; c.number_format = BRL_FMT; c.font = FONTE_BOLD
    pct = ws2.cell(row=r, column=4)
    pct.value = valor / 5829.82
    pct.number_format = "0.0%"
    ws2.row_dimensions[r].height = 20

rt2 = 18
c = ws2.cell(row=rt2, column=1, value="TOTAL SEMANAS (maio)")
c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
c.fill = fill(AZUL_MEDIO); c.border = thin_border(); c.alignment = center()
c2 = ws2.cell(row=rt2, column=2, value="01/05 – 27/05/2026")
c2.font = FONTE_BRANCA; c2.fill = fill(AZUL_MEDIO); c2.border = thin_border(); c2.alignment = center()
c3 = ws2.cell(row=rt2, column=3, value=total_maio_sem)
c3.number_format = BRL_FMT
c3.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
c3.fill = fill(AZUL_MEDIO); c3.border = thin_border(); c3.alignment = center()
for col in range(4, 7):
    c = ws2.cell(row=rt2, column=col)
    c.fill = fill(AZUL_MEDIO); c.border = thin_border()
ws2.row_dimensions[rt2].height = 22

# ─────────────────────────────────────────────────────────────────────────────
#  SEÇÃO 3 – INVESTIMENTO POR DIA (Maio/2026)
# ─────────────────────────────────────────────────────────────────────────────
titulo_secao(ws2, 20, 1, "SEÇÃO 3 — INVESTIMENTO POR DIA — Maio/2026", 6, "2E4A7A")
header_row(ws2, 21, ["Data","Semana","Campanha / Resumo do Dia","Meta Ads (Gasto)","Acumulado Mês",""])

# daily data já consolidado por dia (soma de todas as campanhas)
dados_diarios = [
    ("01/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas22/04",               133.04),
    ("02/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas22/04",               137.12),
    ("03/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas22/04",               143.37),
    ("04/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas22/04 + Vendas04/05", 171.65),
    ("05/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas04/05",               144.83),
    ("06/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas04/05",               142.79),
    ("07/05/2026","Sem. 1", "Camp.02 + Camp.01 + Camp.03 + Vendas04/05",               159.14),
    ("08/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site",              218.05),
    ("09/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site",              197.01),
    ("10/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site",              202.99),
    ("11/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site",              228.70),
    ("12/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site",              277.33),
    ("13/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site",              298.78),
    ("14/05/2026","Sem. 2", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site + Tráf.Perfil",233.02),
    ("15/05/2026","Sem. 3", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site + Tráf.Perfil",194.02),
    ("16/05/2026","Sem. 3", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site + Tráf.Perfil",230.21),
    ("17/05/2026","Sem. 3", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site + Tráf.Perfil",213.29),
    ("18/05/2026","Sem. 3", "Camp.02 + Camp.01 + Camp.03 + Tráfego Site + Tráf.Perfil",248.96),
    ("19/05/2026","Sem. 3", "Camp.02+01+03 + Site + Perfil + Venda Ingressos",         274.50),
    ("20/05/2026","Sem. 3", "Camp.02+01+03 + Site + Perfil + Ingressos + Copa",        269.00),
    ("21/05/2026","Sem. 3", "Camp.02+01+03 + Site + Ingressos + Copa + Vídeos",        305.56),
    ("22/05/2026","Sem. 4", "Camp.02+01+03 + Perfil + Ingressos + Copa + Vídeos",      212.56),
    ("23/05/2026","Sem. 4", "Camp.02+01+03 + Perfil + Ingressos + Copa + Vídeos",      233.67),
    ("24/05/2026","Sem. 4", "Camp.02+01+03 + Perfil + Ingressos + Copa + Vídeos",      252.57),
    ("25/05/2026","Sem. 4", "Camp.02+01+03 + Perfil + Ingressos + Copa + Vídeos",      282.11),
    ("26/05/2026","Sem. 4", "Camp.02+01+03 + Perfil + Ingressos + Copa + Vídeos",      247.02),
    ("27/05/2026","Sem. 4", "Camp.02+01+03 + Perfil + Ingressos + Copa + Vídeos",      135.79),
]

acumulado = 0
for i, (data, semana, resumo, valor) in enumerate(dados_diarios):
    r = 22 + i
    acumulado += valor
    bg = CINZA_CLARO if i % 2 == 0 else BRANCO
    for col in range(1, 7):
        ws2.cell(row=r, column=col).fill   = fill(bg)
        ws2.cell(row=r, column=col).border = thin_border()
        ws2.cell(row=r, column=col).font   = FONTE_NORMAL
        ws2.cell(row=r, column=col).alignment = center()
    ws2.cell(row=r, column=1).value = data
    ws2.cell(row=r, column=2).value = semana
    ws2.cell(row=r, column=3).value = resumo
    ws2.cell(row=r, column=3).alignment = left()
    c4 = ws2.cell(row=r, column=4)
    c4.value = valor; c4.number_format = BRL_FMT; c4.font = FONTE_BOLD
    c5 = ws2.cell(row=r, column=5)
    c5.value = acumulado; c5.number_format = BRL_FMT
    ws2.row_dimensions[r].height = 18

# Total final
rt3 = 22 + len(dados_diarios)
for col in range(1, 7):
    c = ws2.cell(row=rt3, column=col)
    c.fill = fill(AZUL_ESCURO); c.border = thin_border(); c.font = FONTE_BRANCA; c.alignment = center()
ws2.cell(row=rt3, column=1).value = "TOTAL MAIO (01–27)"
c4 = ws2.cell(row=rt3, column=4, value=sum(d[3] for d in dados_diarios))
c4.number_format = BRL_FMT
c4.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
c4.fill = fill(AZUL_ESCURO); c4.border = thin_border(); c4.alignment = center()
ws2.row_dimensions[rt3].height = 22

# Larguras colunas aba2
col_widths2 = [18, 10, 55, 22, 22, 10]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

ws2.freeze_panes = "A22"

# ── Salvar ────────────────────────────────────────────────────────────────────
wb.save("/tmp/controle_investimentos_meta.xlsx")
print("Arquivo gerado com sucesso!")
