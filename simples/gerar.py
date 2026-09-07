#!/usr/bin/env python3
"""Gera CONTROLE_FINANCEIRO_SIMPLES.xlsx
Entrada/saida, projetado, saldo de contas, transferencias, cartao com
parcelas automaticas, resumo mensal e resumo anual. Moeda: Guarani.
Todas as formulas usam virgula como separador de argumentos."""
import datetime as dt
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

SAIDA = Path(__file__).resolve().parent.parent / "CONTROLE_FINANCEIRO_SIMPLES.xlsx"
ANO = 2026

HDR = 5
P1 = 6                                  # primeira linha de dados
N_LANC, N_TRF, N_CONTA, N_COMPRA, N_PARC, N_CAT, N_CARTAO = 1000, 300, 20, 60, 36, 40, 8
F_LANC = P1 + N_LANC - 1        # 1005
F_TRF = P1 + N_TRF - 1          # 305
F_CONTA = P1 + N_CONTA - 1      # 25
F_COMPRA = P1 + N_COMPRA - 1    # 65
F_CAT = P1 + N_CAT - 1          # 45
F_CARTAO = P1 + N_CARTAO - 1    # 13

GS = '"Gs. "#,##0;-"Gs. "#,##0;"Gs. "0'
GS0 = '"Gs. "#,##0;-"Gs. "#,##0;""'
NUM0 = '#,##0;-#,##0;""'
DATA = "DD/MM/YYYY"

AZUL, AZUL_CLARO, CINZA = "1F3864", "D9E2F3", "F2F2F2"
VERDE, VERMELHO, AMARELO, BRANCO = "E2EFDA", "FCE4E4", "FFF2CC", "FFFFFF"

f_titulo = Font(size=16, bold=True, color=AZUL)
f_sub = Font(size=10, italic=True, color="595959")
f_hdr = Font(size=11, bold=True, color=BRANCO)
f_b = Font(size=11, bold=True)
f_n = Font(size=11)
f_auto = Font(size=11, color="7F7F7F", italic=True)
f_kpi = Font(size=14, bold=True, color=AZUL)
f_sec = Font(size=12, bold=True, color=AZUL)

fill_hdr = PatternFill("solid", fgColor=AZUL)
fill_az = PatternFill("solid", fgColor=AZUL_CLARO)
fill_cz = PatternFill("solid", fgColor=CINZA)
fill_vd = PatternFill("solid", fgColor=VERDE)
fill_vm = PatternFill("solid", fgColor=VERMELHO)
fill_am = PatternFill("solid", fgColor=AMARELO)

fina = Side(style="thin", color="BFBFBF")
bd = Border(left=fina, right=fina, top=fina, bottom=fina)
ctr = Alignment(horizontal="center", vertical="center")
esq = Alignment(horizontal="left", vertical="center")
wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

CAT_ENTRADA = ["Salário", "Pró-labore", "Aluguel recebido", "Venda", "Rendimento",
               "Empréstimo recebido", "Reembolso", "Outras entradas"]
CAT_SAIDA = ["Moradia", "Alimentação", "Transporte", "Saúde", "Educação", "Lazer",
             "Vestuário", "Serviços (luz, água, internet)", "Impostos e taxas",
             "Empréstimo / Financiamento", "Juros e tarifas", "Família",
             "Manutenção", "Outras saídas"]
TIPO_CONTA = ["Conta corrente", "Poupança", "Dinheiro / Caixa", "Carteira digital", "Outra"]
CARTOES = ["Cartão 1", "Cartão 2", "Cartão 3", "Cartão 4"]
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

wb = openpyxl.Workbook()


def cel(ws, r, c, v=None, *, fonte=None, fill=None, fmt=None, al=None, borda=True):
    x = ws.cell(row=r, column=c)
    x.value = v
    if fonte: x.font = fonte
    if fill: x.fill = fill
    if fmt: x.number_format = fmt
    if al: x.alignment = al
    if borda: x.border = bd
    return x


def cabecalho(ws, titulo, sub, largura=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=largura)
    cel(ws, 1, 1, titulo, fonte=f_titulo, al=esq, borda=False)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=largura)
    cel(ws, 2, 1, sub, fonte=f_sub, al=esq, borda=False)
    ws.row_dimensions[1].height = 24


def linha_hdr(ws, linha, rotulos, col0=1, altura=30):
    for i, t in enumerate(rotulos):
        cel(ws, linha, col0 + i, t, fonte=f_hdr, fill=fill_hdr, al=ctr)
    ws.row_dimensions[linha].height = altura


def secao(ws, r, texto, ate):
    cel(ws, r, 1, texto, fonte=f_sec, fill=fill_az, al=esq)
    for c in range(2, ate + 1):
        cel(ws, r, c, None, fill=fill_az)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ate)


def larguras(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


def kpi(ws, r, c, rot, formula, fill, largura=2, fmt=GS, fonte=None):
    cel(ws, r, c, rot, fonte=f_b, fill=fill, al=ctr)
    cel(ws, r + 1, c, formula, fonte=fonte or f_kpi, fill=fill, fmt=fmt, al=ctr)
    for j in range(1, largura):
        cel(ws, r, c + j, None, fill=fill); cel(ws, r + 1, c + j, None, fill=fill)
    if largura > 1:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + largura - 1)
        ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + largura - 1)
    ws.row_dimensions[r + 1].height = 24


def dv_lista(ws, faixa, origem, titulo, msg):
    d = DataValidation(type="list", formula1="=" + origem, allow_blank=True, showDropDown=False)
    d.errorTitle, d.error = titulo, msg
    ws.add_data_validation(d); d.add(faixa)


def dv_num(ws, faixa, tipo, op, f1, titulo, msg, f2=None):
    d = DataValidation(type=tipo, operator=op, formula1=f1, formula2=f2, allow_blank=True)
    d.errorTitle, d.error = titulo, msg
    ws.add_data_validation(d); d.add(faixa)


# =============================================================== INSTRUÇÕES
ws = wb.active
ws.title = "INSTRUÇÕES"
ws.sheet_properties.tabColor = AZUL
larguras(ws, {"A": 3, "B": 104})
cabecalho(ws, "CONTROLE FINANCEIRO — CONTAS, CARTÕES E TRANSFERÊNCIAS",
          "Moeda: Guaraní (Gs.)  ·  Ano-base: %d" % ANO, 2)

TEXTO = [
    ("t", "AS 4 REGRAS QUE FAZEM A PLANILHA FECHAR"),
    ("n", "1) Dinheiro que ENTRA ou SAI do seu patrimônio → aba LANÇAMENTOS, sempre dizendo de qual CONTA."),
    ("n", "2) Dinheiro que só MUDA DE LUGAR (conta para conta, conta para cartão) → aba TRANSFERÊNCIAS."),
    ("n", "   Transferência não é entrada nem saída. Ela só mexe no saldo dos dois lados."),
    ("n", "3) Compra no cartão de crédito → aba CARTÕES, uma linha por compra. Nunca em LANÇAMENTOS."),
    ("n", "4) Pagamento da fatura do cartão → aba TRANSFERÊNCIAS, De = sua conta, Para = o cartão."),
    ("n", "   Isso tira o dinheiro da conta e abate a dívida do cartão. NÃO é despesa: a parcela já foi contada."),
    ("a", "Seguindo essas 4 regras é impossível contar o mesmo dinheiro duas vezes."),
    ("", ""),
    ("t", "COMO USAR, PASSO A PASSO"),
    ("n", "1º  Abra a aba CONTAS e cadastre suas contas: nome, banco, tipo e o saldo que elas têm hoje."),
    ("n", "    Cadastre também uma conta chamada \"Dinheiro\" para o que você carrega no bolso."),
    ("n", "2º  Abra a aba LISTAS e escreva o nome dos seus cartões e o limite de cada um."),
    ("n", "3º  Comece a digitar em LANÇAMENTOS. Data, Tipo, Status, Categoria, Descrição, Conta, Valor."),
    ("n", "4º  Compras parceladas no cartão vão em CARTÕES. Pagamentos de fatura vão em TRANSFERÊNCIAS."),
    ("n", "5º  Veja o saldo de cada conta na aba CONTAS, o mês em RESUMO MENSAL e o ano em RESUMO ANUAL."),
    ("", ""),
    ("t", "REALIZADO x PROJETADO"),
    ("n", "REALIZADO = já aconteceu, o dinheiro já entrou ou já saiu. Entra no SALDO ATUAL."),
    ("n", "PROJETADO = ainda vai acontecer. Não entra no saldo atual, entra no SALDO PREVISTO."),
    ("n", "Serve para lançar salário futuro, aluguel do mês que vem, conta de luz que ainda vai chegar."),
    ("", ""),
    ("t", "OS DOIS NÚMEROS QUE VOCÊ VAI OLHAR TODO DIA"),
    ("a", "SALDO ATUAL (aba CONTAS)  =  quanto você TEM agora, somando todas as contas."),
    ("a", "RESULTADO DO MÊS (aba RESUMO MENSAL)  =  Entradas − Saídas − Parcelas de cartão do mês."),
    ("n", "São coisas diferentes e as duas estão certas. O saldo é caixa: o que já passou pela conta."),
    ("n", "O resultado é o mês: mostra a parcela do cartão no mês em que ela vence, mesmo que a fatura"),
    ("n", "só saia da conta depois. É assim que se enxerga se o mês fechou no azul ou no vermelho."),
    ("", ""),
    ("t", "O QUE VOCÊ NUNCA DEVE FAZER"),
    ("x", "Não EXCLUA linhas das tabelas. Para apagar um lançamento, selecione a linha e aperte DELETE"),
    ("x", "   (limpar o conteúdo). Excluir a linha apaga as fórmulas das colunas automáticas."),
    ("x", "Não digite nada nas colunas cinza escritas (auto). Elas se calculam sozinhas."),
    ("x", "Não mexa na aba PARCELAS. Ela é o motor que espalha as parcelas pelos meses."),
    ("x", "Não lance a fatura do cartão como SAÍDA. Isso conta o gasto duas vezes."),
    ("", ""),
    ("t", "AS ABAS"),
    ("n", "CONTAS         ·  suas contas e o saldo de cada uma, atual e previsto."),
    ("n", "LANÇAMENTOS    ·  tudo que entra e sai de verdade do seu bolso ou da conta."),
    ("n", "TRANSFERÊNCIAS ·  dinheiro mudando de lugar. Inclui o pagamento da fatura do cartão."),
    ("n", "CARTÕES        ·  compras no crédito, parcelamento automático, dívida e limite disponível."),
    ("n", "RESUMO MENSAL  ·  um mês por vez, detalhe por categoria, resultado e conferência."),
    ("n", "RESUMO ANUAL   ·  os 12 meses lado a lado, saldo acumulado e categorias mês a mês."),
    ("n", "EXEMPLO        ·  modelos preenchidos de cada caso. Pode apagar a aba quando quiser."),
    ("n", "LISTAS         ·  categorias, tipos de conta, cartões e limites."),
    ("n", "PARCELAS       ·  aba de apoio, automática. Não mexa."),
    ("", ""),
    ("t", "COMO O SALDO DE CADA CONTA É CALCULADO"),
    ("a", "Saldo atual = Saldo inicial + Entradas − Saídas + Transf. recebidas − Transf. enviadas"),
    ("n", "Só entram os REALIZADOS. O Saldo previsto soma por cima os PROJETADOS."),
    ("", ""),
    ("t", "COMO A DÍVIDA DO CARTÃO É CALCULADA"),
    ("a", "Dívida atual = Total comprado − Pagamentos que você lançou em TRANSFERÊNCIAS"),
    ("n", "Se você não quiser lançar os pagamentos de fatura, use a coluna \"Parcelas a vencer\":"),
    ("n", "ela funciona sozinha, só pelo calendário das parcelas, sem depender de nenhum lançamento."),
]
r = 4
for tipo, txt in TEXTO:
    c = cel(ws, r, 2, txt, borda=False, al=wrap)
    if tipo == "t":
        c.font = f_sec; c.fill = fill_az
    elif tipo == "a":
        c.font = f_b; c.fill = fill_am
    elif tipo == "x":
        c.font = Font(size=11, color="9C0006"); c.fill = fill_vm
    else:
        c.font = f_n
    ws.row_dimensions[r].height = 17
    r += 1
ws.sheet_view.showGridLines = False

# ==================================================================== LISTAS
lst = wb.create_sheet("LISTAS")
lst.sheet_properties.tabColor = "808080"
cabecalho(lst, "LISTAS", "Ajuste aqui suas categorias, os nomes dos cartões e os limites.", 8)
larguras(lst, {"A": 34, "B": 12, "C": 2, "D": 20, "E": 18, "F": 2, "G": 20, "H": 2,
               "I": 13, "J": 13, "K": 2, "L": 13, "M": 2, "N": 26})
linha_hdr(lst, HDR, ["Categoria", "Tipo"], 1)
linha_hdr(lst, HDR, ["Cartão de crédito", "Limite (Gs.)"], 4)
linha_hdr(lst, HDR, ["Tipo de conta"], 7)
linha_hdr(lst, HDR, ["Tipo", "Status"], 9)
linha_hdr(lst, HDR, ["Mês"], 12)
linha_hdr(lst, HDR, ["Contas e cartões (auto)"], 14)

CATS = [(c, "ENTRADA") for c in CAT_ENTRADA] + [(c, "SAÍDA") for c in CAT_SAIDA]
for i in range(N_CAT):
    r = P1 + i
    v = CATS[i] if i < len(CATS) else ("", "")
    cel(lst, r, 1, v[0] or None, fonte=f_n, al=esq)
    cel(lst, r, 2, v[1] or None, fonte=f_n, al=ctr,
        fill=fill_vd if v[1] == "ENTRADA" else (fill_vm if v[1] == "SAÍDA" else None))
for i in range(N_CARTAO):
    cel(lst, P1 + i, 4, CARTOES[i] if i < len(CARTOES) else None, fonte=f_n, al=esq)
    cel(lst, P1 + i, 5, None, fonte=f_n, fmt=GS0, al=ctr)
for i in range(len(TIPO_CONTA)):
    cel(lst, P1 + i, 7, TIPO_CONTA[i], fonte=f_n, al=esq)
F_TCONTA = P1 + len(TIPO_CONTA) - 1
for i, v in enumerate(["ENTRADA", "SAÍDA"]):
    cel(lst, P1 + i, 9, v, fonte=f_n, al=ctr)
for i, v in enumerate(["REALIZADO", "PROJETADO"]):
    cel(lst, P1 + i, 10, v, fonte=f_n, al=ctr)
for i, m in enumerate(MESES):
    cel(lst, P1 + i, 12, m, fonte=f_n, al=esq)
# coluna N: contas (das CONTAS) seguidas dos cartoes -> destino de transferencia
for i in range(N_CONTA):
    cel(lst, P1 + i, 14, f'=IF(CONTAS!$A{P1+i}="","",CONTAS!$A{P1+i})',
        fonte=f_auto, fill=fill_cz, al=esq)
for i in range(N_CARTAO):
    cel(lst, P1 + N_CONTA + i, 14, f'=IF($D{P1+i}="","",$D{P1+i})',
        fonte=f_auto, fill=fill_cz, al=esq)
F_TUDO = P1 + N_CONTA + N_CARTAO - 1
dv_num(lst, f"E{P1}:E{F_CARTAO}", "decimal", "greaterThanOrEqual", "0",
       "Limite", "O limite do cartão não pode ser negativo.")
cel(lst, F_CARTAO + 2, 4, "Troque \"Cartão 1\" pelo nome real do seu cartão e coloque o limite.",
    fonte=f_sub, borda=False)
lst.merge_cells(start_row=F_CARTAO + 2, start_column=4, end_row=F_CARTAO + 2, end_column=7)
lst.sheet_view.showGridLines = False

R_CAT = f"LISTAS!$A${P1}:$A${F_CAT}"
R_CARTAO = f"LISTAS!$D${P1}:$D${F_CARTAO}"
R_TCONTA = f"LISTAS!$G${P1}:$G${F_TCONTA}"
R_TIPO = f"LISTAS!$I${P1}:$I${P1+1}"
R_STATUS = f"LISTAS!$J${P1}:$J${P1+1}"
R_MESES = f"LISTAS!$L${P1}:$L${P1+11}"
R_TUDO = f"LISTAS!$N${P1}:$N${F_TUDO}"
R_CONTA = f"CONTAS!$A${P1}:$A${F_CONTA}"

# ==================================================================== CONTAS
con = wb.create_sheet("CONTAS")
con.sheet_properties.tabColor = "375623"
cabecalho(con, "CONTAS E SALDOS",
          "Cadastre cada conta e o saldo que ela tem hoje. Daí em diante a planilha mantém o saldo sozinha.", 14)
larguras(con, {"A": 24, "B": 20, "C": 18, "D": 17, "E": 15, "F": 15, "G": 15,
               "H": 15, "I": 15, "J": 18, "K": 15, "L": 15, "M": 15, "N": 18})

kpi(con, 3, 1, "SALDO ATUAL (todas as contas)", f"=SUM($J${P1}:$J${F_CONTA})", fill_vd, 2)
kpi(con, 3, 3, "SALDO PREVISTO (com projetados)", f"=SUM($N${P1}:$N${F_CONTA})", fill_am, 2)
kpi(con, 3, 5, "DÍVIDA NOS CARTÕES", "=CARTÕES!$G$4", fill_vm, 2)
kpi(con, 3, 7, "POSIÇÃO LÍQUIDA (contas − cartões)", "=$A$4-$E$4", fill_az, 2)

linha_hdr(con, HDR, ["Nome da conta", "Banco / Instituição", "Tipo", "Saldo inicial (Gs.)",
                     "Saldo inicial em", "Entradas (auto)", "Saídas (auto)",
                     "Transf. recebidas (auto)", "Transf. enviadas (auto)", "SALDO ATUAL (auto)",
                     "Entradas prev. (auto)", "Saídas prev. (auto)", "Transf. prev. (auto)",
                     "SALDO PREVISTO (auto)"], altura=42)
con.freeze_panes = "A6"

LAN, TRF = "'LANÇAMENTOS'!", "'TRANSFERÊNCIAS'!"
RL_IDX = f"{LAN}$I${P1}:$I${F_LANC}"
RL_TIPO = f"{LAN}$B${P1}:$B${F_LANC}"
RL_STAT = f"{LAN}$C${P1}:$C${F_LANC}"
RL_CAT = f"{LAN}$D${P1}:$D${F_LANC}"
RL_CONTA = f"{LAN}$F${P1}:$F${F_LANC}"
RL_VAL = f"{LAN}$G${P1}:$G${F_LANC}"
RT_IDX = f"{TRF}$H${P1}:$H${F_TRF}"
RT_STAT = f"{TRF}$B${P1}:$B${F_TRF}"
RT_DE = f"{TRF}$C${P1}:$C${F_TRF}"
RT_PARA = f"{TRF}$D${P1}:$D${F_TRF}"
RT_VAL = f"{TRF}$F${P1}:$F${F_TRF}"

for i in range(N_CONTA):
    r = P1 + i
    a = f"$A{r}"
    cel(con, r, 1, None, fonte=f_n, al=esq)
    cel(con, r, 2, None, fonte=f_n, al=esq)
    cel(con, r, 3, None, fonte=f_n, al=ctr)
    cel(con, r, 4, None, fonte=f_n, fmt=GS0, al=ctr)
    cel(con, r, 5, None, fonte=f_n, fmt=DATA, al=ctr)
    cel(con, r, 6, f'=IF({a}="","",SUMIFS({RL_VAL},{RL_CONTA},{a},{RL_TIPO},"ENTRADA",{RL_STAT},"REALIZADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 7, f'=IF({a}="","",SUMIFS({RL_VAL},{RL_CONTA},{a},{RL_TIPO},"SAÍDA",{RL_STAT},"REALIZADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 8, f'=IF({a}="","",SUMIFS({RT_VAL},{RT_PARA},{a},{RT_STAT},"REALIZADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 9, f'=IF({a}="","",SUMIFS({RT_VAL},{RT_DE},{a},{RT_STAT},"REALIZADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 10, f'=IF({a}="","",$D{r}+$F{r}-$G{r}+$H{r}-$I{r})',
        fonte=f_b, fill=fill_vd, fmt=GS0, al=ctr)
    cel(con, r, 11, f'=IF({a}="","",SUMIFS({RL_VAL},{RL_CONTA},{a},{RL_TIPO},"ENTRADA",{RL_STAT},"PROJETADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 12, f'=IF({a}="","",SUMIFS({RL_VAL},{RL_CONTA},{a},{RL_TIPO},"SAÍDA",{RL_STAT},"PROJETADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 13, f'=IF({a}="","",SUMIFS({RT_VAL},{RT_PARA},{a},{RT_STAT},"PROJETADO")'
                    f'-SUMIFS({RT_VAL},{RT_DE},{a},{RT_STAT},"PROJETADO"))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(con, r, 14, f'=IF({a}="","",$J{r}+$K{r}-$L{r}+$M{r})',
        fonte=f_b, fill=fill_am, fmt=GS0, al=ctr)

rt = F_CONTA + 1
cel(con, rt, 1, "TOTAL", fonte=f_b, fill=fill_az, al=esq)
for c in range(2, 15):
    if c in (3, 5):
        cel(con, rt, c, None, fill=fill_az)
    else:
        cel(con, rt, c, f"=SUM({L(c)}{P1}:{L(c)}{F_CONTA})", fonte=f_b, fill=fill_az, fmt=GS0, al=ctr)

dv_lista(con, f"C{P1}:C{F_CONTA}", R_TCONTA, "Tipo de conta",
         "Escolha um tipo da lista. Você pode acrescentar tipos na aba LISTAS.")
dv_num(con, f"D{P1}:D{F_CONTA}", "decimal", "greaterThanOrEqual", "-999999999999",
       "Saldo inicial", "Digite um número. Pode ser negativo se a conta está no vermelho.")
con.conditional_formatting.add(f"J{P1}:J{rt}", CellIsRule(operator="lessThan", formula=["0"],
                                                          fill=fill_vm, font=Font(bold=True, color="9C0006")))
con.conditional_formatting.add(f"N{P1}:N{rt}", CellIsRule(operator="lessThan", formula=["0"],
                                                          fill=fill_vm, font=Font(bold=True, color="9C0006")))
con[f"A{P1}"].comment = Comment(
    "Escreva um nome curto e único para cada conta.\n"
    "Esse nome é o que aparece nas listas de LANÇAMENTOS e TRANSFERÊNCIAS.\n"
    "Crie também uma conta chamada Dinheiro para o que você carrega no bolso.", "Planilha")
cel(con, rt + 2, 1, "O saldo só conta os REALIZADOS. O saldo previsto soma por cima os PROJETADOS.",
    fonte=f_sub, borda=False, al=esq)
con.sheet_view.showGridLines = False

# =============================================================== LANÇAMENTOS
lan = wb.create_sheet("LANÇAMENTOS")
lan.sheet_properties.tabColor = "2E75B6"
cabecalho(lan, "LANÇAMENTOS",
          "Dinheiro que entra e sai de verdade. Compra no cartão vai em CARTÕES. "
          "Dinheiro trocando de conta vai em TRANSFERÊNCIAS.", 9)
larguras(lan, {"A": 12, "B": 12, "C": 13, "D": 28, "E": 34, "F": 22, "G": 16, "H": 11, "I": 11})

kpi(lan, 3, 1, "Total ENTRADAS", f'=SUMIF($B${P1}:$B${F_LANC},"ENTRADA",$G${P1}:$G${F_LANC})',
    fill_vd, 1)
kpi(lan, 3, 2, "Total SAÍDAS", f'=SUMIF($B${P1}:$B${F_LANC},"SAÍDA",$G${P1}:$G${F_LANC})',
    fill_vm, 1)
kpi(lan, 3, 3, "Diferença", "=$A$4-$B$4", fill_az, 1)
kpi(lan, 3, 4, "Lançamentos digitados", f"=COUNTA($A${P1}:$A${F_LANC})", fill_cz, 1, fmt="#,##0")
kpi(lan, 3, 5, "Espaço livre", f"={N_LANC}-$D$4", fill_cz, 1, fmt="#,##0")

linha_hdr(lan, HDR, ["Data", "Tipo", "Status", "Categoria", "Descrição",
                     "Conta", "Valor (Gs.)", "Mês (auto)", "Índice (auto)"])
lan.freeze_panes = "A6"
for r in range(P1, F_LANC + 1):
    cel(lan, r, 1, None, fonte=f_n, fmt=DATA, al=ctr)
    for c in (2, 3):
        cel(lan, r, c, None, fonte=f_n, al=ctr)
    for c in (4, 5):
        cel(lan, r, c, None, fonte=f_n, al=esq)
    cel(lan, r, 6, None, fonte=f_n, al=ctr)
    cel(lan, r, 7, None, fonte=f_n, fmt=GS, al=ctr)
    cel(lan, r, 8, f'=IF($A{r}="","",MONTH($A{r})&"/"&YEAR($A{r}))',
        fonte=f_auto, fill=fill_cz, al=ctr)
    cel(lan, r, 9, f'=IF($A{r}="",0,YEAR($A{r})*12+MONTH($A{r}))',
        fonte=f_auto, fill=fill_cz, fmt=NUM0, al=ctr)

dv_lista(lan, f"B{P1}:B{F_LANC}", R_TIPO, "Tipo", "Escolha ENTRADA ou SAÍDA na lista.")
dv_lista(lan, f"C{P1}:C{F_LANC}", R_STATUS, "Status",
         "Escolha REALIZADO (já aconteceu) ou PROJETADO (ainda vai acontecer).")
dv_lista(lan, f"D{P1}:D{F_LANC}", R_CAT, "Categoria",
         "Categoria não está na lista. Cadastre na aba LISTAS.")
dv_lista(lan, f"F{P1}:F{F_LANC}", R_CONTA, "Conta",
         "Conta não está cadastrada. Cadastre na aba CONTAS.")
dv_num(lan, f"G{P1}:G{F_LANC}", "decimal", "greaterThan", "0", "Valor",
       "O valor tem que ser maior que zero. Se é saída, escolha Tipo = SAÍDA.")
lan.conditional_formatting.add(f"A{P1}:I{F_LANC}",
                               FormulaRule(formula=[f'$C{P1}="PROJETADO"'], fill=fill_am))
lan[f"C{P1}"].comment = Comment(
    "REALIZADO = já entrou ou já saiu. Conta no saldo.\n"
    "PROJETADO = ainda vai acontecer. Só conta no saldo previsto.", "Planilha")
lan[f"F{P1}"].comment = Comment(
    "De qual conta esse dinheiro saiu ou em qual conta ele entrou.\n"
    "Cadastre suas contas na aba CONTAS.", "Planilha")
lan.sheet_view.showGridLines = False

# ============================================================ TRANSFERÊNCIAS
trf = wb.create_sheet("TRANSFERÊNCIAS")
trf.sheet_properties.tabColor = "7030A0"
cabecalho(trf, "TRANSFERÊNCIAS",
          "Dinheiro mudando de lugar. Não é entrada nem saída — só muda o saldo dos dois lados. "
          "É aqui que entra o pagamento da fatura do cartão.", 9)
larguras(trf, {"A": 12, "B": 13, "C": 24, "D": 24, "E": 34, "F": 16, "G": 11, "H": 11, "I": 30})

kpi(trf, 3, 1, "Total transferido (realizado)",
    f'=SUMIF($B${P1}:$B${F_TRF},"REALIZADO",$F${P1}:$F${F_TRF})', fill_az, 2)
kpi(trf, 3, 3, "Pagamentos de fatura de cartão",
    f'=SUMPRODUCT(($B${P1}:$B${F_TRF}="REALIZADO")*'
    f'(COUNTIF({R_CARTAO},$D${P1}:$D${F_TRF})>0)*$F${P1}:$F${F_TRF})', fill_vm, 2)
kpi(trf, 3, 5, "Transferências digitadas", f"=COUNTA($A${P1}:$A${F_TRF})", fill_cz, 1, fmt="#,##0")
kpi(trf, 3, 6, "Erros a corrigir", f'=COUNTIF($I${P1}:$I${F_TRF},"ERRO*")', fill_am, 1, fmt="#,##0")

linha_hdr(trf, HDR, ["Data", "Status", "De (sai de)", "Para (entra em)", "Descrição",
                     "Valor (Gs.)", "Mês (auto)", "Índice (auto)", "Conferência (auto)"])
trf.freeze_panes = "A6"
for r in range(P1, F_TRF + 1):
    cel(trf, r, 1, None, fonte=f_n, fmt=DATA, al=ctr)
    cel(trf, r, 2, None, fonte=f_n, al=ctr)
    for c in (3, 4):
        cel(trf, r, c, None, fonte=f_n, al=ctr)
    cel(trf, r, 5, None, fonte=f_n, al=esq)
    cel(trf, r, 6, None, fonte=f_n, fmt=GS, al=ctr)
    cel(trf, r, 7, f'=IF($A{r}="","",MONTH($A{r})&"/"&YEAR($A{r}))',
        fonte=f_auto, fill=fill_cz, al=ctr)
    cel(trf, r, 8, f'=IF($A{r}="",0,YEAR($A{r})*12+MONTH($A{r}))',
        fonte=f_auto, fill=fill_cz, fmt=NUM0, al=ctr)
    cel(trf, r, 9,
        f'=IF($A{r}="","",'
        f'IF($C{r}="","ERRO: falta a conta de origem",'
        f'IF($D{r}="","ERRO: falta o destino",'
        f'IF($C{r}=$D{r},"ERRO: origem e destino iguais",'
        f'IF($F{r}="","ERRO: falta o valor",'
        f'IF(COUNTIF({R_CARTAO},$D{r})>0,"Pagamento de fatura de cartão","OK"))))))',
        fonte=f_auto, fill=fill_cz, al=esq)

dv_lista(trf, f"B{P1}:B{F_TRF}", R_STATUS, "Status",
         "Escolha REALIZADO (já aconteceu) ou PROJETADO (ainda vai acontecer).")
dv_lista(trf, f"C{P1}:C{F_TRF}", R_CONTA, "Conta de origem",
         "A origem tem que ser uma conta cadastrada na aba CONTAS.")
dv_lista(trf, f"D{P1}:D{F_TRF}", R_TUDO, "Destino",
         "O destino tem que ser uma conta (aba CONTAS) ou um cartão (aba LISTAS).")
dv_num(trf, f"F{P1}:F{F_TRF}", "decimal", "greaterThan", "0", "Valor",
       "O valor transferido tem que ser maior que zero.")
trf.conditional_formatting.add(f"A{P1}:I{F_TRF}",
                               FormulaRule(formula=[f'LEFT($I{P1},4)="ERRO"'], fill=fill_vm,
                                           font=Font(color="9C0006", bold=True)))
trf.conditional_formatting.add(f"A{P1}:I{F_TRF}",
                               FormulaRule(formula=[f'$B{P1}="PROJETADO"'], fill=fill_am))
trf[f"D{P1}"].comment = Comment(
    "Pode ser outra conta sua ou um CARTÃO.\n"
    "Escolhendo um cartão, a planilha entende que é pagamento de fatura\n"
    "e abate a dívida daquele cartão.", "Planilha")
trf.sheet_view.showGridLines = False

# =================================================================== CARTÕES
car = wb.create_sheet("CARTÕES")
car.sheet_properties.tabColor = "C00000"
cabecalho(car, "CARTÕES DE CRÉDITO — COMPRAS E PARCELAS",
          "Digite a compra uma vez. A planilha divide nas parcelas e joga cada uma no mês certo. "
          "O pagamento da fatura vai na aba TRANSFERÊNCIAS.", 13)
larguras(car, {"A": 13, "B": 16, "C": 30, "D": 24, "E": 16, "F": 10, "G": 14,
               "H": 13, "I": 16, "J": 13, "K": 11, "L": 17, "M": 16})

kpi(car, 3, 1, "Total comprado", f"=SUM($E${P1}:$E${F_COMPRA})", fill_cz, 2)
kpi(car, 3, 3, "Parcelas a vencer", f"=SUM($M${P1}:$M${F_COMPRA})", fill_am, 2)
kpi(car, 3, 5, "Pago (transferências)",
    f'=SUMPRODUCT((COUNTIF({R_CARTAO},{RT_PARA})>0)*({RT_STAT}="REALIZADO")*{RT_VAL})', fill_vd, 2)
kpi(car, 3, 7, "DÍVIDA ATUAL", "=$A$4-$E$4", fill_vm, 2)

linha_hdr(car, HDR, ["Data da compra", "Cartão", "Descrição", "Categoria",
                     "Valor total (Gs.)", "Nº de parcelas", "1ª parcela (opcional)",
                     "Início (auto)", "Valor da parcela (auto)", "Última parcela (auto)",
                     "Índice (auto)", "Parcelas vencidas (auto)", "Parcelas a vencer (auto)"],
           altura=42)
car.freeze_panes = "A6"
HOJE_IDX = "YEAR(TODAY())*12+MONTH(TODAY())"
for r in range(P1, F_COMPRA + 1):
    cel(car, r, 1, None, fonte=f_n, fmt=DATA, al=ctr)
    cel(car, r, 2, None, fonte=f_n, al=ctr)
    for c in (3, 4):
        cel(car, r, c, None, fonte=f_n, al=esq)
    cel(car, r, 5, None, fonte=f_n, fmt=GS, al=ctr)
    cel(car, r, 6, None, fonte=f_n, fmt="0", al=ctr)
    cel(car, r, 7, None, fonte=f_n, fmt=DATA, al=ctr)
    cel(car, r, 8,
        f'=IF($A{r}="","",IF($G{r}<>"",$G{r},'
        f'DATE(YEAR($A{r})+IF(MONTH($A{r})=12,1,0),IF(MONTH($A{r})=12,1,MONTH($A{r})+1),1)))',
        fonte=f_auto, fill=fill_cz, fmt=DATA, al=ctr)
    cel(car, r, 9, f'=IF(OR($E{r}="",$F{r}=""),"",ROUND($E{r}/$F{r},0))',
        fonte=f_auto, fill=fill_cz, fmt=GS, al=ctr)
    cel(car, r, 11, f'=IF($H{r}="",0,YEAR($H{r})*12+MONTH($H{r}))',
        fonte=f_auto, fill=fill_cz, fmt=NUM0, al=ctr)
    cel(car, r, 10,
        f'=IF($K{r}=0,"",($K{r}+$F{r}-1-INT(($K{r}+$F{r}-2)/12)*12)&"/"&INT(($K{r}+$F{r}-2)/12))',
        fonte=f_auto, fill=fill_cz, al=ctr)
    cel(car, r, 12,
        f'=IF($K{r}=0,"",SUMPRODUCT((PARCELAS!$C{r}:$AL{r}>0)*'
        f'(PARCELAS!$C{r}:$AL{r}<={HOJE_IDX})*PARCELAS!$AM{r}:$BV{r}))',
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    cel(car, r, 13, f'=IF($K{r}=0,"",$E{r}-$L{r})', fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)

dv_lista(car, f"B{P1}:B{F_COMPRA}", R_CARTAO, "Cartão",
         "Cartão não está na lista. Cadastre o nome dele na aba LISTAS.")
dv_lista(car, f"D{P1}:D{F_COMPRA}", R_CAT, "Categoria",
         "Categoria não está na lista. Cadastre na aba LISTAS.")
dv_num(car, f"E{P1}:E{F_COMPRA}", "decimal", "greaterThan", "0", "Valor",
       "O valor total da compra tem que ser maior que zero.")
dv_num(car, f"F{P1}:F{F_COMPRA}", "whole", "between", "1", "Parcelas",
       f"O número de parcelas tem que ser de 1 a {N_PARC}.", f2=str(N_PARC))

r0 = F_COMPRA + 3
secao(car, r0, "SALDO POR CARTÃO", 8)
linha_hdr(car, r0 + 1, ["Cartão", "Limite (Gs.)", "Total comprado", "Parcelas vencidas",
                        "Parcelas a vencer", "Pago (transferências)", "Dívida atual",
                        "Limite disponível"], altura=42)
for i in range(N_CARTAO):
    r = r0 + 2 + i
    lr = P1 + i
    a = f"$A{r}"
    cel(car, r, 1, f'=IF(LISTAS!$D{lr}="","",LISTAS!$D{lr})', fonte=f_b, al=esq)
    cel(car, r, 2, f'=IF({a}="","",LISTAS!$E{lr})', fonte=f_n, fmt=GS0, al=ctr)
    for c, col in ((3, "E"), (4, "L"), (5, "M")):
        cel(car, r, c, f'=IF({a}="","",SUMIF($B${P1}:$B${F_COMPRA},{a},${col}${P1}:${col}${F_COMPRA}))',
            fonte=f_n, fmt=GS0, al=ctr)
    cel(car, r, 6, f'=IF({a}="","",SUMIFS({RT_VAL},{RT_PARA},{a},{RT_STAT},"REALIZADO"))',
        fonte=f_n, fmt=GS0, al=ctr)
    cel(car, r, 7, f'=IF({a}="","",$C{r}-$F{r})', fonte=f_b, fill=fill_vm, fmt=GS0, al=ctr)
    cel(car, r, 8, f'=IF(OR({a}="",$B{r}=""),"",$B{r}-$G{r})', fonte=f_b, fill=fill_vd, fmt=GS0, al=ctr)
rtc = r0 + 2 + N_CARTAO
cel(car, rtc, 1, "TOTAL", fonte=f_b, fill=fill_az, al=esq)
for c in range(2, 9):
    cel(car, rtc, c, f"=SUM({L(c)}{r0+2}:{L(c)}{rtc-1})", fonte=f_b, fill=fill_az, fmt=GS0, al=ctr)
car.conditional_formatting.add(f"H{r0+2}:H{rtc}",
                               CellIsRule(operator="lessThan", formula=["0"], fill=fill_vm,
                                          font=Font(bold=True, color="9C0006")))
cel(car, rtc + 2, 1,
    "\"Parcelas a vencer\" funciona sozinha, só pelo calendário. \"Dívida atual\" depende de você "
    "lançar os pagamentos de fatura em TRANSFERÊNCIAS.", fonte=f_sub, borda=False, al=esq)
car.merge_cells(start_row=rtc + 2, start_column=1, end_row=rtc + 2, end_column=8)
car.sheet_view.showGridLines = False

# ================================================================== PARCELAS
par = wb.create_sheet("PARCELAS")
par.sheet_properties.tabColor = "A6A6A6"
cabecalho(par, "PARCELAS (automático) — NÃO MEXA NESTA ABA",
          "Cada linha é uma compra da aba CARTÕES. As colunas espalham as parcelas pelos meses.", 10)
larguras(par, {"A": 30, "B": 22})
linha_hdr(par, HDR, ["Compra", "Valor no mês do RESUMO MENSAL"])
for k in range(1, N_PARC + 1):
    cel(par, HDR, 2 + k, f"Mês {k}", fonte=f_hdr, fill=fill_hdr, al=ctr)
    cel(par, HDR, 38 + k, f"Valor {k}", fonte=f_hdr, fill=fill_hdr, al=ctr)
    par.column_dimensions[L(2 + k)].width = 8
    par.column_dimensions[L(38 + k)].width = 12
par.freeze_panes = "C6"
for r in range(P1, F_COMPRA + 1):
    cel(par, r, 1, f'=IF(CARTÕES!$C{r}="","",CARTÕES!$C{r})', fonte=f_auto, al=esq)
    cel(par, r, 2, f"=SUMPRODUCT((C{r}:AL{r}='RESUMO MENSAL'!$C$4)*AM{r}:BV{r})",
        fonte=f_auto, fill=fill_cz, fmt=GS0, al=ctr)
    for k in range(1, N_PARC + 1):
        cm, cv = 2 + k, 38 + k
        cel(par, r, cm,
            f'=IF(OR(CARTÕES!$K{r}=0,CARTÕES!$F{r}=""),0,'
            f'IF({k}>CARTÕES!$F{r},0,CARTÕES!$K{r}+{k}-1))',
            fonte=f_auto, fmt=NUM0, al=ctr, borda=False)
        cel(par, r, cv,
            f'=IF({L(cm)}{r}=0,0,IF({k}=CARTÕES!$F{r},'
            f'CARTÕES!$E{r}-(CARTÕES!$F{r}-1)*ROUND(CARTÕES!$E{r}/CARTÕES!$F{r},0),'
            f'ROUND(CARTÕES!$E{r}/CARTÕES!$F{r},0)))',
            fonte=f_auto, fmt=NUM0, al=ctr, borda=False)
par.sheet_view.showGridLines = False

MESGRID = f"PARCELAS!$C${P1}:$AL${F_COMPRA}"
VALGRID = f"PARCELAS!$AM${P1}:$BV${F_COMPRA}"

# ============================================================= RESUMO MENSAL
rm = wb.create_sheet("RESUMO MENSAL")
rm.sheet_properties.tabColor = "548235"
cabecalho(rm, "RESUMO MENSAL", "Escolha o ano e o mês nas células amarelas. "
          "Tudo abaixo se recalcula sozinho.", 9)
larguras(rm, {"A": 32, "B": 18, "C": 18, "D": 18, "E": 4, "F": 32, "G": 18, "H": 18, "I": 18})

cel(rm, 3, 1, "Ano", fonte=f_b, fill=fill_az, al=ctr)
cel(rm, 3, 2, "Mês", fonte=f_b, fill=fill_az, al=ctr)
cel(rm, 3, 3, "Índice (auto)", fonte=f_b, fill=fill_cz, al=ctr)
cel(rm, 3, 4, "Mês escolhido", fonte=f_b, fill=fill_az, al=ctr)
cel(rm, 4, 1, ANO, fonte=f_kpi, fill=fill_am, fmt="0", al=ctr)
cel(rm, 4, 2, 1, fonte=f_kpi, fill=fill_am, fmt="0", al=ctr)
cel(rm, 4, 3, "=$A$4*12+$B$4", fonte=f_auto, fill=fill_cz, fmt=NUM0, al=ctr)
cel(rm, 4, 4, f'=INDEX({R_MESES},$B$4)&" de "&$A$4', fonte=f_kpi, fill=fill_az, al=ctr)
dv_num(rm, "A4", "whole", "between", "2000", "Ano", "Digite o ano com 4 dígitos.", f2="2100")
dv_num(rm, "B4", "whole", "between", "1", "Mês", "Digite o mês de 1 a 12.", f2="12")


def sm(tipo, status=None, cat=None, idx="$C$4"):
    f = f"SUMIFS({RL_VAL},{RL_IDX},{idx},{RL_TIPO},\"{tipo}\""
    if status: f += f",{RL_STAT},\"{status}\""
    if cat: f += f",{RL_CAT},{cat}"
    return f + ")"


CARTAO_MES = f"SUMPRODUCT(({MESGRID}=$C$4)*{VALGRID})"

for i, (rot, f, fl) in enumerate([
        ("Entradas REALIZADAS", sm("ENTRADA", "REALIZADO"), fill_vd),
        ("Entradas PROJETADAS", sm("ENTRADA", "PROJETADO"), fill_am),
        ("Saídas REALIZADAS", sm("SAÍDA", "REALIZADO"), fill_vm),
        ("Saídas PROJETADAS", sm("SAÍDA", "PROJETADO"), fill_am)]):
    kpi(rm, 6, 1 + i * 2, rot, "=" + f, fl, 2)

BLOCO = [
    ("Entradas do mês (realizado + projetado)", "=" + sm("ENTRADA"), fill_vd, f_kpi),
    ("Saídas do mês (realizado + projetado)", "=" + sm("SAÍDA"), fill_vm, f_kpi),
    ("Parcelas de cartão que vencem no mês", "=" + CARTAO_MES, fill_vm, f_kpi),
    ("RESULTADO DO MÊS", "=$B$9-$B$10-$B$11", fill_az, Font(size=16, bold=True, color=AZUL)),
]
for i, (rot, f, fl, fo) in enumerate(BLOCO):
    r = 9 + i
    cel(rm, r, 1, rot, fonte=f_b, fill=fl, al=esq)
    cel(rm, r, 2, f, fonte=fo, fill=fl, fmt=GS, al=ctr)
    rm.row_dimensions[r].height = 22
rm.conditional_formatting.add("B12", CellIsRule(operator="lessThan", formula=["0"],
                                                font=Font(size=16, bold=True, color="9C0006")))

# posicao em contas + transferencias do mes
cel(rm, 9, 4, "Saldo em contas hoje", fonte=f_b, fill=fill_vd, al=esq)
rm.merge_cells(start_row=9, start_column=4, end_row=9, end_column=6)
cel(rm, 9, 7, "=CONTAS!$A$4", fonte=f_kpi, fill=fill_vd, fmt=GS, al=ctr)
rm.merge_cells(start_row=9, start_column=7, end_row=9, end_column=9)
cel(rm, 10, 4, "Dívida nos cartões", fonte=f_b, fill=fill_vm, al=esq)
rm.merge_cells(start_row=10, start_column=4, end_row=10, end_column=6)
cel(rm, 10, 7, "=CARTÕES!$G$4", fonte=f_kpi, fill=fill_vm, fmt=GS, al=ctr)
rm.merge_cells(start_row=10, start_column=7, end_row=10, end_column=9)
cel(rm, 11, 4, "Transferências no mês (não mexem no resultado)", fonte=f_b, fill=fill_cz, al=esq)
rm.merge_cells(start_row=11, start_column=4, end_row=11, end_column=6)
cel(rm, 11, 7, f'=SUMIFS({RT_VAL},{RT_IDX},$C$4,{RT_STAT},"REALIZADO")',
    fonte=f_kpi, fill=fill_cz, fmt=GS, al=ctr)
rm.merge_cells(start_row=11, start_column=7, end_row=11, end_column=9)
cel(rm, 12, 4, "POSIÇÃO LÍQUIDA (contas − cartões)", fonte=f_b, fill=fill_az, al=esq)
rm.merge_cells(start_row=12, start_column=4, end_row=12, end_column=6)
cel(rm, 12, 7, "=CONTAS!$G$4", fonte=Font(size=16, bold=True, color=AZUL), fill=fill_az,
    fmt=GS, al=ctr)
rm.merge_cells(start_row=12, start_column=7, end_row=12, end_column=9)

secao(rm, 14, "ENTRADAS POR CATEGORIA", 4)
linha_hdr(rm, 15, ["Categoria", "Realizado", "Projetado", "Total"])
cel(rm, 14, 6, "SAÍDAS POR CATEGORIA", fonte=f_sec, fill=fill_az, al=esq)
for c in (7, 8, 9):
    cel(rm, 14, c, None, fill=fill_az)
rm.merge_cells(start_row=14, start_column=6, end_row=14, end_column=9)
linha_hdr(rm, 15, ["Categoria", "Realizado", "Projetado", "Total"], 6)
for i in range(N_CAT):
    r = 16 + i
    lr = P1 + i
    for base, tipo in ((1, "ENTRADA"), (6, "SAÍDA")):
        cel(rm, r, base, f'=IF(LISTAS!$B{lr}="{tipo}",LISTAS!$A{lr},"")', fonte=f_n, al=esq)
        ref = f"${L(base)}{r}"
        cel(rm, r, base + 1, f'=IF({ref}="","",{sm(tipo, "REALIZADO", ref)})', fonte=f_n, fmt=GS0, al=ctr)
        cel(rm, r, base + 2, f'=IF({ref}="","",{sm(tipo, "PROJETADO", ref)})', fonte=f_n, fmt=GS0, al=ctr)
        cel(rm, r, base + 3, f'=IF({ref}="","",{sm(tipo, None, ref)})', fonte=f_b, fmt=GS0, al=ctr)
rtot = 16 + N_CAT
for base, fl in ((1, fill_vd), (6, fill_vm)):
    cel(rm, rtot, base, "TOTAL", fonte=f_b, fill=fl, al=esq)
    for j in (1, 2, 3):
        cel(rm, rtot, base + j, f"=SUM({L(base+j)}16:{L(base+j)}{rtot-1})",
            fonte=f_b, fill=fl, fmt=GS, al=ctr)

rc = rtot + 2
secao(rm, rc, "CONFERÊNCIA — tudo tem que dar zero", 4)
CONF = [
    ("Entradas: KPI menos soma das categorias", f"=$B$9-$D${rtot}"),
    ("Saídas: KPI menos soma das categorias", f"=$B$10-$I${rtot}"),
    ("Lançamentos sem tipo, sem status ou sem conta",
     f'=SUMPRODUCT(({LAN}$A${P1}:$A${F_LANC}<>"")*'
     f'(({RL_TIPO}="")+({RL_STAT}="")+({RL_CONTA}="")>0))'),
    ("Lançamentos com conta que não existe na aba CONTAS",
     f'=SUMPRODUCT(({RL_CONTA}<>"")*(COUNTIF({R_CONTA},{RL_CONTA})=0))'),
    ("Transferências com erro (ver coluna Conferência)",
     f'=COUNTIF({TRF}$I${P1}:$I${F_TRF},"ERRO*")'),
    ("Transferências: total enviado menos total recebido",
     f'=SUMIF({RT_STAT},"REALIZADO",{RT_VAL})'
     f'-SUM(CONTAS!$H${P1}:$H${F_CONTA})-CARTÕES!$E$4'),
    ("Compras de cartão sem nº de parcelas",
     f'=SUMPRODUCT((CARTÕES!$A${P1}:$A${F_COMPRA}<>"")*(CARTÕES!$F${P1}:$F${F_COMPRA}=""))'),
    ("Parcelas: soma das parcelas menos total comprado",
     f"=SUMPRODUCT(({MESGRID}>0)*{VALGRID})-SUM(CARTÕES!$E${P1}:$E${F_COMPRA})"),
]
for i, (rot, f) in enumerate(CONF):
    rr = rc + 1 + i
    cel(rm, rr, 1, rot, fonte=f_n, al=esq)
    rm.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=3)
    cel(rm, rr, 4, f, fonte=f_b, fmt="#,##0", al=ctr)
    rm.conditional_formatting.add(f"D{rr}", CellIsRule(operator="notEqual", formula=["0"],
                                                       fill=fill_vm, font=Font(bold=True, color="9C0006")))
    rm.conditional_formatting.add(f"D{rr}", CellIsRule(operator="equal", formula=["0"],
                                                       fill=fill_vd, font=Font(bold=True, color="375623")))
cel(rm, rc + len(CONF) + 2, 1,
    "Tudo em verde e em zero = a planilha está fechando. Qualquer número diferente de zero fica vermelho.",
    fonte=f_sub, borda=False, al=esq)
rm.freeze_panes = "A6"
rm.sheet_view.showGridLines = False

# ============================================================== RESUMO ANUAL
ra = wb.create_sheet("RESUMO ANUAL")
ra.sheet_properties.tabColor = "833C00"
cabecalho(ra, "RESUMO ANUAL", "Os 12 meses do ano lado a lado. Troque o ano na célula amarela.", 10)
larguras(ra, {"A": 32})
for c in range(2, 15):
    ra.column_dimensions[L(c)].width = 15
cel(ra, 3, 1, "Ano", fonte=f_b, fill=fill_az, al=ctr)
cel(ra, 4, 1, ANO, fonte=f_kpi, fill=fill_am, fmt="0", al=ctr)
dv_num(ra, "A4", "whole", "between", "2000", "Ano", "Digite o ano com 4 dígitos.", f2="2100")

linha_hdr(ra, HDR, ["Mês", "Entradas realizadas", "Entradas projetadas", "Entradas total",
                    "Saídas realizadas", "Saídas projetadas", "Saídas total",
                    "Cartão (parcelas)", "Resultado do mês", "Saldo acumulado",
                    "Transferências"], altura=42)
for i in range(12):
    r = P1 + i
    idx = f"($A$4*12+{i+1})"
    cel(ra, r, 1, f'=INDEX({R_MESES},{i+1})', fonte=f_b, fill=fill_cz, al=esq)
    cel(ra, r, 2, "=" + sm("ENTRADA", "REALIZADO", idx=idx), fonte=f_n, fmt=GS0, al=ctr)
    cel(ra, r, 3, "=" + sm("ENTRADA", "PROJETADO", idx=idx), fonte=f_n, fmt=GS0, al=ctr)
    cel(ra, r, 4, f"=B{r}+C{r}", fonte=f_b, fill=fill_vd, fmt=GS0, al=ctr)
    cel(ra, r, 5, "=" + sm("SAÍDA", "REALIZADO", idx=idx), fonte=f_n, fmt=GS0, al=ctr)
    cel(ra, r, 6, "=" + sm("SAÍDA", "PROJETADO", idx=idx), fonte=f_n, fmt=GS0, al=ctr)
    cel(ra, r, 7, f"=E{r}+F{r}", fonte=f_b, fill=fill_vm, fmt=GS0, al=ctr)
    cel(ra, r, 8, f"=SUMPRODUCT(({MESGRID}={idx})*{VALGRID})", fonte=f_n, fill=fill_vm, fmt=GS0, al=ctr)
    cel(ra, r, 9, f"=D{r}-G{r}-H{r}", fonte=f_b, fill=fill_az, fmt=GS, al=ctr)
    cel(ra, r, 10, (f"=I{r}" if i == 0 else f"=J{r-1}+I{r}"), fonte=f_b, fmt=GS, al=ctr)
    cel(ra, r, 11, f'=SUMIFS({RT_VAL},{RT_IDX},{idx})', fonte=f_n, fill=fill_cz, fmt=GS0, al=ctr)
rt = P1 + 12
cel(ra, rt, 1, "TOTAL DO ANO", fonte=f_b, fill=fill_az, al=esq)
for c in list(range(2, 10)) + [11]:
    cel(ra, rt, c, f"=SUM({L(c)}{P1}:{L(c)}{rt-1})", fonte=f_b, fill=fill_az, fmt=GS, al=ctr)
cel(ra, rt, 10, f"=J{rt-1}", fonte=f_b, fill=fill_az, fmt=GS, al=ctr)
for col in ("I", "J"):
    ra.conditional_formatting.add(f"{col}{P1}:{col}{rt}",
                                  CellIsRule(operator="lessThan", formula=["0"],
                                             font=Font(bold=True, color="9C0006"), fill=fill_vm))

blocos = []
r0 = rt + 3
for titulo, tipo in (("SAÍDAS POR CATEGORIA, MÊS A MÊS", "SAÍDA"),
                     ("ENTRADAS POR CATEGORIA, MÊS A MÊS", "ENTRADA")):
    secao(ra, r0, titulo, 14)
    linha_hdr(ra, r0 + 1, ["Categoria"] + [m[:3] for m in MESES] + ["Total do ano"])
    for i in range(N_CAT):
        r = r0 + 2 + i
        lr = P1 + i
        cel(ra, r, 1, f'=IF(LISTAS!$B{lr}="{tipo}",LISTAS!$A{lr},"")', fonte=f_n, al=esq)
        for m in range(12):
            cel(ra, r, 2 + m,
                f'=IF($A{r}="","",{sm(tipo, None, f"$A{r}", idx=f"($A$4*12+{m+1})")})',
                fonte=f_n, fmt=GS0, al=ctr)
        cel(ra, r, 14, f'=IF($A{r}="","",SUM(B{r}:M{r}))', fonte=f_b, fmt=GS0, al=ctr)
    rtx = r0 + 2 + N_CAT
    cel(ra, rtx, 1, "TOTAL " + ("SAÍDAS" if tipo == "SAÍDA" else "ENTRADAS"), fonte=f_b,
        fill=fill_vm if tipo == "SAÍDA" else fill_vd, al=esq)
    for c in range(2, 15):
        cel(ra, rtx, c, f"=SUM({L(c)}{r0+2}:{L(c)}{rtx-1})", fonte=f_b,
            fill=fill_vm if tipo == "SAÍDA" else fill_vd, fmt=GS0, al=ctr)
    blocos.append(rtx)
    r0 = rtx + 3
rsc, rec = blocos

r2 = r0
secao(ra, r2, "CONFERÊNCIA DO ANO — tudo tem que dar zero", 4)
CONF2 = [
    ("Entradas: total do ano menos soma das categorias", f"=D{rt}-N{rec}"),
    ("Saídas: total do ano menos soma das categorias", f"=G{rt}-N{rsc}"),
    ("Saldo acumulado bate com a soma dos resultados", f"=J{rt}-I{rt}"),
]
for i, (rot, f) in enumerate(CONF2):
    rr = r2 + 1 + i
    cel(ra, rr, 1, rot, fonte=f_n, al=esq)
    ra.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=3)
    cel(ra, rr, 4, f, fonte=f_b, fmt="#,##0", al=ctr)
    ra.conditional_formatting.add(f"D{rr}", CellIsRule(operator="notEqual", formula=["0"],
                                                       fill=fill_vm, font=Font(bold=True, color="9C0006")))
    ra.conditional_formatting.add(f"D{rr}", CellIsRule(operator="equal", formula=["0"],
                                                       fill=fill_vd, font=Font(bold=True, color="375623")))
ra.freeze_panes = "B6"
ra.sheet_view.showGridLines = False

# =================================================================== EXEMPLO
ex = wb.create_sheet("EXEMPLO")
ex.sheet_properties.tabColor = "FFC000"
cabecalho(ex, "EXEMPLO — COMO PREENCHER CADA CASO",
          "Esta aba é só modelo. Nada aqui entra nas contas. Pode apagar a aba quando não precisar mais.", 8)
larguras(ex, {"A": 13, "B": 13, "C": 13, "D": 26, "E": 30, "F": 22, "G": 16, "H": 52})

secao(ex, 4, "MODELOS PARA A ABA CONTAS", 8)
linha_hdr(ex, 5, ["Nome da conta", "Banco / Instituição", "Tipo", "Saldo inicial (Gs.)",
                  "Saldo inicial em", "", "", "Por que assim"])
EX_CONTA = [
    ("Conta Ueno", "Banco Ueno", "Conta corrente", 12000000, dt.date(2025, 12, 31), "", "",
     "Saldo que a conta tinha no dia em que você começou a usar a planilha."),
    ("Dinheiro", "—", "Dinheiro / Caixa", 800000, dt.date(2025, 12, 31), "", "",
     "O que você carrega no bolso. Também é uma conta."),
    ("Poupança", "Banco Ueno", "Poupança", 25000000, dt.date(2025, 12, 31), "", "",
     "Guardar dinheiro aqui é TRANSFERÊNCIA, não é saída."),
]
for i, linha in enumerate(EX_CONTA):
    r = 6 + i
    for j, v in enumerate(linha):
        fmt = GS if j == 3 else (DATA if j == 4 else None)
        cel(ex, r, 1 + j, v or None, fonte=f_n if j < 7 else f_sub, fmt=fmt,
            al=ctr if j in (2, 3, 4) else esq)
    ex.row_dimensions[r].height = 16

r0 = 6 + len(EX_CONTA) + 2
secao(ex, r0, "MODELOS PARA A ABA LANÇAMENTOS", 8)
linha_hdr(ex, r0 + 1, ["Data", "Tipo", "Status", "Categoria", "Descrição", "Conta",
                       "Valor (Gs.)", "Por que assim"])
EX_LANC = [
    (dt.date(2026, 1, 5), "ENTRADA", "REALIZADO", "Salário", "Salário de janeiro", "Conta Ueno", 8500000,
     "Já caiu na conta. REALIZADO, e o saldo da Conta Ueno sobe."),
    (dt.date(2026, 1, 10), "SAÍDA", "REALIZADO", "Alimentação", "Supermercado", "Conta Ueno", 780000,
     "Pagou no débito, saiu da conta na hora. Saldo da Conta Ueno cai."),
    (dt.date(2026, 1, 11), "SAÍDA", "REALIZADO", "Transporte", "Combustível", "Dinheiro", 150000,
     "Pagou em espécie. Conta = Dinheiro."),
    (dt.date(2026, 2, 5), "ENTRADA", "PROJETADO", "Salário", "Salário de fevereiro", "Conta Ueno", 8500000,
     "Ainda não caiu. PROJETADO: entra no saldo previsto, não no saldo atual."),
    (dt.date(2026, 2, 15), "SAÍDA", "PROJETADO", "Moradia", "Aluguel de fevereiro", "Conta Ueno", 2500000,
     "Gasto certo que ainda vai acontecer."),
]
for i, linha in enumerate(EX_LANC):
    r = r0 + 2 + i
    for j, v in enumerate(linha):
        fmt = DATA if j == 0 else (GS if j == 6 else None)
        cel(ex, r, 1 + j, v, fonte=f_n if j < 7 else f_sub, fmt=fmt,
            al=ctr if j in (0, 1, 2, 5, 6) else esq)
    ex.row_dimensions[r].height = 16

r1 = r0 + 2 + len(EX_LANC) + 2
secao(ex, r1, "MODELOS PARA A ABA TRANSFERÊNCIAS", 8)
linha_hdr(ex, r1 + 1, ["Data", "Status", "De (sai de)", "Para (entra em)", "Descrição", "",
                       "Valor (Gs.)", "Por que assim"])
EX_TRF = [
    (dt.date(2026, 1, 15), "REALIZADO", "Conta Ueno", "Poupança", "Guardando para a reserva", "", 2000000,
     "Não é saída: o dinheiro continua seu. Só saiu de uma conta e entrou na outra."),
    (dt.date(2026, 2, 10), "REALIZADO", "Conta Ueno", "Cartão 1", "Pagamento da fatura de fevereiro", "", 500000,
     "Pagamento de fatura. Tira da conta e abate a dívida do cartão. NÃO é despesa."),
    (dt.date(2026, 1, 20), "REALIZADO", "Conta Ueno", "Dinheiro", "Saque no caixa", "", 1000000,
     "Sacar dinheiro é transferência da conta para o Dinheiro."),
    (dt.date(2026, 3, 10), "PROJETADO", "Conta Ueno", "Cartão 1", "Fatura de março", "", 500000,
     "Pagamento futuro. Só mexe no saldo previsto."),
]
for i, linha in enumerate(EX_TRF):
    r = r1 + 2 + i
    for j, v in enumerate(linha):
        fmt = DATA if j == 0 else (GS if j == 6 else None)
        cel(ex, r, 1 + j, v or None, fonte=f_n if j < 7 else f_sub, fmt=fmt,
            al=ctr if j in (0, 1, 2, 3, 6) else esq)
    ex.row_dimensions[r].height = 16

r2 = r1 + 2 + len(EX_TRF) + 2
secao(ex, r2, "MODELOS PARA A ABA CARTÕES", 8)
linha_hdr(ex, r2 + 1, ["Data da compra", "Cartão", "Descrição", "Categoria",
                       "Valor total (Gs.)", "Nº parcelas", "1ª parcela", "O que a planilha faz"])
EX_CART = [
    (dt.date(2026, 1, 18), "Cartão 1", "Geladeira", "Manutenção", 6000000, 12, None,
     "1ª parcela em branco: começa em fev/2026, 12 parcelas de Gs. 500.000."),
    (dt.date(2026, 3, 2), "Cartão 1", "Passagem aérea", "Lazer", 4500000, 6, dt.date(2026, 4, 1),
     "1ª parcela digitada: 6 parcelas de Gs. 750.000, de abr a set/2026."),
    (dt.date(2026, 2, 8), "Cartão 2", "Farmácia", "Saúde", 350000, 1, None,
     "Compra à vista no crédito: 1 parcela, cai inteira no mês seguinte."),
    (dt.date(2026, 5, 4), "Cartão 2", "Notebook", "Educação", 10000000, 7, None,
     "Gs. 10.000.000 em 7x não dá número redondo. A última parcela é ajustada para fechar exato."),
]
for i, linha in enumerate(EX_CART):
    r = r2 + 2 + i
    for j, v in enumerate(linha):
        fmt = DATA if j in (0, 6) else (GS if j == 4 else ("0" if j == 5 else None))
        cel(ex, r, 1 + j, v, fonte=f_n if j < 7 else f_sub, fmt=fmt,
            al=ctr if j in (0, 1, 4, 5, 6) else esq)
    ex.row_dimensions[r].height = 16

r3 = r2 + 2 + len(EX_CART) + 2
secao(ex, r3, "CASOS QUE CONFUNDEM — O QUE FAZER", 8)
DUVIDAS = [
    ("Paguei a fatura do cartão", "TRANSFERÊNCIAS: De = sua conta, Para = o cartão. "
     "Tira da conta e abate a dívida. Nunca lance como SAÍDA — a parcela já foi contada."),
    ("Comprei no cartão e paguei à vista (1x)", "Aba CARTÕES, com Nº de parcelas = 1."),
    ("Comprei no débito", "Aba LANÇAMENTOS, SAÍDA, escolhendo a Conta. Débito sai da conta na hora, "
     "não é cartão de crédito."),
    ("Passei dinheiro da conta para a poupança", "TRANSFERÊNCIAS. Não é saída, o dinheiro continua seu."),
    ("Saquei dinheiro no caixa", "TRANSFERÊNCIAS: De = conta, Para = Dinheiro."),
    ("Recebi um empréstimo", "ENTRADA, categoria Empréstimo recebido, na conta que recebeu. "
     "As parcelas que você vai pagar são SAÍDA em Empréstimo / Financiamento, com status PROJETADO."),
    ("Paguei juros ou tarifa do banco", "SAÍDA, categoria Juros e tarifas. Isso sim é despesa de verdade."),
    ("Parcelei uma compra e depois quitei tudo", "Mude o Nº de parcelas para o que você realmente pagou "
     "e ajuste o Valor total. A planilha recalcula sozinha."),
    ("Errei um lançamento", "Selecione a linha inteira e aperte DELETE para limpar. "
     "Nunca use Excluir linha, senão as fórmulas automáticas somem."),
    ("O saldo da conta não bate com o banco", "Confira: (1) todo lançamento tem Conta preenchida; "
     "(2) o saldo inicial está certo; (3) as faturas de cartão pagas foram lançadas em TRANSFERÊNCIAS."),
]
linha_hdr(ex, r3 + 1, ["Situação", "O que fazer"])
for c in range(3, 9):
    cel(ex, r3 + 1, c, None, fonte=f_hdr, fill=fill_hdr)
ex.merge_cells(start_row=r3 + 1, start_column=2, end_row=r3 + 1, end_column=8)
for i, (a, b) in enumerate(DUVIDAS):
    r = r3 + 2 + i
    cel(ex, r, 1, a, fonte=f_b, al=wrap)
    cel(ex, r, 2, b, fonte=f_n, al=wrap)
    for c in range(3, 9):
        cel(ex, r, c, None)
    ex.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ex.row_dimensions[r].height = 32
ex.sheet_view.showGridLines = False

ORDEM = ["INSTRUÇÕES", "CONTAS", "LANÇAMENTOS", "TRANSFERÊNCIAS", "CARTÕES",
         "RESUMO MENSAL", "RESUMO ANUAL", "EXEMPLO", "PARCELAS", "LISTAS"]
wb._sheets = [wb[n] for n in ORDEM]
wb.active = 0
wb.save(SAIDA)
print("gerado:", SAIDA, f"{SAIDA.stat().st_size/1024:.0f} KB")
print("abas:", wb.sheetnames)
