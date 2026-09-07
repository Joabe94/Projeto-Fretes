#!/usr/bin/env python3
"""Preenche uma copia da planilha com dados conhecidos, recalcula no LibreOffice
e confere cada numero contra um calculo independente feito em Python."""
import datetime as dt, shutil, sys, random
from pathlib import Path
import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "gerador"))
from recalc_lo import recalc

LIMPO = RAIZ / "CONTROLE_FINANCEIRO_SIMPLES.xlsx"
TESTE = RAIZ / "simples" / "_teste.xlsx"
LIN_1, FIM_LANC, FIM_COMPRA, N_PARC = 6, 1005, 65, 36

# ----------------------------------------------------------- dados de teste
rnd = random.Random(20260907)
CAT_E = ["Salário", "Venda", "Aluguel recebido", "Rendimento"]
CAT_S = ["Moradia", "Alimentação", "Transporte", "Saúde", "Lazer", "Impostos e taxas"]
MEIOS = ["Dinheiro", "Conta corrente", "Débito", "Boleto", "Transferência"]

LANC = []
for i in range(240):
    d = dt.date(2026, rnd.randint(1, 12), rnd.randint(1, 28))
    ent = rnd.random() < 0.35
    LANC.append(dict(data=d, tipo="ENTRADA" if ent else "SAÍDA",
                     status="REALIZADO" if rnd.random() < 0.6 else "PROJETADO",
                     cat=rnd.choice(CAT_E if ent else CAT_S),
                     desc=f"Teste {i+1}", meio=rnd.choice(MEIOS),
                     valor=rnd.randrange(50, 9000) * 1000))
# alguns lancamentos em 2027 para provar que o filtro de ano funciona
for i in range(20):
    LANC.append(dict(data=dt.date(2027, rnd.randint(1, 12), 10), tipo="SAÍDA",
                     status="REALIZADO", cat="Lazer", desc=f"Outro ano {i+1}",
                     meio="Dinheiro", valor=rnd.randrange(50, 900) * 1000))

COMPRAS = [
    dict(data=dt.date(2026, 1, 18), cartao="Cartão 1", desc="Geladeira", cat="Manutenção",
         total=6000000, n=12, primeira=None),
    dict(data=dt.date(2026, 3, 2), cartao="Cartão 1", desc="Passagem", cat="Lazer",
         total=4500000, n=6, primeira=dt.date(2026, 4, 1)),
    dict(data=dt.date(2026, 2, 8), cartao="Cartão 2", desc="Farmácia", cat="Saúde",
         total=350000, n=1, primeira=None),
    dict(data=dt.date(2026, 5, 4), cartao="Cartão 2", desc="Notebook", cat="Educação",
         total=10000000, n=7, primeira=None),
    dict(data=dt.date(2026, 12, 20), cartao="Cartão 1", desc="Presente virada de ano",
         cat="Família", total=3333333, n=5, primeira=None),   # cruza o ano
    dict(data=dt.date(2025, 11, 10), cartao="Cartão 3", desc="Compra do ano passado",
         cat="Moradia", total=2400000, n=24, primeira=None),  # comeca antes de 2026
]

def idx(d): return d.year * 12 + d.month

def inicio(c):
    if c["primeira"]: return c["primeira"]
    a, m = c["data"].year, c["data"].month
    return dt.date(a + 1, 1, 1) if m == 12 else dt.date(a, m + 1, 1)

def parcelas(c):
    """devolve [(indice_mes, valor)] exatamente como a planilha deve calcular"""
    base = round(c["total"] / c["n"])
    i0 = idx(inicio(c))
    out = []
    for k in range(1, c["n"] + 1):
        v = c["total"] - (c["n"] - 1) * base if k == c["n"] else base
        out.append((i0 + k - 1, v))
    return out

# ----------------------------------------------------------------- preencher
shutil.copy(LIMPO, TESTE)
wb = openpyxl.load_workbook(TESTE)
lan, car = wb["LANÇAMENTOS"], wb["CARTÕES"]
for i, l in enumerate(LANC):
    r = LIN_1 + i
    lan.cell(r, 1).value = l["data"]; lan.cell(r, 2).value = l["tipo"]
    lan.cell(r, 3).value = l["status"]; lan.cell(r, 4).value = l["cat"]
    lan.cell(r, 5).value = l["desc"]; lan.cell(r, 6).value = l["meio"]
    lan.cell(r, 7).value = l["valor"]
for i, c in enumerate(COMPRAS):
    r = LIN_1 + i
    car.cell(r, 1).value = c["data"]; car.cell(r, 2).value = c["cartao"]
    car.cell(r, 3).value = c["desc"]; car.cell(r, 4).value = c["cat"]
    car.cell(r, 5).value = c["total"]; car.cell(r, 6).value = c["n"]
    car.cell(r, 7).value = c["primeira"]
# usa mes de referencia fixo no RESUMO MENSAL
wb["RESUMO MENSAL"]["A4"].value = 2026
wb["RESUMO MENSAL"]["B4"].value = 3
wb["RESUMO ANUAL"]["A4"].value = 2026
wb.save(TESTE)

print("recalculando no LibreOffice...")
res = recalc(TESTE)
if res.get("error"): print("ERRO:", res["error"]); sys.exit(1)
print(f"  {res['total_formulas']} formulas, {res['seconds']}s, {res['total_errors']} com erro")
if res["total_errors"]:
    print("  CELULAS COM ERRO:")
    for k, v in list(res["error_summary"].items())[:20]: print("   ", k, v[:12])

# ------------------------------------------------------------------ conferir
wv = openpyxl.load_workbook(TESTE, data_only=True)
ok = falhas = 0
def conf(nome, esperado, obtido, tol=0.5):
    global ok, falhas
    o = obtido if isinstance(obtido, (int, float)) else 0
    bom = abs(float(esperado) - float(o)) <= tol
    print(("  [OK   ] " if bom else "  [FALHA] ") + nome.ljust(52)
          + f"esperado {esperado:>16,.0f}  obtido {o:>16,.0f}".replace(",", "."))
    if bom: ok += 1
    else: falhas += 1

def soma(tipo=None, status=None, cat=None, ano=None, mes=None):
    t = 0
    for l in LANC:
        if tipo and l["tipo"] != tipo: continue
        if status and l["status"] != status: continue
        if cat and l["cat"] != cat: continue
        if ano and l["data"].year != ano: continue
        if mes and l["data"].month != mes: continue
        t += l["valor"]
    return t

def cartao_mes(i):
    return sum(v for c in COMPRAS for (m, v) in parcelas(c) if m == i)

rm, ra, cr = wv["RESUMO MENSAL"], wv["RESUMO ANUAL"], wv["CARTÕES"]

print("\n[1] RESUMO MENSAL — marco/2026")
I3 = 2026 * 12 + 3
conf("Entradas realizadas", soma("ENTRADA", "REALIZADO", ano=2026, mes=3), rm["A7"].value)
conf("Entradas projetadas", soma("ENTRADA", "PROJETADO", ano=2026, mes=3), rm["C7"].value)
conf("Saídas realizadas", soma("SAÍDA", "REALIZADO", ano=2026, mes=3), rm["E7"].value)
conf("Saídas projetadas", soma("SAÍDA", "PROJETADO", ano=2026, mes=3), rm["G7"].value)
conf("Entradas do mês (total)", soma("ENTRADA", ano=2026, mes=3), rm["B9"].value)
conf("Saídas do mês (total)", soma("SAÍDA", ano=2026, mes=3), rm["B10"].value)
conf("Parcelas de cartão do mês", cartao_mes(I3), rm["B11"].value)
conf("Resultado do mês", soma("ENTRADA", ano=2026, mes=3) - soma("SAÍDA", ano=2026, mes=3)
     - cartao_mes(I3), rm["B12"].value)
for cat in ["Salário", "Venda"]:
    lin = next(r for r in range(16, 56) if rm.cell(r, 1).value == cat)
    conf(f"Entrada por categoria · {cat}", soma("ENTRADA", cat=cat, ano=2026, mes=3),
         rm.cell(lin, 4).value)
for cat in ["Moradia", "Alimentação", "Lazer"]:
    lin = next(r for r in range(16, 56) if rm.cell(r, 6).value == cat)
    conf(f"Saída por categoria · {cat}", soma("SAÍDA", cat=cat, ano=2026, mes=3),
         rm.cell(lin, 9).value)
print("  conferência interna do mês:")
for r in range(58, 63):
    rot = rm.cell(r, 1).value
    if rot: conf("  " + str(rot)[:48], 0, rm.cell(r, 4).value)

print("\n[2] RESUMO ANUAL — 2026")
acum = 0
for m in range(1, 13):
    r = 5 + m
    e = soma("ENTRADA", ano=2026, mes=m); s = soma("SAÍDA", ano=2026, mes=m)
    c = cartao_mes(2026 * 12 + m); acum += e - s - c
    conf(f"{m:02d}/2026 entradas", e, ra.cell(r, 4).value)
    conf(f"{m:02d}/2026 saídas", s, ra.cell(r, 7).value)
    conf(f"{m:02d}/2026 cartão", c, ra.cell(r, 8).value)
    conf(f"{m:02d}/2026 resultado", e - s - c, ra.cell(r, 9).value)
    conf(f"{m:02d}/2026 acumulado", acum, ra.cell(r, 10).value)
conf("Total do ano · entradas", soma("ENTRADA", ano=2026), ra.cell(18, 4).value)
conf("Total do ano · saídas", soma("SAÍDA", ano=2026), ra.cell(18, 7).value)
conf("Total do ano · cartão", sum(cartao_mes(2026 * 12 + m) for m in range(1, 13)),
     ra.cell(18, 8).value)
conf("Total do ano · resultado", soma("ENTRADA", ano=2026) - soma("SAÍDA", ano=2026)
     - sum(cartao_mes(2026 * 12 + m) for m in range(1, 13)), ra.cell(18, 9).value)
conf("Saldo acumulado final", acum, ra.cell(18, 10).value)
print("  conferência interna do ano:")
for r in range(ra.max_row - 6, ra.max_row + 1):
    rot = ra.cell(r, 1).value
    if isinstance(rot, str) and rot.startswith(("Entradas:", "Saídas:", "Saldo")):
        conf("  " + rot[:48], 0, ra.cell(r, 4).value)

print("\n[3] CARTÕES")
conf("Total comprado", sum(c["total"] for c in COMPRAS), cr["A4"].value)
hoje_i = dt.date.today().year * 12 + dt.date.today().month
pago = sum(v for c in COMPRAS for (m, v) in parcelas(c) if m <= hoje_i)
conf("Já pago até hoje", pago, cr["B4"].value)
conf("Falta pagar", sum(c["total"] for c in COMPRAS) - pago, cr["C4"].value)
for i, c in enumerate(COMPRAS):
    r = LIN_1 + i
    p = parcelas(c)
    conf(f"Compra {i+1} · soma das parcelas = total", c["total"], sum(v for _, v in p))
    conf(f"Compra {i+1} · valor da parcela", round(c["total"] / c["n"]), cr.cell(r, 9).value)
    conf(f"Compra {i+1} · já pago", sum(v for m, v in p if m <= hoje_i), cr.cell(r, 12).value)
    conf(f"Compra {i+1} · falta pagar", sum(v for m, v in p if m > hoje_i), cr.cell(r, 13).value)
r0 = FIM_COMPRA + 3
for i in range(4):
    r = r0 + 2 + i
    nome = cr.cell(r, 1).value
    if not nome: continue
    tot = sum(c["total"] for c in COMPRAS if c["cartao"] == nome)
    conf(f"Saldo por cartão · {nome} · comprado", tot, cr.cell(r, 2).value)
    conf(f"Saldo por cartão · {nome} · falta",
         sum(v for c in COMPRAS if c["cartao"] == nome for m, v in parcelas(c) if m > hoje_i),
         cr.cell(r, 4).value)

print("\n[4] PARCELAS — grade automática")
pa = wv["PARCELAS"]
for i, c in enumerate(COMPRAS):
    r = LIN_1 + i
    esp = parcelas(c)
    got = [(pa.cell(r, 2 + k).value, pa.cell(r, 38 + k).value) for k in range(1, N_PARC + 1)]
    got = [(m, v) for m, v in got if m]
    conf(f"Compra {i+1} · nº de parcelas geradas", len(esp), len(got))
    conf(f"Compra {i+1} · 1º mês", esp[0][0], got[0][0] if got else 0)
    conf(f"Compra {i+1} · último mês", esp[-1][0], got[-1][0] if got else 0)
    conf(f"Compra {i+1} · soma na grade", c["total"], sum(v for _, v in got))
linhas_vazias = sum(1 for r in range(LIN_1 + len(COMPRAS), FIM_COMPRA + 1)
                    for k in range(1, N_PARC + 1) if pa.cell(r, 2 + k).value)
conf("Linhas sem compra não geram parcela", 0, linhas_vazias)

print("\n" + "=" * 96)
print(f"RESULTADO: {ok} conferências OK, {falhas} falhas, {res['total_errors']} células com erro")
print("=" * 96)
sys.exit(1 if falhas or res["total_errors"] else 0)
