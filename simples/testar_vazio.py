#!/usr/bin/env python3
"""Confere o arquivo VAZIO: recalcula, exige zero erro de formula, zero numero
diferente de zero nos resumos, e confirma que todas as listas suspensas existem."""
import shutil, sys
from pathlib import Path
import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "gerador"))
from recalc_lo import recalc

LIMPO = RAIZ / "CONTROLE_FINANCEIRO_SIMPLES.xlsx"
COPIA = RAIZ / "simples" / "_vazio.xlsx"
shutil.copy(LIMPO, COPIA)

print("recalculando o arquivo vazio...")
res = recalc(COPIA)
print(f"  {res['total_formulas']} formulas, {res['seconds']}s, {res['total_errors']} com erro")
for k, v in list(res.get("error_summary", {}).items())[:10]:
    print("   ", k, v[:12])

wb = openpyxl.load_workbook(COPIA, data_only=True)
wf = openpyxl.load_workbook(COPIA)
ok = falhas = 0
def conf(nome, esperado, obtido):
    global ok, falhas
    bom = esperado == obtido
    print(("  [OK   ] " if bom else "  [FALHA] ") + nome.ljust(56)
          + f"esperado {esperado!r:>14}  obtido {obtido!r:>14}")
    if bom: ok += 1
    else: falhas += 1

print("\n[1] Abas presentes")
esperadas = ["INSTRUÇÕES", "LANÇAMENTOS", "CARTÕES", "RESUMO MENSAL",
             "RESUMO ANUAL", "EXEMPLO", "PARCELAS", "LISTAS"]
conf("Todas as abas na ordem certa", esperadas, wb.sheetnames)

print("\n[2] Nada aparece como número no arquivo vazio")
rm, ra, cr = wb["RESUMO MENSAL"], wb["RESUMO ANUAL"], wb["CARTÕES"]
for cel, rot in [("A7", "Entradas realizadas"), ("C7", "Entradas projetadas"),
                 ("E7", "Saídas realizadas"), ("G7", "Saídas projetadas"),
                 ("B9", "Entradas do mês"), ("B10", "Saídas do mês"),
                 ("B11", "Cartão do mês"), ("B12", "Resultado do mês")]:
    conf(f"RESUMO MENSAL {cel} · {rot}", 0, rm[cel].value)
for cel, rot in [("A4", "Total comprado"), ("B4", "Já pago"), ("C4", "Falta pagar")]:
    conf(f"CARTÕES {cel} · {rot}", 0, cr[cel].value)
for m in range(12):
    r = 6 + m
    for c, rot in [(4, "entradas"), (7, "saídas"), (8, "cartão"), (9, "resultado"), (10, "acumulado")]:
        v = ra.cell(r, c).value
        if v != 0:
            conf(f"RESUMO ANUAL linha {r} col {c} ({rot})", 0, v)
conf("RESUMO ANUAL · 12 meses todos zerados", True,
     all(ra.cell(6 + m, c).value == 0 for m in range(12) for c in (4, 7, 8, 9, 10)))

print("\n[3] Conferências internas fecham em zero")
for r in range(57, 63):
    rot = rm.cell(r, 1).value
    if isinstance(rot, str) and rot[:1].isalpha() and rm.cell(r, 4).value is not None:
        conf("MENSAL · " + rot[:44], 0, rm.cell(r, 4).value)
for r in range(ra.max_row - 5, ra.max_row + 1):
    rot = ra.cell(r, 1).value
    if isinstance(rot, str) and ra.cell(r, 4).value is not None:
        conf("ANUAL · " + rot[:44], 0, ra.cell(r, 4).value)

print("\n[4] Listas suspensas (dropdowns)")
ESPERADO = {
    "LANÇAMENTOS": {"B6:B1005": "Tipo", "C6:C1005": "Status", "D6:D1005": "Categoria",
                    "F6:F1005": "Meio de pagamento", "G6:G1005": "Valor"},
    "CARTÕES": {"B6:B65": "Cartão", "D6:D65": "Categoria", "E6:E65": "Valor total",
                "F6:F65": "Nº de parcelas"},
    "RESUMO MENSAL": {"A4": "Ano", "B4": "Mês"},
    "RESUMO ANUAL": {"A4": "Ano"},
}
for aba, faixas in ESPERADO.items():
    tem = set()
    for dv in wf[aba].data_validations.dataValidation:
        tem.update(str(x) for x in dv.sqref.ranges)
    for faixa, rot in faixas.items():
        conf(f"{aba} · {faixa} · {rot}", True, faixa in tem)

print("\n[5] Colunas automáticas têm fórmula em todas as linhas")
lan, car, par = wf["LANÇAMENTOS"], wf["CARTÕES"], wf["PARCELAS"]
conf("LANÇAMENTOS · coluna Mês (H)", 1000,
     sum(1 for r in range(6, 1006) if str(lan.cell(r, 8).value or "").startswith("=")))
conf("LANÇAMENTOS · coluna Índice (I)", 1000,
     sum(1 for r in range(6, 1006) if str(lan.cell(r, 9).value or "").startswith("=")))
for c, rot in [(8, "Início"), (9, "Valor da parcela"), (10, "Última parcela"),
               (11, "Índice"), (12, "Já pago"), (13, "Falta pagar")]:
    conf(f"CARTÕES · coluna {rot}", 60,
         sum(1 for r in range(6, 66) if str(car.cell(r, c).value or "").startswith("=")))
conf("PARCELAS · grade completa (60 x 36 x 2)", 4320,
     sum(1 for r in range(6, 66) for k in range(1, 37)
         if str(par.cell(r, 2 + k).value or "").startswith("=")
         and str(par.cell(r, 38 + k).value or "").startswith("=")) * 2)

print("\n[6] Nenhum dado de exemplo dentro das abas de trabalho")
conf("LANÇAMENTOS sem nenhuma data digitada", 0,
     sum(1 for r in range(6, 1006) if lan.cell(r, 1).value is not None))
conf("CARTÕES sem nenhuma compra digitada", 0,
     sum(1 for r in range(6, 66) if car.cell(r, 1).value is not None))

print("\n" + "=" * 96)
print(f"RESULTADO: {ok} conferências OK, {falhas} falhas, {res['total_errors']} células com erro")
print("=" * 96)
sys.exit(1 if falhas or res["total_errors"] else 0)
