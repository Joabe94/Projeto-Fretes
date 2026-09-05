#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Primeiro uso real: parte do arquivo LIMPO e preenche tudo do zero.

Prova que o arquivo vazio funciona: cadastros, lancamentos das nove operacoes,
cronograma de parcelas, pagamento de parcela e reserva de meta - conferindo
cada numero na mao.
"""
from __future__ import annotations

import datetime as dt
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "gerador"))

import openpyxl

from comum import LINHA_DADOS as DAT, gs
from dados import sub
from recalc_lo import recalc

LIMPO = "SISTEMA_FINANCEIRO_PESSOAL_V1_LIMPO.xlsx"
TMP = "/tmp/primeiro_uso.xlsx"
FALHAS, OKS = [], []


def ok(nome, esperado, obtido, tol=0.5):
    try:
        bom = abs(float(esperado or 0) - float(obtido or 0)) <= tol
    except (TypeError, ValueError):
        bom = str(esperado) == str(obtido)
    (OKS if bom else FALHAS).append(nome)
    def _f(x):
        if isinstance(x, float) and abs(x) < 1:
            return f"{x * 100:.1f}%"
        return gs(x) if isinstance(x, (int, float)) else x
    e, o = _f(esperado), _f(obtido)
    print(f"  [{'OK   ' if bom else 'FALHA'}] {nome:52s} esperado {str(e):>18}  obtido {str(o):>18}")


def escrever(ws, linha, valores, fmt_datas=()):
    for col, v in valores.items():
        c = ws.cell(linha, col, v)
        if col in fmt_datas:
            c.number_format = "DD/MM/YYYY"


def main():
    shutil.copy(LIMPO, TMP)
    wb = openpyxl.load_workbook(TMP)
    print("=" * 100)
    print("PRIMEIRO USO: PREENCHENDO O ARQUIVO LIMPO DO ZERO")
    print("=" * 100)

    # ---------------------------------------------------------------- cadastros
    ins = wb["CAD_Instituicoes"]
    escrever(ins, DAT, {2: "Ueno Bank", 3: "Banco", 7: "Conta e tarjeta"})
    escrever(ins, DAT + 1, {2: "Efectivo", 3: "Nao aplica", 7: "Dinheiro fisico"})

    con = wb["CAD_Contas"]
    escrever(con, DAT, {2: "Conta Ueno", 3: "Ueno Bank", 5: "Conta corrente",
                        6: 5_000_000, 7: dt.date(2025, 12, 31), 16: "Ativo"}, {7})
    escrever(con, DAT + 1, {2: "Efectivo", 3: "Efectivo", 5: "Carteira/dinheiro",
                            6: 500_000, 7: dt.date(2025, 12, 31), 16: "Ativo"}, {7})

    car = wb["CAD_Cartoes"]
    escrever(car, DAT, {2: "Tarjeta Ueno", 3: "Ueno Bank", 5: 10_000_000, 6: 1_000_000,
                        7: dt.date(2025, 12, 31), 8: 25, 9: 5, 22: "Ativo"}, {7})

    inv = wb["CAD_Investimentos"]
    escrever(inv, DAT, {2: "Fondo Mutuo Ueno", 3: "Ueno Bank", 5: "Fondo Mutuo",
                        6: 2_000_000, 7: dt.date(2025, 12, 31), 8: "Taxa anual",
                        9: 0.09, 10: 0, 24: "Ativo"}, {7})

    cmp_ = wb["CAD_Compromissos"]
    escrever(cmp_, DAT, {2: "Moto parcelada", 3: "Dividas e Financiamentos",
                         5: "Parcela de compra", 7: 12, 8: 500_000,
                         10: dt.date(2026, 3, 10), 11: 10, 12: "Mensal", 13: 0,
                         14: "CONTA", 15: "CON-000001", 25: "Ativo"}, {10})

    met = wb["CAD_Metas"]
    escrever(met, DAT, {2: "Reserva de emergencia", 3: 3_000_000,
                        4: dt.date(2027, 12, 31), 5: "CON-000001", 7: "NAO",
                        13: "Ativa"}, {4})

    # ------------------------------------------------------- cronograma (12x)
    par = wb["PARCELAS"]
    for i in range(12):
        venc = dt.date(2026, 3 + i, 10) if 3 + i <= 12 else dt.date(2027, 3 + i - 12, 10)
        escrever(par, DAT + i, {2: "CMP-000001", 3: i + 1, 4: venc, 5: 500_000}, {4})

    # ------------------------------------------------------------ lancamentos
    lan = wb["LANCAMENTOS"]
    linhas = [
        dict(data=dt.date(2026, 1, 30), desc="Salario de janeiro", tipo="RECEITA",
             st="REALIZADO", val=8_000_000, ot="EXTERNO", oi="", dtp="CONTA",
             di="CON-000001", cat="CAT-000010", sb=sub(10, "Salario"), fp="Transferencia"),
        dict(data=dt.date(2026, 2, 5), desc="Aluguel de fevereiro", tipo="DESPESA",
             st="REALIZADO", val=1_500_000, ot="CONTA", oi="CON-000001", dtp="EXTERNO",
             di="", cat="CAT-000003", sb=sub(3, "Aluguel"), fp="Transferencia"),
        dict(data=dt.date(2026, 2, 10), desc="Supermercado", tipo="COMPRA_CARTAO",
             st="REALIZADO", val=400_000, ot="CARTAO", oi="CAR-000001", dtp="EXTERNO",
             di="", cat="CAT-000001", sb=sub(1, "Supermercado"), fp="Credito"),
        dict(data=dt.date(2026, 2, 5), desc="Pagamento da fatura (divida inicial)",
             tipo="PAGAMENTO_CARTAO", st="REALIZADO", val=1_000_000, ot="CONTA",
             oi="CON-000001", dtp="CARTAO", di="CAR-000001", cat="CAT-000013",
             sb=sub(13, "Entre contas proprias"), fp="Transferencia"),
        dict(data=dt.date(2026, 2, 20), desc="Saque para efectivo", tipo="TRANSFERENCIA",
             st="REALIZADO", val=300_000, ot="CONTA", oi="CON-000001", dtp="CONTA",
             di="CON-000002", cat="CAT-000013", sb=sub(13, "Saque/Deposito"), fp="Dinheiro"),
        dict(data=dt.date(2026, 2, 25), desc="Aporte no Fondo Mutuo", tipo="APORTE",
             st="REALIZADO", val=1_000_000, ot="CONTA", oi="CON-000001",
             dtp="INVESTIMENTO", di="INV-000001", cat="CAT-000009", sb=sub(9, "Aporte"),
             fp="Transferencia"),
        dict(data=dt.date(2026, 3, 10), desc="Moto parcelada - parcela 1/12",
             tipo="PAGAMENTO_PARCELA", st="REALIZADO", val=500_000, ot="CONTA",
             oi="CON-000001", dtp="EXTERNO", di="", cat="CAT-000008",
             sb=sub(8, "Parcela de compra"), fp="Transferencia",
             cm="CMP-000001", pc="PAR-000001"),
    ]
    for i, l in enumerate(linhas):
        escrever(lan, DAT + i, {2: l["data"], 3: l["desc"], 4: l["tipo"], 5: l["st"],
                                6: l["val"], 7: l["ot"], 8: l["oi"], 9: l["dtp"],
                                10: l["di"], 11: l["cat"], 12: l["sb"], 13: l["fp"],
                                14: l.get("cm", ""), 15: l.get("pc", "")}, {2})

    # --------------------------------------------------------------- meta
    mov = wb["MOV_METAS"]
    escrever(mov, DAT, {2: dt.date(2026, 3, 15), 3: "MET-000001", 4: "RESERVA",
                        5: 1_000_000, 8: "Primeira reserva"}, {2})

    wb.save(TMP)
    rc = recalc(TMP, 600)
    print(f"\nRecalculo: {rc['status']}, {rc['total_errors']} celulas de erro em "
          f"{rc['total_formulas']} formulas\n")
    if rc["status"] != "success":
        print(rc)
        return 1

    v = openpyxl.load_workbook(TMP, data_only=True)
    c, k, i_, m, cm, p = (v["CAD_Contas"], v["CAD_Cartoes"], v["CAD_Investimentos"],
                          v["CAD_Metas"], v["CAD_Compromissos"], v["PARCELAS"])

    print("IDs gerados automaticamente:")
    for aba, col, n in [("CAD_Contas", 1, 2), ("CAD_Cartoes", 1, 1),
                        ("CAD_Investimentos", 1, 1), ("CAD_Compromissos", 1, 1),
                        ("CAD_Metas", 1, 1), ("PARCELAS", 1, 3), ("LANCAMENTOS", 1, 3)]:
        ids = [v[aba].cell(DAT + j, col).value for j in range(n)]
        print(f"    {aba:20s} {ids}")

    print("\nConferencia:")
    ok("Saldo da Conta Ueno", 8_700_000, c.cell(DAT, 10).value)
    ok("Saldo do Efectivo", 800_000, c.cell(DAT + 1, 10).value)
    ok("Reservado para metas na Conta Ueno", 1_000_000, c.cell(DAT, 11).value)
    ok("Saldo disponivel da Conta Ueno", 7_700_000, c.cell(DAT, 12).value)
    ok("Divida da Tarjeta Ueno", 400_000, k.cell(DAT, 12).value)
    ok("Limite disponivel da Tarjeta Ueno", 9_600_000, k.cell(DAT, 14).value)
    ok("Saldo do Fondo Mutuo", 3_000_000, i_.cell(DAT, 15).value)
    ok("Principal do Fondo Mutuo", 3_000_000, i_.cell(DAT, 16).value)
    ok("Valor reservado na meta", 1_000_000, m.cell(DAT, 8).value)
    ok("Percentual da meta", 1_000_000 / 3_000_000, m.cell(DAT, 9).value, tol=0.001)
    ok("Status da parcela 1/12", "PAGA", p.cell(DAT, 15).value)
    ok("Status da parcela 2/12 (venceu 10/04/2026)", "ATRASADA", p.cell(DAT + 1, 15).value)
    ok("Status da parcela 7/12 (vence 10/09/2026)", "ABERTA", p.cell(DAT + 6, 15).value)
    ok("Status da parcela 12/12 (vence 10/02/2027)", "PROJETADA", p.cell(DAT + 11, 15).value)
    ok("Parcelas atrasadas detectadas", 5,
       sum(1 for j in range(12) if p.cell(DAT + j, 15).value == "ATRASADA"))
    ok("Rotulo da parcela 1", "1/12", p.cell(DAT, 9).value)
    ok("Parcela atual do compromisso", "2/12", cm.cell(DAT, 23).value)
    ok("Valor total do compromisso", 6_000_000, cm.cell(DAT, 9).value)
    ok("Saldo devedor do compromisso", 5_500_000, cm.cell(DAT, 22).value)

    lanv = v["LANCAMENTOS"]
    ok("Transferencia nao virou receita", 0, lanv.cell(DAT + 4, 30).value)
    ok("Transferencia nao virou despesa", 0, lanv.cell(DAT + 4, 31).value)
    ok("Transferencia fora do fluxo de caixa", 0,
       (lanv.cell(DAT + 4, 28).value or 0) + (lanv.cell(DAT + 4, 29).value or 0))
    ok("Aporte nao virou despesa", 0, lanv.cell(DAT + 5, 31).value)
    ok("Pagamento de fatura nao virou despesa", 0, lanv.cell(DAT + 3, 31).value)
    ok("Compra no cartao virou despesa", 400_000, lanv.cell(DAT + 2, 31).value)
    ok("Compra no cartao fora do fluxo de caixa", 0,
       (lanv.cell(DAT + 2, 28).value or 0) + (lanv.cell(DAT + 2, 29).value or 0))
    erros = sum(1 for j in range(20)
                if str(lanv.cell(DAT + j, 36).value or "").startswith("ERRO"))
    ok("Lancamentos com erro de validacao", 0, erros)

    pat = {}
    for r in range(1, 40):
        kk = v["PATRIMONIO"].cell(r, 1).value
        if isinstance(kk, str):
            pat[kk] = v["PATRIMONIO"].cell(r, 2).value
    ok("Patrimonio liquido financeiro", 6_600_000, pat.get("Patrimonio liquido FINANCEIRO"))

    ts = v["TESTES"]
    falhas = exec_ = None
    for r in range(1, 12):
        if ts.cell(r, 1).value == "TESTES COM FALHA":
            falhas, exec_ = ts.cell(r, 2).value, ts.cell(r, 5).value
    ok(f"Bateria interna ({exec_} testes)", 0, falhas)
    if falhas:
        for r in range(1, ts.max_row + 1):
            if ts.cell(r, 7).value == "FALHA":
                print(f"      FALHA -> {ts.cell(r, 3).value}: esperado "
                      f"{gs(ts.cell(r, 4).value)} obtido {gs(ts.cell(r, 5).value)}")

    print("\n" + "=" * 100)
    print(f"RESULTADO: {len(OKS)} verificacoes OK, {len(FALHAS)} falhas")
    print("=" * 100)
    for f in FALHAS:
        print("  FALHA:", f)
    os.path.exists(TMP) and os.remove(TMP)
    return len(FALHAS)


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
