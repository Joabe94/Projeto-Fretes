# -*- coding: utf-8 -*-
"""Motor de calculo independente em Python.

Reimplementa as regras financeiras a partir dos dados brutos, sem usar as
formulas do Excel. Serve de referencia para a auditoria (item 46/47): os
numeros daqui sao comparados com os que o Excel calcula.
"""
from __future__ import annotations

import datetime as dt
from collections import defaultdict

from comum import CLASSIFICACAO, HOJE, SINAL_MOV_META, add_meses, fim_do_mes
import dados as D


def _ativo(l):
    return l["status"] in ("REALIZADO", "PROJETADO")


def saldos_contas(lancs, status=("REALIZADO",), ate: dt.date | None = None):
    saldo = {}
    for c in D.CONTAS:
        saldo[c[0]] = c[4]
    for l in lancs:
        if l["status"] not in status:
            continue
        if ate and l["data"] > ate:
            continue
        if l["d_tipo"] == "CONTA" and l["d_id"] in saldo:
            saldo[l["d_id"]] += l["valor"]
        if l["o_tipo"] == "CONTA" and l["o_id"] in saldo:
            saldo[l["o_id"]] -= l["valor"]
    return saldo


def dividas_cartoes(lancs, status=("REALIZADO",), ate: dt.date | None = None):
    div = {c[0]: c[4] for c in D.CARTOES}
    for l in lancs:
        if l["status"] not in status:
            continue
        if ate and l["data"] > ate:
            continue
        if l["o_tipo"] == "CARTAO" and l["o_id"] in div:
            div[l["o_id"]] += l["valor"]
        if l["d_tipo"] == "CARTAO" and l["d_id"] in div:
            div[l["d_id"]] -= l["valor"]
    return div


def saldos_investimentos(lancs, status=("REALIZADO",), ate: dt.date | None = None):
    """Retorna dict id -> (saldo, principal, rendimento_no_saldo, aportes, resgates, rend)."""
    ini = {i[0]: i[4] for i in D.INVESTIMENTOS}
    ap = defaultdict(int); rg = defaultdict(int); rd = defaultdict(int)
    for l in lancs:
        if l["status"] not in status:
            continue
        if ate and l["data"] > ate:
            continue
        if l["d_tipo"] == "INVESTIMENTO" and l["d_id"] in ini:
            (rd if l["tipo"] == "RENDIMENTO" else ap)[l["d_id"]] += l["valor"]
        if l["o_tipo"] == "INVESTIMENTO" and l["o_id"] in ini:
            rg[l["o_id"]] += l["valor"]
    out = {}
    for iid, v0 in ini.items():
        saldo = v0 + ap[iid] + rd[iid] - rg[iid]
        principal = max(0, v0 + ap[iid] - rg[iid])
        out[iid] = dict(saldo=saldo, principal=principal, rend_no_saldo=saldo - principal,
                        aportes=ap[iid], resgates=rg[iid], rendimento=rd[iid])
    return out


def saldos_metas(movs, ate: dt.date | None = None):
    s = defaultdict(int)
    for m in movs:
        if ate and m["data"] > ate:
            continue
        s[m["meta"]] += SINAL_MOV_META[m["tipo"]] * m["valor"]
    return {m["id"]: s.get(m["id"], 0) for m in D.METAS}


def reservado_por_conta(movs, metas, ate=None):
    conta_de = {m["id"]: m["conta"] for m in metas}
    r = defaultdict(int)
    for m in movs:
        if ate and m["data"] > ate:
            continue
        r[conta_de.get(m["meta"], "")] += SINAL_MOV_META[m["tipo"]] * m["valor"]
    return r


def status_parcelas(parcelas, lancs, hoje=HOJE):
    pagto_real = defaultdict(int)
    for l in lancs:
        if l["parcela"] and l["status"] == "REALIZADO":
            pagto_real[l["parcela"]] += l["valor"]
    fim_mes = fim_do_mes(hoje)
    out = {}
    for p in parcelas:
        if p["pago_antes"] == "SIM" or pagto_real.get(p["id"], 0) > 0:
            st = "PAGA"
        elif p["venc"] < hoje:
            st = "ATRASADA"
        elif p["venc"] <= fim_mes:
            st = "ABERTA"
        else:
            st = "PROJETADA"
        out[p["id"]] = st
    return out


def resultado_periodo(lancs, ini, fim, status="REALIZADO"):
    r = dict(receita=0, despesa=0, rec_fin=0, entrada=0, saida=0)
    for l in lancs:
        if l["status"] != status or not (ini <= l["data"] <= fim):
            continue
        cl = CLASSIFICACAO[l["tipo"]]
        if cl == "RECEITA":
            r["receita"] += l["valor"]
        elif cl == "DESPESA":
            r["despesa"] += l["valor"]
        elif cl == "RECEITA_FINANCEIRA":
            r["rec_fin"] += l["valor"]
        if l["d_tipo"] == "CONTA" and l["o_tipo"] != "CONTA":
            r["entrada"] += l["valor"]
        if l["o_tipo"] == "CONTA" and l["d_tipo"] != "CONTA":
            r["saida"] += l["valor"]
    return r


def patrimonio(lancs, movs, parcelas, ate, status=("REALIZADO",)):
    contas = sum(saldos_contas(lancs, status, ate).values())
    inv = sum(v["saldo"] for v in saldos_investimentos(lancs, status, ate).values())
    cart = sum(dividas_cartoes(lancs, status, ate).values())
    st = status_parcelas(parcelas, lancs, hoje=ate if ate else HOJE)
    devedor = sum(p["valor"] for p in parcelas if st[p["id"]] != "PAGA")
    bens = sum(b[3] for b in D.BENS)
    return dict(contas=contas, investimentos=inv, cartoes=cart, compromissos=devedor,
                bens=bens,
                pl_financeiro=contas + inv - cart - devedor,
                pl_total=contas + inv + bens - cart - devedor)


def fluxo_mensal(lancs, meses_ini: dt.date, n: int):
    """Lista de dicts por mes com entradas/saidas realizadas e projetadas."""
    saldo_ini_total = sum(c[4] for c in D.CONTAS)
    linhas, saldo = [], saldo_ini_total
    d = dt.date(meses_ini.year, meses_ini.month, 1)
    # movimentos anteriores ao inicio da grade
    for l in lancs:
        if l["status"] in ("REALIZADO", "PROJETADO") and l["data"] < d:
            if l["d_tipo"] == "CONTA" and l["o_tipo"] != "CONTA":
                saldo += l["valor"]
            if l["o_tipo"] == "CONTA" and l["d_tipo"] != "CONTA":
                saldo -= l["valor"]
    for _ in range(n):
        ini, fim = d, fim_do_mes(d)
        rr = resultado_periodo(lancs, ini, fim, "REALIZADO")
        rp = resultado_periodo(lancs, ini, fim, "PROJETADO")
        ini_saldo = saldo
        saldo = saldo + rr["entrada"] - rr["saida"] + rp["entrada"] - rp["saida"]
        linhas.append(dict(ano=d.year, mes=d.month, ini=ini, fim=fim,
                           saldo_ini=ini_saldo,
                           ent_real=rr["entrada"], sai_real=rr["saida"],
                           ent_proj=rp["entrada"], sai_proj=rp["saida"],
                           saldo_fim=saldo))
        d = add_meses(d, 1)
    return linhas


def evolucao_patrimonial(lancs, parcelas, ini: dt.date, n: int, hoje=HOJE):
    """Mesma definicao usada na aba PATRIMONIO: realizado + projetado ate o fim do mes."""
    st = status_parcelas(parcelas, lancs, hoje=hoje)
    bens = sum(b[3] for b in D.BENS)
    linhas, d = [], dt.date(ini.year, ini.month, 1)
    for _ in range(n):
        fim = fim_do_mes(d)
        amb = ("REALIZADO", "PROJETADO")
        contas = sum(saldos_contas(lancs, amb, fim).values())
        inv = sum(v["saldo"] for v in saldos_investimentos(lancs, amb, fim).values())
        cart = sum(dividas_cartoes(lancs, amb, fim).values())
        dev = sum(p["valor"] for p in parcelas
                  if p["venc"] > fim or (p["venc"] <= fim and st[p["id"]] == "ATRASADA"))
        linhas.append(dict(ano=d.year, mes=d.month, fim=fim, contas=contas,
                           investimentos=inv, cartoes=cart, compromissos=dev, bens=bens,
                           pl_fin=contas + inv - cart - dev,
                           pl_total=contas + inv + bens - cart - dev))
        d = add_meses(d, 1)
    return linhas
