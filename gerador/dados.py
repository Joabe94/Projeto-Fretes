# -*- coding: utf-8 -*-
"""Dataset ficticio deterministico para teste completo do sistema (itens 43 a 45).

Gera cadastros, cronograma de parcelas, movimentacoes de metas e lancamentos
cobrindo jan/2026 a dez/2027 (realizado ate 05/09/2026, projetado adiante).
"""
from __future__ import annotations

import datetime as dt
import random

from comum import (DATA_INICIO, HOJE, add_meses, anomes, dia_do_mes, fim_do_mes,
                   id_fmt)

RNG = random.Random(20260905)   # semente fixa -> dataset reproduzivel

FIM_DADOS = dt.date(2027, 12, 31)


# ---------------------------------------------------------------------------
# CADASTROS
# ---------------------------------------------------------------------------
INSTITUICOES = [
    # id, nome, tipo, obs
    (id_fmt("INS", 1), "Ueno Bank", "Banco", "Conta corrente e tarjeta principal"),
    (id_fmt("INS", 2), "Banco Continental", "Banco", "Caja de ahorro, tarjeta e ahorro a plazo"),
    (id_fmt("INS", 3), "Banco Atlas", "Banco", "Fondo mutuo"),
    (id_fmt("INS", 4), "Efectivo", "Nao aplica", "Dinheiro fisico em maos"),
    (id_fmt("INS", 5), "Financiera El Comercio", "Financeira", "Emprestimo pessoal"),
    (id_fmt("INS", 6), "Inmobiliaria Santa Rita", "Outra", "Financiamento do terreno"),
]

CONTAS = [
    # id, nome, instituicao_id, tipo, saldo_inicial, data_saldo, status, obs
    (id_fmt("CON", 1), "Cuenta Corriente Ueno", id_fmt("INS", 1), "Conta corrente",
     12_500_000, DATA_INICIO, "Ativo", "Conta principal de movimento"),
    (id_fmt("CON", 2), "Caja de Ahorro Continental", id_fmt("INS", 2), "Conta poupanca",
     28_000_000, DATA_INICIO, "Ativo", "Conta de reservas e metas"),
    (id_fmt("CON", 3), "Efectivo (Billetera)", id_fmt("INS", 4), "Carteira/dinheiro",
     1_200_000, DATA_INICIO, "Ativo", "Dinheiro fisico"),
]

CARTOES = [
    # id, nome, instituicao_id, limite, divida_inicial, data, fechamento, vencimento, status, obs
    (id_fmt("CAR", 1), "Tarjeta Ueno Visa", id_fmt("INS", 1), 15_000_000, 3_000_000,
     DATA_INICIO, 25, 5, "Ativo", "Cartao de uso diario"),
    (id_fmt("CAR", 2), "Tarjeta Continental Mastercard", id_fmt("INS", 2), 3_000_000,
     1_500_000, DATA_INICIO, 20, 1, "Ativo", "Cartao secundario / assinaturas"),
]

INVESTIMENTOS = [
    # id, nome, inst, tipo, valor_inicial, data_inicial, base, taxa, prazo_meses, status, obs
    (id_fmt("INV", 1), "Fondo Mutuo Atlas Renta", id_fmt("INS", 3), "Fondo Mutuo",
     8_000_000, DATA_INICIO, "Taxa anual", 0.095, 0, "Ativo",
     "Liquidez diaria, sem vencimento; horizonte de projecao de 12 meses"),
    (id_fmt("INV", 2), "Ahorro a Plazo Continental 12m", id_fmt("INS", 2), "Ahorro a Plazo",
     0, dt.date(2026, 2, 15), "Taxa anual", 0.11, 12, "Ativo",
     "Constituido com aporte de Gs. 10.000.000 em 15/02/2026"),
]

CATEGORIAS = [
    (id_fmt("CAT", 1), "Alimentacao", "DESPESA"),
    (id_fmt("CAT", 2), "Transporte", "DESPESA"),
    (id_fmt("CAT", 3), "Moradia", "DESPESA"),
    (id_fmt("CAT", 4), "Saude", "DESPESA"),
    (id_fmt("CAT", 5), "Educacao", "DESPESA"),
    (id_fmt("CAT", 6), "Lazer", "DESPESA"),
    (id_fmt("CAT", 7), "Compras", "DESPESA"),
    (id_fmt("CAT", 8), "Dividas e Financiamentos", "DESPESA"),
    (id_fmt("CAT", 9), "Investimentos", "NEUTRO"),
    (id_fmt("CAT", 10), "Receitas", "RECEITA"),
    (id_fmt("CAT", 11), "Impostos e Taxas", "DESPESA"),
    (id_fmt("CAT", 12), "Outros", "DESPESA"),
    (id_fmt("CAT", 13), "Transferencias", "NEUTRO"),
]

_SUBS = {
    1: ["Supermercado", "Restaurante", "Delivery", "Lanches", "Feira"],
    2: ["Combustivel", "Uber/Taxi", "Onibus", "Manutencao", "Seguro veicular"],
    3: ["Aluguel", "Energia (ANDE)", "Agua (ESSAP)", "Internet", "Condominio", "Manutencao do lar"],
    4: ["Plano de saude", "Farmacia", "Consultas", "Exames", "Odontologia"],
    5: ["Mensalidade", "Cursos", "Livros", "Material escolar"],
    6: ["Viagens", "Streaming", "Eventos", "Bar e cafe"],
    7: ["Vestuario", "Eletronicos", "Casa e decoracao", "Presentes"],
    8: ["Parcela de terreno", "Parcela de emprestimo", "Juros e multas", "Parcela de compra"],
    9: ["Aporte", "Resgate", "Rendimento"],
    10: ["Salario", "Aguinaldo (13o)", "Honorarios/Freelance", "Rendimento financeiro",
         "Reembolso", "Outras receitas"],
    11: ["IVA", "IRP", "Tarifas bancarias", "Taxas municipais"],
    12: ["Diversos", "Nao classificado"],
    13: ["Entre contas proprias", "Saque/Deposito"],
}

SUBCATEGORIAS = []
_n = 0
for _c, _lista in _SUBS.items():
    for _s in _lista:
        _n += 1
        SUBCATEGORIAS.append((id_fmt("SUB", _n), id_fmt("CAT", _c), _s))

SUB = {}  # (cat_num, nome) -> sub_id
for _sid, _cid, _nome in SUBCATEGORIAS:
    SUB[(int(_cid.split("-")[1]), _nome)] = _sid


def sub(cat_num: int, nome: str) -> str:
    return SUB[(cat_num, nome)]


def cat(cat_num: int) -> str:
    return id_fmt("CAT", cat_num)


BENS = [
    # id, nome, tipo, valor, data, compromisso_vinculado, obs
    (id_fmt("BEM", 1), "Terreno Santa Rita (400 m2)", "Imovel", 108_000_000,
     dt.date(2023, 2, 10), id_fmt("CMP", 1), "Ativo nao financeiro; financiado em 60 parcelas"),
]

# ---------------------------------------------------------------------------
# COMPROMISSOS E CRONOGRAMA DE PARCELAS
# ---------------------------------------------------------------------------
COMPROMISSOS = [
    # id, nome, cat, sub, qtd, valor_parcela, primeira, dia, period, pagas_antes,
    # entidade_tipo, entidade_id, status, obs
    dict(id=id_fmt("CMP", 1), nome="Terreno Santa Rita", cat=cat(8),
         sub=sub(8, "Parcela de terreno"), qtd=60, valor=1_800_000,
         primeira=dt.date(2023, 2, 10), dia=10, period="Mensal", pagas_antes=35,
         ent_tipo="CONTA", ent_id=id_fmt("CON", 1), status="Ativo",
         obs="Parcelas 1 a 35 quitadas antes do inicio do sistema (ate dez/2025)"),
    dict(id=id_fmt("CMP", 2), nome="Notebook parcelado (Tarjeta Ueno)", cat=cat(7),
         sub=sub(7, "Eletronicos"), qtd=10, valor=850_000,
         primeira=dt.date(2026, 3, 15), dia=15, period="Mensal", pagas_antes=0,
         ent_tipo="CARTAO", ent_id=id_fmt("CAR", 1), status="Ativo",
         obs="Compra parcelada: cada parcela e lancada como COMPRA_CARTAO na data do vencimento"),
    dict(id=id_fmt("CMP", 3), nome="Emprestimo El Comercio", cat=cat(8),
         sub=sub(8, "Parcela de emprestimo"), qtd=12, valor=1_100_000,
         primeira=dt.date(2026, 1, 20), dia=20, period="Mensal", pagas_antes=0,
         ent_tipo="CONTA", ent_id=id_fmt("CON", 1), status="Ativo",
         obs="Parcela 8 (20/08/2026) nao paga -> teste de parcela ATRASADA"),
]

# Quantas parcelas foram efetivamente pagas dentro do sistema (geram lancamento)
PAGAS_NO_SISTEMA = {id_fmt("CMP", 1): 8,    # parcelas 36 a 43 (jan a ago/2026)
                    id_fmt("CMP", 2): 6,    # parcelas 1 a 6 (mar a ago/2026)
                    id_fmt("CMP", 3): 7}    # parcelas 1 a 7 (jan a jul/2026); a 8 fica atrasada


def gerar_parcelas():
    parcelas = []
    n = 0
    for c in COMPROMISSOS:
        for i in range(1, c["qtd"] + 1):
            n += 1
            venc = add_meses(c["primeira"], i - 1)
            venc = dia_do_mes(venc.year, venc.month, c["dia"])
            parcelas.append(dict(
                id=id_fmt("PAR", n), cmp=c["id"], nome=c["nome"], num=i, total=c["qtd"],
                rotulo=f"{i}/{c['qtd']}", venc=venc, valor=c["valor"],
                pago_antes="SIM" if i <= c["pagas_antes"] else "NAO",
                obs=""))
    return parcelas


PARCELAS = gerar_parcelas()
PARC_POR_CMP = {}
for p in PARCELAS:
    PARC_POR_CMP.setdefault(p["cmp"], []).append(p)

# ---------------------------------------------------------------------------
# METAS
# ---------------------------------------------------------------------------
METAS = [
    dict(id=id_fmt("MET", 1), nome="Comprar auto", objetivo=80_000_000,
         prazo=dt.date(2027, 12, 31), conta=id_fmt("CON", 2), inv="",
         fluxo="SIM", status="Ativa", conclusao="",
         obs="Reserva logica na Caja de Ahorro Continental"),
    dict(id=id_fmt("MET", 2), nome="Reserva de Emergencia", objetivo=30_000_000,
         prazo=dt.date(2026, 12, 31), conta=id_fmt("CON", 2), inv="",
         fluxo="NAO", status="Ativa", conclusao="",
         obs="Nao considerada no fluxo projetado (item 28 = NAO)"),
    dict(id=id_fmt("MET", 3), nome="Viagem em familia", objetivo=12_000_000,
         prazo=dt.date(2026, 6, 30), conta=id_fmt("CON", 2), inv="",
         fluxo="NAO", status="Cancelada", conclusao=dt.date(2026, 5, 15),
         obs="Cancelada em 15/05/2026; saldo transferido para a meta Comprar auto"),
    dict(id=id_fmt("MET", 4), nome="Trocar notebook", objetivo=8_500_000,
         prazo=dt.date(2026, 3, 31), conta=id_fmt("CON", 1), inv="",
         fluxo="NAO", status="Concluida", conclusao=dt.date(2026, 3, 15),
         obs="Concluida em 15/03/2026; saldo liberado para o saldo disponivel"),
]


def gerar_mov_metas():
    movs = []
    n = 0

    def add(data, meta, tipo, valor, destino="", meta_dest="", obs=""):
        nonlocal n
        n += 1
        movs.append(dict(id=id_fmt("MOV", n), data=data, meta=meta, tipo=tipo,
                         valor=valor, destino=destino, meta_dest=meta_dest, obs=obs))

    # MET-000004 "Trocar notebook": acumula e conclui em 15/03/2026
    add(dt.date(2026, 1, 20), id_fmt("MET", 4), "RESERVA", 4_250_000,
        obs="Reserva inicial para troca de notebook")
    add(dt.date(2026, 2, 20), id_fmt("MET", 4), "RESERVA", 4_250_000,
        obs="Segunda reserva - meta atingida")
    add(dt.date(2026, 3, 15), id_fmt("MET", 4), "LIBERACAO_CONCLUSAO", 8_500_000,
        destino="Saldo disponivel",
        obs="Meta concluida: valor liberado para o saldo disponivel e usado na compra parcelada")

    # MET-000001 "Comprar auto": reserva mensal de 1.500.000 (jan a ago/2026)
    for m in range(1, 9):
        add(dia_do_mes(2026, m, 28), id_fmt("MET", 1), "RESERVA", 1_500_000,
            obs=f"Reserva mensal {m:02d}/2026")
    # MET-000002 "Reserva de Emergencia": 800.000/mes (jan a ago/2026)
    for m in range(1, 9):
        add(dia_do_mes(2026, m, 28), id_fmt("MET", 2), "RESERVA", 800_000,
            obs=f"Reserva mensal {m:02d}/2026")
    # MET-000003 "Viagem": reservas jan a abr e cancelamento em 15/05/2026
    for m in range(1, 5):
        add(dia_do_mes(2026, m, 28), id_fmt("MET", 3), "RESERVA", 2_000_000,
            obs=f"Reserva mensal {m:02d}/2026")
    add(dt.date(2026, 5, 15), id_fmt("MET", 3), "TRANSF_SAIDA", 8_000_000,
        destino="Outra meta", meta_dest=id_fmt("MET", 1),
        obs="Cancelamento da meta: saldo transferido para Comprar auto")
    add(dt.date(2026, 5, 15), id_fmt("MET", 1), "TRANSF_ENTRADA", 8_000_000,
        obs="Recebido do cancelamento da meta Viagem em familia")
    movs.sort(key=lambda x: (x["data"], x["id"]))
    return movs


MOV_METAS = gerar_mov_metas()

# ---------------------------------------------------------------------------
# RECORRENCIAS
# ---------------------------------------------------------------------------
RECORRENCIAS = [
    dict(id=id_fmt("REC", 1), desc="Salario mensal", tipo="RECEITA", valor=12_000_000,
         period="Mensal", dia=30, inicio=dt.date(2026, 1, 30), ocorr=24,
         cat=cat(10), sub=sub(10, "Salario"), o_tipo="EXTERNO", o_id="",
         d_tipo="CONTA", d_id=id_fmt("CON", 1), forma="Transferencia", status="Ativa"),
    dict(id=id_fmt("REC", 2), desc="Honorarios contabeis (freelance)", tipo="RECEITA",
         valor=2_800_000, period="Mensal", dia=20, inicio=dt.date(2026, 1, 20), ocorr=24,
         cat=cat(10), sub=sub(10, "Honorarios/Freelance"), o_tipo="EXTERNO", o_id="",
         d_tipo="CONTA", d_id=id_fmt("CON", 1), forma="Transferencia", status="Ativa"),
    dict(id=id_fmt("REC", 3), desc="Aluguel", tipo="DESPESA", valor=2_000_000,
         period="Mensal", dia=5, inicio=dt.date(2026, 1, 5), ocorr=24,
         cat=cat(3), sub=sub(3, "Aluguel"), o_tipo="CONTA", o_id=id_fmt("CON", 1),
         d_tipo="EXTERNO", d_id="", forma="Transferencia", status="Ativa"),
    dict(id=id_fmt("REC", 4), desc="Plano de saude", tipo="DESPESA", valor=650_000,
         period="Mensal", dia=10, inicio=dt.date(2026, 1, 10), ocorr=24,
         cat=cat(4), sub=sub(4, "Plano de saude"), o_tipo="CONTA", o_id=id_fmt("CON", 1),
         d_tipo="EXTERNO", d_id="", forma="Debito automatico", status="Ativa"),
    dict(id=id_fmt("REC", 5), desc="Energia ANDE", tipo="DESPESA", valor=380_000,
         period="Mensal", dia=18, inicio=dt.date(2026, 1, 18), ocorr=24,
         cat=cat(3), sub=sub(3, "Energia (ANDE)"), o_tipo="CONTA", o_id=id_fmt("CON", 1),
         d_tipo="EXTERNO", d_id="", forma="Debito automatico", status="Ativa"),
    dict(id=id_fmt("REC", 6), desc="Internet fibra", tipo="COMPRA_CARTAO", valor=250_000,
         period="Mensal", dia=12, inicio=dt.date(2026, 1, 12), ocorr=24,
         cat=cat(3), sub=sub(3, "Internet"), o_tipo="CARTAO", o_id=id_fmt("CAR", 1),
         d_tipo="EXTERNO", d_id="", forma="Credito", status="Ativa"),
    dict(id=id_fmt("REC", 7), desc="Streaming", tipo="COMPRA_CARTAO", valor=90_000,
         period="Mensal", dia=8, inicio=dt.date(2026, 1, 8), ocorr=24,
         cat=cat(6), sub=sub(6, "Streaming"), o_tipo="CARTAO", o_id=id_fmt("CAR", 2),
         d_tipo="EXTERNO", d_id="", forma="Credito", status="Ativa"),
    dict(id=id_fmt("REC", 8), desc="Aguinaldo (13o salario)", tipo="RECEITA",
         valor=12_000_000, period="Anual", dia=20, inicio=dt.date(2026, 12, 20), ocorr=2,
         cat=cat(10), sub=sub(10, "Aguinaldo (13o)"), o_tipo="EXTERNO", o_id="",
         d_tipo="CONTA", d_id=id_fmt("CON", 1), forma="Transferencia", status="Ativa"),
    dict(id=id_fmt("REC", 9), desc="Transferencia mensal para poupanca", tipo="TRANSFERENCIA",
         valor=1_500_000, period="Mensal", dia=28, inicio=dt.date(2026, 1, 28), ocorr=24,
         cat=cat(13), sub=sub(13, "Entre contas proprias"), o_tipo="CONTA",
         o_id=id_fmt("CON", 1), d_tipo="CONTA", d_id=id_fmt("CON", 2),
         forma="Transferencia", status="Ativa"),
    dict(id=id_fmt("REC", 10), desc="Saque para efectivo", tipo="TRANSFERENCIA",
         valor=600_000, period="Mensal", dia=3, inicio=dt.date(2026, 1, 3), ocorr=24,
         cat=cat(13), sub=sub(13, "Saque/Deposito"), o_tipo="CONTA", o_id=id_fmt("CON", 1),
         d_tipo="CONTA", d_id=id_fmt("CON", 3), forma="Dinheiro", status="Ativa"),
    dict(id=id_fmt("REC", 11), desc="Aporte mensal no Fondo Mutuo", tipo="APORTE",
         valor=1_000_000, period="Mensal", dia=25, inicio=dt.date(2026, 3, 25), ocorr=22,
         cat=cat(9), sub=sub(9, "Aporte"), o_tipo="CONTA", o_id=id_fmt("CON", 1),
         d_tipo="INVESTIMENTO", d_id=id_fmt("INV", 1), forma="Transferencia", status="Ativa"),
]


# ---------------------------------------------------------------------------
# LANCAMENTOS
# ---------------------------------------------------------------------------
_seq = 0


def _novo_id():
    global _seq
    _seq += 1
    return id_fmt("LAN", _seq)


def _status(data: dt.date, forcado: str | None = None) -> str:
    if forcado:
        return forcado
    return "REALIZADO" if data <= HOJE else "PROJETADO"


def novo_lanc(data, desc, tipo, valor, o_tipo="EXTERNO", o_id="", d_tipo="EXTERNO",
              d_id="", categoria="", subcategoria="", forma="", status=None,
              cmp="", parcela="", meta="", rec="", rel="", obs=""):
    return dict(id=_novo_id(), data=data, desc=desc, tipo=tipo,
                status=_status(data, status), valor=int(round(valor)),
                o_tipo=o_tipo, o_id=o_id, d_tipo=d_tipo, d_id=d_id,
                cat=categoria, sub=subcategoria, forma=forma, cmp=cmp,
                parcela=parcela, meta=meta, rec=rec, rel=rel, obs=obs)


def _datas_recorrencia(r):
    """Expande a regra de recorrencia em datas de ocorrencia."""
    datas, d = [], r["inicio"]
    passo = {"Mensal": 1, "Bimestral": 2, "Trimestral": 3, "Semestral": 6, "Anual": 12}
    for i in range(r["ocorr"]):
        if r["period"] in passo:
            base = add_meses(r["inicio"], i * passo[r["period"]])
            d = dia_do_mes(base.year, base.month, r["dia"])
        elif r["period"] == "Semanal":
            d = r["inicio"] + dt.timedelta(days=7 * i)
        elif r["period"] == "Quinzenal":
            d = r["inicio"] + dt.timedelta(days=15 * i)
        elif r["period"] == "Diario":
            d = r["inicio"] + dt.timedelta(days=i)
        else:
            d = add_meses(r["inicio"], i)
        if d > FIM_DADOS:
            break
        datas.append(d)
    return datas


def _gerar_recorrentes(lancs):
    for r in RECORRENCIAS:
        for d in _datas_recorrencia(r):
            lancs.append(novo_lanc(
                d, r["desc"], r["tipo"], r["valor"], r["o_tipo"], r["o_id"],
                r["d_tipo"], r["d_id"], r["cat"], r["sub"], r["forma"],
                rec=r["id"], obs="Gerado pela recorrencia " + r["id"]))


def _gerar_parcelas_lanc(lancs):
    """Um lancamento por parcela nao quitada antes do sistema.

    Parcelas vencidas e pagas -> REALIZADO. Parcelas futuras (e a atrasada) -> PROJETADO.
    Nunca dois lancamentos para a mesma parcela (teste de duplicidade, item 45).
    """
    cmp_por_id = {c["id"]: c for c in COMPROMISSOS}
    for cid, plist in PARC_POR_CMP.items():
        c = cmp_por_id[cid]
        pagas_sis = PAGAS_NO_SISTEMA[cid]
        pagas_antes = c["pagas_antes"]
        for p in plist:
            if p["pago_antes"] == "SIM":
                continue                      # quitada antes do sistema: sem lancamento
            idx = p["num"] - pagas_antes      # 1..n dentro do sistema
            paga = idx <= pagas_sis
            if c["ent_tipo"] == "CARTAO":
                tipo, o_tipo, o_id, forma = "COMPRA_CARTAO", "CARTAO", c["ent_id"], "Credito"
            else:
                tipo, o_tipo, o_id, forma = "PAGAMENTO_PARCELA", "CONTA", c["ent_id"], "Transferencia"
            st = "REALIZADO" if paga else "PROJETADO"
            lancs.append(novo_lanc(
                p["venc"], f"{c['nome']} - parcela {p['rotulo']}", tipo, p["valor"],
                o_tipo, o_id, "EXTERNO", "", c["cat"], c["sub"], forma, status=st,
                cmp=cid, parcela=p["id"],
                obs="Pagamento da parcela" if paga else "Parcela ainda nao paga"))


def _gerar_investimentos(lancs):
    inv1, inv2 = id_fmt("INV", 1), id_fmt("INV", 2)
    c1, c2 = id_fmt("CON", 1), id_fmt("CON", 2)

    # Constituicao do Ahorro a Plazo
    lancs.append(novo_lanc(dt.date(2026, 2, 15), "Constituicao Ahorro a Plazo 12m",
                           "APORTE", 10_000_000, "CONTA", c2, "INVESTIMENTO", inv2,
                           cat(9), sub(9, "Aporte"), "Transferencia",
                           obs="Aporte inicial; nao e despesa"))
    # Rendimento e resgate do Ahorro a Plazo no vencimento (projetado)
    venc2 = dt.date(2027, 2, 15)
    lancs.append(novo_lanc(venc2, "Rendimento Ahorro a Plazo (vencimento)", "RENDIMENTO",
                           1_100_000, "EXTERNO", "", "INVESTIMENTO", inv2,
                           cat(10), sub(10, "Rendimento financeiro"), "Transferencia",
                           obs="Rendimento de 11% a.a. sobre Gs. 10.000.000"))
    lancs.append(novo_lanc(venc2, "Resgate total Ahorro a Plazo", "RESGATE", 11_100_000,
                           "INVESTIMENTO", inv2, "CONTA", c2, cat(9), sub(9, "Resgate"),
                           "Transferencia",
                           obs="Resgate total: principal + rendimento ja reconhecido"))

    # Fondo Mutuo: rendimento mensal composto sobre o saldo, competencia no fim do mes
    taxa_m = (1 + 0.095) ** (1 / 12) - 1
    saldo = 8_000_000
    aportes = {d: 1_000_000 for r in RECORRENCIAS if r["id"] == id_fmt("REC", 11)
               for d in _datas_recorrencia(r)}
    resgate_data, resgate_valor = dt.date(2026, 7, 20), 3_000_000
    lancs.append(novo_lanc(resgate_data, "Resgate parcial Fondo Mutuo Atlas", "RESGATE",
                           resgate_valor, "INVESTIMENTO", inv1, "CONTA", c1,
                           cat(9), sub(9, "Resgate"), "Transferencia",
                           obs="Resgate parcial; principal nao e receita"))
    d = dt.date(2026, 1, 1)
    while d <= FIM_DADOS:
        fim = fim_do_mes(d)
        for ad, av in aportes.items():
            if ad.year == d.year and ad.month == d.month:
                saldo += av
        if resgate_data.year == d.year and resgate_data.month == d.month:
            saldo -= resgate_valor
        rend = int(round(saldo * taxa_m, -3))
        lancs.append(novo_lanc(fim, "Rendimento Fondo Mutuo Atlas", "RENDIMENTO", rend,
                               "EXTERNO", "", "INVESTIMENTO", inv1, cat(10),
                               sub(10, "Rendimento financeiro"), "Transferencia",
                               obs="Rendimento mensal capitalizado no proprio fundo"))
        saldo += rend
        d = add_meses(d, 1)


def _gerar_variaveis(lancs):
    car1, car2 = id_fmt("CAR", 1), id_fmt("CAR", 2)
    c1, c3 = id_fmt("CON", 1), id_fmt("CON", 3)
    d = dt.date(2026, 1, 1)
    while d <= FIM_DADOS:
        ano, mes = d.year, d.month
        detalhado = fim_do_mes(d) <= HOJE          # meses ja fechados -> lancamento a lancamento
        if detalhado:
            for i, dia in enumerate((6, 14, 22, 28), start=1):
                lancs.append(novo_lanc(dia_do_mes(ano, mes, dia),
                                       f"Supermercado (compra {i})", "COMPRA_CARTAO",
                                       RNG.randrange(600, 951) * 1000, "CARTAO", car1,
                                       "EXTERNO", "", cat(1), sub(1, "Supermercado"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 9), "Feira do bairro", "DESPESA",
                                   RNG.randrange(150, 251) * 1000, "CONTA", c3,
                                   "EXTERNO", "", cat(1), sub(1, "Feira"), "Dinheiro"))
            for dia in (7, 21):
                lancs.append(novo_lanc(dia_do_mes(ano, mes, dia), "Combustivel", "COMPRA_CARTAO",
                                       RNG.randrange(300, 421) * 1000, "CARTAO", car1,
                                       "EXTERNO", "", cat(2), sub(2, "Combustivel"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 16), "Restaurante", "COMPRA_CARTAO",
                                   RNG.randrange(120, 231) * 1000, "CARTAO", car1,
                                   "EXTERNO", "", cat(1), sub(1, "Restaurante"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 24), "Delivery", "DESPESA",
                                   RNG.randrange(80, 161) * 1000, "CONTA", c3,
                                   "EXTERNO", "", cat(1), sub(1, "Delivery"), "Dinheiro"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 11), "Farmacia", "COMPRA_CARTAO",
                                   RNG.randrange(80, 181) * 1000, "CARTAO", car2,
                                   "EXTERNO", "", cat(4), sub(4, "Farmacia"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 19), "Uber / Taxi", "DESPESA",
                                   RNG.randrange(50, 91) * 1000, "CONTA", c3,
                                   "EXTERNO", "", cat(2), sub(2, "Uber/Taxi"), "Dinheiro"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 23), "Compras online", "COMPRA_CARTAO",
                                   RNG.randrange(250, 601) * 1000, "CARTAO", car2,
                                   "EXTERNO", "", cat(7), sub(7, "Casa e decoracao"), "Credito"))
            if mes in (3, 6, 8, 11):
                lancs.append(novo_lanc(dia_do_mes(ano, mes, 17), "Vestuario", "COMPRA_CARTAO",
                                       RNG.randrange(350, 801) * 1000, "CARTAO", car2,
                                       "EXTERNO", "", cat(7), sub(7, "Vestuario"), "Credito"))
        else:
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 14), "Supermercado (estimativa do mes)",
                                   "COMPRA_CARTAO", 2_300_000, "CARTAO", car1, "EXTERNO", "",
                                   cat(1), sub(1, "Supermercado"), "Credito",
                                   obs="Projecao com base na media dos meses realizados"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 7), "Combustivel (estimativa do mes)",
                                   "COMPRA_CARTAO", 720_000, "CARTAO", car1, "EXTERNO", "",
                                   cat(2), sub(2, "Combustivel"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 16), "Alimentacao fora (estimativa)",
                                   "DESPESA", 400_000, "CONTA", c3, "EXTERNO", "",
                                   cat(1), sub(1, "Restaurante"), "Dinheiro"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 28), "Supermercado fim de mes (estimativa)",
                                   "COMPRA_CARTAO", 780_000, "CARTAO", car1, "EXTERNO", "",
                                   cat(1), sub(1, "Supermercado"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 23), "Compras online (estimativa)",
                                   "COMPRA_CARTAO", 400_000, "CARTAO", car2, "EXTERNO", "",
                                   cat(7), sub(7, "Casa e decoracao"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 11), "Farmacia (estimativa)",
                                   "COMPRA_CARTAO", 130_000, "CARTAO", car2, "EXTERNO", "",
                                   cat(4), sub(4, "Farmacia"), "Credito"))
            lancs.append(novo_lanc(dia_do_mes(ano, mes, 19), "Transporte diverso (estimativa)",
                                   "DESPESA", 180_000, "CONTA", c3, "EXTERNO", "",
                                   cat(2), sub(2, "Uber/Taxi"), "Dinheiro"))
        lancs.append(novo_lanc(fim_do_mes(d), "Tarifas bancarias", "DESPESA", 35_000,
                               "CONTA", c1, "EXTERNO", "", cat(11), sub(11, "Tarifas bancarias"),
                               "Debito automatico"))
        d = add_meses(d, 1)


def _proximo_dia(base: dt.date, dia: int, estrito: bool = False) -> dt.date:
    """Primeira data com o dia informado que seja >= base (> base se estrito)."""
    cand = dia_do_mes(base.year, base.month, dia)
    if cand < base or (estrito and cand <= base):
        prox = add_meses(dt.date(base.year, base.month, 1), 1)
        cand = dia_do_mes(prox.year, prox.month, dia)
    return cand


def _gerar_pagamento_cartoes(lancs):
    """Fecha cada fatura e gera o pagamento CONTA -> CARTAO (operacao neutra)."""
    conta_pag = {id_fmt("CAR", 1): id_fmt("CON", 1), id_fmt("CAR", 2): id_fmt("CON", 1)}
    for cid, nome, _inst, _lim, div_ini, _dt, fech, venc, *_ in CARTOES:
        compras = [l for l in lancs
                   if l["o_tipo"] == "CARTAO" and l["o_id"] == cid and l["status"] != "CANCELADO"]
        # 1) fatura que carrega a divida inicial
        f0 = _proximo_dia(DATA_INICIO - dt.timedelta(days=31), fech)
        while _proximo_dia(f0 + dt.timedelta(days=1), fech) <= DATA_INICIO:
            f0 = _proximo_dia(f0 + dt.timedelta(days=1), fech)
        v0 = _proximo_dia(f0, venc, estrito=True)
        if div_ini:
            lancs.append(novo_lanc(v0, f"Pagamento fatura {nome} (divida inicial)",
                                   "PAGAMENTO_CARTAO", div_ini, "CONTA", conta_pag[cid],
                                   "CARTAO", cid, cat(13), sub(13, "Entre contas proprias"),
                                   "Transferencia",
                                   obs="Quitacao da divida inicial; nao e despesa"))
        # 2) faturas seguintes
        fech_ant = f0
        while True:
            fech_atual = _proximo_dia(fech_ant + dt.timedelta(days=1), fech)
            if fech_atual > FIM_DADOS:
                break
            total = sum(l["valor"] for l in compras if fech_ant < l["data"] <= fech_atual)
            if total:
                vdata = _proximo_dia(fech_atual, venc, estrito=True)
                lancs.append(novo_lanc(
                    vdata, f"Pagamento fatura {nome} (fech. {fech_atual.strftime('%d/%m/%Y')})",
                    "PAGAMENTO_CARTAO", total, "CONTA", conta_pag[cid], "CARTAO", cid,
                    cat(13), sub(13, "Entre contas proprias"), "Transferencia",
                    obs="Reduz o saldo da conta e a divida do cartao; nao gera nova despesa"))
            fech_ant = fech_atual


def _gerar_casos_de_erro(lancs):
    """Item 44: lancamento errado, cancelamento e estorno."""
    lancs.append(novo_lanc(dt.date(2026, 4, 12), "Supermercado (lancamento em duplicidade)",
                           "COMPRA_CARTAO", 685_000, "CARTAO", id_fmt("CAR", 1),
                           "EXTERNO", "", cat(1), sub(1, "Supermercado"), "Credito",
                           status="CANCELADO",
                           obs="CANCELADO: digitado duas vezes. Mantido no historico, sem efeito financeiro"))
    lancs.append(novo_lanc(dt.date(2026, 6, 8), "Reembolso de despesa medica", "RECEITA",
                           420_000, "EXTERNO", "", "CONTA", id_fmt("CON", 1),
                           cat(10), sub(10, "Reembolso"), "Transferencia",
                           obs="Devolucao do plano de saude"))


def gerar_lancamentos():
    global _seq
    _seq = 0
    lancs: list[dict] = []
    _gerar_recorrentes(lancs)
    _gerar_parcelas_lanc(lancs)
    _gerar_investimentos(lancs)
    _gerar_variaveis(lancs)
    _gerar_casos_de_erro(lancs)
    _gerar_pagamento_cartoes(lancs)
    lancs.sort(key=lambda x: (x["data"], x["id"]))
    for i, l in enumerate(lancs, start=1):       # renumera na ordem cronologica
        l["id"] = id_fmt("LAN", i)
    return lancs


LANCAMENTOS = gerar_lancamentos()
