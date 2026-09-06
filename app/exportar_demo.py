#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Exporta o dataset ficticio do Excel para JSON, no formato que o app web usa.

Garante que o exemplo do sistema web e exatamente o mesmo que foi auditado no
Excel: mesma semente, mesmos 602 lancamentos, mesmos saldos.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "gerador"))

import comum as K
import dados as D


def d(x):
    return x.isoformat() if isinstance(x, (dt.date, dt.datetime)) else x


def limpo(obj: dict) -> dict:
    """Remove campos vazios para encolher o JSON embutido na pagina."""
    return {k: v for k, v in obj.items() if v not in ("", None)}


def montar():
    cfg = dict(moeda="Gs.", dataRef=d(K.HOJE), inicio=d(K.DATA_INICIO),
               ano=K.ANO_BASE, mes=K.HOJE.month, horizonte=K.HORIZONTE_MESES,
               reservaNoFluxo="SIM", saldoMin=4_000_000, pctCartao=0.80,
               diasAlerta=7, diasInv=60, pctRenda=0.40, pctDesp=0.25)

    dados = {
        "config": cfg,
        "instituicoes": [limpo(dict(id=i[0], nome=i[1], tipo=i[2], obs=i[3]))
                         for i in D.INSTITUICOES],
        "contas": [limpo(dict(id=c[0], nome=c[1], instituicaoId=c[2], tipo=c[3],
                              saldoInicial=c[4], dataSaldoInicial=d(c[5]),
                              status=c[6], obs=c[7])) for c in D.CONTAS],
        "cartoes": [limpo(dict(id=c[0], nome=c[1], instituicaoId=c[2], limite=c[3],
                               dividaInicial=c[4], dataDivida=d(c[5]),
                               diaFechamento=c[6], diaVencimento=c[7],
                               status=c[8], obs=c[9])) for c in D.CARTOES],
        "investimentos": [limpo(dict(id=i[0], nome=i[1], instituicaoId=i[2], tipo=i[3],
                                     valorInicial=i[4], dataInicial=d(i[5]),
                                     baseRendimento=i[6], taxa=i[7], prazoMeses=i[8],
                                     status=i[9], obs=i[10]))
                          for i in D.INVESTIMENTOS],
        "categorias": [limpo(dict(id=c[0], nome=c[1], tipoPadrao=c[2]))
                       for c in D.CATEGORIAS],
        "subcategorias": [limpo(dict(id=s[0], categoriaId=s[1], nome=s[2]))
                          for s in D.SUBCATEGORIAS],
        "compromissos": [limpo(dict(id=c["id"], nome=c["nome"], categoriaId=c["cat"],
                                    subcategoriaId=c["sub"], qtd=c["qtd"],
                                    valorParcela=c["valor"], primeira=d(c["primeira"]),
                                    dia=c["dia"], periodicidade=c["period"],
                                    pagasAntes=c["pagas_antes"], entidadeTipo=c["ent_tipo"],
                                    entidadeId=c["ent_id"], status=c["status"],
                                    obs=c["obs"])) for c in D.COMPROMISSOS],
        "parcelas": [limpo(dict(id=p["id"], compromissoId=p["cmp"], num=p["num"],
                                vencimento=d(p["venc"]), valor=p["valor"], obs=p["obs"]))
                     for p in D.PARCELAS],
        "metas": [limpo(dict(id=m["id"], nome=m["nome"], objetivo=m["objetivo"],
                             prazo=d(m["prazo"]), contaId=m["conta"],
                             investimentoId=m["inv"], considerarFluxo=m["fluxo"],
                             status=m["status"], conclusao=d(m["conclusao"]),
                             obs=m["obs"])) for m in D.METAS],
        "movMetas": [limpo(dict(id=m["id"], data=d(m["data"]), metaId=m["meta"],
                                tipo=m["tipo"], valor=m["valor"], destino=m["destino"],
                                metaDestinoId=m["meta_dest"], obs=m["obs"]))
                     for m in D.MOV_METAS],
        "recorrencias": [limpo(dict(id=r["id"], descricao=r["desc"], tipo=r["tipo"],
                                    valor=r["valor"], periodicidade=r["period"],
                                    dia=r["dia"], inicio=d(r["inicio"]),
                                    ocorrencias=r["ocorr"], categoriaId=r["cat"],
                                    subcategoriaId=r["sub"], origemTipo=r["o_tipo"],
                                    origemId=r["o_id"], destinoTipo=r["d_tipo"],
                                    destinoId=r["d_id"], forma=r["forma"],
                                    status=r["status"])) for r in D.RECORRENCIAS],
        "bens": [limpo(dict(id=b[0], nome=b[1], tipo=b[2], valor=b[3],
                            dataAquisicao=d(b[4]), compromissoId=b[5], status="Ativo",
                            obs=b[6])) for b in D.BENS],
        "lancamentos": [limpo(dict(id=l["id"], data=d(l["data"]), descricao=l["desc"],
                                   tipo=l["tipo"], status=l["status"], valor=l["valor"],
                                   origemTipo=l["o_tipo"], origemId=l["o_id"],
                                   destinoTipo=l["d_tipo"], destinoId=l["d_id"],
                                   categoriaId=l["cat"], subcategoriaId=l["sub"],
                                   forma=l["forma"], compromissoId=l["cmp"],
                                   parcelaId=l["parcela"], metaId=l["meta"],
                                   recorrenciaId=l["rec"], relacionadoId=l["rel"],
                                   obs=l["obs"])) for l in D.LANCAMENTOS],
    }
    return dados


if __name__ == "__main__":
    dados = montar()
    destino = os.path.join(AQUI, "dados_demo.json")
    with open(destino, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, separators=(",", ":"))
    tam = os.path.getsize(destino)
    print(f"{destino}  {tam/1024:.0f} KB")
    for k, v in dados.items():
        if isinstance(v, list):
            print(f"   {k:16s} {len(v)}")
