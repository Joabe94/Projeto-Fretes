#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera os arquivos Excel do Sistema Financeiro Pessoal V1.

    python3 build.py            gera os dois arquivos (exemplo e limpo)
    python3 build.py exemplo    gera apenas o arquivo com dados ficticios
    python3 build.py limpo      gera apenas o arquivo vazio
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "gerador"))

from planilha import construir

EXEMPLO = "SISTEMA_FINANCEIRO_PESSOAL_V1_EXEMPLO.xlsx"
LIMPO = "SISTEMA_FINANCEIRO_PESSOAL_V1_LIMPO.xlsx"

if __name__ == "__main__":
    modo = sys.argv[1].lower() if len(sys.argv) > 1 else "ambos"
    if modo in ("ambos", "exemplo"):
        L = construir(EXEMPLO, limpo=False)
        print(f"Gerado: {EXEMPLO}  ({len(L.wb.sheetnames)} abas, {len(L.nomes)} nomes)")
    if modo in ("ambos", "limpo"):
        L = construir(LIMPO, limpo=True)
        print(f"Gerado: {LIMPO}  ({len(L.wb.sheetnames)} abas, {len(L.nomes)} nomes)")
