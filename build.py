#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera o arquivo Excel do Sistema Financeiro Pessoal V1."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "gerador"))
from comum import ARQUIVO_SAIDA
from planilha import construir

if __name__ == "__main__":
    destino = sys.argv[1] if len(sys.argv) > 1 else ARQUIVO_SAIDA
    L = construir(destino)
    print(f"Gerado: {destino}")
    print(f"Abas: {len(L.wb.sheetnames)}  |  Intervalos nomeados: {len(L.nomes)}")
