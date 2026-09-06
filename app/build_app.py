#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Monta o arquivo final do app, embutindo o dataset de exemplo."""
import json
import os
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
FONTE = os.path.join(AQUI, "index.html")
DEMO = os.path.join(AQUI, "dados_demo.json")
DIST = os.path.join(AQUI, "dist")
SAIDA = os.path.join(DIST, "caja-guarani.html")
MARCA = "/*__DEMO__*/ null"

if __name__ == "__main__":
    if not os.path.exists(DEMO) or "--regerar" in sys.argv:
        subprocess.run([sys.executable, os.path.join(AQUI, "exportar_demo.py")], check=True)
    html = open(FONTE, encoding="utf-8").read()
    if MARCA not in html:
        sys.exit("marcador do dataset de exemplo nao encontrado em index.html")
    demo = json.load(open(DEMO, encoding="utf-8"))
    # </script> dentro de uma string quebraria o bloco; nao ocorre, mas garantimos
    js = json.dumps(demo, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    os.makedirs(DIST, exist_ok=True)
    open(SAIDA, "w", encoding="utf-8").write(html.replace(MARCA, js))
    print(f"{SAIDA}  {os.path.getsize(SAIDA)/1024:.0f} KB  "
          f"({len(demo['lancamentos'])} lancamentos de exemplo)")
