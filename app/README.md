# Caja Guaraní — versão web

App de página única, sem build, que reproduz o mesmo modelo financeiro auditado
no Excel. Publicado como Artifact do claude.ai, com os dados salvos no banco de
dados do próprio artefato (capacidade `db`) e cópia de segurança pela capacidade
`downloads`.

| Arquivo | Papel |
|---|---|
| `index.html` | Fonte do app (11 blocos de script, sem framework nem CDN de código) |
| `exportar_demo.py` | Exporta o dataset fictício de `gerador/dados.py` para JSON |
| `build_app.py` | Embute o JSON em `index.html` e grava `dist/caja-guarani.html` |
| `teste_app.mjs` | Abre o app no Chromium e confere 62 números contra o Excel |

```bash
python3 app/build_app.py --regerar   # gera dist/caja-guarani.html
node app/teste_app.mjs               # teste de fumaça + conferência numérica
```

## Arquitetura

- **Armazenamento** (`Store`): documentos em blocos no banco do artefato.
  `perfis/<id>` guarda o metadado; `perfis/<id>/dados/cadastros` os cadastros;
  `perfis/<id>/dados/{lancamentos,parcelas,movMetas}_<n>` os fatos, em blocos de
  150 itens para ficar bem abaixo do limite de 256 KiB por documento. Sem banco
  disponível, cai para `localStorage`.
- **Motor** (`Calc.derivar`): recebe os dados brutos e devolve tudo o que as
  telas mostram — saldos, dívidas, ciclos de cartão, status de parcela, reservas
  de meta, fluxo diário/semanal/mensal/anual, projeção, patrimônio e evolução.
- **Conferência** (`Calc.testes`): 41 verificações que recalculam cada número por
  um caminho diferente. A tela Conferência mostra o resultado ao usuário.
- **Telas** (`Views`): 20 rotas renderizadas por string, com delegação de eventos.

## Equivalência com o Excel

`teste_app.mjs` compara 62 valores do app com os que a auditoria do Excel
apurou: saldos das 3 contas, dívida e fatura dos 2 cartões, saldo/principal/
rendimento dos 2 investimentos, as 4 metas, a contagem dos 82 status de parcela,
o compromisso do terreno em 44/60, patrimônio, fluxo e projeção. Todos batem.
