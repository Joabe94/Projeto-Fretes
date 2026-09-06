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
| `teste_app.mjs` | Abre o app no Chromium: confere os números contra o Excel, a ajuda das 20 telas, o guia e o layout de telefone |

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
- **Telas** (`Views`): 21 rotas renderizadas por string, com delegação de eventos.
- **Extrato** (`Calc.extrato`): movimentos de uma conta, cartão ou investimento em
  ordem de data, com saldo corrente a partir do saldo inicial. No cartão a lógica
  inverte — comprar aumenta a dívida, pagar reduz — e as colunas mudam de nome.
- **Ajuda** (`AJUDA`): para cada uma das 20 telas, o que ela faz, como se liga às
  outras (com link que navega) e os sintomas mais prováveis com o que verificar.
- **Guia** (`Guia`): três telas de boas-vindas explicando a regra origem/destino,
  o cartão de crédito e realizado/projetado; depois uma lista de 10 passos no
  painel que se marcam sozinhos conforme os dados aparecem.
- **Telefone**: barra inferior de 5 abas com folha para os menus longos, botão
  flutuante de lançamento e, abaixo de 860px, toda tabela vira lista de cartões —
  o papel de cada coluna no cartão é deduzido da própria definição da tabela,
  sem duplicar código por tela.

## Equivalência com o Excel

`teste_app.mjs` faz 116 verificações. Compara os valores do app com os que a
auditoria do Excel apurou (saldos das 3 contas, dívida e fatura dos 2 cartões,
saldo/principal/rendimento dos 2 investimentos, as 4 metas, a contagem dos 82
status de parcela, o terreno em 44/60, patrimônio, fluxo e projeção); percorre as
20 telas; abre a ajuda de cada uma exigindo as três seções preenchidas; cria um
perfil vazio e percorre o guia até o primeiro passo se marcar sozinho; e, num
viewport de 390x844, confere a barra inferior, a folha de menu, os cartões no
lugar das tabelas, a altura de toque dos botões e a ausência de rolagem
horizontal. Também confere que o saldo final do extrato de cada conta, cartão e
investimento é idêntico ao saldo do cadastro, que os filtros de lançamentos
funcionam e que cada tipo de cadastro pode ser criado do zero.
