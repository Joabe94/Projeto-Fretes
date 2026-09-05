# Sistema Financeiro Pessoal V1 — Excel (Guaraní, Paraguai)

Sistema completo de controle financeiro pessoal em Excel, sem macros, com moeda
**Guaraní paraguaio (Gs.)**. Todo o cálculo é feito por fórmulas nativas; o
arquivo abre em qualquer Excel ou LibreOffice sem aviso de segurança.

**Entregáveis** — 29 abas, 40.226 fórmulas, 285 intervalos nomeados cada:

| Arquivo | Conteúdo |
|---|---|
| `SISTEMA_FINANCEIRO_PESSOAL_V1_LIMPO.xlsx` | Vazio. Só as 13 categorias e 54 subcategorias padrão. É por onde se começa. |
| `SISTEMA_FINANCEIRO_PESSOAL_V1_EXEMPLO.xlsx` | 602 lançamentos fictícios de jan/2026 a jan/2028, para ver o sistema funcionando. |

> **Nunca use "Excluir linha" nem classifique (sort) as abas de dados.** Os IDs são
> gerados pela posição da linha, então excluir uma linha desloca todos os IDs
> abaixo dela e os lançamentos passam a apontar para o registro errado, sem aviso.
> Para apagar um registro, selecione as células **amarelas** e pressione **Delete**
> (limpar conteúdo). Para anular um lançamento sem perder o histórico, use
> `Status = CANCELADO`.

## Como o sistema funciona

Existe **um único livro de operações** (aba `LANCAMENTOS`). Saldos, dívidas,
parcelas, metas, fluxo de caixa, relatórios e patrimônio são todos derivados dele.
Nenhum número é digitado duas vezes.

Cada lançamento tem uma **origem** e um **destino**, cada um sendo `CONTA`,
`CARTAO`, `INVESTIMENTO` ou `EXTERNO`. Essa única regra elimina a duplicidade:

| Operação | Origem → Destino | Receita/Despesa | Caixa |
|---|---|---|---|
| Receita | EXTERNO → CONTA | Receita | Entrada |
| Despesa | CONTA → EXTERNO | Despesa | Saída |
| Transferência | CONTA → CONTA | Neutro | Não afeta |
| Compra no cartão | CARTAO → EXTERNO | Despesa | Não afeta |
| Pagamento de fatura | CONTA → CARTAO | Neutro | Saída |
| Aporte | CONTA → INVESTIMENTO | Neutro | Saída |
| Resgate | INVESTIMENTO → CONTA | Neutro | Entrada |
| Rendimento | EXTERNO → INVESTIMENTO | Receita financeira | Não afeta |

Metas são **reserva lógica**: registradas em `MOV_METAS`, reduzem o
*Saldo_Disponivel* da conta vinculada sem tocar no saldo bancário.

## Estrutura das abas

| Camada | Abas |
|---|---|
| Painel e entrada | `DASHBOARD`, `LANCAR`, `GERADOR` |
| Operações | `LANCAMENTOS`, `PARCELAS`, `MOV_METAS` |
| Cadastros | `CAD_Instituicoes`, `CAD_Contas`, `CAD_Cartoes`, `CAD_Investimentos`, `CAD_Categorias`, `CAD_Subcategorias`, `CAD_Compromissos`, `CAD_Metas`, `CAD_Recorrencias`, `CAD_Bens` |
| Análise | `FLUXO_CAIXA`, `PROJECAO`, `REL_Mensal`, `REL_Anual`, `PATRIMONIO`, `INDICADORES`, `ALERTAS` |
| Controle e docs | `TESTES`, `MANUAL`, `ARQUITETURA`, `CFG_Sistema`, `CFG_Listas`, `AUX` |

IDs automáticos por prefixo: `LAN`, `CON`, `CAR`, `INV`, `CMP`, `PAR`, `MET`,
`CAT`, `SUB`, `INS`, `REC`, `MOV`, `BEM`. O usuário nunca digita um ID.

## Regenerar e testar

```bash
pip install openpyxl                                    # dependência única
python3 build.py                                        # gera os dois arquivos
python3 gerador/recalc_lo.py SISTEMA_FINANCEIRO_PESSOAL_V1_EXEMPLO.xlsx 900   # recalcula
python3 auditoria.py                                    # auditoria do arquivo de exemplo
python3 auditoria_limpo.py                              # confere o arquivo vazio
python3 teste_primeiro_uso.py                           # preenche o vazio do zero
python3 testes_mutacao.py                               # 14 cenários de erro e duplicidade
```

`recalc_lo.py` exige `libreoffice-calc` instalado.

## Arquivos

| Arquivo | Papel |
|---|---|
| `build.py` | Orquestra a geração do arquivo |
| `gerador/comum.py` | Constantes, estilos, formatos e domínios |
| `gerador/dados.py` | Dataset fictício determinístico (602 lançamentos) |
| `gerador/planilha.py` | Construção das 29 abas |
| `gerador/motor.py` | Reimplementação das regras em Python (referência de auditoria) |
| `gerador/recalc_lo.py` | Recalcula o arquivo e reporta células de erro |
| `auditoria.py` | Compara o que o Excel calculou contra o motor Python |
| `auditoria_limpo.py` | Confere que o arquivo vazio está zerado e sem erro |
| `teste_primeiro_uso.py` | Preenche o arquivo vazio do zero e confere cada número |
| `testes_mutacao.py` | 14 cenários de edição, cancelamento, erro e duplicidade |

## Resultado da validação

| Verificação | Resultado |
|---|---|
| Células de erro de fórmula (ambos os arquivos) | 0 de 40.226 |
| Bateria interna (aba `TESTES`) | 53 testes no exemplo, 0 falhas |
| Auditoria independente do exemplo | 422 verificações, 0 falhas |
| Cobertura das listas suspensas | 59 colunas de domínio, 0 sem lista |
| Conferência do arquivo limpo | 29 verificações, 0 falhas |
| Primeiro uso a partir do arquivo limpo | 29 verificações, 0 falhas |
| Cenários de mutação | 14 de 14 com o comportamento esperado |

## Limites da V1

Sem macros: gerar cronogramas e ocorrências é copiar blocos prontos da aba
`GERADOR`. As fórmulas cobrem 3.000 lançamentos, 1.200 parcelas, 600 movimentos
de meta e 40 registros por cadastro. Não classifique (sort) nem exclua linhas das
abas de dados — os IDs derivam da posição da linha; use `Status = CANCELADO`.
