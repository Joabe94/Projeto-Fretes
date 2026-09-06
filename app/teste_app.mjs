/* Teste de fumaça do app: abre o arquivo no Chromium, cria o perfil de exemplo
   e confere os números contra os valores auditados na planilha Excel. */
import { chromium } from "playwright-core";
import { fileURLToPath } from "url";
import path from "path";

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const ARQ = "file://" + path.join(AQUI, "dist", "caja-guarani.html");
const erros = [], falhas = [], oks = [];

const gs = (n) => "Gs. " + Math.round(n).toLocaleString("de-DE");
function conf(nome, esperado, obtido, tol = 0.5) {
  const ok = typeof esperado === "number"
    ? Math.abs(esperado - obtido) <= tol : String(esperado) === String(obtido);
  (ok ? oks : falhas).push(nome);
  const f = (x) => (typeof x === "number" ? gs(x) : x);
  console.log(`  [${ok ? "OK   " : "FALHA"}] ${nome.padEnd(52)} esperado ${String(f(esperado)).padStart(17)}  obtido ${String(f(obtido)).padStart(17)}`);
}

const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium-1194/chrome-linux/chrome" });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
page.on("pageerror", (e) => erros.push("pageerror: " + e.message));
page.on("console", (m) => { if (m.type() === "error") erros.push("console: " + m.text()); });

console.log("=".repeat(104));
console.log("TESTE DO APP NO NAVEGADOR");
console.log("=".repeat(104));
await page.goto(ARQ, { waitUntil: "load" });
await page.waitForSelector(".onb-card", { timeout: 15000 });
console.log("\n[1] Tela inicial carregou");
conf("Título da tela inicial", "Caja Guaraní", await page.textContent(".onb-top h1"));

await page.click('[data-act="novo-perfil-demo"]');
await page.waitForSelector(".app .side", { timeout: 20000 });
console.log("\n[2] Perfil de exemplo criado e painel aberto");

const v = await page.evaluate(() => {
  const V = App.v, p = V.patrimonio;
  const cont = {};
  for (const x of V.parcelas) cont[x.status] = (cont[x.status] || 0) + 1;
  const t = V.testes;
  return {
    contas: V.contas.map((c) => [c.id, c.saldo, c.reservado, c.disponivel, c.projetado]),
    cartoes: V.cartoes.map((k) => [k.id, k.divida, k.fatura, k.vencimento]),
    inv: V.investimentos.map((i) => [i.id, i.saldo, i.principal, i.rendimento, i.rendProjetado]),
    metas: V.metas.map((m) => [m.id, m.reservado, m.situacao]),
    parcStatus: cont,
    cmp: V.compromissos.map((c) => [c.nome, c.atual, c.devedor, c.proximoVencimento]),
    pat: [p.ativos, p.passivos, p.plFin, p.plTotal],
    fluxo1: V.fluxo.slice(0, 3).map((f) => [f.compet, f.entR, f.saiR, f.saldoFim]),
    proj1: V.projecao[0].saldo,
    lanc: V.L.length, errosVal: V.L.filter((l) => l.validacao !== "OK").length,
    testes: t.length, testesFalha: t.filter((x) => !x.ok).map((x) => `${x.grupo}: ${x.nome}`),
    alertas: V.alertas.length,
  };
});

console.log("\n[3] Saldos de contas");
conf("Cuenta Corriente Ueno", 13910000, v.contas[0][1]);
conf("Caja de Ahorro Continental", 30000000, v.contas[1][1]);
conf("Efectivo (Billetera)", 3381000, v.contas[2][1]);
conf("Reservado na conta de metas", 26400000, v.contas[1][2]);
conf("Disponível na conta de metas", 3600000, v.contas[1][3]);

console.log("\n[4] Cartões");
conf("Dívida Tarjeta Ueno Visa", 849000, v.cartoes[0][1]);
conf("Dívida Tarjeta Continental", 446000, v.cartoes[1][1]);
conf("Fatura do ciclo aberto (Ueno)", 4969000, v.cartoes[0][2]);
conf("Vencimento da fatura (Ueno)", "2026-10-05", v.cartoes[0][3]);

console.log("\n[5] Investimentos");
conf("Saldo Fondo Mutuo Atlas", 11615000, v.inv[0][1]);
conf("Principal Fondo Mutuo", 11000000, v.inv[0][2]);
conf("Rendimento realizado Fondo Mutuo", 615000, v.inv[0][3]);
conf("Saldo Ahorro a Plazo", 10000000, v.inv[1][1]);
conf("Rendimento projetado Ahorro a Plazo", 1100000, v.inv[1][4]);

console.log("\n[6] Metas e parcelas");
conf("Meta Comprar auto", 20000000, v.metas[0][1]);
conf("Meta Reserva de Emergencia", 6400000, v.metas[1][1]);
conf("Situação da Reserva de Emergencia", "EM RISCO", v.metas[1][2]);
conf("Parcelas pagas", 56, v.parcStatus.PAGA || 0);
conf("Parcelas abertas", 3, v.parcStatus.ABERTA || 0);
conf("Parcelas atrasadas", 1, v.parcStatus.ATRASADA || 0);
conf("Parcelas projetadas", 22, v.parcStatus.PROJETADA || 0);
conf("Terreno: parcela atual", "44/60", v.cmp[0][1]);
conf("Terreno: saldo devedor", 30600000, v.cmp[0][2]);

console.log("\n[7] Patrimônio e fluxo");
conf("Total de ativos", 176906000, v.pat[0]);
conf("Total de passivos", 40795000, v.pat[1]);
conf("Patrimônio financeiro", 28111000, v.pat[2]);
conf("Patrimônio total", 136111000, v.pat[3]);
conf("Fluxo jan/2026: entradas", 14800000, v.fluxo1[0][1]);
conf("Fluxo jan/2026: saldo final", 45663000, v.fluxo1[0][3]);
conf("Projeção set/2026", 55446000, v.proj1);

console.log("\n[8] Integridade");
conf("Lançamentos carregados", 602, v.lanc);
conf("Lançamentos com erro de validação", 0, v.errosVal);
conf("Testes internos com falha", 0, v.testesFalha.length);
console.log(`         (${v.testes} testes internos executados, ${v.alertas} alertas ativos)`);
if (v.testesFalha.length) v.testesFalha.forEach((f) => console.log("         FALHA -> " + f));

console.log("\n[9] Navegação por todas as telas");
const rotas = ["painel", "alertas", "indicadores", "lancamentos", "compromissos", "metas",
  "recorrencias", "contas", "cartoes", "investimentos", "categorias", "instituicoes", "bens",
  "fluxo", "projecao", "relMensal", "relAnual", "patrimonio", "testes", "config"];
for (const r of rotas) {
  const antes = erros.length;
  await page.click(`[data-act="ir"][data-rota="${r}"]`);
  await page.waitForTimeout(90);
  const txt = (await page.textContent(".page")) || "";
  conf("Tela " + r, true, txt.length > 120 && erros.length === antes);
}

console.log("\n[10] Ações");
await page.click('[data-act="ir"][data-rota="compromissos"]');
await page.waitForTimeout(120);
await page.click('[data-act="toggle"]');
await page.waitForTimeout(120);
conf("Cronograma abre", true, (await page.locator("table").count()) > 0);
await page.click('[data-act="ir"][data-rota="lancamentos"]');
await page.waitForTimeout(120);
await page.click('[data-act="novo"][data-tipo="lancamento"]');
await page.waitForSelector("#modal-form");
conf("Formulário de lançamento abre", true, (await page.locator('[data-campo="tipo"]').count()) === 1);
await page.selectOption('[data-campo="tipo"]', "TRANSFERENCIA");
await page.waitForTimeout(150);
conf("Origem muda com o tipo", "CONTA", await page.inputValue('[data-campo="origemTipo"]'));
conf("Destino muda com o tipo", "CONTA", await page.inputValue('[data-campo="destinoTipo"]'));
await page.click('[data-act="fechar-modal"]');
await page.waitForTimeout(100);

const antesLanc = await page.evaluate(() => App.d.lancamentos.length);
await page.evaluate(() => {
  const p = App.v.parcelas.find((x) => x.status === "ABERTA");
  App.pagarParcela(p.id);
});
await page.waitForSelector("#modal-form");
await page.click('[data-act="modal-ok"]');
await page.waitForTimeout(250);
const dep = await page.evaluate(() => ({
  lanc: App.d.lancamentos.length,
  pagas: App.v.parcelas.filter((p) => p.status === "PAGA").length,
  falhas: App.v.testes.filter((t) => !t.ok).length,
  dupl: App.v.parcelas.filter((p) => p.qtdReal > 1).length,
}));
conf("Pagar parcela não duplica lançamento", antesLanc, dep.lanc, 0);
conf("Parcela vira paga", 57, dep.pagas, 0);
conf("Nenhuma parcela com pagamento duplicado", 0, dep.dupl, 0);
conf("Bateria interna segue sem falha", 0, dep.falhas, 0);

console.log("\n[11] Ajuda de tela");
let ajudaOk = 0, ajudaVazia = [];
for (const r of rotas) {
  await page.click(`[data-act="ir"][data-rota="${r}"]`);
  await page.waitForTimeout(60);
  await page.click('[data-act="ajuda"]');
  await page.waitForSelector("#modal-form", { timeout: 4000 });
  const t = await page.textContent(".modal-h h3");
  const secs = await page.locator(".ajuda-sec").count();
  const ligs = await page.locator(".ajuda-lig button").count();
  const probs = await page.locator(".ajuda-prob div").count();
  if (t.startsWith("Ajuda —") && secs === 3 && ligs >= 1 && probs >= 1) ajudaOk++;
  else ajudaVazia.push(`${r} (secoes ${secs}, ligacoes ${ligs}, problemas ${probs})`);
  await page.click('[data-act="fechar-modal"]');
  await page.waitForTimeout(40);
}
conf("Ajuda completa nas 20 telas", 20, ajudaOk, 0);
if (ajudaVazia.length) ajudaVazia.forEach((x) => console.log("         incompleta -> " + x));
await page.click(`[data-act="ir"][data-rota="contas"]`);
await page.waitForTimeout(60);
await page.click('[data-act="ajuda"]');
await page.waitForSelector(".ajuda-lig button");
await page.click('.ajuda-lig button');
await page.waitForTimeout(200);
conf("Link da ajuda navega para a tela citada", true,
  (await page.evaluate(() => App.rota)) !== "contas");
await page.click('[data-act="fechar-modal"]').catch(() => {});
await page.waitForTimeout(100);

console.log("\n[12] Guia de primeiros passos (perfil novo do zero)");
await page.evaluate(() => App.criarPerfil(false));
await page.waitForSelector(".modal-h", { timeout: 8000 });
conf("Boas-vindas abrem no perfil vazio", true,
  (await page.textContent(".modal-h h3")).includes("Como funciona"));
await page.click('[data-act="modal-ok"]'); await page.waitForTimeout(120);
await page.click('[data-act="modal-ok"]'); await page.waitForTimeout(120);
conf("Botão Voltar aparece a partir do 2º passo", true,
  (await page.locator('[data-act="guia-voltar"]').count()) > 0);
await page.click('[data-act="modal-ok"]'); await page.waitForTimeout(220);
conf("Guia aparece no painel", true, (await page.locator(".guia").count()) === 1);
conf("Guia lista os 10 passos", 10, await page.locator(".passo").count(), 0);
const prog0 = await page.textContent(".guia-h .tag");
conf("Progresso começa em zero essenciais", true, prog0.includes("0 de"));
await page.click('[data-act="guia-fazer"][data-passo="conta"]');
await page.waitForSelector("#modal-form", { timeout: 5000 });
conf("“Fazer agora” abre o cadastro certo", true,
  (await page.textContent(".modal-h h3")).toLowerCase().includes("conta"));
await page.fill('[name="nome"]', "Conta teste");
await page.fill('[name="saldoInicial"]', "5000000");
await page.click('[data-act="modal-ok"]');
await page.waitForTimeout(300);
await page.evaluate(() => { App.rota = "painel"; App.render(); });
await page.waitForTimeout(200);
const prog1 = await page.textContent(".guia-h .tag");
conf("Passo se marca sozinho ao cadastrar", true, prog1.includes("1 de"));
conf("Saldo do perfil novo confere", 5000000,
  await page.evaluate(() => App.v.patrimonio.contas));
conf("Perfil novo passa na conferência", 0,
  await page.evaluate(() => App.v.testes.filter((t) => !t.ok).length), 0);
await page.click('[data-act="guia-fechar"]');
await page.waitForTimeout(200);
conf("Esconder o guia funciona", 0, await page.locator(".guia").count(), 0);

console.log("\n[13] Telefone (390 x 844)");
await page.setViewportSize({ width: 390, height: 844 });
await page.evaluate(() => { App.rota = "lancamentos"; App.render(); });
await page.waitForTimeout(250);
conf("Menu lateral escondido", false, await page.locator(".side").isVisible());
conf("Barra inferior visível", true, await page.locator(".mob-bar").isVisible());
conf("Botão flutuante visível", true, await page.locator(".fab").isVisible());
await page.evaluate(() => { App.abrirPerfil(App.perfis[0].id); });
await page.waitForTimeout(400);
await page.evaluate(() => { App.rota = "lancamentos"; App.render(); });
await page.waitForTimeout(250);
conf("Tabela vira lista de cartões", false, await page.locator(".tbl-wrap").first().isVisible());
conf("Cartões aparecem", true, await page.locator(".cards").first().isVisible());
conf("Um cartão por lançamento", true, (await page.locator(".cardrow").count()) > 20);
const larguraOk = await page.evaluate(() =>
  document.documentElement.scrollWidth <= window.innerWidth + 1);
conf("Nada estoura a largura da tela", true, larguraOk);
await page.click('[data-act="folha"][data-folha="mais"]');
await page.waitForTimeout(200);
conf("Folha “Mais” abre", true, await page.locator(".sheet").isVisible());
await page.click('.sheet-item[data-rota="metas"]');
await page.waitForTimeout(250);
conf("Folha navega e fecha", "metas", await page.evaluate(() => App.rota));
conf("Folha fechou", 0, await page.locator(".sheet").count(), 0);
const alvos = await page.evaluate(() => {
  const r = [];
  for (const b of document.querySelectorAll(".mob-bar button, .fab, .cr-acts .btn")) {
    const x = b.getBoundingClientRect(); if (x.height && x.height < 34) r.push(b.textContent.trim());
  }
  return r;
});
conf("Todo botão tem altura de toque >= 34px", 0, alvos.length, 0);
await page.screenshot({ path: path.join(AQUI, "dist", "telefone.png"), fullPage: false });
await page.evaluate(() => { App.rota = "painel"; App.render(); });
await page.waitForTimeout(250);
await page.screenshot({ path: path.join(AQUI, "dist", "telefone-painel.png"), fullPage: false });

await page.setViewportSize({ width: 1440, height: 1000 });

console.log("\n[13b] Filtros de lançamentos");
await page.evaluate(() => { App.rota = "lancamentos"; App.render(); });
await page.waitForTimeout(200);
conf("Barra de filtros aparece", 6, await page.locator("[data-filtro]").count(), 0);
const todos = await page.locator(".tbl-wrap tbody tr").count();
await page.selectOption('[data-filtro="status"]', "PROJETADO");
await page.waitForTimeout(250);
const soProj = await page.locator(".tbl-wrap tbody tr").count();
conf("Filtrar por status reduz a lista", true, soProj > 0 && soProj < todos);
await page.click('[data-act="limpar-filtros"]');
await page.waitForTimeout(220);
conf("Limpar filtros volta a lista inteira", todos, await page.locator(".tbl-wrap tbody tr").count(), 0);

console.log("\n[14] Extrato de conta, cartão e investimento");
await page.evaluate(() => { App.rota = "contas"; App.render(); });
await page.waitForTimeout(150);
await page.click('.btn-link[data-act="extrato"][data-tipo="CONTA"]');
await page.waitForTimeout(250);
conf("Clicar no nome da conta abre o extrato", "extrato", await page.evaluate(() => App.rota));
const ex = await page.evaluate(() => {
  const r = {};
  for (const [t, lista] of [["CONTA", App.v.contas], ["CARTAO", App.v.cartoes],
                            ["INVESTIMENTO", App.v.investimentos]]) {
    r[t] = lista.map((e) => {
      const x = Calc.extrato(App.v, t, e.id, new Set(["REALIZADO"]));
      const oficial = t === "CONTA" ? e.saldo : t === "CARTAO" ? e.divida : e.saldo;
      return { nome: e.nome, extrato: x.saldoFinal, oficial, movs: x.linhas.length };
    });
  }
  return r;
});
for (const [t, itens] of Object.entries(ex))
  for (const i of itens)
    conf(`Saldo do extrato = saldo do cadastro (${i.nome})`, i.oficial, i.extrato);
conf("Extrato da conta principal tem movimentos", true, ex.CONTA[0].movs > 50);
const linhasVis = await page.locator(".tbl-wrap tbody tr").count();
conf("Extrato lista as linhas", true, linhasVis > 20);
await page.click('[data-act="extrato-status"][data-st="TODOS"]');
await page.waitForTimeout(220);
const comProj = await page.locator(".tbl-wrap tbody tr").count();
conf("Filtro “com projetados” mostra mais linhas", true, comProj > linhasVis);
const meses = await page.locator('[data-extrato="mes"] option').count();
conf("Filtro de mês foi montado", true, meses > 5);
await page.selectOption('[data-extrato="mes"]', { index: 1 });
await page.waitForTimeout(220);
const doMes = await page.locator(".tbl-wrap tbody tr").count();
conf("Filtrar por mês reduz a lista", true, doMes > 0 && doMes < comProj);
await page.click('[data-act="ajuda"]');
await page.waitForSelector("#modal-form");
conf("Extrato tem ajuda própria", true,
  (await page.textContent(".modal-h h3")).includes("Extrato"));
await page.click('[data-act="fechar-modal"]');
await page.waitForTimeout(120);
await page.evaluate(() => { App.rota = "cartoes"; App.render(); });
await page.waitForTimeout(150);
await page.click('.btn-link[data-act="extrato"][data-tipo="CARTAO"]');
await page.waitForTimeout(250);
const cab = await page.evaluate(() => Array.from(document.querySelectorAll(".tbl-wrap th")).map((x) => x.textContent.trim()));
conf("No cartão as colunas viram Compra/Pagamento/Dívida", true,
  cab.includes("Compra") && cab.includes("Pagamento") && cab.includes("Dívida"));

console.log("\n[15] Cadastrar do zero cada tipo (regressão do bug do campo Tipo)");
await page.evaluate(() => App.criarPerfil(false));
await page.waitForSelector(".modal-h", { timeout: 8000 });
for (let i = 0; i < 3; i++) { await page.click('[data-act="modal-ok"]'); await page.waitForTimeout(110); }
const cadastros = [
  ["instituicao", "instituicoes", { nome: "Ueno Bank" }, { tipo: "Banco" }],
  ["conta", "contas", { nome: "Cuenta Ueno", saldoInicial: "7000000" }, { tipo: "Conta corrente" }],
  ["cartao", "cartoes", { nome: "Visa Ueno", limite: "9000000", dividaInicial: "500000" }, {}],
  ["investimento", "investimentos", { nome: "Fondo Atlas", valorInicial: "3000000" }, { tipo: "Fondo Mutuo" }],
  ["categoria", "categorias", { nome: "Pets" }, { tipoPadrao: "DESPESA" }],
  ["bem", "bens", { nome: "Moto", valor: "20000000" }, { tipo: "Veículo" }],
  ["meta", "metas", { nome: "Viagem", objetivo: "15000000" }, {}],
];
for (const [tipo, lista, textos, selects] of cadastros) {
  const antes = await page.evaluate((l) => App.d[l].length, lista);
  await page.evaluate((t) => App.editar(t, null), tipo);
  await page.waitForSelector("#modal-form", { timeout: 5000 });
  for (const [k, val] of Object.entries(selects))
    await page.selectOption(`[name="${k}"]`, val).catch(() => {});
  for (const [k, val] of Object.entries(textos)) await page.fill(`[name="${k}"]`, val);
  await page.click('[data-act="modal-ok"]');
  await page.waitForTimeout(260);
  const dep = await page.evaluate((l) => App.d[l].length, lista);
  const fechou = (await page.locator("#modal-form").count()) === 0;
  conf(`Criar ${tipo}`, antes + 1, dep, 0);
  if (!fechou) { falhas.push(`Modal de ${tipo} não fechou`); await page.click('[data-act="fechar-modal"]'); }
}
const criados = await page.evaluate(() => ({
  inst: App.d.instituicoes[0], conta: App.d.contas[0],
  saldo: App.v.patrimonio.contas, falhas: App.v.testes.filter((t) => !t.ok).length,
}));
conf("Instituição guardou o tipo escolhido", "Banco", criados.inst.tipo);
conf("Instituição recebeu ID automático", "INS-000001", criados.inst.id);
conf("Conta guardou nome e saldo", "Cuenta Ueno", criados.conta.nome);
conf("Saldo do perfil novo", 7000000, criados.saldo);
conf("Perfil montado do zero passa na conferência", 0, criados.falhas, 0);
await page.evaluate(() => { App.rota = "contas"; App.render(); });
await page.waitForTimeout(200);
await page.click('.btn-link[data-act="extrato"][data-tipo="CONTA"]');
await page.waitForTimeout(250);
conf("Extrato de conta recém-criada abre sem erro", "extrato", await page.evaluate(() => App.rota));
await page.evaluate(() => App.abrirPerfil(App.perfis[0].id));
await page.waitForTimeout(400);

await page.click('[data-act="ir"][data-rota="painel"]');
await page.waitForTimeout(300);
await page.screenshot({ path: path.join(AQUI, "dist", "painel.png"), fullPage: false });
await page.emulateMedia({ colorScheme: "dark" });
await page.waitForTimeout(200);
await page.screenshot({ path: path.join(AQUI, "dist", "painel-escuro.png"), fullPage: false });

console.log("\n" + "=".repeat(104));
console.log(`RESULTADO: ${oks.length} verificações OK, ${falhas.length} falhas, ${erros.length} erros de página`);
console.log("=".repeat(104));
falhas.forEach((f) => console.log("  FALHA: " + f));
erros.slice(0, 12).forEach((e) => console.log("  ERRO: " + e));
await browser.close();
process.exit(falhas.length || erros.length ? 1 : 0);
