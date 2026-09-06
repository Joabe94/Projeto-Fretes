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
