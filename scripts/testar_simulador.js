// Confere se o simulador da apresentação reproduz o resumo K4–K11 da planilha.
// Uso: node scripts/testar_simulador.js [index.html]
const fs = require('fs');
const path = process.argv[2] || require('path').join(__dirname, '..', 'index.html');
const html = fs.readFileSync(path, 'utf8');
const json = html.match(/<script id="site-data" type="application\/json">([\s\S]*?)<\/script>/)[1];
let D = JSON.parse(json);
const sim = html.match(/\/\*SIM-START\*\/([\s\S]*?)\/\*SIM-END\*\//)[1];
const ctx = {};
new Function('ctx', sim + '\nctx.simular = simular; ctx.parametrosDaPlanilha = parametrosDaPlanilha; ctx.desembrulhar = desembrulhar;')(ctx);
D = ctx.desembrulhar(D);
const s = ctx.simular(D, ctx.parametrosDaPlanilha(D));
let max = 0;
for (const k of ['K4','K5','K6','K7','K8','K9','K10','K11']) {
  const d = s.res[k] - D.resumo[k]; max = Math.max(max, Math.abs(d));
  console.log(k.padEnd(4), D.resumo[k].toFixed(2).padStart(18), s.res[k].toFixed(2).padStart(18), d.toFixed(4).padStart(12));
}
// conferência mês a mês do saldo com juros
let maxM = 0;
for (let i = 0; i <= D.corteIdx; i++) maxM = Math.max(maxM, Math.abs(s.AN[i] - D.encontro.AN[i]));
console.log('diferença máxima no resumo:', max.toFixed(4), '| no saldo com juros mês a mês:', maxM.toFixed(4));
if (max > 0.01) { console.error('FALHOU: o simulador não reproduz a planilha.'); process.exit(1); }
console.log('OK');
