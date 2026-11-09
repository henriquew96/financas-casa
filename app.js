'use strict';
// Não deixa o app funcionar dentro de outro site (proteção contra "clickjacking": um site falso
// mostrando o app por baixo para fazer você clicar sem perceber).
if (window.top !== window.self) {
  document.documentElement.innerHTML = '<body style="font:16px system-ui;padding:24px">Abra o Finanças da Casa pelo endereço oficial.</body>';
  throw new Error('app aberto dentro de outro site');
}
/* =====================================================================
   Finanças da Casa — fase 2 (nuvem opcional, via config.js)
   Dados em localStorage. Importação lida 100% no navegador.
   ===================================================================== */
const CHAVE = 'financas-casa-v1';
const PDFJS = 'libs/pdf.min.js';                 // pdf.js 3.11.174, hospedado aqui (conferido em 2 fontes)
const PDFJS_WORKER = 'libs/pdf.worker.min.js';
const SHEETJS = 'libs/xlsx.full.min.js';         // SheetJS 0.20.3 (corrige CVE-2023-30533 e CVE-2024-22363)

const MESES = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
const MESES_C = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
const MES_IDX = {jan:1,fev:2,feb:2,mar:3,abr:4,apr:4,mai:5,may:5,jun:6,jul:7,ago:8,aug:8,set:9,sep:9,out:10,oct:10,nov:11,dez:12,dec:12};
const SEMANA = ['dom','seg','ter','qua','qui','sex','sáb'];
const RE_MES_TXT = '(?:jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)';

const CATEGORIAS_PADRAO = ['Moradia','Contas da casa','Mercado','Alimentação fora','Transporte','Saúde','Educação','Lazer','Viagem','Assinaturas','Compras','Vestuário','Beleza e cuidados','Pets','Presentes','Taxas e juros','Reserva e investimentos','Outros'];
const REGRAS_BASE = {
  'Alimentação fora': ['ifood','ifd','rappi','ubereats','uber eats','restaurante','lanchonete','padaria','pizzaria','pizza','burger','hamburg','mcdonald','mc donald','burger king','subway','outback','starbucks','cafeteria','cafe','churrascaria','sushi','temaki','acai','sorvet','doceria','confeitaria','ze delivery','habibs','giraffas','spoleto','coco bambu','madero','pastel','marmita','refeicao'],
  'Mercado': ['supermerc','mercado','mercearia','carrefour','assai','atacad','pao de acucar','hortifruti','sacolao','hipermerc','st marche','zaffari','angeloni','condor','savegnago','muffato','bistek','giassi','sams club','makro','tenda atacado','mambo','oxxo','minuto pao','swift','acougue','emporio','feira'],
  'Transporte': ['uber','99app','99 pop','99 tecnologia','99 taxi','cabify','posto','auto posto','ipiranga','shell','petrobras','br mania','combust','estacion','sem parar','conectcar','veloe','taggy','metro','cptm','sptrans','bilhete unico','pedagio','zul digital','buser','localiza','movida','unidas','detran','ipva','oficina','borracharia','lava rapido'],
  'Saúde': ['drogaria','farmacia','droga raia','raia','drogasil','pacheco','pague menos','panvel','drogao','hospital','clinica','laborat','odonto','dentist','unimed','amil','sulamerica','hapvida','psicolog','consulta','plano de saude','smart fit','smartfit','bluefit','academia','gympass','wellhub','totalpass'],
  'Assinaturas': ['netflix','spotify','prime video','amazon prime','primevideo','disney','hbo','globoplay','youtube','deezer','apple com','icloud','google one','google storage','microsoft','xbox','playstation','chatgpt','openai','anthropic','claude ai','paramount','crunchyroll','kindle unlimited','audible','canva','dropbox','adobe'],
  'Compras': ['amazon','mercadolivre','mercado livre','magalu','magazine luiza','shopee','aliexpress','americanas','casas bahia','ponto frio','kabum','submarino','leroy merlin','tok stok','tokstok','ikea','camicado','fast shop','havan','temu','kalunga','geladeira','fogao','notebook','brastemp','consul','electrolux','samsung','dell','lenovo','apple store','eletro'],
  'Vestuário': ['renner','riachuelo','zara','shein','centauro','netshoes','nike','adidas','hering','marisa','pernambucanas','arezzo','vivara','track field','decathlon','youcom'],
  'Contas da casa': ['enel','cemig','copel','celesc','light s a','energisa','equatorial','neoenergia','coelba','cpfl','sabesp','cedae','copasa','sanepar','embasa','comgas','naturgy','vivo','claro','tim celular','oi fibra','net servicos','sky','iptu','internet','energia','luz','agua','gas','celular','telefone'],
  'Moradia': ['aluguel','condominio','quinto andar','quintoandar','imobiliaria','diarista'],
  'Lazer': ['cinema','cinemark','kinoplex','ingresso','sympla','eventim','ticketmaster','teatro','show','parque','steam','nintendo','epic games'],
  'Viagem': ['hotel','airbnb','booking','decolar','latam','gol linhas','voegol','azul linhas','voeazul','hoteis','pousada','hurb','123milhas','smiles'],
  'Educação': ['escola','colegio','faculdade','universidade','curso','udemy','alura','coursera','livraria','saraiva','duolingo','papelaria'],
  'Pets': ['petz','cobasi','petshop','pet shop','petlove','veterin','agropecuaria'],
  'Beleza e cuidados': ['salao','barbearia','barber','cabeleir','boticario','natura','sephora','avon','manicure','estetica','depila','perfumaria','eudora'],
  'Taxas e juros': ['anuidade','tarifa','iof','juros','encargo','multa','mora'],
  'Reserva e investimentos': ['reserva','investimento','poupanca','aplicacao','tesouro direto','previdencia'],
};

/* ---------- utilidades ---------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pad = n => String(Math.trunc(+n) || 0).padStart(2, '0');
const round2 = v => Math.round((+v || 0) * 100) / 100;
const fmtBRL = new Intl.NumberFormat('pt-BR', {style:'currency', currency:'BRL'});
const fmtCurto = new Intl.NumberFormat('pt-BR', {style:'currency', currency:'BRL', notation:'compact', maximumFractionDigits:1});
const fmtNumC = new Intl.NumberFormat('pt-BR', {notation:'compact', maximumFractionDigits:1});
const brl = v => fmtBRL.format(v || 0);
const brlC = v => Math.abs(v) < 1000 ? fmtBRL.format(Math.round(v || 0)).replace(/,00$/, '') : fmtCurto.format(v);
const num2 = v => (+v || 0).toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2});
const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-5);
const hojeISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const mesAtual = () => hojeISO().slice(0, 7);
function somaMes(m, n) { const [a, b] = m.split('-').map(Number); const t = a * 12 + (b - 1) + n; return `${Math.floor(t / 12)}-${pad(t % 12 + 1)}`; }
function difMeses(de, ate) { const [a1, b1] = de.split('-').map(Number), [a2, b2] = ate.split('-').map(Number); return (a2 - a1) * 12 + (b2 - b1); }
const nomeMes = m => { const [a, b] = m.split('-').map(Number); return `${MESES[b - 1]} ${a}`; };
const nomeMesC = m => { const [a, b] = m.split('-').map(Number); return `${MESES_C[b - 1]}/${String(a).slice(2)}`; };
const mesSo = m => MESES[+m.slice(5, 7) - 1];
const RE_ISO = /^\d{4}-\d{2}-\d{2}$/;
const dataCurta = iso => RE_ISO.test(iso) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}` : '';
const dataLonga = iso => RE_ISO.test(iso) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : '';
const diaSemana = iso => SEMANA[new Date(iso + 'T12:00:00').getDay()];
const diasNoMes = m => { const [a, b] = m.split('-').map(Number); return new Date(a, b, 0).getDate(); };
const dataNoMes = (M, dia) => `${M}-${pad(Math.min(Math.max(1, +dia || 1), diasNoMes(M)))}`;
const semAcento = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '');
const norm = s => semAcento(String(s || '').toLowerCase()).replace(/[^a-z0-9]+/g, ' ').trim();
const clampDia = v => Math.min(31, Math.max(1, parseInt(v, 10) || 1));
let toastT;
function toast(msg, ms = 3200) { const t = $('#toast'); t.textContent = msg; t.classList.add('on'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('on'), ms); }

/* ---------- estado ---------- */
function estadoVazio() {
  return { versao: 1, pessoas: [{id:'p1', nome:'Pessoa 1'}, {id:'p2', nome:'Pessoa 2'}], receitas: [], fixas: [], cartoes: [], lancamentos: [], regras: [], categorias: CATEGORIAS_PADRAO.slice(), tema: 'auto', exemplo: false };
}
function migrar(o) {
  if (!o || typeof o !== 'object' || Array.isArray(o)) o = {};
  const b = estadoVazio();
  for (const k of Object.keys(b)) if (o[k] === undefined) o[k] = b[k];
  return sanearEstado(o);
}

/* ---------- validação de tudo que vem de fora (nuvem, backup, cópia do aparelho) ----------
   Cada item é reconstruído campo a campo com tipo e formato conferidos: texto vira texto
   (sem caracteres de controle e com tamanho máximo), número vira número, data tem de ser
   AAAA-MM-DD, e só valores conhecidos passam em campos como "meio" e "pessoa".
   Assim um texto malicioso nunca chega à tela como código. */
const RE_IDOK = /^[A-Za-z0-9_-]{1,64}$/, RE_MESOK = /^\d{4}-\d{2}$/, RE_DIAOK = /^\d{4}-\d{2}-\d{2}$/;
const MEIOS_LANC = ['cartao', 'pix', 'debito', 'dinheiro', 'boleto', 'vr', 'va'], ORIGENS = ['manual', 'import', 'notificacao', 'exemplo'];
const txt = (v, n) => String(v ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, n);
const numOk = v => { const n = Number(v); return Number.isFinite(n) ? round2(n) : 0; };
const intOk = (v, a, b) => Math.min(b, Math.max(a, parseInt(v, 10) || a));
const idOk = v => RE_IDOK.test(String(v ?? '')) ? String(v) : '';
const pesOk = (v, casal) => (v === 'p1' || v === 'p2' || (casal && v === 'casal')) ? v : 'p1';
const mesOk = v => RE_MESOK.test(v) ? v : '';
function sanear(tipo, d) {
  if (!d || typeof d !== 'object') return null;
  if (tipo === 'regra') { const termo = norm(d.termo).slice(0, 60); return termo ? { termo, categoria: txt(d.categoria, 40) || 'Outros', ordem: numOk(d.ordem) } : null; }
  if (tipo === 'config') {
    const pes = Array.isArray(d.pessoas) ? d.pessoas : [];
    const nome = id => txt(pes.find(p => p && p.id === id)?.nome, 30);
    const cats = Array.isArray(d.categorias) ? [...new Set(d.categorias.map(c => txt(c, 40)).filter(Boolean))].slice(0, 60) : [];
    if (!cats.includes('Outros')) cats.push('Outros');
    return { pessoas: [{ id: 'p1', nome: nome('p1') || 'Pessoa 1' }, { id: 'p2', nome: nome('p2') || 'Pessoa 2' }], categorias: cats };
  }
  const id = idOk(d.id); if (!id) return null;
  if (tipo === 'lancamento') {
    if (!RE_DIAOK.test(d.data) || !RE_MESOK.test(d.mesRef)) return null;
    const parcelas = intOk(d.parcelas, 1, 72);
    return { id, data: d.data, descricao: txt(d.descricao, 160) || 'Sem descrição', valor: numOk(d.valor), categoria: txt(d.categoria, 40) || 'Outros', pessoaId: pesOk(d.pessoaId), meio: MEIOS_LANC.includes(d.meio) ? d.meio : 'pix', cartaoId: idOk(d.cartaoId), parcelas, parcelaAtual: intOk(d.parcelaAtual, 1, parcelas), mesRef: d.mesRef, mesRefAuto: !!d.mesRefAuto, origem: ORIGENS.includes(d.origem) ? d.origem : 'manual', obs: txt(d.obs, 300), criadoEm: RE_DIAOK.test(d.criadoEm) ? d.criadoEm : '' };
  }
  if (tipo === 'receita') return { id, pessoaId: pesOk(d.pessoaId), tipo: ['salario', 'vr', 'va', 'extra'].includes(d.tipo) ? d.tipo : 'extra', descricao: txt(d.descricao, 80), valor: numOk(d.valor), dia: intOk(d.dia, 1, 31), inicio: mesOk(d.inicio), fim: mesOk(d.fim) };
  if (tipo === 'fixa') return { id, descricao: txt(d.descricao, 80) || 'Conta', valor: numOk(d.valor), dia: intOk(d.dia, 1, 31), categoria: txt(d.categoria, 40) || 'Outros', pessoaId: pesOk(d.pessoaId, true), meio: ['conta', 'cartao', 'vr', 'va'].includes(d.meio) ? d.meio : 'conta', cartaoId: idOk(d.cartaoId), inicio: mesOk(d.inicio), fim: mesOk(d.fim) };
  if (tipo === 'cartao') return { id, nome: txt(d.nome, 40) || 'Cartão', pessoaId: pesOk(d.pessoaId), final: /^\d{4}$/.test(d.final) ? d.final : '', fechamento: intOk(d.fechamento, 1, 31), vencimento: intOk(d.vencimento, 1, 31), limite: Math.max(0, numOk(d.limite)) };
  return null;
}
function sanearEstado(o) {
  for (const [lista, tipo] of [['lancamentos', 'lancamento'], ['receitas', 'receita'], ['fixas', 'fixa'], ['cartoes', 'cartao']]) o[lista] = (Array.isArray(o[lista]) ? o[lista] : []).map(x => sanear(tipo, x)).filter(Boolean);
  const rs = Array.isArray(o.regras) ? o.regras : [];
  o.regras = rs.map((r, i) => sanear('regra', { ...(r || {}), ordem: r && r.ordem ? r.ordem : rs.length - i })).filter(Boolean);
  const cfg = sanear('config', { pessoas: o.pessoas, categorias: o.categorias });
  o.pessoas = cfg.pessoas; o.categorias = cfg.categorias;
  o.tema = ['auto', 'claro', 'escuro'].includes(o.tema) ? o.tema : 'auto';
  o.exemplo = !!o.exemplo;
  o.versao = 1;
  return o;
}
function carregar() {
  let txt = null;
  try { txt = localStorage.getItem(CHAVE); } catch (e) {}
  if (txt) {
    try { return migrar(JSON.parse(txt)); }
    catch (e) { try { localStorage.setItem(CHAVE + '-corrompido-' + Date.now(), txt); } catch (_) {} }
  }
  return estadoVazio();
}
let avisouSalvar = false;
// grava só neste aparelho
function salvarLocal() {
  REGRAS_CACHE = null;
  try { localStorage.setItem(CHAVE, JSON.stringify(S)); }
  catch (e) { if (!avisouSalvar) { avisouSalvar = true; alert('Este navegador não deixou salvar os dados (janela anônima ou armazenamento bloqueado). Baixe um backup em Ajustes para não perder nada.'); } }
}
// grava neste aparelho e, se a casa estiver na nuvem, agenda o envio
function salvar() { salvarLocal(); agendarEnvio(); }
let REGRAS_CACHE = null;
let S = carregar();

/* ---------- estado da tela ---------- */
let ABA = 'painel';
let MES = mesAtual();
let FILTRO = { q: '', pessoa: '', meio: '', cat: '', visao: 'compra' };
let IMP = null;
let CARREGANDO = false;

/* ---------- consultas ---------- */
const pessoaNome = id => id === 'casal' ? 'Casal' : (S.pessoas.find(p => p.id === id)?.nome || '—');
const cartaoPorId = id => S.cartoes.find(c => c.id === id) || null;
const nomeCartao = id => { const c = cartaoPorId(id); return c ? c.nome : (id ? 'Cartão removido' : 'Cartão sem cadastro'); };
const ROT_MEIO = { cartao:'Cartão', pix:'Pix', debito:'Débito', dinheiro:'Dinheiro', boleto:'Boleto/conta', conta:'Conta/boleto/Pix', vr:'VR', va:'VA' };
const rotMeio = (meio, cid) => meio === 'cartao' ? nomeCartao(cid) : (ROT_MEIO[meio] || meio);
const ehVale = meio => meio === 'vr' || meio === 'va';
// quem está usando este aparelho (na nuvem, cada login é uma das pessoas da casa)
const EU = () => (typeof CASA !== 'undefined' && CASA && S.pessoas.some(p => p.id === CASA.pessoa_id)) ? CASA.pessoa_id : S.pessoas[0].id;
const catOps = extra => { const l = S.categorias.slice(); if (extra && !l.includes(extra)) l.unshift(extra); return l.map(c => [c, c]); };
const opcoesCat = sel => catOps(sel).map(([v]) => `<option value="${esc(v)}" ${v === sel ? 'selected' : ''}>${esc(v)}</option>`).join('');

/* ---------- regras de cálculo ---------- */
// Mês de VENCIMENTO da fatura em que uma compra feita em `iso` cai.
function mesFaturaDe(iso, c) {
  const [a, m, d] = iso.split('-').map(Number);
  let mf = `${a}-${pad(m)}`;
  if (d >= (+c.fechamento || 1)) mf = somaMes(mf, 1);           // a partir do fechamento, vai para a próxima
  return (+c.vencimento > +c.fechamento) ? mf : somaMes(mf, 1);  // vencimento antes do fechamento = mês seguinte
}
function calcMesRef(l) {
  if (l.meio !== 'cartao') return l.data.slice(0, 7);
  const c = cartaoPorId(l.cartaoId);
  return c ? mesFaturaDe(l.data, c) : somaMes(l.data.slice(0, 7), 1);
}
const ativo = (it, M) => (!it.inicio || it.inicio <= M) && (!it.fim || M <= it.fim);
// Número da parcela que cai no mês M (ou null se não cai).
function ocorrencia(l, M) {
  if (!l.mesRef) return null;
  const k = difMeses(l.mesRef, M); if (k < 0) return null;
  const p = (l.parcelaAtual || 1) + k;
  return p > (l.parcelas || 1) ? null : p;
}
function resumoMes(M) {
  const R = { M, receitas: [], receitaTotal: 0, rendaPessoa: {}, vr: {credito:0, gasto:0}, va: {credito:0, gasto:0}, fixasContaTotal: 0, fixasCartaoTotal: 0, faturas: new Map(), lancCartaoTotal: 0, avistaTotal: 0, parcelasTotal: 0, cats: {}, pessoas: {} };
  const soma = (o, k, v) => { o[k] = (o[k] || 0) + v; };
  const fatura = id => { if (!R.faturas.has(id)) R.faturas.set(id, { id, cartao: cartaoPorId(id), itens: 0, total: 0, parcelas: 0 }); return R.faturas.get(id); };
  for (const r of S.receitas) if (ativo(r, M)) {
    if (ehVale(r.tipo)) R[r.tipo].credito += r.valor;
    else { R.receitas.push(r); R.receitaTotal += r.valor; soma(R.rendaPessoa, r.pessoaId, r.valor); }
  }
  for (const f of S.fixas) {
    if (f.meio === 'cartao') {
      const c = cartaoPorId(f.cartaoId);
      for (const C of [somaMes(M, -2), somaMes(M, -1), M]) {
        if (!ativo(f, C)) continue;
        const mf = c ? mesFaturaDe(dataNoMes(C, f.dia), c) : somaMes(C, 1);
        if (mf !== M) continue;
        const ft = fatura(f.cartaoId || ''); ft.itens++; ft.total += f.valor;
        R.fixasCartaoTotal += f.valor; soma(R.cats, f.categoria, f.valor); soma(R.pessoas, f.pessoaId, f.valor);
      }
    } else if (ativo(f, M)) {
      if (ehVale(f.meio)) { R[f.meio].gasto += f.valor; continue; }
      R.fixasContaTotal += f.valor; soma(R.cats, f.categoria, f.valor); soma(R.pessoas, f.pessoaId, f.valor);
    }
  }
  for (const l of S.lancamentos) {
    const p = ocorrencia(l, M); if (!p) continue;
    if (ehVale(l.meio)) { R[l.meio].gasto += l.valor; continue; }
    if (l.meio === 'cartao') {
      const ft = fatura(l.cartaoId || ''); ft.itens++; ft.total += l.valor; R.lancCartaoTotal += l.valor;
      if ((l.parcelas || 1) > 1) { R.parcelasTotal += l.valor; ft.parcelas += l.valor; }
    } else R.avistaTotal += l.valor;
    soma(R.cats, l.categoria || 'Outros', l.valor); soma(R.pessoas, l.pessoaId, l.valor);
  }
  R.fixasTotal = R.fixasContaTotal + R.fixasCartaoTotal;
  R.variaveisTotal = R.lancCartaoTotal + R.avistaTotal;
  R.comprometido = R.fixasTotal + R.variaveisTotal;
  R.livre = R.receitaTotal - R.comprometido;
  return R;
}

/* ---------- categorização ---------- */
const REGRAS_PADRAO_LISTA = Object.entries(REGRAS_BASE).flatMap(([cat, ts]) => ts.map(t => ({ termo: norm(t), categoria: cat }))).sort((a, b) => b.termo.length - a.termo.length);
function regras() { return REGRAS_CACHE || (REGRAS_CACHE = [...S.regras.map(r => ({ termo: norm(r.termo), categoria: r.categoria })).filter(r => r.termo), ...REGRAS_PADRAO_LISTA]); }
const casa = (d, t) => t.length <= 4 ? d.includes(' ' + t + ' ') : d.includes(' ' + t);
function categorizar(desc) {
  const d = ' ' + norm(desc) + ' ';
  for (const r of regras()) if (casa(d, r.termo)) return S.categorias.includes(r.categoria) ? r.categoria : 'Outros';
  return 'Outros';
}
function categoriaAuto(desc, meio) {
  const c = categorizar(desc);
  if (c === 'Outros' && meio === 'vr') return 'Alimentação fora';
  if (c === 'Outros' && meio === 'va') return 'Mercado';
  return c;
}
const STOP = new Set(['de','do','da','dos','das','e','em','no','na','com','compra','compras','pagamento','pag','pg','ltda','me','sa','eireli','epp','pix','para','br','www','com br']);
function termoDe(desc) {
  const ws = norm(desc).split(' ').filter(Boolean);
  let i = 0;
  while (i < ws.length && (STOP.has(ws[i]) || /\d/.test(ws[i]) || ws[i].length < 2)) i++;
  const out = []; let sig = 0;
  for (; i < ws.length && sig < 2; i++) { const w = ws[i]; if (/\d/.test(w)) break; out.push(w); if (!STOP.has(w) && w.length >= 2) sig++; }
  while (out.length && STOP.has(out[out.length - 1])) out.pop();
  return out.join(' ');
}
function aprenderRegra(desc, cat) {
  const termo = termoDe(desc); if (!termo) return;
  S.regras = S.regras.filter(r => norm(r.termo) !== termo);
  S.regras.unshift({ termo, categoria: cat, ordem: Date.now() });
  REGRAS_CACHE = null;
}

/* =====================================================================
   TELAS
   ===================================================================== */
function aplicarTema() {
  const r = document.documentElement;
  if (S.tema === 'claro') r.dataset.theme = 'light'; else if (S.tema === 'escuro') r.dataset.theme = 'dark'; else delete r.dataset.theme;
}
function rotuloMes() {
  const r = difMeses(mesAtual(), MES);
  const tag = r === 0 ? 'este mês' : r === 1 ? 'mês que vem' : r === -1 ? 'mês passado' : r > 0 ? `daqui a ${r} meses` : `há ${-r} meses`;
  return `<b>${esc(nomeMes(MES))}</b><small>${tag}</small>`;
}
function render() {
  aplicarTema();
  renderPendente = false;
  const modo = modoTela();
  pintarNuvem();
  $('#abas').hidden = modo !== 'app';
  $('#mesRot').innerHTML = rotuloMes();
  $('.mesnav').style.visibility = (modo !== 'app' || ABA === 'importar' || ABA === 'ajustes') ? 'hidden' : '';
  if (modo !== 'app') {
    $('#fab').hidden = true;
    $('#app').innerHTML = { carregando: vCarregando, entrar: vEntrar, casa: vCasa, senha: vNovaSenha }[modo]();
    return;
  }
  $$('#abas [data-aba]').forEach(b => b.setAttribute('aria-current', b.dataset.aba === ABA ? 'page' : 'false'));
  const v = { painel: vPainel, lancamentos: vLancamentos, importar: vImportar, cadastros: vCadastros, ajustes: vAjustes }[ABA];
  $('#app').innerHTML = (S.exemplo ? `<div class="aviso"><span><b>Dados de exemplo.</b> Explore à vontade — nada aqui é real.</span><button class="btn peq" data-acao="sair-exemplo">Apagar exemplo e começar</button></div>` : '') + v();
  if (ABA === 'lancamentos') renderListaLanc();
  if (ABA === 'importar' && IMP) atualizarResumoImp();
  $('#fab').hidden = !(ABA === 'painel' || ABA === 'lancamentos');
}
// atualização vinda do outro celular: não redesenha a tela no meio de uma digitação
let renderPendente = false;
function renderRemoto() {
  const a = document.activeElement;
  if (a && a.closest && a.closest('#app') && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName)) { renderPendente = true; return; }
  render();
}

/* ---------- Painel ---------- */
function avisoConvite() {
  if (!CASA || (CASA.membros || 1) >= 2) return '';
  const outra = S.pessoas.find(p => p.id !== CASA.pessoa_id)?.nome || 'a outra pessoa';
  return `<div class="aviso"><span>Falta ${esc(outra)} entrar na casa: criar a própria conta no app e digitar o código <b style="font-family:ui-monospace,Consolas,monospace;letter-spacing:.06em">${esc(fmtCodigo(CASA.codigo_convite))}</b></span><button class="btn peq" data-acao="copiar-codigo">Copiar código</button></div>`;
}
function vBoasVindas() {
  const passos = [
    ['Renda e vales', 'Salário de cada um, VR e VA.', 'nova-receita', S.receitas.length],
    ['Cartões de crédito', 'Nome, titular, dia de fechamento e de vencimento.', 'novo-cartao', S.cartoes.length],
    ['Contas fixas', 'Aluguel, condomínio, luz, internet, assinaturas…', 'nova-fixa', S.fixas.length],
    ['Importar a última fatura', 'CSV, Excel, OFX, PDF ou HTML do banco.', 'ir-importar', S.lancamentos.length],
  ];
  const prox = passos.findIndex(p => !p[3]);
  return avisoConvite() + `<section class="card">
    <h2>Vamos montar as finanças da casa</h2>
    <p class="sub">Quatro passos e o painel já mostra quanto da renda de cada mês está comprometido — com contas fixas, cartão e parcelas.</p>
    <ol class="passos">${passos.map(([t, d, a, ok], i) => `<li class="${ok ? 'ok' : ''}"><span class="n">${ok ? '✓' : i + 1}</span><span><b>${t}</b><br><span class="sub">${d}</span></span><button class="btn ${i === prox ? 'primario' : ''}" data-acao="${a}">${ok ? 'Adicionar mais' : 'Começar'}</button></li>`).join('')}</ol>
    ${CASA ? '' : '<div class="aviso" style="margin-top:14px"><span>Quer ver como fica antes de preencher?</span><button class="btn" data-acao="exemplo">Carregar dados de exemplo</button></div>'}
  </section>`;
}
function hbar(rot, v, max, total) {
  const pct = total > 0 ? Math.round(v / total * 100) : 0;
  return `<div class="hbar" data-tip="${esc(rot)}\n${brl(v)} · ${pct}% do comprometido"><span class="hb-rot">${esc(rot)}</span><span class="hb-trilho"><span class="hb-barra" style="width:${(v / max * 100).toFixed(1)}%"></span></span><span class="hb-val">${brl(v)}</span></div>`;
}
function kpiVale(rot, x) {
  if (!x.credito && !x.gasto) return '';
  const saldo = x.credito - x.gasto, uso = x.credito ? Math.min(100, x.gasto / x.credito * 100) : 100;
  return `<div class="card kpi"><div class="rot">${rot}</div><div class="val ${saldo < 0 ? 'neg' : ''}">${saldo < 0 ? '⚠ ' : ''}${brl(saldo)}</div>
    <div class="medidor ${saldo < 0 ? 'estouro' : ''}" role="img" aria-label="${Math.round(uso)}% usado"><span style="width:${uso.toFixed(1)}%"></span></div>
    <div class="det"><span>Crédito do mês<b>${brl(x.credito)}</b></span><span>Já gasto<b>${brl(x.gasto)}</b></span>${saldo < 0 ? '<span class="neg">Passou do saldo do vale</span>' : ''}</div></div>`;
}
function graficoProjecao() {
  const meses = Array.from({ length: 6 }, (_, i) => { const m = somaMes(MES, i); return { m, r: resumoMes(m) }; });
  const max = Math.max(1, ...meses.map(x => Math.max(x.r.receitaTotal, x.r.comprometido)));
  const pc = v => (v / max * 100).toFixed(2) + '%';
  return `<div class="proj">${meses.map(({ m, r }) => {
    const tip = `${nomeMes(m)}\nRenda: ${brl(r.receitaTotal)}\nFixas: ${brl(r.fixasTotal)}\nVariáveis: ${brl(r.variaveisTotal)}\nSobra: ${brl(r.livre)}`;
    return `<div class="proj-col" tabindex="0" data-tip="${esc(tip)}" role="img" aria-label="${esc(tip.replace(/\n/g, ', '))}">
      <div class="proj-area"><div class="proj-pilha">${r.variaveisTotal > 0 ? `<span class="seg c-var" style="height:${pc(r.variaveisTotal)}"></span>` : ''}${r.fixasTotal > 0 ? `<span class="seg c-fixas" style="height:${pc(r.fixasTotal)}"></span>` : ''}</div>${r.receitaTotal > 0 ? `<i class="proj-renda" style="bottom:${pc(r.receitaTotal)}"></i>` : ''}</div>
      <div class="proj-mes">${nomeMesC(m)}</div>
      <div class="proj-sobra ${r.livre < 0 ? 'neg' : ''}">${r.livre < 0 ? '⚠ −' : ''}<span class="rs">R$ </span>${Math.abs(r.livre) < 1000 ? Math.round(Math.abs(r.livre)) : fmtNumC.format(Math.abs(r.livre))}</div></div>`;
  }).join('')}</div>`;
}
function vPainel() {
  if (!S.receitas.length && !S.fixas.length && !S.lancamentos.length && !S.cartoes.length) return vBoasVindas();
  const R = resumoMes(MES);
  const renda = R.receitaTotal, comp = R.comprometido, livre = R.livre;
  const pct = renda > 0 ? Math.round(comp / renda * 100) : null;
  const rel = difMeses(mesAtual(), MES);
  const falta = [];
  if (!S.receitas.length) falta.push(['renda e vales', 'nova-receita']);
  if (!S.cartoes.length) falta.push(['cartões', 'novo-cartao']);
  if (!S.fixas.length) falta.push(['contas fixas', 'nova-fixa']);
  let h = '';
  h += avisoConvite();
  if (falta.length) h += `<div class="aviso"><span>Para o painel ficar completo, falta cadastrar: ${falta.map(([t, a]) => `<button class="link" data-acao="${a}">${t}</button>`).join(', ')}.</span></div>`;

  const segs = [];
  if (R.fixasTotal > 0) segs.push(`<span class="seg c-fixas" style="flex:${+R.fixasTotal || 0} 1 0" data-tip="Contas fixas\n${brl(R.fixasTotal)}"></span>`);
  if (R.variaveisTotal > 0) segs.push(`<span class="seg c-var" style="flex:${+R.variaveisTotal || 0} 1 0" data-tip="Variáveis\n${brl(R.variaveisTotal)}"></span>`);
  if (livre > 0) segs.push(`<span class="seg c-livre" style="flex:${+livre || 0} 1 0" data-tip="Sobra\n${brl(livre)}"></span>`);
  const marcaRenda = comp > renda && renda > 0 ? `<i style="left:${(renda / comp * 100).toFixed(2)}%" data-tip="Renda do mês\n${brl(renda)}"></i>` : '';
  h += `<section class="card hero">
    <div class="hero-top">
      <div>
        <div class="rot">Sobra prevista em ${esc(nomeMes(MES))}</div>
        <div class="hero-num ${livre < 0 ? 'neg' : ''}">${livre < 0 ? '⚠ ' : ''}${brl(livre)}</div>
        <p class="sub">${renda ? `${brl(comp)} já comprometidos de ${brl(renda)} de renda (${pct}%).` : 'Cadastre a renda para ver quanto sobra.'}${livre < 0 && renda ? ' <b class="neg">Os gastos já passam da renda do mês.</b>' : ''}</p>
      </div>
      ${rel === 0 ? `<button class="btn" data-acao="mes" data-d="1">Ver o mês que vem →</button>` : `<button class="btn" data-acao="mes-hoje">Voltar para este mês</button>`}
    </div>
    ${segs.length ? `<div class="barra-comp" role="img" aria-label="Fixas ${brl(R.fixasTotal)}, variáveis ${brl(R.variaveisTotal)}, sobra ${brl(Math.max(livre, 0))}">${segs.join('')}${marcaRenda}</div>` : ''}
    <div class="leg">
      <span><i class="ponto c-fixas"></i>Fixas <b>${brl(R.fixasTotal)}</b></span>
      <span><i class="ponto c-var"></i>Variáveis <b>${brl(R.variaveisTotal)}</b></span>
      <span><i class="ponto c-livre"></i>Sobra <b>${brl(Math.max(livre, 0))}</b></span>
      ${marcaRenda ? '<span><i class="traco" style="width:2px;height:12px;border:0;background:var(--tx)"></i>Renda</span>' : ''}
    </div>
  </section>`;

  const rendaDet = S.pessoas.filter(p => R.rendaPessoa[p.id]).map(p => `<span>${esc(p.nome)}<b>${brl(R.rendaPessoa[p.id])}</b></span>`).join('');
  h += `<div class="grid kpis">
    <div class="card kpi"><div class="rot">Renda do mês</div><div class="val">${brl(renda)}</div><div class="det">${rendaDet || '<span class="muted">Nenhuma renda neste mês</span>'}</div></div>
    <div class="card kpi"><div class="rot"><i class="ponto c-fixas"></i>Contas fixas</div><div class="val">${brl(R.fixasTotal)}</div><div class="det"><span>Na conta<b>${brl(R.fixasContaTotal)}</b></span><span>No cartão<b>${brl(R.fixasCartaoTotal)}</b></span></div></div>
    <div class="card kpi"><div class="rot"><i class="ponto c-var"></i>Variáveis</div><div class="val">${brl(R.variaveisTotal)}</div><div class="det"><span>No cartão<b>${brl(R.lancCartaoTotal)}</b></span><span>Pix, débito, dinheiro<b>${brl(R.avistaTotal)}</b></span><span>…das quais parcelas<b>${brl(R.parcelasTotal)}</b></span></div></div>
    ${kpiVale('VR · refeição', R.vr)}${kpiVale('VA · alimentação', R.va)}
  </div>`;

  const fats = [...R.faturas.values()].sort((a, b) => (a.cartao?.vencimento || 99) - (b.cartao?.vencimento || 99));
  if (fats.length || S.cartoes.length) {
    h += `<section class="card"><h3>Faturas que vencem em ${esc(mesSo(MES))}<span class="muted">${brl(R.fixasCartaoTotal + R.lancCartaoTotal)}</span></h3>
      <div class="lista">${fats.length ? fats.map(ft => {
        const c = ft.cartao;
        return `<button class="linha" data-acao="ver-cartao" data-id="${esc(ft.id)}">
          <span class="l-ico">${esc((c?.nome || '?').slice(0, 2).toUpperCase())}</span>
          <span class="l-desc"><b>${esc(nomeCartao(ft.id))}${c ? ` <span class="muted">· ${esc(pessoaNome(c.pessoaId))}</span>` : ''}</b><small>${c ? `vence ${pad(Math.min(c.vencimento, diasNoMes(MES)))}/${MES.slice(5)} · ` : ''}${ft.itens} ${ft.itens === 1 ? 'item' : 'itens'}${ft.parcelas ? ` · parcelas ${brl(ft.parcelas)}` : ''}</small></span>
          <span class="l-valor">${brl(ft.total)}${c && c.limite ? `<small>limite ${brlC(c.limite)}</small>` : ''}</span></button>`;
      }).join('') : '<p class="sub">Nenhuma fatura vence neste mês.</p>'}</div>
      ${S.cartoes.length && rel === 0 ? `<p class="sub" style="margin-top:10px">Comprando hoje: ${S.cartoes.map(c => `${esc(c.nome)} → fatura de ${esc(mesSo(mesFaturaDe(hojeISO(), c)))}`).join(' · ')}</p>` : ''}
    </section>`;
  }

  h += `<section class="card"><h3>Próximos 6 meses · o que já está comprometido</h3>${graficoProjecao()}
    <div class="leg" style="margin-top:10px"><span><i class="ponto c-fixas"></i>Contas fixas</span><span><i class="ponto c-var"></i>Variáveis já lançadas</span><span><i class="traco"></i>Renda prevista</span></div>
    <p class="sub" style="margin-top:8px">Os meses à frente mostram só o que já está garantido: contas fixas, parcelas e compras que caem nessas faturas. O número embaixo é a sobra.</p></section>`;

  const cats = Object.entries(R.cats).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
  const topo = cats.slice(0, 8); const resto = cats.slice(8).reduce((s, [, v]) => s + v, 0);
  if (resto > 0) topo.push(['Outras categorias', resto]);
  const maxC = Math.max(1, ...topo.map(x => x[1]));
  const pes = [...S.pessoas.map(p => [p.nome, R.pessoas[p.id] || 0]), ['Casal', R.pessoas.casal || 0]].filter(x => x[1] > 0);
  const maxP = Math.max(1, ...pes.map(x => x[1]));
  h += `<div class="grid dois">
    <section class="card"><h3>Para onde vai o dinheiro</h3>${topo.length ? topo.map(([c, v]) => hbar(c, v, maxC, comp)).join('') : '<p class="sub">Nada comprometido neste mês.</p>'}</section>
    <section class="card"><h3>Por pessoa</h3>${pes.length ? pes.map(([n, v]) => hbar(n, v, maxP, comp)).join('') : '<p class="sub">Nada comprometido neste mês.</p>'}<p class="sub" style="margin-top:8px">“Casal” são as contas fixas divididas. VR e VA ficam de fora.</p></section>
  </div>`;

  const fixasMes = S.fixas.filter(f => ativo(f, MES)).sort((a, b) => a.dia - b.dia);
  if (fixasMes.length) {
    const hojeDia = +hojeISO().slice(8);
    h += `<section class="card"><h3>Contas fixas de ${esc(mesSo(MES))}<button class="btn peq" data-acao="nova-fixa">+ Conta fixa</button></h3><div class="lista">${fixasMes.map(f => {
      const passou = rel === 0 && hojeDia > f.dia;
      return `<button class="linha ${passou ? 'passou' : ''}" data-acao="editar-fixa" data-id="${esc(f.id)}">
        <span class="l-ico">${pad(f.dia)}<small>${MESES_C[+MES.slice(5) - 1]}</small></span>
        <span class="l-desc"><b>${esc(f.descricao)}</b><small>${esc(f.categoria)} · ${esc(pessoaNome(f.pessoaId))} · ${f.meio === 'cartao' ? 'no cartão ' + esc(nomeCartao(f.cartaoId)) : esc(ROT_MEIO[f.meio] || '')}${passou ? ' · dia já passou' : ''}</small></span>
        <span class="l-valor">${brl(f.valor)}</span></button>`;
    }).join('')}</div></section>`;
  }
  return h;
}

/* ---------- Lançamentos ---------- */
function vLancamentos() {
  const meios = [['', 'Todos os meios'], ['cartao', 'Todos os cartões'], ...S.cartoes.map(c => ['cartao:' + c.id, 'Cartão ' + c.nome]), ['pix', 'Pix'], ['debito', 'Débito'], ['dinheiro', 'Dinheiro'], ['boleto', 'Boleto/conta'], ['vr', 'VR'], ['va', 'VA']];
  if (FILTRO.meio && !meios.some(m => m[0] === FILTRO.meio)) meios.push([FILTRO.meio, nomeCartao(FILTRO.meio.slice(7))]);
  const op = (lista, sel) => lista.map(([v, t]) => `<option value="${esc(v)}" ${v === sel ? 'selected' : ''}>${esc(t)}</option>`).join('');
  return `<section class="card">
    <div class="ferr">
      <input type="search" data-filtro="q" placeholder="Buscar descrição" value="${esc(FILTRO.q)}" aria-label="Buscar">
      <select data-filtro="pessoa" aria-label="Pessoa">${op([['', 'Todas as pessoas'], ...S.pessoas.map(p => [p.id, p.nome])], FILTRO.pessoa)}</select>
      <select data-filtro="meio" aria-label="Meio de pagamento">${op(meios, FILTRO.meio)}</select>
      <select data-filtro="cat" aria-label="Categoria">${op([['', 'Todas as categorias'], ...catOps()], FILTRO.cat)}</select>
      <button class="btn primario" data-acao="novo-lanc">+ Lançamento</button>
    </div>
    <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin:14px 0 4px">
      <div class="seg-ctl" role="group" aria-label="Como listar">
        <button data-acao="visao" data-v="compra" aria-pressed="${FILTRO.visao === 'compra'}">Compras do mês</button>
        <button data-acao="visao" data-v="pagamento" aria-pressed="${FILTRO.visao === 'pagamento'}">O que paga no mês</button>
      </div>
      ${FILTRO.q || FILTRO.pessoa || FILTRO.meio || FILTRO.cat ? '<button class="link" data-acao="limpar-filtros">Limpar filtros</button>' : ''}
    </div>
    <p class="sub" id="resumoLanc" style="margin:8px 0 4px"></p>
    <div class="lista" id="listaLanc"></div>
  </section>`;
}
function renderListaLanc() {
  const q = norm(FILTRO.q), pag = FILTRO.visao === 'pagamento';
  const itens = [];
  for (const l of S.lancamentos) {
    let p = 1;
    if (pag) { p = ocorrencia(l, MES); if (!p) continue; }
    else if (l.data.slice(0, 7) !== MES) continue;
    if (FILTRO.pessoa && l.pessoaId !== FILTRO.pessoa) continue;
    if (FILTRO.meio) {
      if (FILTRO.meio === 'cartao') { if (l.meio !== 'cartao') continue; }
      else if (FILTRO.meio.startsWith('cartao:')) { if (l.meio !== 'cartao' || (l.cartaoId || '') !== FILTRO.meio.slice(7)) continue; }
      else if (l.meio !== FILTRO.meio) continue;
    }
    if (FILTRO.cat && l.categoria !== FILTRO.cat) continue;
    if (q && !norm(l.descricao + ' ' + (l.obs || '')).includes(q)) continue;
    const mc = l.data.slice(0, 7);
    let ordem = 0, grupo = `${diaSemana(l.data)}, ${dataCurta(l.data)}`;
    if (pag && mc < MES) {
      if ((l.parcelas || 1) > 1) { ordem = 2; grupo = 'Parcelas de compras anteriores'; }
      else { ordem = 1; grupo = `Compras de ${mesSo(mc)} que caem neste mês`; }
    }
    const valor = pag ? l.valor : round2(l.valor * (l.parcelas || 1));
    itens.push({ l, p, ordem, grupo, valor });
  }
  itens.sort((a, b) => a.ordem - b.ordem || b.l.data.localeCompare(a.l.data) || a.l.descricao.localeCompare(b.l.descricao));
  const tot = itens.reduce((s, x) => s + (ehVale(x.l.meio) ? 0 : x.valor), 0);
  const totV = itens.reduce((s, x) => s + (ehVale(x.l.meio) ? x.valor : 0), 0);
  $('#resumoLanc').innerHTML = itens.length ? `<b>${itens.length}</b> lançamento${itens.length > 1 ? 's' : ''} · <b>${brl(tot)}</b>${totV ? ` + ${brl(totV)} em VR/VA` : ''} · ${pag ? 'compras no cartão aparecem no mês da fatura' : 'compras parceladas aparecem pelo valor total'}` : '';
  if (!itens.length) {
    $('#listaLanc').innerHTML = `<div class="vazio">${S.lancamentos.length ? 'Nada com esses filtros neste mês.' : 'Nenhum lançamento ainda. Use “+ Lançamento”, importe uma fatura ou cole uma notificação do banco.'}</div>`;
    return;
  }
  let h = '', ant = null;
  for (const { l, p, grupo, valor } of itens) {
    if (grupo !== ant) { h += `<div class="grupo">${esc(grupo)}</div>`; ant = grupo; }
    const meta = [esc(l.categoria), esc(pessoaNome(l.pessoaId)), esc(rotMeio(l.meio, l.cartaoId))];
    if ((l.parcelas || 1) > 1) meta.push(pag ? `parcela ${p}/${l.parcelas}` : `${l.parcelas}x de ${brl(l.valor)}`);
    if (l.origem === 'import') meta.push('importado'); else if (l.origem === 'notificacao') meta.push('notificação');
    h += `<button class="linha" data-acao="editar-lanc" data-id="${esc(l.id)}"><span class="l-ico">${esc(l.data.slice(8))}<small>${MESES_C[+l.data.slice(5, 7) - 1]}</small></span><span class="l-desc"><b>${esc(l.descricao)}</b><small>${meta.join(' · ')}</small></span><span class="l-valor ${valor < 0 ? 'cred' : ''}">${brl(valor)}</span></button>`;
  }
  $('#listaLanc').innerHTML = h;
}

/* ---------- Cadastros ---------- */
function vCadastros() {
  const rotTipo = { salario: 'Salário', vr: 'VR · refeição', va: 'VA · alimentação', extra: 'Renda extra' };
  const periodo = it => it.inicio && it.fim && it.inicio === it.fim ? `só em ${nomeMesC(it.inicio)}` : [it.inicio ? 'desde ' + nomeMesC(it.inicio) : '', it.fim ? 'até ' + nomeMesC(it.fim) : ''].filter(Boolean).join(' ');
  let h = `<section class="card"><h3>Renda e vales<button class="btn peq" data-acao="nova-receita">+ Renda ou vale</button></h3>`;
  if (!S.receitas.length) h += `<p class="sub">Cadastre o salário líquido de cada um e os vales (VR/VA). Bônus e freelas também entram aqui, marcados como “só neste mês”.</p>`;
  for (const p of S.pessoas) {
    const its = S.receitas.filter(r => r.pessoaId === p.id).sort((a, b) => (a.tipo === 'salario' ? 0 : 1) - (b.tipo === 'salario' ? 0 : 1));
    if (!its.length) continue;
    h += `<div class="pessoa-cab">${esc(p.nome)}</div><div class="lista">${its.map(r => `<button class="linha ${ativo(r, MES) ? '' : 'inativo'}" data-acao="editar-receita" data-id="${esc(r.id)}">
      <span class="l-ico">${esc(r.tipo === 'vr' ? 'VR' : r.tipo === 'va' ? 'VA' : r.tipo === 'extra' ? '+' : 'R$')}</span>
      <span class="l-desc"><b>${esc(r.descricao || rotTipo[r.tipo])}</b><small>${esc(rotTipo[r.tipo])} · dia ${esc(r.dia)} · ${esc(periodo(r))}${ativo(r, MES) ? '' : ' · fora deste mês'}</small></span>
      <span class="l-valor">${brl(r.valor)}</span></button>`).join('')}</div>`;
  }
  h += `</section>`;

  h += `<section class="card"><h3>Cartões de crédito<button class="btn peq" data-acao="novo-cartao">+ Cartão</button></h3>`;
  h += S.cartoes.length ? `<div class="lista">${S.cartoes.map(c => `<button class="linha" data-acao="editar-cartao" data-id="${esc(c.id)}">
      <span class="l-ico">${esc(c.nome.slice(0, 2).toUpperCase())}</span>
      <span class="l-desc"><b>${esc(c.nome)}${c.final ? ` <span class="muted">•• ${esc(c.final)}</span>` : ''}</b><small>${esc(pessoaNome(c.pessoaId))} · fecha dia ${esc(c.fechamento)} · vence dia ${esc(c.vencimento)}</small></span>
      <span class="l-valor">${c.limite ? brl(c.limite) + '<small>limite</small>' : ''}</span></button>`).join('')}</div>`
    : `<p class="sub">Com o dia de fechamento e de vencimento, cada compra vai para a fatura certa — e as parcelas se espalham pelos meses seguintes.</p>`;
  h += `</section>`;

  const fixas = S.fixas.slice().sort((a, b) => a.dia - b.dia);
  const totFix = fixas.filter(f => ativo(f, MES)).reduce((s, f) => s + f.valor, 0);
  h += `<section class="card"><h3>Contas fixas${fixas.length ? `<span class="muted">${brl(totFix)} em ${esc(mesSo(MES))}</span>` : ''}<button class="btn peq" data-acao="nova-fixa">+ Conta fixa</button></h3>`;
  h += fixas.length ? `<div class="lista">${fixas.map(f => `<button class="linha ${ativo(f, MES) ? '' : 'inativo'}" data-acao="editar-fixa" data-id="${esc(f.id)}">
      <span class="l-ico">${pad(f.dia)}<small>dia</small></span>
      <span class="l-desc"><b>${esc(f.descricao)}</b><small>${esc(f.categoria)} · ${esc(pessoaNome(f.pessoaId))} · ${f.meio === 'cartao' ? 'cartão ' + esc(nomeCartao(f.cartaoId)) : esc(ROT_MEIO[f.meio] || '')}${f.inicio || f.fim ? ' · ' + esc(periodo(f)) : ''}</small></span>
      <span class="l-valor">${brl(f.valor)}</span></button>`).join('')}</div>`
    : `<p class="sub">Aluguel, condomínio, luz, água, internet, celular, plano de saúde, academia, assinaturas… Tudo que se repete todo mês com o mesmo valor (ou quase).</p>`;
  h += `</section>`;
  return h;
}

/* ---------- Ajustes ---------- */
function vAjustes() {
  const outra = S.pessoas[1]?.nome || 'a outra pessoa';
  return `<section class="card"><h3>Pessoas</h3><div class="cfg" style="margin:0">${S.pessoas.map((p, i) => `<label class="campo"><span class="lbl">Pessoa ${i + 1}</span><input type="text" data-pessoa="${i}" value="${esc(p.nome)}"></label>`).join('')}</div></section>
  <section class="card"><h3>Aparência</h3><div class="seg-ctl" role="group" aria-label="Tema">${[['auto', 'Automático'], ['claro', 'Claro'], ['escuro', 'Escuro']].map(([v, t]) => `<button data-acao="tema" data-v="${v}" aria-pressed="${S.tema === v}">${t}</button>`).join('')}</div></section>
  <section class="card"><h3>Categorias</h3><div class="chips">${S.categorias.map(c => `<span class="chip">${esc(c)}${c !== 'Outros' ? `<button data-acao="rem-cat" data-v="${esc(c)}" aria-label="Remover ${esc(c)}">×</button>` : ''}</span>`).join('')}</div>
    <div class="linha-add"><input type="text" id="novaCat" placeholder="Nova categoria" aria-label="Nova categoria"><button class="btn" data-acao="add-cat">Adicionar</button></div></section>
  <section class="card"><h3>Regras de categoria</h3>
    <p class="sub">Quando a descrição tem o termo, o lançamento já entra na categoria. Suas regras valem antes das ${REGRAS_PADRAO_LISTA.length} regras prontas — e são criadas sozinhas quando você corrige uma categoria.</p>
    <div class="regras">${S.regras.length ? S.regras.map((r, i) => `<div class="regra"><code title="${esc(r.termo)}">${esc(r.termo)}</code><span class="muted">→</span><select data-regra="${i}" aria-label="Categoria da regra">${opcoesCat(r.categoria)}</select><button class="btn peq" data-acao="rem-regra" data-i="${i}" aria-label="Remover regra">×</button></div>`).join('') : '<p class="sub muted">Nenhuma regra sua ainda.</p>'}</div>
    <div class="linha-add"><input type="text" id="novaRegraTermo" placeholder="Termo (ex.: padaria do joão)" aria-label="Termo"><select id="novaRegraCat" aria-label="Categoria">${opcoesCat('Outros')}</select><button class="btn" data-acao="add-regra">Adicionar</button></div>
    <div class="botoes"><button class="btn peq" data-acao="recategorizar">Reaplicar regras nos lançamentos em “Outros”</button></div></section>
  ${secaoConta()}
  <section class="card"><h3>Seus dados</h3>
    <p class="sub">${CASA ? `Os dados ficam na nuvem da casa, iguais nos dois celulares, e com uma cópia aqui para funcionar sem internet. O backup é uma garantia extra.` : 'Tudo fica salvo só neste navegador, neste aparelho. Baixe um backup de vez em quando — é também o jeito de levar os dados para outro aparelho.'}</p>
    <div class="botoes"><button class="btn" data-acao="exportar-json">Baixar backup (.json)</button><button class="btn" data-acao="importar-json">Restaurar backup</button><button class="btn" data-acao="exportar-csv">Lançamentos em planilha (.csv)</button><button class="btn perigo" data-acao="apagar-tudo">${CASA ? 'Apagar todos os dados da casa' : 'Apagar tudo'}</button></div>
    <input type="file" id="arqBackup" accept=".json,application/json" hidden>
    ${!CASA && !S.lancamentos.length && !S.receitas.length ? `<p class="sub" style="margin-top:12px"><button class="link" data-acao="exemplo">Carregar dados de exemplo</button> para conhecer o app.</p>` : ''}
  </section>
  <section class="card"><h3>Fases do app</h3><ol class="fases">
    <li><b>Fase 1 — no navegador.</b> Cadastros, painel do mês, projeção de 6 meses, importação de faturas e leitura de notificações coladas.</li>
    <li><b>Fase 2 — no celular dos dois.</b> Nuvem com login: você e ${esc(outra)} lançam cada um do seu celular e veem os mesmos números. Instala na tela de início como app.${CASA ? ' <b>Ativa.</b>' : ''}</li>
    <li><b>Fase 3 — compras chegando sozinhas.</b> No iPhone, pelo app Atalhos: automação de “Transação” do Apple Pay e de SMS/e-mail de compra do banco mandando o texto para cá.</li>
  </ol></section>`;
}
function secaoConta() {
  if (!NUVEM_ATIVA) return '';
  if (!CASA) return `<section class="card"><h3>Conta e casa</h3><p class="sub">Você não está conectado à casa neste aparelho.</p><div class="botoes"><button class="btn primario" data-acao="ir-login">Entrar</button></div></section>`;
  const membros = CASA.membros || 1;
  return `<section class="card" id="secConta"><h3>Conta e casa</h3>
    <div class="det-lista">
      <span>Conta</span><b>${esc(SESSAO?.user?.email || ULT.email || '—')}</b>
      <span>Casa</span><b>${esc(CASA.nome)}</b>
      <span>Você é</span><b>${esc(pessoaNome(CASA.pessoa_id))}</b>
      <span>Pessoas na casa</span><b>${membros} de 2</b>
      <span>Sincronização</span><b>${esc((ROT_NUVEM[NUVEM.estado] || ['—'])[0])}${NUVEM.msg && NUVEM.estado === 'erro' ? ` — ${esc(NUVEM.msg)}` : ''}</b>
    </div>
    ${membros < 2 ? `<p class="sub" style="margin-top:14px">Código para ${esc(S.pessoas.find(p => p.id !== CASA.pessoa_id)?.nome || 'a outra pessoa')} entrar na casa (é só criar a própria conta no app e digitar este código):</p>` : `<p class="sub" style="margin-top:14px">Código de convite (só serve se alguém sair da casa):</p>`}
    <div class="codigo"><code>${esc(fmtCodigo(CASA.codigo_convite))}</code><button class="btn peq" data-acao="copiar-codigo">Copiar</button><button class="btn peq" data-acao="trocar-codigo">Gerar outro</button></div>
    <div class="botoes"><button class="btn" data-acao="sincronizar">Sincronizar agora</button><button class="btn" data-acao="sair">Sair deste aparelho</button></div>
  </section>`;
}

/* =====================================================================
   FORMULÁRIOS
   ===================================================================== */
function campoHTML(c) {
  const id = 'f_' + c.nome;
  let inp;
  if (c.tipo === 'checkbox') return `<label class="campo chk" data-campo="${c.nome}"><input type="checkbox" name="${c.nome}" ${c.valor ? 'checked' : ''}> <span>${esc(c.label)}</span></label>`;
  if (c.tipo === 'select') {
    const ops = c.opcoes.slice();
    if (c.valor && !ops.some(([v]) => String(v) === String(c.valor))) ops.unshift([c.valor, c.valor]);
    inp = `<select id="${id}" name="${c.nome}">${ops.map(([v, t]) => `<option value="${esc(v)}" ${String(v) === String(c.valor) ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
  } else if (c.tipo === 'dinheiro') {
    inp = `<div class="dinheiro"><span>R$</span><input id="${id}" name="${c.nome}" type="text" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${c.valor === '' || c.valor == null ? '' : esc(num2(c.valor))}"></div>`;
  } else if (c.tipo === 'number') {
    inp = `<input id="${id}" name="${c.nome}" type="number" inputmode="numeric" step="1" ${c.min != null ? `min="${c.min}"` : ''} ${c.max != null ? `max="${c.max}"` : ''} value="${esc(c.valor ?? '')}">`;
  } else {
    inp = `<input id="${id}" name="${c.nome}" type="${c.tipo}" value="${esc(c.valor ?? '')}" ${c.placeholder ? `placeholder="${esc(c.placeholder)}"` : ''} ${c.inputmode ? `inputmode="${c.inputmode}"` : ''} ${c.maxlength ? `maxlength="${c.maxlength}"` : ''} ${c.tipo === 'text' ? 'autocomplete="off"' : ''}>`;
  }
  return `<div class="campo ${c.largo ? 'largo' : ''}" data-campo="${c.nome}"><label for="${id}">${esc(c.label)}</label>${inp}${c.dica ? `<small class="dica" data-dica="${c.nome}"></small>` : ''}</div>`;
}
function abrirForm(cfg) {
  const dlg = $('#modal');
  dlg.innerHTML = `<form class="form" novalidate><h2>${esc(cfg.titulo)}</h2>${cfg.info ? `<p class="info sub">${cfg.info}</p>` : ''}${cfg.campos.map(campoHTML).join('')}<p class="erro" role="alert" hidden></p><div class="acoes-form">${cfg.aoExcluir ? '<button type="button" class="btn perigo" data-f="excluir">Excluir</button>' : ''}<span class="sp"></span><button type="button" class="btn" data-f="cancelar">Cancelar</button><button type="submit" class="btn primario">Salvar</button></div></form>`;
  const form = dlg.querySelector('form'), erro = form.querySelector('.erro');
  const ler = () => { const o = {}; for (const c of cfg.campos) { const el = form.elements[c.nome]; if (el) o[c.nome] = c.tipo === 'checkbox' ? el.checked : el.value; } return o; };
  const atualizar = ev => {
    if (cfg.aoMudar) cfg.aoMudar(ler(), form, ev);
    const v = ler();
    for (const c of cfg.campos) if (c.mostrarSe) form.querySelector(`[data-campo="${c.nome}"]`).hidden = !c.mostrarSe(v);
  };
  form.addEventListener('input', atualizar);
  form.addEventListener('change', atualizar);
  form.addEventListener('submit', e => {
    e.preventDefault();
    const v = ler(), msg = cfg.validar && cfg.validar(v);
    if (msg) { erro.textContent = msg; erro.hidden = false; return; }
    cfg.aoSalvar(v); salvar(); dlg.close(); render();
  });
  form.querySelector('[data-f="cancelar"]').onclick = () => dlg.close();
  const bx = form.querySelector('[data-f="excluir"]');
  if (bx) bx.onclick = () => { if (!confirm(cfg.confirmaExcluir || 'Excluir este item?')) return; cfg.aoExcluir(); salvar(); dlg.close(); render(); toast('Excluído'); };
  atualizar(null);
  dlg.showModal();
  if (!matchMedia('(pointer:coarse)').matches) form.querySelector('input:not([type=checkbox]),select')?.focus();
}
const opcoesMeioLanc = () => [...S.cartoes.map(c => ['cartao:' + c.id, `Cartão ${c.nome}${c.final ? ' •• ' + c.final : ''}`]), ...(S.cartoes.length ? [] : [['cartao:', 'Cartão de crédito']]), ['pix', 'Pix'], ['debito', 'Débito'], ['dinheiro', 'Dinheiro'], ['boleto', 'Boleto / conta'], ['vr', 'VR (refeição)'], ['va', 'VA (alimentação)']];

function formLancamento(l, preset) {
  const novo = !l;
  const b = l || Object.assign({ data: hojeISO(), descricao: '', valor: 0, meio: 'cartao', cartaoId: (S.cartoes.find(c => c.pessoaId === EU()) || S.cartoes[0])?.id || '', parcelas: 1, pessoaId: EU(), categoria: '', obs: '' }, preset || {});
  const meioVal = b.meio === 'cartao' ? 'cartao:' + (b.cartaoId || '') : b.meio;
  const total = round2(b.valor * (b.parcelas || 1));
  const faturaFixa = l && l.origem === 'import' && l.meio === 'cartao';
  let catTocada = false, pesTocada = !novo;
  const info = [];
  if (preset && preset.origem === 'notificacao') info.push('Li a notificação assim — confira e salve.');
  if (l && (l.parcelas || 1) > 1) info.push(`Compra em ${l.parcelas}x de ${brl(l.valor)}.${faturaFixa ? ` Veio da fatura de ${esc(nomeMes(l.mesRef))} como parcela ${l.parcelaAtual}/${l.parcelas}.` : ''} Mudar aqui muda todas as parcelas.`);
  abrirForm({
    titulo: novo ? 'Novo lançamento' : 'Editar lançamento',
    info: info.join(' '),
    campos: [
      { nome: 'descricao', label: 'Descrição', tipo: 'text', valor: b.descricao, placeholder: 'Ex.: Mercado, iFood, Farmácia', largo: true },
      { nome: 'valor', label: 'Valor total', tipo: 'dinheiro', valor: total || '' },
      { nome: 'data', label: 'Data da compra', tipo: 'date', valor: b.data },
      { nome: 'meio', label: 'Pago com', tipo: 'select', opcoes: opcoesMeioLanc(), valor: meioVal, dica: true },
      { nome: 'parcelas', label: 'Parcelas', tipo: 'number', min: 1, max: 48, valor: b.parcelas || 1, dica: true, mostrarSe: v => v.meio.startsWith('cartao') },
      { nome: 'pessoaId', label: 'Quem comprou', tipo: 'select', opcoes: S.pessoas.map(p => [p.id, p.nome]), valor: b.pessoaId },
      { nome: 'categoria', label: 'Categoria', tipo: 'select', opcoes: catOps(), valor: b.categoria || categoriaAuto(b.descricao, b.meio) },
      { nome: 'lembrar', label: 'Sempre usar esta categoria', tipo: 'checkbox', valor: true, mostrarSe: v => catTocada && !!termoDe(v.descricao) && v.categoria !== categorizar(v.descricao) },
      { nome: 'obs', label: 'Observação (opcional)', tipo: 'text', valor: b.obs || '', largo: true },
    ],
    aoMudar: (v, form, ev) => {
      const alvo = ev && ev.target && ev.target.name;
      if (alvo === 'categoria') catTocada = true;
      if (alvo === 'pessoaId') pesTocada = true;
      const meioSimples = v.meio.startsWith('cartao') ? 'cartao' : v.meio;
      if (novo && !catTocada && (alvo === 'descricao' || alvo === 'meio')) form.elements.categoria.value = categoriaAuto(v.descricao, meioSimples);
      if (alvo === 'meio' && !pesTocada && v.meio.startsWith('cartao:')) { const c = cartaoPorId(v.meio.slice(7)); if (c) form.elements.pessoaId.value = c.pessoaId; }
      const n = parseInt(form.elements.parcelas.value, 10) || 1, tot = parseValor(form.elements.valor.value);
      form.querySelector('[data-dica="parcelas"]').textContent = n > 1 && tot > 0 ? `${n}x de ${brl(tot / n)}` : '';
      let dm = '';
      if (v.meio.startsWith('cartao') && !faturaFixa) { const c = cartaoPorId(v.meio.slice(7)); if (c && /^\d{4}-\d{2}-\d{2}$/.test(v.data)) dm = `Cai na fatura de ${mesSo(mesFaturaDe(v.data, c))}${n > 1 ? ` e segue por ${n} meses` : ''}.`; }
      form.querySelector('[data-dica="meio"]').textContent = dm;
      const lb = form.querySelector('[data-campo="lembrar"] span');
      if (lb) lb.textContent = `Sempre classificar “${termoDe(v.descricao)}” como ${form.elements.categoria.value}`;
    },
    validar: v => !v.descricao.trim() ? 'Preencha a descrição.' : !(parseValor(v.valor) > 0) ? 'Informe um valor maior que zero.' : !/^\d{4}-\d{2}-\d{2}$/.test(v.data) ? 'Informe a data.' : null,
    aoSalvar: v => {
      const tot = parseValor(v.valor);
      const [meio, cid] = v.meio.startsWith('cartao:') ? ['cartao', v.meio.slice(7)] : [v.meio, ''];
      const parcelas = meio === 'cartao' ? Math.max(1, Math.min(48, parseInt(v.parcelas, 10) || 1)) : 1;
      const o = l || { id: uid(), origem: (preset && preset.origem) || 'manual', criadoEm: hojeISO() };
      const manter = faturaFixa && meio === 'cartao';
      Object.assign(o, { descricao: v.descricao.trim(), data: v.data, meio, cartaoId: cid, pessoaId: v.pessoaId, categoria: v.categoria, obs: v.obs.trim(), parcelas, valor: round2(tot / parcelas) });
      if (manter) { o.parcelaAtual = Math.min(o.parcelaAtual || 1, parcelas); }
      else { o.parcelaAtual = 1; o.mesRefAuto = true; o.mesRef = calcMesRef(o); }
      if (v.lembrar && catTocada && termoDe(o.descricao) && o.categoria !== categorizar(o.descricao)) aprenderRegra(o.descricao, o.categoria);
      if (!l) S.lancamentos.push(o);
      toast(l ? 'Lançamento atualizado' : `Lançamento salvo${meio === 'cartao' ? ' · fatura de ' + mesSo(o.mesRef) : ''}`);
    },
    aoExcluir: l ? () => { S.lancamentos = S.lancamentos.filter(x => x.id !== l.id); } : null,
    confirmaExcluir: l && (l.parcelas || 1) > 1 ? 'Excluir a compra inteira, com todas as parcelas?' : 'Excluir este lançamento?',
  });
}
function formReceita(r) {
  const b = r || { pessoaId: EU(), tipo: 'salario', descricao: '', valor: '', dia: 5, inicio: MES, fim: '' };
  const unico = !!(b.inicio && b.fim && b.inicio === b.fim);
  abrirForm({
    titulo: r ? 'Editar renda' : 'Nova renda ou vale',
    info: 'VR e VA ficam separados: têm saldo próprio no painel e não entram na sobra do salário. Para parar de contar uma renda, preencha “Até” em vez de excluir — assim os meses passados continuam certos.',
    campos: [
      { nome: 'pessoaId', label: 'De quem', tipo: 'select', opcoes: S.pessoas.map(p => [p.id, p.nome]), valor: b.pessoaId },
      { nome: 'tipo', label: 'Tipo', tipo: 'select', opcoes: [['salario', 'Salário (líquido)'], ['vr', 'Vale-refeição (VR)'], ['va', 'Vale-alimentação (VA)'], ['extra', 'Renda extra / outra']], valor: b.tipo },
      { nome: 'descricao', label: 'Descrição (opcional)', tipo: 'text', valor: b.descricao, placeholder: 'Ex.: Salário, Freela, 13º', largo: true },
      { nome: 'valor', label: 'Valor por mês', tipo: 'dinheiro', valor: b.valor },
      { nome: 'dia', label: 'Dia em que cai', tipo: 'number', min: 1, max: 31, valor: b.dia },
      { nome: 'inicio', label: 'A partir de', tipo: 'month', valor: b.inicio },
      { nome: 'fim', label: 'Até (vazio = sem fim)', tipo: 'month', valor: unico ? '' : b.fim, mostrarSe: v => !v.unico },
      { nome: 'unico', label: 'Só neste mês (bônus, 13º, freela)', tipo: 'checkbox', valor: unico },
    ],
    validar: v => !(parseValor(v.valor) > 0) ? 'Informe o valor.' : !/^\d{4}-\d{2}$/.test(v.inicio) ? 'Informe o mês de início (aaaa-mm).' : (!v.unico && v.fim && v.fim < v.inicio) ? 'O mês final está antes do inicial.' : null,
    aoSalvar: v => {
      const o = r || { id: uid() };
      Object.assign(o, { pessoaId: v.pessoaId, tipo: v.tipo, descricao: v.descricao.trim(), valor: round2(parseValor(v.valor)), dia: clampDia(v.dia), inicio: v.inicio, fim: v.unico ? v.inicio : (v.fim || '') });
      if (!r) S.receitas.push(o);
      toast('Renda salva');
    },
    aoExcluir: r ? () => { S.receitas = S.receitas.filter(x => x.id !== r.id); } : null,
  });
}
function formCartao(c) {
  const b = c || { nome: '', pessoaId: EU(), final: '', fechamento: '', vencimento: '', limite: '' };
  abrirForm({
    titulo: c ? 'Editar cartão' : 'Novo cartão de crédito',
    info: 'O dia do fechamento é o “melhor dia de compra”: o que é comprado desse dia em diante já vai para a fatura seguinte.',
    campos: [
      { nome: 'nome', label: 'Nome do cartão', tipo: 'text', valor: b.nome, placeholder: 'Ex.: Nubank, Itaú Click', largo: true },
      { nome: 'pessoaId', label: 'Titular', tipo: 'select', opcoes: S.pessoas.map(p => [p.id, p.nome]), valor: b.pessoaId },
      { nome: 'final', label: 'Últimos 4 dígitos (opcional)', tipo: 'text', valor: b.final, placeholder: '0000', inputmode: 'numeric', maxlength: 4 },
      { nome: 'fechamento', label: 'Dia do fechamento', tipo: 'number', min: 1, max: 31, valor: b.fechamento },
      { nome: 'vencimento', label: 'Dia do vencimento', tipo: 'number', min: 1, max: 31, valor: b.vencimento, dica: true },
      { nome: 'limite', label: 'Limite (opcional)', tipo: 'dinheiro', valor: b.limite || '' },
    ],
    aoMudar: (v, form) => {
      const f = +v.fechamento, ve = +v.vencimento;
      form.querySelector('[data-dica="vencimento"]').textContent = f >= 1 && f <= 31 && ve >= 1 && ve <= 31 ? `Uma compra feita hoje cai na fatura de ${mesSo(mesFaturaDe(hojeISO(), { fechamento: f, vencimento: ve }))}.` : '';
    },
    validar: v => !v.nome.trim() ? 'Dê um nome ao cartão.' : !(+v.fechamento >= 1 && +v.fechamento <= 31) ? 'Informe o dia do fechamento (1 a 31).' : !(+v.vencimento >= 1 && +v.vencimento <= 31) ? 'Informe o dia do vencimento (1 a 31).' : (v.final && !/^\d{4}$/.test(v.final.trim())) ? 'Os últimos dígitos são 4 números.' : null,
    aoSalvar: v => {
      const o = c || { id: uid() };
      const lim = parseValor(v.limite);
      Object.assign(o, { nome: v.nome.trim(), pessoaId: v.pessoaId, final: v.final.trim(), fechamento: +v.fechamento, vencimento: +v.vencimento, limite: lim > 0 ? round2(lim) : 0 });
      if (!c) S.cartoes.push(o);
      for (const l of S.lancamentos) if (l.meio === 'cartao' && l.cartaoId === o.id && l.mesRefAuto) l.mesRef = calcMesRef(l);
      toast('Cartão salvo');
    },
    aoExcluir: c ? () => { S.cartoes = S.cartoes.filter(x => x.id !== c.id); } : null,
    confirmaExcluir: 'Excluir este cartão? Os lançamentos dele continuam, marcados como “cartão removido”.',
  });
}
function formFixa(f) {
  const b = f || { descricao: '', categoria: 'Moradia', valor: '', dia: 10, pessoaId: 'casal', meio: 'conta', cartaoId: '', inicio: MES, fim: '' };
  const meioVal = b.meio === 'cartao' ? 'cartao:' + b.cartaoId : b.meio;
  let catTocada = !!f;
  abrirForm({
    titulo: f ? 'Editar conta fixa' : 'Nova conta fixa',
    info: f ? 'Se a conta acabou (ou mudou de valor), preencha “Até” e crie a nova — assim os meses passados continuam certos.' : '',
    campos: [
      { nome: 'descricao', label: 'Descrição', tipo: 'text', valor: b.descricao, placeholder: 'Ex.: Aluguel, Internet, Netflix', largo: true },
      { nome: 'valor', label: 'Valor por mês', tipo: 'dinheiro', valor: b.valor },
      { nome: 'dia', label: 'Dia do vencimento/cobrança', tipo: 'number', min: 1, max: 31, valor: b.dia },
      { nome: 'categoria', label: 'Categoria', tipo: 'select', opcoes: catOps(), valor: b.categoria },
      { nome: 'pessoaId', label: 'De quem', tipo: 'select', opcoes: [...S.pessoas.map(p => [p.id, p.nome]), ['casal', 'Casal (dividida)']], valor: b.pessoaId },
      { nome: 'meio', label: 'Paga com', tipo: 'select', opcoes: [['conta', 'Conta / boleto / Pix'], ...S.cartoes.map(c => ['cartao:' + c.id, 'Cartão ' + c.nome]), ['vr', 'VR'], ['va', 'VA']], valor: meioVal, dica: true, largo: true },
      { nome: 'inicio', label: 'A partir de', tipo: 'month', valor: b.inicio },
      { nome: 'fim', label: 'Até (vazio = sem fim)', tipo: 'month', valor: b.fim },
    ],
    aoMudar: (v, form, ev) => {
      const alvo = ev && ev.target && ev.target.name;
      if (alvo === 'categoria') catTocada = true;
      if (!catTocada && alvo === 'descricao') { const c = categorizar(v.descricao); if (c !== 'Outros') form.elements.categoria.value = c; }
      form.querySelector('[data-dica="meio"]').textContent = v.meio.startsWith('cartao:') ? 'Entra todo mês na fatura desse cartão (conta como fixa, não como variável).' : ehVale(v.meio) ? 'Sai do saldo do vale, não do salário.' : '';
    },
    validar: v => !v.descricao.trim() ? 'Preencha a descrição.' : !(parseValor(v.valor) > 0) ? 'Informe o valor.' : (v.inicio && !/^\d{4}-\d{2}$/.test(v.inicio)) ? 'Mês de início no formato aaaa-mm.' : (v.fim && v.inicio && v.fim < v.inicio) ? 'O mês final está antes do inicial.' : null,
    aoSalvar: v => {
      const o = f || { id: uid() };
      const [meio, cid] = v.meio.startsWith('cartao:') ? ['cartao', v.meio.slice(7)] : [v.meio, ''];
      Object.assign(o, { descricao: v.descricao.trim(), valor: round2(parseValor(v.valor)), dia: clampDia(v.dia), categoria: v.categoria, pessoaId: v.pessoaId, meio, cartaoId: cid, inicio: v.inicio || '', fim: v.fim || '' });
      if (!f) S.fixas.push(o);
      toast('Conta fixa salva');
    },
    aoExcluir: f ? () => { S.fixas = S.fixas.filter(x => x.id !== f.id); } : null,
  });
}

/* =====================================================================
   IMPORTAÇÃO (tudo roda aqui no navegador)
   ===================================================================== */
const scripts = {};
function carregarScript(src) {
  if (!scripts[src]) scripts[src] = new Promise((ok, falha) => {
    const s = document.createElement('script'); s.src = src; s.async = true; s.crossOrigin = 'anonymous';
    s.onload = ok; s.onerror = () => { delete scripts[src]; falha(new Error('sem internet para carregar o leitor de arquivo')); };
    document.head.appendChild(s);
  });
  return scripts[src];
}
function decodificar(buf) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
  catch (e) { return new TextDecoder('windows-1252').decode(buf); }
}
function parseValor(v) {
  if (typeof v === 'number') return v;
  let s = String(v ?? '').trim(); if (!s) return NaN;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (/\sD$/i.test(s)) { neg = true; s = s.slice(0, -1); } else if (/\sC$/i.test(s)) s = s.slice(0, -1);
  if (/-\s*$/.test(s)) { neg = true; s = s.replace(/-\s*$/, ''); }
  if (/^[^\d]*-/.test(s)) neg = true;
  s = s.replace(/[^\d.,]/g, '');
  if (!/\d/.test(s)) return NaN;
  const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc > ld) s = s.replace(/\./g, '').replace(',', '.');
  else if (ld > lc) {
    if (lc === -1 && s.length - ld - 1 === 3 && !/^0\./.test(s)) s = s.replace(/\./g, '');
    else s = s.replace(/,/g, '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? NaN : (neg ? -n : n);
}
function parseData(v, mesRef) {
  if (v instanceof Date) return isNaN(v) ? null : `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 864e5)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
    return null;
  }
  const s = semAcento(String(v || '').trim().toLowerCase());
  let m, a, me, d;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) [a, me, d] = [+m[1], +m[2], +m[3]];
  else if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(?!\d)/))) [d, me, a] = [+m[1], +m[2], +m[3] < 100 ? 2000 + +m[3] : +m[3]];
  else if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})(?![\d\/.\-])/))) [d, me] = [+m[1], +m[2]];
  else if ((m = s.match(/^(\d{1,2})\s*(?:de\s+)?([a-z]{3})[a-z]*\.?(?:\s*(?:de\s+)?(\d{4}))?/)) && MES_IDX[m[2]]) [d, me, a] = [+m[1], MES_IDX[m[2]], m[3] ? +m[3] : 0];
  else return null;
  if (!(me >= 1 && me <= 12 && d >= 1 && d <= 31)) return null;
  if (!a) { const [ra, rm] = (mesRef || mesAtual()).split('-').map(Number); a = me > rm ? ra - 1 : ra; }
  return `${a}-${pad(me)}-${pad(d)}`;
}
function extrairParcela(desc, raw) {
  let m;
  const s = String(raw ?? '').trim();
  if (s && (m = s.match(/(\d{1,2})\s*(?:\/|de)\s*(\d{1,2})/))) { const a = +m[1], b = +m[2]; if (a >= 1 && b >= 2 && a <= b) return { desc, atual: a, total: b }; }
  for (const re of [/\s*[-–]?\s*\(?\bparc(?:ela)?\.?\s*(\d{1,2})\s*(?:\/|de)\s*(\d{1,2})\)?/i, /\s+(\d{1,2})\s*\/\s*(\d{1,2})\s*$/]) {
    m = desc.match(re);
    if (m) { const a = +m[1], b = +m[2]; if (a >= 1 && b >= 2 && a <= b && b <= 72) return { desc: (desc.slice(0, m.index) + desc.slice(m.index + m[0].length)).replace(/\s*[-–]\s*$/, '').trim(), atual: a, total: b }; }
  }
  return { desc, atual: 1, total: 1 };
}
const limparDesc = s => String(s || '').replace(/[•*·.]{3,}\s*\d{4}/g, ' ').replace(/^\d{1,2}:\d{2}(?::\d{2})?\s+/, '').replace(/^[-–\s]+/, '').replace(/\s*(?:R\$|BRL)\s*$/i, '').replace(/\s+/g, ' ').trim();

/* --- texto em linhas (PDF, texto colado, HTML sem tabela) --- */
const RE_DATA_INICIO = new RegExp(`^(\\d{4}-\\d{2}-\\d{2}|\\d{1,2}[\\/.\\-]\\d{1,2}(?:[\\/.\\-]\\d{2,4})?|\\d{1,2}\\s+(?:de\\s+)?${RE_MES_TXT}[a-zç]*\\.?(?:\\s+(?:de\\s+)?\\d{4})?)(?:\\s+|$)`, 'i');
const RE_VALOR_FIM = /(?:^|\s)((?:-\s*)?(?:R\$\s*)?(?:-\s*)?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}(?:\s*-|\s+[CD])?)\s*$/i;
// quebra linhas com duas colunas de lançamentos (comum em PDF de fatura): só quando a data é seguida de texto
const RE_DATA_COM_TEXTO = `(?:\\d{1,2}[\\/.]\\d{1,2}(?:[\\/.]\\d{2,4})?\\s+[A-Za-zÀ-ú*]|\\d{1,2}\\s+${RE_MES_TXT}\\b)`;
const RE_QUEBRA_GRANDE = new RegExp(`\\s{3,}(?=${RE_DATA_COM_TEXTO})`, 'gi');
const RE_QUEBRA_VALOR = new RegExp(`(\\d,\\d{2}(?:\\s?[-CD])?)\\s+(?=${RE_DATA_COM_TEXTO})`, 'gi');
function dividirColunas(l) { return String(l).replace(RE_QUEBRA_GRANDE, '\n').replace(RE_QUEBRA_VALOR, '$1\n').split('\n'); }
function linhasParaTransacoes(brutas) {
  const linhas = [];
  for (const raw of brutas) for (const parte of dividirColunas(raw)) { const l = parte.replace(/\s+/g, ' ').trim(); if (l) linhas.push(l); }
  const out = [];
  for (let i = 0; i < linhas.length; i++) {
    const md = linhas[i].match(RE_DATA_INICIO); if (!md) continue;
    let texto = linhas[i], j = i;
    while (!RE_VALOR_FIM.test(texto.slice(md[0].length)) && j < i + 3 && j + 1 < linhas.length && !RE_DATA_INICIO.test(linhas[j + 1])) { j++; texto += ' ' + linhas[j]; }
    const resto = texto.slice(md[0].length);
    const mv = resto.match(RE_VALOR_FIM); if (!mv) continue;
    const desc = limparDesc(resto.slice(0, mv.index));
    if (!desc || desc.length < 2 || !/[a-z]/i.test(desc)) continue;
    out.push({ dataRaw: md[1], descricao: desc, valorRaw: mv[1] });
    i = j;
  }
  return out;
}

/* --- tabelas (CSV, Excel, HTML) --- */
function dividirCSV(l, d) {
  const out = []; let cur = '', q = false;
  for (let i = 0; i < l.length; i++) {
    const c = l[i];
    if (q) { if (c === '"') { if (l[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
    else if (c === '"') q = true;
    else if (c === d) { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out.map(s => s.trim());
}
function csvParaTabela(t) {
  const linhas = t.replace(/^﻿/, '').split(/\r?\n/).filter(l => l.trim());
  if (!linhas.length) return [];
  const am = linhas.slice(0, 30);
  const media = d => am.reduce((s, l) => s + dividirCSV(l, d).length - 1, 0) / am.length;
  let melhor = ',', nota = 0;
  for (const d of [';', '\t', '|', ',']) { const n = media(d); if (n >= 1 && n > nota + (d === ',' ? 0.5 : 0)) { melhor = d; nota = n; } }
  if (!nota) return linhas.map(l => [l]);
  return linhas.map(l => dividirCSV(l, melhor));
}
const H_DATA = /^(data|date|dt|dia)\b|data (da |de |do )?(compra|transa|lanc|mov)/;
const H_DESC = /descri|historic|title|titulo|estabelecimento|lancamento|detalhe|memo|nome|loja|transac|favorecido/;
const H_VALOR = /valor|amount|quantia|montante|r\$|preco|^total/;
const H_PARC = /parcela|prestac/;
function detectarMapa(tab) {
  for (let r = 0; r < Math.min(tab.length, 40); r++) {
    const cel = (tab[r] || []).map(c => semAcento(String(c ?? '').toLowerCase().trim()));
    const iData = cel.findIndex(c => c.length < 40 && H_DATA.test(c));
    if (iData < 0) continue;
    let iValor = -1;
    cel.forEach((c, i) => { if (i !== iData && c.length < 40 && H_VALOR.test(c) && !/us\$|usd|dolar|saldo/.test(c) && (iValor < 0 || /r\$|brl|real/.test(c))) iValor = i; });
    if (iValor < 0) continue;
    let iDesc = cel.findIndex((c, i) => i !== iData && i !== iValor && c.length < 40 && H_DESC.test(c));
    const iParc = cel.findIndex((c, i) => i !== iData && i !== iValor && i !== iDesc && H_PARC.test(c));
    if (iDesc < 0) iDesc = colunaTexto(tab, r + 1, [iData, iValor, iParc]);
    return { cab: r, data: iData, desc: iDesc, valor: iValor, parcela: iParc };
  }
  return adivinharMapa(tab);
}
function colunaTexto(tab, de, fora) {
  const amostra = tab.slice(de, de + 40); const ncol = Math.max(0, ...amostra.map(r => r.length));
  let melhor = -1, nota = 0;
  for (let i = 0; i < ncol; i++) {
    if (fora.includes(i)) continue;
    const n = amostra.reduce((s, r) => { const c = String(r[i] ?? ''); return s + (/[a-z]/i.test(c) ? c.length : 0); }, 0);
    if (n > nota) { nota = n; melhor = i; }
  }
  return melhor;
}
function adivinharMapa(tab) {
  const am = tab.slice(0, 60); const ncol = Math.max(0, ...am.map(r => r.length));
  if (ncol < 3) return null;
  const pareceData = c => (typeof c === 'number' ? c > 30000 && c < 60000 && Number.isInteger(c) : /^\s*(\d{4}-\d{2}-\d{2}|\d{1,2}[\/.\-]\d{1,2}([\/.\-]\d{2,4})?|\d{1,2}\s+[a-z]{3})/i.test(String(c ?? ''))) && !!parseData(c);
  const pareceValor = c => typeof c === 'number' ? true : /^\s*\(?-?\s*(R\$\s*)?-?\s*[\d.]*\d[.,]\d{2}\)?\s*-?\s*[CD]?\s*$/i.test(String(c ?? ''));
  const conta = f => Array.from({ length: ncol }, (_, i) => am.filter(r => f(r[i])).length);
  const cd = conta(pareceData), cv = conta(pareceValor);
  const iData = cd.indexOf(Math.max(...cd)); if (cd[iData] < 2) return null;
  let iValor = -1, nv = 1;
  cv.forEach((n, i) => { if (i !== iData && n > nv) { nv = n; iValor = i; } });
  if (iValor < 0) return null;
  const iDesc = colunaTexto(tab, 0, [iData, iValor]);
  const primeira = am.findIndex(r => pareceData(r[iData]));
  return { cab: primeira - 1, data: iData, desc: iDesc, valor: iValor, parcela: -1 };
}
function tabelaParaTransacoes(tab, m) {
  if (!m) return [];
  const out = [];
  for (let r = m.cab + 1; r < tab.length; r++) {
    const row = tab[r] || [];
    const dr = row[m.data];
    if (dr === undefined || dr === null || dr === '') continue;
    out.push({ dataRaw: dr, descricao: m.desc >= 0 ? String(row[m.desc] ?? '') : '', valorRaw: row[m.valor], parcelaRaw: m.parcela >= 0 ? row[m.parcela] : '' });
  }
  return out;
}

/* --- leitores por formato --- */
async function pdfParaLinhas(buf) {
  await carregarScript(PDFJS);
  pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
  const tarefa = pdfjsLib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true, useSystemFonts: false, enableXfa: false }); // isEvalSupported:false fecha a falha CVE-2024-4367
  tarefa.onPassword = (atualizar, motivo) => {
    const s = prompt(motivo === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD ? 'Senha incorreta. Tente de novo:' : 'Este PDF tem senha (muitos bancos usam parte do CPF). Digite a senha do PDF:');
    if (s === null) tarefa.destroy(); else atualizar(s);
  };
  const pdf = await tarefa.promise;
  const linhas = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const pg = await pdf.getPage(p);
    const tc = await pg.getTextContent();
    const itens = tc.items.filter(it => it.str && it.str.trim()).map(it => ({ s: it.str, x: it.transform[4], y: it.transform[5], w: it.width || 0 }));
    itens.sort((a, b) => b.y - a.y || a.x - b.x);
    let atual = [], yAt = null;
    const fechar = () => {
      if (!atual.length) return;
      atual.sort((a, b) => a.x - b.x);
      let s = '', fim = null;
      for (const it of atual) { if (fim !== null) { const gap = it.x - fim; s += gap > 22 ? '    ' : (gap > 0.8 ? ' ' : ''); } s += it.s; fim = it.x + it.w; }
      linhas.push(s); atual = [];
    };
    for (const it of itens) { if (yAt === null || Math.abs(it.y - yAt) > 3) { fechar(); yAt = it.y; } atual.push(it); }
    fechar();
  }
  return linhas;
}
async function xlsxParaFonte(buf, nome) {
  await carregarScript(SHEETJS);
  const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
  let melhor = null, pts = -1;
  for (const n of wb.SheetNames) {
    const tab = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
    const qtd = tabelaParaTransacoes(tab, detectarMapa(tab)).length;
    if (qtd > pts) { pts = qtd; melhor = tab; }
  }
  melhor = melhor || [];
  return { nome, tabela: melhor, textoBruto: melhor.slice(0, 80).map(r => r.join(' ')).join('\n') };
}
const ent = s => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'");
function ofxParaFonte(t, nome) {
  const ofxTipo = /<CREDITCARDMSGSRSV1>|<CCSTMTRS>/i.test(t) ? 'cartao' : 'banco';
  const transacoes = t.split(/<STMTTRN>/i).slice(1).map(b => {
    const g = tag => { const m = b.match(new RegExp('<' + tag + '>\\s*([^<\\r\\n]*)', 'i')); return m ? m[1].trim() : ''; };
    const dt = g('DTPOSTED').replace(/\D/g, '').slice(0, 8);
    const memo = ent(g('MEMO')), nm = ent(g('NAME'));
    const desc = memo && nm && !norm(memo).includes(norm(nm)) && !norm(nm).includes(norm(memo)) ? `${nm} ${memo}` : (memo.length >= nm.length ? memo : nm);
    return { dataRaw: dt.length === 8 ? `${dt.slice(0, 4)}-${dt.slice(4, 6)}-${dt.slice(6, 8)}` : '', descricao: desc, valorRaw: g('TRNAMT') };
  }).filter(x => x.dataRaw);
  return { nome, transacoes, ofxTipo, textoBruto: t };
}
function htmlParaFonte(texto, nome) {
  const doc = new DOMParser().parseFromString(texto, 'text/html');
  const tabelas = [...doc.querySelectorAll('table')].map(t => [...t.rows].map(r => [...r.cells].map(c => c.textContent.replace(/\s+/g, ' ').trim())));
  let melhor = null, pts = 0;
  for (const tb of tabelas) { const n = tabelaParaTransacoes(tb, detectarMapa(tb)).length; if (n > pts) { pts = n; melhor = tb; } }
  doc.querySelectorAll('br,p,div,tr,li,h1,h2,h3,h4,h5,h6,section,article').forEach(el => el.append('\n'));
  doc.querySelectorAll('td,th,span').forEach(el => el.append(' '));
  const corpo = doc.body ? doc.body.textContent : texto;
  const linhas = corpo.split('\n');
  if (melhor && pts >= linhasParaTransacoes(linhas).length) return { nome, tabela: melhor, textoBruto: corpo };
  return { nome, linhas, textoBruto: corpo };
}
function textoParaFonte(texto, nome) {
  const linhas = texto.replace(/^﻿/, '').split(/\r?\n/);
  const tab = csvParaTabela(texto);
  const nT = tab.length && tab[0].length > 1 ? tabelaParaTransacoes(tab, detectarMapa(tab)).length : 0;
  const nL = linhasParaTransacoes(linhas).length;
  if (nT > 0 && nT >= nL) return { nome, tabela: tab, textoBruto: texto };
  return { nome, linhas, textoBruto: texto };
}
async function lerArquivo(file) {
  const nome = file.name || 'arquivo';
  const ext = (nome.split('.').pop() || '').toLowerCase();
  const buf = await file.arrayBuffer();
  if (ext === 'pdf') { const linhas = await pdfParaLinhas(buf); return { nome, linhas, textoBruto: linhas.join('\n') }; }
  if (['xlsx', 'xls', 'xlsm', 'ods'].includes(ext)) return xlsxParaFonte(buf, nome);
  const texto = decodificar(buf);
  if (['ofx', 'qfx'].includes(ext) || /<OFX>|OFXHEADER/i.test(texto.slice(0, 3000))) return ofxParaFonte(texto, nome);
  if (['html', 'htm'].includes(ext) || /<table[\s>]|<html[\s>]/i.test(texto.slice(0, 5000))) return htmlParaFonte(texto, nome);
  return textoParaFonte(texto, nome);
}

/* --- montagem da conferência --- */
function palpiteDestino(f) {
  const txt = f.textoBruto || '';
  for (const c of S.cartoes) if (c.final && new RegExp('(final|[*•x]{3,})\\s*:?\\s*' + c.final + '(?!\\d)', 'i').test(txt)) return 'cartao:' + c.id;
  if (f.ofxTipo === 'banco') return 'conta';
  if (f.ofxTipo !== 'cartao' && /extrato|saldo anterior|saldo do dia/i.test(txt) && !/fatura/i.test(txt)) return 'conta';
  return 'cartao:' + (S.cartoes[0]?.id || '');
}
function iniciarImport(fonte) {
  IMP = { fonte, nome: fonte.nome, destino: palpiteDestino(fonte), mesFatura: null, pessoaId: EU(), inverter: false, mapa: fonte.tabela ? detectarMapa(fonte.tabela) : null, ajustes: {} };
  if (fonte.tabela && !IMP.mapa) IMP.mapa = { cab: 0, data: 0, desc: 1, valor: 2, parcela: -1 };
  construirLinhas(true);
}
function brutasImp() {
  if (IMP.fonte.transacoes) return IMP.fonte.transacoes;
  if (IMP.fonte.tabela) return tabelaParaTransacoes(IMP.fonte.tabela, IMP.mapa);
  return linhasParaTransacoes(IMP.fonte.linhas || []);
}
const chaveLanc = l => [norm(l.descricao), round2(l.valor).toFixed(2), l.data, l.meio === 'cartao' ? 'c:' + (l.cartaoId || '') : l.meio, l.parcelas || 1].join('|');
function procurarParecido(x) {
  const t = Date.parse(x.data);
  for (const e of S.lancamentos) {
    if (Math.abs(e.valor - x.valor) > 0.009 || (e.parcelas || 1) !== (x.parcelas || 1)) continue;
    const mesmoMeio = x.meio === 'cartao' ? e.meio === 'cartao' && (!x.cartaoId || !e.cartaoId || e.cartaoId === x.cartaoId)
      : ehVale(x.meio) ? e.meio === x.meio : !['cartao', 'vr', 'va'].includes(e.meio);
    if (!mesmoMeio) continue;
    if (x.parcelas > 1 && norm(e.descricao) === norm(x.descricao) && ocorrencia(e, x.mesRef) === x.parcelaAtual) return { tipo: 'prevista', e };
    if (Math.abs(Date.parse(e.data) - t) <= 3 * 864e5) return { tipo: 'possivel', e };
  }
  return null;
}
// A linha da fatura/extrato é uma conta fixa já cadastrada? (evita contar duas vezes)
function procurarFixa(x) {
  const nd = ' ' + norm(x.descricao) + ' ', t = Date.parse(x.data);
  for (const f of S.fixas) {
    const mesmoMeio = x.meio === 'cartao' ? f.meio === 'cartao' && (!x.cartaoId || !f.cartaoId || f.cartaoId === x.cartaoId)
      : ehVale(x.meio) ? f.meio === x.meio : !['cartao', 'vr', 'va'].includes(f.meio);
    if (!mesmoMeio || !(ativo(f, x.data.slice(0, 7)) || ativo(f, x.mesRef))) continue;
    const dif = Math.abs(f.valor - x.valor);
    const palavra = termoDe(f.descricao).split(' ')[0] || '';
    const nomeBate = palavra.length >= 3 && nd.includes(' ' + palavra);
    const perto = Math.abs(Date.parse(dataNoMes(x.data.slice(0, 7), f.dia)) - t) <= 7 * 864e5;
    if ((nomeBate && dif <= Math.max(5, f.valor * 0.15)) || (dif < 0.01 && perto)) return f;
  }
  return null;
}
function meioDeExtrato(desc) {
  const n = norm(desc);
  if (/\bpix\b/.test(n)) return 'pix';
  if (/\bsaque\b/.test(n)) return 'dinheiro';
  if (/boleto|pagamento|pagto|pgto|debito automatico|conta de/.test(n)) return 'boleto';
  return 'debito';
}
function construirLinhas(primeira) {
  const [tipo, cid = ''] = IMP.destino.split(':');
  const cartao = tipo === 'cartao' ? cartaoPorId(cid) : null;
  const brutos = brutasImp();
  const ler = mesRef => brutos.map((b, i) => {
    const pc = extrairParcela(String(b.descricao ?? '').replace(/\s+/g, ' ').trim(), b.parcelaRaw);
    return { i, data: parseData(b.dataRaw, mesRef), descricao: limparDesc(pc.desc), valor: parseValor(b.valorRaw), atual: pc.atual, total: pc.total };
  }).filter(l => l.data && isFinite(l.valor) && l.valor !== 0 && l.descricao);
  let L = ler(IMP.mesFatura || mesAtual());
  if (primeira) {
    const nNeg = L.filter(l => l.valor < 0).length;
    if (nNeg > L.length / 2) IMP.inverter = true;
  }
  if (primeira && cartao) IMP.pessoaId = cartao.pessoaId;
  if (tipo === 'cartao' && !IMP.mesFatura) {
    const datas = L.filter(l => l.total === 1).map(l => l.data).sort();
    const ult = datas[datas.length - 1] || L.map(l => l.data).sort().pop() || hojeISO();
    IMP.mesFatura = cartao ? mesFaturaDe(ult, cartao) : somaMes(ult.slice(0, 7), 1);
    L = ler(IMP.mesFatura);
  }
  const inv = IMP.inverter ? -1 : 1;
  const chaves = new Map(S.lancamentos.map(e => [chaveLanc(e), e]));
  IMP.linhas = L.map(l => {
    const valor = round2(l.valor * inv);
    const meio = tipo === 'cartao' ? 'cartao' : tipo === 'conta' ? meioDeExtrato(l.descricao) : tipo;
    const x = { i: l.i, data: l.data, descricao: l.descricao, valor, meio, cartaoId: tipo === 'cartao' ? cid : '', mesRef: tipo === 'cartao' ? IMP.mesFatura : l.data.slice(0, 7), parcelaAtual: l.atual, parcelas: l.total };
    let status = 'novo', dupDe = '';
    const nd = norm(l.descricao);
    if (tipo === 'conta' && valor > 0 && /(pagamento|pagto|pgto|pag)\b.*\bfatura|fatura\b.*cart|cartao de credito|pagto cartao/.test(nd)) status = 'fatura';
    else if (valor < 0) status = tipo === 'cartao' ? 'credito' : 'entrada';
    else if (/^(pagamento|pgto|pagto)\b.*(receb|efetuad)|^saldo\b|^total\b/.test(nd)) status = 'credito';
    else {
      const e = chaves.get(chaveLanc(x));
      if (e) { status = (e.parcelaAtual || 1) !== x.parcelaAtual ? 'prevista' : 'dup'; dupDe = `Já existe: ${e.descricao} · ${dataLonga(e.data)} · ${brl(e.valor)}`; }
      else {
        const f = x.parcelas === 1 ? procurarFixa(x) : null;
        if (f) { status = 'fixa'; dupDe = `Conta fixa cadastrada: ${f.descricao} · dia ${f.dia} · ${brl(f.valor)}`; }
        else { const p = procurarParecido(x); if (p) { status = p.tipo; dupDe = `Já existe: ${p.e.descricao} · ${dataLonga(p.e.data)} · ${brl(p.e.valor)}`; } }
      }
    }
    const aj = IMP.ajustes[l.i] || {};
    return { ...x, status, dupDe, categoria: aj.categoria || categoriaAuto(l.descricao, meio), sel: aj.sel ?? (status === 'novo') };
  });
}
function vImportar() {
  return `<section class="card">
    <h3>Importar fatura ou extrato</h3>
    <p class="sub">CSV, Excel (.xlsx/.xls), OFX, PDF ou HTML. O arquivo é lido aqui mesmo, no navegador — o conteúdo não sai do seu computador.</p>
    <label class="drop" id="drop" style="margin-top:12px">
      <input type="file" id="arqImport" accept=".csv,.txt,.xlsx,.xls,.xlsm,.ods,.ofx,.qfx,.pdf,.html,.htm" aria-label="Escolher arquivo da fatura ou extrato">
      ${CARREGANDO ? `<b>Lendo ${esc(CARREGANDO)}…</b><span class="sub">PDF e Excel podem levar alguns segundos</span>` : `<b>Escolher arquivo</b><span class="sub">ou arraste para cá</span>`}
    </label>
    <details style="margin-top:12px"><summary>Ou cole o texto da fatura</summary>
      <textarea id="txtImport" rows="6" style="margin-top:8px" placeholder="05/11 PADARIA PAO QUENTE 45,90&#10;06/11 UBER *TRIP 18,70&#10;07/11 MAGAZINE LUIZA 03/10 150,00"></textarea>
      <button class="btn" data-acao="import-texto" style="margin-top:8px">Ler texto colado</button>
    </details>
  </section>
  ${IMP ? previewHTML() : ''}
  <section class="card">
    <h3>Colar notificação do banco</h3>
    <p class="sub">Copie a notificação da compra (ex.: “Compra de R$ 45,90 aprovada em PADARIA… cartão final 1234”) e cole aqui. O app separa valor, loja, cartão e parcelas para você conferir.</p>
    <textarea id="txtNotif" rows="3" style="margin-top:10px" placeholder="Compra de R$ 45,90 APROVADA em PADARIA PAO QUENTE para o cartão com final 1234."></textarea>
    <button class="btn primario" data-acao="ler-notif" style="margin-top:8px">Interpretar e lançar</button>
  </section>`;
}
const ROT_STATUS = { dup: 'já lançado', prevista: 'parcela já prevista', possivel: 'parece já lançado', fixa: 'conta fixa já prevista', credito: 'pagamento/crédito', fatura: 'pagamento de fatura', entrada: 'entrada de dinheiro' };
function previewHTML() {
  const L = IMP.linhas, tipo = IMP.destino.split(':')[0];
  const destinos = [...S.cartoes.map(c => ['cartao:' + c.id, `Fatura ${c.nome} · ${pessoaNome(c.pessoaId)}`]), ...(S.cartoes.length ? [] : [['cartao:', 'Fatura de cartão de crédito']]), ['conta', 'Extrato da conta (Pix, débito)'], ['vr', 'Extrato do VR'], ['va', 'Extrato do VA']];
  const baseM = IMP.mesFatura || mesAtual(); const mesesOp = []; for (let i = -4; i <= 3; i++) mesesOp.push(somaMes(baseM, i));
  let mapaH = '';
  if (IMP.fonte.tabela && IMP.mapa) {
    const tab = IMP.fonte.tabela, cab = IMP.mapa.cab;
    const ncol = Math.max(1, ...tab.slice(0, 50).map(r => r.length));
    const am = tab[cab + 1] || [];
    const nomeCol = i => { const n = cab >= 0 ? String((tab[cab] || [])[i] ?? '').trim() : ''; const a = String(am[i] ?? '').trim().slice(0, 18); return `${String.fromCharCode(65 + (i % 26))}${n ? ' · ' + n : ''}${a && a !== n ? ` (ex.: ${a})` : ''}`; };
    const sel = (campo, rot, opcional) => `<label class="campo"><span class="lbl">${rot}</span><select data-imp="mapa-${campo}">${opcional ? `<option value="-1" ${IMP.mapa[campo] === -1 ? 'selected' : ''}>— nenhuma —</option>` : ''}${Array.from({ length: ncol }, (_, i) => `<option value="${i}" ${IMP.mapa[campo] === i ? 'selected' : ''}>${esc(nomeCol(i))}</option>`).join('')}</select></label>`;
    mapaH = `<details ${L.length ? '' : 'open'}><summary>Colunas da planilha</summary><div class="cfg">${sel('data', 'Data')}${sel('desc', 'Descrição')}${sel('valor', 'Valor')}${sel('parcela', 'Parcela', true)}</div></details>`;
  }
  const todos = L.length && L.every(l => l.sel);
  return `<section class="card" id="preview">
    <h3>Conferir: ${esc(IMP.nome)}<button class="btn peq" data-acao="import-cancelar">Cancelar</button></h3>
    <div class="cfg">
      <label class="campo"><span class="lbl">O que é este arquivo?</span><select data-imp="destino">${destinos.map(([v, t]) => `<option value="${esc(v)}" ${v === IMP.destino ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select></label>
      ${tipo === 'cartao' ? `<label class="campo"><span class="lbl">Fatura com vencimento em</span><select data-imp="mesFatura">${mesesOp.map(m => `<option value="${m}" ${m === IMP.mesFatura ? 'selected' : ''}>${nomeMes(m)}</option>`).join('')}</select></label>` : ''}
      <label class="campo"><span class="lbl">De quem são as compras?</span><select data-imp="pessoa">${S.pessoas.map(p => `<option value="${p.id}" ${p.id === IMP.pessoaId ? 'selected' : ''}>${esc(p.nome)}</option>`).join('')}</select></label>
      <label class="campo chk"><input type="checkbox" data-imp="inverter" ${IMP.inverter ? 'checked' : ''}> <span>Inverter sinais (se os gastos vierem negativos)</span></label>
    </div>
    ${mapaH}
    <p class="sub" id="impResumo" style="margin:4px 0 10px"></p>
    ${L.length ? `<div class="tab-wrap"><table class="prev"><thead><tr><th><input type="checkbox" data-imp="todos" ${todos ? 'checked' : ''} aria-label="Marcar todos"></th><th>Data</th><th>Descrição</th><th>Parcela</th><th>Categoria</th><th class="num">Valor</th></tr></thead><tbody>
      ${L.map((l, k) => `<tr class="${l.sel ? '' : 'off'}" data-k="${k}"><td><input type="checkbox" data-imp-sel="${k}" ${l.sel ? 'checked' : ''} aria-label="Importar esta linha"></td><td style="white-space:nowrap">${dataLonga(l.data)}</td><td>${esc(l.descricao)}${l.status !== 'novo' ? ` <span class="tag" ${l.dupDe ? `title="${esc(l.dupDe)}"` : ''}>${ROT_STATUS[l.status]}</span>` : ''}</td><td>${l.parcelas > 1 ? `${l.parcelaAtual}/${l.parcelas}` : ''}</td><td><select data-imp-cat="${k}" aria-label="Categoria">${opcoesCat(l.categoria)}</select></td><td class="num ${l.valor < 0 ? 'cred' : ''}">${brl(l.valor)}</td></tr>`).join('')}
    </tbody></table></div>` : `<div class="vazio">Não encontrei lançamentos. ${IMP.fonte.tabela ? 'Confira as colunas acima.' : 'Tente exportar em CSV, Excel ou OFX — ou cole o texto da fatura.'}</div>`}
    <div class="acoes-form" style="margin-top:12px"><span class="sp"></span><button class="btn primario" data-acao="import-confirmar" id="btnImportar">Importar</button></div>
  </section>`;
}
function atualizarResumoImp() {
  if (!IMP || !$('#impResumo')) return;
  const L = IMP.linhas, sel = L.filter(l => l.sel);
  const tot = sel.reduce((s, l) => s + l.valor, 0);
  const fora = L.length - sel.length, novasParc = sel.filter(l => l.parcelas > l.parcelaAtual).length;
  $('#impResumo').innerHTML = `${L.length} linha${L.length === 1 ? '' : 's'} lida${L.length === 1 ? '' : 's'} · <b>${sel.length}</b> para importar (${brl(tot)})${fora ? ` · ${fora} de fora (já lançadas, pagamentos ou desmarcadas)` : ''}${novasParc ? ` · ${novasParc} parcelada${novasParc > 1 ? 's' : ''}: as próximas parcelas entram sozinhas nos meses seguintes` : ''}`;
  const b = $('#btnImportar');
  if (b) { b.textContent = sel.length ? `Importar ${sel.length} lançamento${sel.length > 1 ? 's' : ''}` : 'Nada marcado'; b.disabled = !sel.length; }
  $$('#preview tr[data-k]').forEach(tr => tr.classList.toggle('off', !L[+tr.dataset.k].sel));
  const t = $('[data-imp="todos"]'); if (t) t.checked = L.length > 0 && L.every(l => l.sel);
}
function rerenderPreview() { const p = $('#preview'); if (p) { p.outerHTML = previewHTML(); atualizarResumoImp(); } }
function mudarImp(t) {
  const k = t.dataset.imp;
  if (k === 'todos') { IMP.linhas.forEach(l => { l.sel = t.checked; (IMP.ajustes[l.i] ||= {}).sel = t.checked; }); $$('[data-imp-sel]').forEach(c => { c.checked = t.checked; }); atualizarResumoImp(); return; }
  if (k === 'pessoa') { IMP.pessoaId = t.value; return; }
  if (k === 'destino') { IMP.destino = t.value; const c = cartaoPorId(t.value.split(':')[1]); if (c) IMP.pessoaId = c.pessoaId; IMP.mesFatura = null; IMP.ajustes = {}; }
  else if (k === 'mesFatura') IMP.mesFatura = t.value;
  else if (k === 'inverter') { IMP.inverter = t.checked; IMP.ajustes = {}; }
  else if (k.startsWith('mapa-')) { IMP.mapa[k.slice(5)] = +t.value; IMP.ajustes = {}; }
  construirLinhas(false);
  rerenderPreview();
}
async function processarArquivo(file) {
  if (!file) return;
  CARREGANDO = file.name; IMP = null; render();
  try {
    const fonte = await lerArquivo(file);
    iniciarImport(fonte);
    if (!IMP.linhas.length && !fonte.tabela) { IMP = null; toast('Não encontrei lançamentos nesse arquivo. Tente CSV, Excel ou OFX — ou cole o texto da fatura.', 6000); }
  } catch (err) {
    console.error(err); IMP = null;
    toast(err && err.name === 'PasswordException' ? 'PDF com senha: não foi aberto.' : 'Não consegui ler o arquivo: ' + (err && err.message || err), 6000);
  }
  CARREGANDO = false; render();
  $('#preview')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}
function confirmarImport() {
  const [tipo, cid = ''] = IMP.destino.split(':');
  const sel = IMP.linhas.filter(l => l.sel); if (!sel.length) return;
  for (const l of sel) {
    S.lancamentos.push({ id: uid(), data: l.data, descricao: l.descricao, valor: l.valor, categoria: l.categoria, pessoaId: IMP.pessoaId, meio: l.meio, cartaoId: l.cartaoId, parcelaAtual: l.parcelaAtual, parcelas: l.parcelas, mesRef: l.mesRef, mesRefAuto: false, origem: 'import', criadoEm: hojeISO(), obs: '' });
    const aj = IMP.ajustes[l.i];
    if (aj && aj.categoria && aj.categoria !== categoriaAuto(l.descricao, l.meio)) aprenderRegra(l.descricao, aj.categoria);
  }
  salvar();
  const n = sel.length, mf = IMP.mesFatura;
  IMP = null;
  if (tipo === 'cartao') { MES = mf; FILTRO = { q: '', pessoa: '', cat: '', meio: 'cartao:' + cid, visao: 'pagamento' }; }
  else FILTRO = { q: '', pessoa: '', cat: '', meio: '', visao: 'compra' };
  ABA = 'lancamentos'; render(); window.scrollTo(0, 0);
  toast(`${n} lançamento${n > 1 ? 's' : ''} importado${n > 1 ? 's' : ''}`);
}

/* --- notificação colada --- */
function interpretarNotificacao(txt) {
  const t = String(txt || '').replace(/\s+/g, ' ').trim();
  if (!t) return null;
  const n = semAcento(t.toLowerCase());
  const mv = t.match(/R\$\s*(-?\s*\d{1,3}(?:\.\d{3})*,\d{2}|-?\s*\d+,\d{2}|\d+\.\d{2})/i);
  if (!mv) return null;
  const total = Math.abs(parseValor(mv[1]));
  let parcelas = 1; const mp = n.match(/\b(?:em\s+)?(\d{1,2})\s*x\b/); if (mp && +mp[1] >= 2 && +mp[1] <= 48) parcelas = +mp[1];
  const mf = n.match(/final\s*(?:do cartao\s*)?:?\s*(\d{4})/);
  let cartao = mf ? S.cartoes.find(c => c.final === mf[1]) : null;
  let meio;
  if (/\bpix\b/.test(n)) meio = 'pix';
  else if (/refeicao|\bvr\b|vale refei|ticket restaurante/.test(n)) meio = 'vr';
  else if (/alimentacao|\bva\b|vale alim/.test(n)) meio = 'va';
  else if (/debito/.test(n) && !/credito/.test(n)) meio = 'debito';
  else if (/boleto|pagamento de conta/.test(n)) meio = 'boleto';
  else meio = 'cartao';
  if (meio === 'cartao' && !cartao && S.cartoes.length === 1) cartao = S.cartoes[0];
  const semValor = t.replace(mv[0], ' ').replace(/\s+/g, ' ');
  const FIM = '(?=\\s+(?:para|pra|no|na|com|foi|via|as|às|dia|hoje|valor|cartao|cartão)\\b|[,;!]|\\.\\s|\\.$|$)';
  let est = '', m;
  if (meio === 'pix' && (m = semValor.match(/\b(?:para|pra)\s+(.+?)(?=\s+(?:foi|com|em|no|na|via)\b|[,;!]|\.\s|\.$|$)/i))) est = 'Pix para ' + m[1];
  else if ((m = semValor.match(/\b(?:local|estabelecimento|loja)\s*:\s*(.+?)(?=[,;!]|\.\s|\.$|$)/i))) est = m[1];
  else if ((m = semValor.match(new RegExp('\\bem\\s+(?!\\d{1,2}\\s*x\\b)(?!\\d{1,2}\\/)(.+?)' + FIM, 'i')))) est = m[1];
  else { const caps = semValor.match(/\b[A-Z0-9][A-Z0-9*&.'\-]*(?:\s+[A-Z0-9*&.'\-]+)+\b/g) || semValor.match(/\b[A-Z][A-Z0-9*&.'\-]{3,}\b/g); if (caps) est = caps.sort((a, b) => b.length - a.length)[0]; }
  est = est.replace(/\s+/g, ' ').trim().replace(/^(o|a|os|as)\s+/i, '').replace(/[.]+$/, '');
  let data = hojeISO();
  const md = t.match(/\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/); if (md) { const d = parseData(md[0]); if (d && d <= hojeISO()) data = d; }
  const donoVale = ehVale(meio) ? S.receitas.find(r => r.tipo === meio)?.pessoaId : null;
  return { descricao: est || 'Compra (notificação)', valor: round2(total / parcelas), parcelas: meio === 'cartao' ? parcelas : 1, meio, cartaoId: meio === 'cartao' ? (cartao?.id || S.cartoes[0]?.id || '') : '', pessoaId: (meio === 'cartao' && cartao?.pessoaId) || donoVale || EU(), data, categoria: categoriaAuto(est, meio), origem: 'notificacao' };
}

/* =====================================================================
   BACKUP E EXEMPLO
   ===================================================================== */
function baixar(nome, conteudo, tipo) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([conteudo], { type: tipo })); a.download = nome;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 800);
}
function exportarCSV() {
  // texto começando com = + - @ vira fórmula no Excel ("CSV injection"): um apóstrofo na frente desarma
  const q = s => { let t = String(s ?? ''); if (/^[=+\-@\t\r]/.test(t)) t = "'" + t; return `"${t.replace(/"/g, '""')}"`; };
  const linhas = [['data', 'descricao', 'valor_parcela', 'parcelas', 'parcela_inicial', 'mes_da_fatura', 'categoria', 'pessoa', 'meio', 'cartao', 'origem'].join(';')];
  for (const l of S.lancamentos.slice().sort((a, b) => a.data.localeCompare(b.data))) linhas.push([dataLonga(l.data), q(l.descricao), num2(l.valor), l.parcelas || 1, l.parcelaAtual || 1, l.mesRef, q(l.categoria), q(pessoaNome(l.pessoaId)), ROT_MEIO[l.meio] || l.meio, l.meio === 'cartao' ? q(nomeCartao(l.cartaoId)) : '', l.origem || ''].join(';'));
  baixar(`financas-casa-lancamentos-${hojeISO()}.csv`, '﻿' + linhas.join('\r\n'), 'text/csv;charset=utf-8');
}
async function restaurarBackup(file) {
  if (!file) return;
  try {
    const o = JSON.parse(await file.text());
    if (!o || !Array.isArray(o.lancamentos) || !Array.isArray(o.receitas)) throw new Error('formato');
    if (!confirm(`Substituir os dados atuais pelo backup “${file.name}” (${o.lancamentos.length} lançamentos)?`)) return;
    if (CASA && !confirm('Atenção: o backup também substitui os dados da casa na nuvem, para vocês dois. Continuar?')) return;
    const tema = S.tema; S = migrar(o); S.tema = tema; S.exemplo = false; salvar(); IMP = null; render(); toast('Backup restaurado');
  } catch (e) { toast('Esse arquivo não parece um backup deste app.', 5000); }
}
function carregarExemplo() {
  const s = estadoVazio(); s.exemplo = true;
  const H = 'p1', C = 'p2', M0 = mesAtual(), hojeDia = +hojeISO().slice(8);
  const dt = (off, dia) => dataNoMes(somaMes(M0, off), off === 0 ? Math.min(dia, hojeDia) : dia);
  const ini = somaMes(M0, -4);
  s.cartoes = [
    { id: 'c1', nome: 'Nubank', pessoaId: H, final: '1234', fechamento: 3, vencimento: 10, limite: 8000 },
    { id: 'c2', nome: 'Itaú', pessoaId: C, final: '5678', fechamento: 26, vencimento: 5, limite: 6000 },
  ];
  s.receitas = [
    { id: uid(), pessoaId: H, tipo: 'salario', descricao: 'Salário', valor: 6800, dia: 5, inicio: ini, fim: '' },
    { id: uid(), pessoaId: C, tipo: 'salario', descricao: 'Salário', valor: 4300, dia: 5, inicio: ini, fim: '' },
    { id: uid(), pessoaId: C, tipo: 'vr', descricao: 'VR', valor: 880, dia: 1, inicio: ini, fim: '' },
    { id: uid(), pessoaId: C, tipo: 'va', descricao: 'VA', valor: 650, dia: 1, inicio: ini, fim: '' },
    { id: uid(), pessoaId: H, tipo: 'extra', descricao: 'Freela', valor: 1200, dia: 20, inicio: M0, fim: M0 },
  ];
  const fx = (descricao, valor, dia, categoria, pessoaId, meio, cartaoId = '') => ({ id: uid(), descricao, valor, dia, categoria, pessoaId, meio, cartaoId, inicio: ini, fim: '' });
  s.fixas = [
    fx('Aluguel', 2300, 10, 'Moradia', 'casal', 'conta'),
    fx('Condomínio', 520, 10, 'Moradia', 'casal', 'conta'),
    fx('Energia', 210, 15, 'Contas da casa', 'casal', 'conta'),
    fx('Plano de saúde', 690, 8, 'Saúde', 'casal', 'conta'),
    fx('Reserva de emergência', 500, 6, 'Reserva e investimentos', 'casal', 'conta'),
    fx('Academia', 129.9, 5, 'Saúde', C, 'conta'),
    fx('Internet', 120, 20, 'Contas da casa', 'casal', 'cartao', 'c1'),
    fx('Celular 1', 65, 12, 'Contas da casa', H, 'cartao', 'c1'),
    fx('Spotify', 34.9, 22, 'Assinaturas', H, 'cartao', 'c1'),
    fx('Celular 2', 55, 12, 'Contas da casa', C, 'cartao', 'c2'),
    fx('Netflix', 44.9, 18, 'Assinaturas', 'casal', 'cartao', 'c2'),
  ];
  const L = [
    [-3, 14, 'GELADEIRA BRASTEMP FROST FREE', 4200, 'cartao', 'c1', H, 10],
    [-1, 2, 'NOTEBOOK DELL INSPIRON', 3000, 'cartao', 'c2', C, 6],
    [-1, 20, 'LATAM AIRLINES', 1680, 'cartao', 'c1', H, 4],
    [-1, 4, 'IFOOD *RESTAURANTE SABOR', 68.9, 'cartao', 'c1', H],
    [-1, 6, 'UBER *TRIP', 23.4, 'cartao', 'c1', H],
    [-1, 8, 'SUPERMERCADO BOM PRECO', 312.55, 'cartao', 'c2', C],
    [-1, 11, 'DROGASIL', 54.2, 'cartao', 'c2', C],
    [-1, 12, 'POSTO IPIRANGA', 220, 'cartao', 'c1', H],
    [-1, 15, 'RESTAURANTE DONA MARIA', 42, 'vr', '', C],
    [-1, 18, 'CINEMARK', 96, 'cartao', 'c1', H],
    [-1, 21, 'MERCADO LIVRE', 189.9, 'cartao', 'c2', C],
    [-1, 22, 'ASSAI ATACADISTA', 428.3, 'va', '', C],
    [-1, 24, 'PIX FEIRA DO BAIRRO', 65, 'pix', '', H],
    [-1, 27, 'RENNER', 259.8, 'cartao', 'c2', C],
    [-1, 29, 'IFOOD *PIZZARIA', 89.9, 'cartao', 'c1', H],
    [-1, 30, 'PETZ', 142, 'cartao', 'c2', C],
    [0, 1, 'PADARIA PAO QUENTE', 28.5, 'cartao', 'c1', H],
    [0, 2, 'UBER *TRIP', 18.7, 'cartao', 'c1', H],
    [0, 3, 'SUPERMERCADO BOM PRECO', 276.4, 'cartao', 'c2', C],
    [0, 3, 'RESTAURANTE SABOR CASEIRO', 38.5, 'vr', '', C],
    [0, 4, 'SHELL SELECT', 180, 'cartao', 'c1', H],
    [0, 5, 'DROGARIA SAO PAULO', 37.8, 'debito', '', H],
    [0, 5, 'AMAZON MARKETPLACE', 349, 'cartao', 'c1', H, 3],
    [0, 6, 'IFOOD *HAMBURGUERIA', 72.3, 'cartao', 'c2', C],
    [0, 6, 'CARREFOUR', 198.65, 'va', '', C],
    [0, 7, 'BARBEARIA DO ZE', 60, 'pix', '', H],
  ];
  S = s; REGRAS_CACHE = null; // categorizar e calcMesRef leem S
  s.lancamentos = L.map(([off, dia, descricao, total, meio, cartaoId, pessoaId, parcelas = 1]) => {
    const l = { id: uid(), data: dt(off, dia), descricao, valor: round2(total / parcelas), parcelas, parcelaAtual: 1, meio, cartaoId, pessoaId, categoria: categoriaAuto(descricao, meio), origem: 'exemplo', mesRefAuto: true, obs: '', criadoEm: hojeISO() };
    l.mesRef = calcMesRef(l);
    return l;
  });
  salvar(); MES = M0; IMP = null; ABA = 'painel'; render(); window.scrollTo(0, 0);
  toast('Dados de exemplo carregados');
}

/* =====================================================================
   EVENTOS
   ===================================================================== */
/* =====================================================================
   NUVEM — login, casa compartilhada e sincronização (Supabase)
   Sem config.js preenchido, o app funciona só neste aparelho, como na fase 1.
   Modelo: cada item (lançamento, renda, fixa, cartão, regra, config) é uma
   linha em "itens". O app compara o estado local com a última versão que
   sabe estar na nuvem (ULT) e envia só a diferença; recebe a do outro
   celular em tempo real.
   ===================================================================== */
const CFG = window.FINANCAS_CONFIG || {};
const NUVEM_ATIVA = !!(CFG.supabaseUrl && CFG.supabaseChave);
const SUPABASE_JS = 'libs/supabase.js';          // supabase-js 2.45.4, hospedado aqui (conferido em 2 fontes)
const CHAVE_NUVEM = 'financas-casa-nuvem-v1';
const ROT_NUVEM = { ok: ['Sincronizado', 'ok', '✓'], enviando: ['Sincronizando…', 'and', '↻'], conectando: ['Conectando…', 'and', '↻'], offline: ['Sem internet — guardado no aparelho', 'off', '!'], erro: ['Erro ao sincronizar', 'erro', '!'] };
const LISTA_DO_TIPO = { lancamento: 'lancamentos', receita: 'receitas', fixa: 'fixas', cartao: 'cartoes' };
const ultVazio = () => ({ casaId: null, casa: null, email: '', cursor: '', itens: {} });
function carregarUlt() { try { return Object.assign(ultVazio(), JSON.parse(localStorage.getItem(CHAVE_NUVEM) || '{}')); } catch (e) { return ultVazio(); } }
function salvarUlt() { try { localStorage.setItem(CHAVE_NUVEM, JSON.stringify(ULT)); } catch (e) {} }
let ULT = carregarUlt();
let SB = null, SESSAO = null, CANAL = null, RECUPERANDO = false, AUTH_MODO = 'entrar';
let CASA = NUVEM_ATIVA ? ULT.casa : null;            // começa com a casa que este aparelho já conhecia (abre sem esperar a internet)
let NUVEM = { estado: NUVEM_ATIVA ? 'conectando' : 'off', msg: '' };
let envioTimer = null, enviando = false, enviarDeNovo = false;

const fmtCodigo = c => String(c || '').replace(/^(.{4})(.+)$/, '$1-$2');
const temDadosLocais = () => !S.exemplo && !!(S.lancamentos.length || S.receitas.length || S.fixas.length || S.cartoes.length);
const urlBase = () => location.origin + location.pathname;
function canon(v) {
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  if (v && typeof v === 'object') return '{' + Object.keys(v).sort().filter(k => v[k] !== undefined).map(k => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
  return JSON.stringify(v === undefined ? null : v);
}
function entidades(s) {
  const m = new Map();
  const por = (tipo, id, dados) => m.set(tipo + ':' + id, { tipo, id: String(id), dados });
  for (const [tipo, lista] of Object.entries(LISTA_DO_TIPO)) for (const x of s[lista]) { const d = sanear(tipo, x); if (d) por(tipo, d.id, d); }
  for (const r of s.regras) { const d = sanear('regra', r); if (d) por('regra', d.termo, d); }
  por('config', 'casa', sanear('config', { pessoas: s.pessoas, categorias: s.categorias }));
  return m;
}
// Linha vinda da nuvem → formato conferido (ou null, se for inválida: aí é ignorada)
function normalizarLinha(row) {
  if (!row || typeof row !== 'object') return null;
  const tipo = row.tipo, id = String(row.id ?? '');
  if (!['lancamento', 'receita', 'fixa', 'cartao', 'regra', 'config'].includes(tipo)) return null;
  if (tipo === 'config' && id !== 'casa') return null;
  if (row.excluido) return { tipo, id, excluido: true, dados: null };
  const base = row.dados && typeof row.dados === 'object' && !Array.isArray(row.dados) ? row.dados : {};
  const dados = sanear(tipo, LISTA_DO_TIPO[tipo] ? { ...base, id } : tipo === 'regra' ? { ...base, termo: id } : base);
  if (!dados || (LISTA_DO_TIPO[tipo] && dados.id !== id) || (tipo === 'regra' && dados.termo !== id)) return null;
  return { tipo, id, excluido: false, dados };
}
function aplicarLinha(row) {
  const lista = LISTA_DO_TIPO[row.tipo];
  if (lista) {
    const i = S[lista].findIndex(x => String(x.id) === row.id);
    if (row.excluido) { if (i >= 0) S[lista].splice(i, 1); }
    else if (i >= 0) S[lista][i] = row.dados; else S[lista].push(row.dados);
  } else if (row.tipo === 'regra') {
    S.regras = S.regras.filter(r => norm(r.termo) !== row.id);
    if (!row.excluido) { S.regras.push(row.dados); S.regras.sort((a, b) => (b.ordem || 0) - (a.ordem || 0)); }
  } else if (row.tipo === 'config' && !row.excluido) {
    S.pessoas = row.dados.pessoas; S.categorias = row.dados.categorias;
  }
}
// Junta linhas vindas da nuvem. Mudança local ainda não enviada vence (sobe no próximo envio).
function receberLinhas(rows) {
  if (!rows || !rows.length) return false;
  const locais = entidades(S);
  let mudou = false;
  for (const bruta of rows) {
    const row = normalizarLinha(bruta);
    if (!row) continue;
    const k = row.tipo + ':' + row.id;
    const serv = row.excluido ? null : canon(row.dados);
    const loc = locais.has(k) ? canon(locais.get(k).dados) : null;
    const ult = ULT.itens[k] ?? null;
    if (loc !== ult && loc !== serv) continue;
    if (serv === null) delete ULT.itens[k]; else ULT.itens[k] = serv;
    if (loc !== serv) { aplicarLinha(row); mudou = true; }
  }
  salvarUlt();
  if (mudou) salvarLocal();
  return mudou;
}
function definirNuvem(estado, msg = '') { NUVEM = { estado, msg }; pintarNuvem(); }
function pintarNuvem() {
  const el = $('#nuvem'); if (!el) return;
  const r = ROT_NUVEM[NUVEM.estado];
  el.hidden = !NUVEM_ATIVA || !r || (!CASA && NUVEM.estado !== 'conectando');
  if (!r) return;
  el.className = 'nuvem ' + r[1];
  el.querySelector('i').textContent = r[2];
  el.querySelector('span').textContent = r[0];
  el.title = r[0] + (NUVEM.msg ? ' — ' + NUVEM.msg : '');
  el.setAttribute('aria-label', el.title);
}
function agendarEnvio(ms = 700) {
  if (!NUVEM_ATIVA || !CASA || !SB) return;
  clearTimeout(envioTimer); envioTimer = setTimeout(enviar, ms);
}
async function enviar() {
  if (!CASA || !SB) return;
  if (enviando) { enviarDeNovo = true; return; }
  const locais = entidades(S), linhas = [];
  for (const [k, e] of locais) { const j = canon(e.dados); if (ULT.itens[k] !== j) linhas.push({ k, j, row: { casa_id: CASA.id, tipo: e.tipo, id: e.id, dados: e.dados, excluido: false } }); }
  for (const k of Object.keys(ULT.itens)) if (!locais.has(k)) { const i = k.indexOf(':'); linhas.push({ k, j: null, row: { casa_id: CASA.id, tipo: k.slice(0, i), id: k.slice(i + 1), dados: {}, excluido: true } }); }
  if (!linhas.length) { definirNuvem(navigator.onLine ? 'ok' : 'offline'); return; }
  if (!navigator.onLine) { definirNuvem('offline'); return; }
  enviando = true; definirNuvem('enviando');
  try {
    for (let i = 0; i < linhas.length; i += 300) {
      const lote = linhas.slice(i, i + 300);
      const { error } = await SB.from('itens').upsert(lote.map(l => l.row), { onConflict: 'casa_id,tipo,id' });
      if (error) throw error;
      for (const l of lote) { if (l.j === null) delete ULT.itens[l.k]; else ULT.itens[l.k] = l.j; }
      salvarUlt();
    }
    definirNuvem('ok');
  } catch (e) { console.error(e); definirNuvem(navigator.onLine ? 'erro' : 'offline', e.message || String(e)); }
  enviando = false;
  if (enviarDeNovo) { enviarDeNovo = false; agendarEnvio(100); }
}
async function puxarLinhas(desde) {
  const todas = [];
  for (let de = 0; ; de += 1000) {
    let q = SB.from('itens').select('tipo,id,dados,excluido,atualizado_em').eq('casa_id', CASA.id).order('atualizado_em').order('tipo').order('id').range(de, de + 999);
    if (desde) q = q.gte('atualizado_em', desde);
    const { data, error } = await q;
    if (error) throw error;
    todas.push(...data);
    if (data.length < 1000) break;
  }
  return todas;
}
async function sincronizar(completo) {
  if (!SB || !CASA || !SESSAO) return;
  try {
    definirNuvem('enviando');
    const { data: casa } = await SB.rpc('minha_casa');   // atualiza código e quantas pessoas já entraram
    if (casa && casa.id === CASA.id) { CASA = casa; ULT.casa = casa; }
    const linhas = await puxarLinhas(completo ? null : (ULT.cursor || null));
    for (const l of linhas) if (l.atualizado_em > (ULT.cursor || '')) ULT.cursor = l.atualizado_em;
    if (receberLinhas(linhas)) renderRemoto(); else salvarUlt();
    await enviar();
  } catch (e) { console.error(e); definirNuvem(navigator.onLine ? 'erro' : 'offline', e.message || String(e)); }
}
function assinar() {
  if (!SB || !CASA) return;
  if (CANAL) { SB.removeChannel(CANAL); CANAL = null; }
  CANAL = SB.channel('casa-' + CASA.id)
    .on('postgres_changes', { event: '*', schema: 'public', table: 'itens', filter: 'casa_id=eq.' + CASA.id }, p => { const r = p.new; if (r && r.tipo && receberLinhas([r])) renderRemoto(); })
    .subscribe();
}
// Prepara este aparelho para a casa. manterLocal = os dados que já estão aqui sobem para a casa.
function prepararBase(manterLocal) {
  if (!manterLocal || S.exemplo) { const tema = S.tema; S = estadoVazio(); S.tema = tema; }
  ULT = ultVazio(); ULT.casaId = CASA.id; ULT.casa = CASA; ULT.email = SESSAO?.user?.email || '';
  ULT.itens['config:casa'] = canon(entidades(S).get('config:casa').dados); // nomes e categorias da casa vencem os deste aparelho
  salvarUlt(); salvarLocal();
}
async function iniciarNuvem() {
  if (!NUVEM_ATIVA) return;
  try { await carregarScript(SUPABASE_JS); }
  catch (e) { definirNuvem('offline', 'sem internet'); render(); return; }
  SB = supabase.createClient(CFG.supabaseUrl, CFG.supabaseChave, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true, storageKey: 'financas-casa-auth' } });
  SB.auth.onAuthStateChange((ev, s) => {
    SESSAO = s;
    if (ev === 'PASSWORD_RECOVERY') { RECUPERANDO = true; render(); }
    if (ev === 'SIGNED_OUT' && CASA) { CASA = null; render(); }
  });
  const { data } = await SB.auth.getSession();
  SESSAO = data.session;
  if (SESSAO) await carregarCasa();
  else { CASA = null; definirNuvem('off'); render(); }
}
async function carregarCasa() {
  let casa = null;
  try { const { data, error } = await SB.rpc('minha_casa'); if (error) throw error; casa = data; }
  catch (e) { console.warn(e); casa = ULT.casa && ULT.email === SESSAO?.user?.email ? ULT.casa : null; definirNuvem(navigator.onLine ? 'erro' : 'offline', e.message || String(e)); }
  CASA = casa || null;
  if (!CASA) { render(); return; }
  if (ULT.casaId !== CASA.id) prepararBase(temDadosLocais() && confirm('Este aparelho já tem lançamentos guardados que não estão na casa. Juntar com os da casa?\n\nOK = juntar · Cancelar = descartar os deste aparelho'));
  ULT.casa = CASA; ULT.email = SESSAO?.user?.email || ULT.email; salvarUlt();
  render();
  assinar();
  await sincronizar(true);
}
function traduzErro(e) {
  const m = String(e && (e.message || e.error_description) || e || '');
  if (/invalid login credentials/i.test(m)) return 'E-mail ou senha incorretos.';
  if (/email not confirmed/i.test(m)) return 'Falta confirmar o e-mail: abra a mensagem de confirmação que chegou e toque no link.';
  if (/already registered|already been registered/i.test(m)) return 'Esse e-mail já tem conta — use “Entrar”.';
  if (/signups? not allowed|signup is disabled/i.test(m)) return 'Novos cadastros estão fechados neste app.';
  if (/rate limit|security purposes|too many/i.test(m)) return 'Muitas tentativas seguidas. Espere um minuto e tente de novo.';
  if (/password should be|weak password/i.test(m)) return 'Senha fraca: use pelo menos 8 caracteres.';
  if (/failed to fetch|network|load failed/i.test(m)) return 'Sem conexão com o servidor. Confira a internet.';
  return m || 'Algo deu errado.';
}
const erroForm = (f, msg) => { const p = f.querySelector('.erro'); if (p) { p.textContent = msg || ''; p.hidden = !msg; } };
function ocupar(f, sim, texto) { const b = f.querySelector('button[type=submit]'); if (!b) return; if (sim) { b.dataset.txt = b.textContent; b.textContent = texto; b.disabled = true; } else { b.textContent = b.dataset.txt || b.textContent; b.disabled = false; } }

function vCarregando() { return `<section class="card entrar"><h2>Conectando…</h2><p class="sub">Abrindo as finanças da casa.</p></section>`; }
function vEntrar() {
  const criar = AUTH_MODO === 'criar';
  return `<section class="card entrar">
    <h2>${criar ? 'Criar sua conta' : 'Entrar'}</h2>
    <p class="sub">${criar ? 'Cada pessoa cria a própria conta, com o próprio e-mail. Depois as duas contas se ligam na mesma casa.' : 'Entre com seu e-mail e senha para ver as finanças da casa.'}</p>
    <form data-form="entrar" class="form-entrar" novalidate>
      <label class="campo"><span class="lbl">E-mail</span><input type="email" name="email" autocomplete="email" inputmode="email" autocapitalize="off" value="${esc(ULT.email || '')}"></label>
      <label class="campo"><span class="lbl">Senha${criar ? ' (mínimo 8 caracteres)' : ''}</span><input type="password" name="senha" autocomplete="${criar ? 'new-password' : 'current-password'}"></label>
      <p class="erro" role="alert" hidden></p>
      <p class="ok-msg" hidden></p>
      <button class="btn primario" type="submit">${criar ? 'Criar conta' : 'Entrar'}</button>
    </form>
    <div class="botoes" style="justify-content:space-between;margin-top:16px">
      <button class="link" data-acao="auth-modo">${criar ? 'Já tenho conta — entrar' : 'Primeira vez? Criar conta'}</button>
      ${criar ? '' : '<button class="link" data-acao="auth-esqueci">Esqueci a senha</button>'}
    </div></section>`;
}
function vCasa() {
  return `<section class="card entrar"><h2>Ligar sua conta a uma casa</h2>
    <p class="sub">Você entrou como <b>${esc(SESSAO?.user?.email || '')}</b>. Quem chega primeiro cria a casa; a outra pessoa entra com o código que aparece depois.</p>
    <div class="grid dois" style="margin-top:14px">
      <form data-form="criar-casa" class="card sub-card" novalidate><h3>Criar a casa</h3>
        <label class="campo"><span class="lbl">Seu nome</span><input type="text" name="meu" value="${esc(S.pessoas[0]?.nome || '')}" autocomplete="given-name"></label>
        <label class="campo"><span class="lbl">Nome da outra pessoa</span><input type="text" name="outro" value="${esc(S.pessoas[1]?.nome || '')}"></label>
        <label class="campo"><span class="lbl">Nome da casa</span><input type="text" name="nome" value="Nossa casa"></label>
        ${temDadosLocais() ? '<p class="sub">Os lançamentos que já estão neste aparelho vão junto para a casa.</p>' : ''}
        <p class="erro" role="alert" hidden></p>
        <button class="btn primario" type="submit">Criar e gerar código</button>
      </form>
      <form data-form="entrar-casa" class="card sub-card" novalidate><h3>Já tenho um código</h3>
        <label class="campo"><span class="lbl">Código de convite</span><input type="text" name="codigo" placeholder="ABCD-2345" autocapitalize="characters" autocomplete="off" spellcheck="false" maxlength="12"></label>
        <p class="sub">A pessoa que criou a casa encontra o código em Ajustes.</p>
        <p class="erro" role="alert" hidden></p>
        <button class="btn primario" type="submit">Entrar na casa</button>
      </form>
    </div>
    <div class="botoes"><button class="link" data-acao="sair">Sair desta conta</button></div></section>`;
}
function vNovaSenha() {
  return `<section class="card entrar"><h2>Nova senha</h2><p class="sub">Escolha a nova senha da sua conta.</p>
    <form data-form="nova-senha" class="form-entrar" novalidate>
      <label class="campo"><span class="lbl">Nova senha (mínimo 8 caracteres)</span><input type="password" name="senha" autocomplete="new-password"></label>
      <p class="erro" role="alert" hidden></p>
      <button class="btn primario" type="submit">Salvar senha</button>
    </form></section>`;
}
function modoTela() {
  if (!NUVEM_ATIVA) return 'app';
  if (RECUPERANDO) return 'senha';
  if (CASA) return 'app';
  if (!SB) return NUVEM.estado === 'conectando' ? 'carregando' : 'entrar';
  if (!SESSAO) return 'entrar';
  return 'casa';
}
const FORMS = {
  async entrar(f) {
    const email = f.email.value.trim(), senha = f.senha.value, criar = AUTH_MODO === 'criar';
    erroForm(f, ''); const ok = f.querySelector('.ok-msg'); ok.hidden = true;
    if (!SB) { erroForm(f, 'Sem conexão com o servidor. Confira a internet e recarregue.'); return; }
    if (!/^\S+@\S+\.\S+$/.test(email)) { erroForm(f, 'Digite um e-mail válido.'); return; }
    if (criar ? senha.length < 8 : !senha) { erroForm(f, criar ? 'Use uma senha com pelo menos 8 caracteres.' : 'Digite a senha.'); return; }
    ocupar(f, true, criar ? 'Criando…' : 'Entrando…');
    try {
      if (criar) {
        const { data, error } = await SB.auth.signUp({ email, password: senha, options: { emailRedirectTo: urlBase() } });
        if (error) throw error;
        ULT.email = email; salvarUlt();
        if (data.session) { SESSAO = data.session; await carregarCasa(); return; }
        AUTH_MODO = 'entrar'; render();
        const m = $('[data-form="entrar"] .ok-msg'); m.textContent = `Conta criada. Abra o e-mail que chegou em ${email}, toque no link de confirmação e depois volte aqui e entre com a senha.`; m.hidden = false;
        return;
      }
      const { data, error } = await SB.auth.signInWithPassword({ email, password: senha });
      if (error) throw error;
      SESSAO = data.session; ULT.email = email; salvarUlt();
      await carregarCasa();
    } catch (e) { ocupar(f, false); erroForm(f, traduzErro(e)); }
  },
  async 'criar-casa'(f) {
    const meu = f.meu.value.trim(), outro = f.outro.value.trim();
    if (!meu || !outro) { erroForm(f, 'Preencha os dois nomes.'); return; }
    erroForm(f, ''); ocupar(f, true, 'Criando…');
    try {
      const { data, error } = await SB.rpc('criar_casa', { p_nome: f.nome.value.trim() });
      if (error) throw error;
      CASA = data;
      prepararBase(true);
      delete ULT.itens['config:casa'];   // quem cria a casa define os nomes e as categorias
      S.pessoas = [{ id: 'p1', nome: meu }, { id: 'p2', nome: outro }];
      salvarLocal(); salvarUlt();
      ABA = 'painel'; render(); assinar(); await enviar();
      toast('Casa criada. Passe o código para ' + outro + '.', 5000);
    } catch (e) { ocupar(f, false); erroForm(f, traduzErro(e)); }
  },
  async 'entrar-casa'(f) {
    const codigo = f.codigo.value.trim();
    if (codigo.replace(/[^A-Za-z0-9]/g, '').length < 6) { erroForm(f, 'Digite o código de 8 letras e números.'); return; }
    erroForm(f, ''); ocupar(f, true, 'Entrando…');
    try {
      const { data, error } = await SB.rpc('entrar_casa', { p_codigo: codigo });
      if (error) throw error;
      CASA = data;
      prepararBase(temDadosLocais() && confirm('Este aparelho já tem lançamentos guardados. Juntar com os da casa?\n\nOK = juntar · Cancelar = descartar os deste aparelho'));
      ABA = 'painel'; render(); assinar(); await sincronizar(true); render();
      toast('Pronto! Você entrou em ' + CASA.nome + '.', 4000);
    } catch (e) { ocupar(f, false); erroForm(f, traduzErro(e)); }
  },
  async 'nova-senha'(f) {
    const senha = f.senha.value;
    if (senha.length < 8) { erroForm(f, 'Use pelo menos 8 caracteres.'); return; }
    ocupar(f, true, 'Salvando…');
    try {
      const { error } = await SB.auth.updateUser({ password: senha });
      if (error) throw error;
      RECUPERANDO = false; history.replaceState(null, '', urlBase());
      toast('Senha alterada.');
      await carregarCasa();
    } catch (e) { ocupar(f, false); erroForm(f, traduzErro(e)); }
  },
};
async function sairDaConta() {
  if (CASA && !confirm('Sair deste aparelho? Os dados continuam na casa; aqui eles são apagados até você entrar de novo.')) return;
  clearTimeout(envioTimer);
  if (CASA && SB) { await enviar(); if (NUVEM.estado === 'offline' || NUVEM.estado === 'erro') { if (!confirm('Há mudanças que ainda não subiram para a nuvem (sem conexão). Sair mesmo assim e perder essas mudanças?')) return; } }
  if (CANAL && SB) { SB.removeChannel(CANAL); CANAL = null; }
  try { if (SB) await SB.auth.signOut(); } catch (e) {}
  SESSAO = null; CASA = null; IMP = null;
  const tema = S.tema; S = estadoVazio(); S.tema = tema; salvarLocal();
  const email = ULT.email; ULT = ultVazio(); ULT.email = email; salvarUlt();
  AUTH_MODO = 'entrar'; ABA = 'painel'; definirNuvem('off'); render();
}

const ACOES = {
  'auth-modo': () => { AUTH_MODO = AUTH_MODO === 'criar' ? 'entrar' : 'criar'; render(); },
  'auth-esqueci': async () => {
    const f = $('[data-form="entrar"]'), email = f.email.value.trim();
    if (!/^\S+@\S+\.\S+$/.test(email)) { erroForm(f, 'Digite seu e-mail acima e toque de novo em “Esqueci a senha”.'); return; }
    try { const { error } = await SB.auth.resetPasswordForEmail(email, { redirectTo: urlBase() }); if (error) throw error; erroForm(f, ''); const m = f.querySelector('.ok-msg'); m.textContent = `Se houver uma conta com ${email}, chega um e-mail com o link para criar uma nova senha.`; m.hidden = false; }
    catch (e) { erroForm(f, traduzErro(e)); }
  },
  'ir-login': () => { AUTH_MODO = 'entrar'; CASA = null; render(); },
  'ver-conta': () => { ABA = 'ajustes'; render(); $('#secConta')?.scrollIntoView({ block: 'start' }); },
  sair: () => sairDaConta(),
  sincronizar: async () => { await sincronizar(true); render(); toast(NUVEM.estado === 'ok' ? 'Tudo sincronizado' : (ROT_NUVEM[NUVEM.estado] || ['—'])[0]); },
  'copiar-codigo': async () => { try { await navigator.clipboard.writeText(fmtCodigo(CASA.codigo_convite)); toast('Código copiado'); } catch (e) { toast('Código: ' + fmtCodigo(CASA.codigo_convite), 6000); } },
  'trocar-codigo': async () => {
    if (!confirm('Gerar um código novo? O antigo para de funcionar.')) return;
    try { const { data, error } = await SB.rpc('trocar_codigo'); if (error) throw error; CASA.codigo_convite = data; ULT.casa = CASA; salvarUlt(); render(); toast('Código novo gerado'); }
    catch (e) { toast(traduzErro(e), 5000); }
  },
  mes: t => { MES = somaMes(MES, +t.dataset.d); render(); },
  'mes-hoje': () => { MES = mesAtual(); render(); },
  'novo-lanc': () => formLancamento(),
  'editar-lanc': t => { const l = S.lancamentos.find(x => x.id === t.dataset.id); if (l) formLancamento(l); },
  'nova-receita': () => formReceita(),
  'editar-receita': t => { const r = S.receitas.find(x => x.id === t.dataset.id); if (r) formReceita(r); },
  'novo-cartao': () => formCartao(),
  'editar-cartao': t => { const c = S.cartoes.find(x => x.id === t.dataset.id); if (c) formCartao(c); },
  'nova-fixa': () => formFixa(),
  'editar-fixa': t => { const f = S.fixas.find(x => x.id === t.dataset.id); if (f) formFixa(f); },
  'ver-cartao': t => { FILTRO = { q: '', pessoa: '', cat: '', meio: 'cartao:' + t.dataset.id, visao: 'pagamento' }; ABA = 'lancamentos'; render(); window.scrollTo(0, 0); },
  'ir-importar': () => { ABA = 'importar'; render(); window.scrollTo(0, 0); },
  visao: t => { FILTRO.visao = t.dataset.v; render(); },
  'limpar-filtros': () => { FILTRO = { q: '', pessoa: '', meio: '', cat: '', visao: FILTRO.visao }; render(); },
  exemplo: () => { if (CASA) { toast('Com a casa na nuvem, os dados de exemplo ficam desligados.'); return; } if ((S.lancamentos.length || S.receitas.length || S.fixas.length) && !confirm('Isso substitui os dados atuais pelos de exemplo. Continuar?')) return; carregarExemplo(); },
  'sair-exemplo': () => { if (!confirm('Apagar os dados de exemplo e começar do zero?')) return; const tema = S.tema; S = estadoVazio(); S.tema = tema; salvar(); MES = mesAtual(); IMP = null; ABA = 'painel'; render(); },
  'apagar-tudo': () => {
    if (CASA) { if (!confirm('Apagar TODOS os lançamentos, rendas, contas fixas e cartões da casa — para vocês dois, nos dois celulares? Não dá para desfazer; baixe um backup antes.') || !confirm('Confirma mesmo? Isso apaga os dados da casa inteira.')) return; }
    else if (!confirm('Apagar TODOS os dados deste navegador? Não dá para desfazer — baixe um backup antes, se quiser.')) return;
    const { tema, pessoas, categorias } = S; S = estadoVazio(); Object.assign(S, { tema, pessoas, categorias }); salvar(); IMP = null; render(); toast('Tudo apagado');
  },
  'exportar-json': () => baixar(`financas-casa-backup-${hojeISO()}.json`, JSON.stringify(S, null, 1), 'application/json'),
  'importar-json': () => $('#arqBackup').click(),
  'exportar-csv': exportarCSV,
  'import-texto': () => { const t = $('#txtImport').value; if (!t.trim()) { toast('Cole o texto da fatura primeiro.'); return; } iniciarImport(textoParaFonte(t, 'texto colado')); if (!IMP.linhas.length && !IMP.fonte.tabela) { IMP = null; toast('Não achei linhas com data, descrição e valor nesse texto.', 5000); return; } render(); $('#preview')?.scrollIntoView({ behavior: 'smooth', block: 'start' }); },
  'import-cancelar': () => { IMP = null; render(); },
  'import-confirmar': () => confirmarImport(),
  'ler-notif': () => { const p = interpretarNotificacao($('#txtNotif').value); if (!p) { toast('Não achei um valor em R$ nessa notificação.'); return; } formLancamento(null, p); },
  tema: t => { S.tema = t.dataset.v; salvar(); render(); },
  'add-cat': () => { const v = $('#novaCat').value.trim(); if (!v) return; if (S.categorias.some(c => norm(c) === norm(v))) { toast('Essa categoria já existe.'); return; } S.categorias.splice(S.categorias.length - 1, 0, v); salvar(); render(); },
  'rem-cat': t => { const c = t.dataset.v; const n = S.lancamentos.filter(l => l.categoria === c).length + S.fixas.filter(f => f.categoria === c).length; if (n && !confirm(`${n} item(ns) usam “${c}”. Eles passam para “Outros”. Continuar?`)) return; S.categorias = S.categorias.filter(x => x !== c); S.lancamentos.forEach(l => { if (l.categoria === c) l.categoria = 'Outros'; }); S.fixas.forEach(f => { if (f.categoria === c) f.categoria = 'Outros'; }); S.regras = S.regras.filter(r => r.categoria !== c); salvar(); render(); },
  'add-regra': () => { const termo = norm($('#novaRegraTermo').value); if (!termo) return; S.regras = S.regras.filter(r => norm(r.termo) !== termo); S.regras.unshift({ termo, categoria: $('#novaRegraCat').value, ordem: Date.now() }); salvar(); render(); toast('Regra criada'); },
  'rem-regra': t => { S.regras.splice(+t.dataset.i, 1); salvar(); render(); },
  recategorizar: () => { let n = 0; for (const l of S.lancamentos) if (l.categoria === 'Outros') { const c = categoriaAuto(l.descricao, l.meio); if (c !== 'Outros') { l.categoria = c; n++; } } salvar(); render(); toast(n ? `${n} lançamento(s) recategorizado(s)` : 'Nenhum lançamento mudou'); },
};
document.addEventListener('click', e => {
  const t = e.target.closest('[data-acao],[data-aba]'); if (!t) return;
  if (t.closest('#modal')) return;
  if (t.dataset.aba) { ABA = t.dataset.aba; render(); window.scrollTo(0, 0); return; }
  const f = ACOES[t.dataset.acao]; if (f) { e.preventDefault(); f(t, e); }
});
document.addEventListener('input', e => { const t = e.target; if (t.matches('[data-filtro]')) { FILTRO[t.dataset.filtro] = t.value; renderListaLanc(); } });
document.addEventListener('change', e => {
  const t = e.target;
  if (t.closest('#modal')) return;
  if (t.id === 'arqImport') { const f = t.files[0]; t.value = ''; processarArquivo(f); return; }
  if (t.id === 'arqBackup') { const f = t.files[0]; t.value = ''; restaurarBackup(f); return; }
  if (t.matches('[data-filtro]')) { FILTRO[t.dataset.filtro] = t.value; if (t.tagName === 'SELECT') render(); return; }
  if (t.dataset.imp) { mudarImp(t); return; }
  if (t.dataset.impSel !== undefined) { const L = IMP.linhas[+t.dataset.impSel]; L.sel = t.checked; (IMP.ajustes[L.i] ||= {}).sel = t.checked; atualizarResumoImp(); return; }
  if (t.dataset.impCat !== undefined) {
    const k = +t.dataset.impCat, L = IMP.linhas[k], n = norm(L.descricao);
    L.categoria = t.value; (IMP.ajustes[L.i] ||= {}).categoria = t.value;
    IMP.linhas.forEach((o, j) => { if (j !== k && norm(o.descricao) === n) { o.categoria = t.value; (IMP.ajustes[o.i] ||= {}).categoria = t.value; const s = $(`[data-imp-cat="${j}"]`); if (s) s.value = t.value; } });
    return;
  }
  if (t.dataset.pessoa !== undefined) { const v = t.value.trim(); if (v) { S.pessoas[+t.dataset.pessoa].nome = v; salvar(); toast('Nome atualizado'); } else t.value = S.pessoas[+t.dataset.pessoa].nome; return; }
  if (t.dataset.regra !== undefined) { S.regras[+t.dataset.regra].categoria = t.value; salvar(); toast('Regra atualizada'); return; }
});
document.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.id === 'novaCat') ACOES['add-cat'](); if (e.key === 'Enter' && e.target.id === 'novaRegraTermo') ACOES['add-regra'](); });
// arrastar arquivo
document.addEventListener('dragover', e => { if (e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files')) { e.preventDefault(); $('#drop')?.classList.add('sobre'); } });
document.addEventListener('dragleave', e => { if (!e.relatedTarget) $('#drop')?.classList.remove('sobre'); });
document.addEventListener('drop', e => {
  if (!e.dataTransfer || !e.dataTransfer.files.length) return;
  e.preventDefault(); $('#drop')?.classList.remove('sobre');
  if (ABA !== 'importar') { ABA = 'importar'; }
  processarArquivo(e.dataTransfer.files[0]);
});
// fechar o formulário clicando fora
$('#modal').addEventListener('click', e => { if (e.target === e.currentTarget) e.currentTarget.close(); });
// dica flutuante dos gráficos
const tip = $('#tip');
function mostrarTip(el, x, y) {
  tip.textContent = el.dataset.tip; tip.classList.add('on');
  const r = tip.getBoundingClientRect();
  let left = x + 14, top = y + 14;
  if (left + r.width > innerWidth - 8) left = x - r.width - 14;
  if (top + r.height > innerHeight - 8) top = y - r.height - 14;
  tip.style.left = Math.max(8, left) + 'px'; tip.style.top = Math.max(8, top) + 'px';
}
document.addEventListener('pointermove', e => { const el = e.target.closest && e.target.closest('[data-tip]'); if (el) mostrarTip(el, e.clientX, e.clientY); else tip.classList.remove('on'); });
document.addEventListener('focusin', e => { const el = e.target.closest && e.target.closest('[data-tip]'); if (el) { const r = el.getBoundingClientRect(); mostrarTip(el, r.left + r.width / 2, r.top + 10); } });
document.addEventListener('focusout', () => tip.classList.remove('on'));
addEventListener('scroll', () => tip.classList.remove('on'), { passive: true });

// formulários de login e da casa
document.addEventListener('submit', e => {
  const f = e.target, tipo = f && f.dataset && f.dataset.form;
  if (!tipo || !FORMS[tipo]) return;
  e.preventDefault(); FORMS[tipo](f);
});
// nuvem: volta da internet, app voltando para a frente (o iPhone pausa o app em segundo plano)
addEventListener('online', () => { if (CASA && SB) sincronizar(false); });
addEventListener('offline', () => { if (CASA) definirNuvem('offline'); });
document.addEventListener('visibilitychange', () => { if (!document.hidden && CASA && SB) sincronizar(false); });
document.addEventListener('focusout', () => setTimeout(() => {
  const a = document.activeElement;
  if (renderPendente && !(a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName))) render();
}, 60));
// modo app (funciona sem internet e pode ser instalado na tela de início)
if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
  navigator.serviceWorker.register('sw.js').catch(e => console.warn('service worker', e));
}

render();
iniciarNuvem();
