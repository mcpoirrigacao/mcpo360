/* =========================================================================
   MCPO 360 — casca única dos apps MCPO (MAQCAMPO Irrigação)
   - um login Microsoft (registro Azure "MCPO API Github")
   - menu com os módulos liberados para o e-mail logado
   - Central de acessos: e-mail + módulos + nível, num único arquivo
     (SharePoint MCPOIRRIGACAO › Documentos Compartilhados › MCPO360 › acessos.json)
   - libera/retira as pastas do SharePoint de cada módulo conforme a Central
   - registro leve de uso (quem abriu qual módulo, por dia)
   Os módulos ficam em modulos/<pasta>/ e perguntam o nível a esta casca
   (window.MCPO360.nivel). Aberto fora daqui, um módulo volta para cá.
   ========================================================================= */
'use strict';

const CFG = {
  clientId: 'a02d9ccd-9650-4855-99ce-3ad5ec98c381',
  tenant: 'aa796efe-0373-4718-9911-6cbf0bfa260f',
  host: 'maqcampoirrigacao.sharepoint.com',
  site: '/sites/MCPOIRRIGACAO',
  bib: 'Documentos Compartilhados',
  pasta: 'MCPO360',
  pastaCrm: 'MCPO360 CRM Comercial',   // só quem tem CRM Comercial lê (senha da planilha)
  scopes: ['Files.ReadWrite', 'Sites.ReadWrite.All', 'User.Read'],
};
const GRAPH = 'https://graph.microsoft.com/v1.0';
const BASE = typeof location !== 'undefined' ? location.origin + location.pathname.replace(/[^/]*$/, '') : '';
const REDIRECT = BASE;   // https://mcpoirrigacao.github.io/mcpo360/ (cadastrado no Azure)
const REG = '_mcpo_permissoes.json';
const VERSAO = '202611271050';   // muda a cada publicação: força o navegador a buscar os módulos novos

/* ---------------- módulos ----------------
   niveis: do mais alto para o mais baixo.
   pasta: onde o módulo guarda os dados (link de compartilhamento e/ou caminhos na biblioteca).
   papel(nivel): permissão na própria pasta. le: pastas de outros módulos que ele lê. */
const LINK_PAINEL = 'https://maqcampoirrigacao.sharepoint.com/:f:/s/MCPOIRRIGACAO/IgCd79cL3HR6SoDH79MxwSZRAWG2ngJVooBVRaxIJiapAiA?e=jWzmFv';
const LINK_CRMSAM = 'https://maqcampoirrigacao.sharepoint.com/:f:/s/MCPOIRRIGACAO/IgBGAtqGFxaXRrJhUQ8lzqOJAd3QQQZ_zeb-i3Ox7QMpddU?e=Q8SRTl';
const MODULOS = [
  { id: 'painel', dir: 'painel', sigla: 'PA', curto: 'Painel', nome: 'Painel 360', grupo: 'Visão', desc: 'Visão geral dos apps para a diretoria',
    niveis: ['admin', 'comum'], pasta: { link: LINK_PAINEL, caminhos: ['Painel'] },
    papel: n => n === 'admin' ? 'write' : 'read', le: () => ['assistencia', 'pecas', 'compras', 'obras', 'entregas', 'orcamento', 'crmsam', 'followup', 'treinamentos'] },
  { id: 'assistencia', dir: 'assistencia', sigla: 'AT', curto: 'Assist.', nome: 'Assistência Técnica', grupo: 'Operação', desc: 'OS, faturamento e comissão dos técnicos',
    niveis: ['admin', 'comum'], pasta: { caminhos: ['assistencia'] }, papel: () => 'write',
    acesso: [['assistencia', 'assistencia_acesso.json']] },
  { id: 'pecas', dir: 'pecas', sigla: 'PÇ', curto: 'Peças', nome: 'Peças', grupo: 'Operação', desc: 'Faturamento de peças por loja',
    niveis: ['admin', 'comum'], pasta: { caminhos: ['Pecas Varejo', 'Faturamento'] }, papel: () => 'write',
    acesso: [['Pecas Varejo', 'pecas_varejo_acesso.json'], ['Pecas Varejo', 'faturamento_acesso.json'], ['Faturamento', 'faturamento_acesso.json']] },
  { id: 'compras', dir: 'compras', sigla: 'CE', curto: 'Compras', nome: 'Compras e estoque', grupo: 'Operação', desc: 'BI de compras e estoque de peças',
    niveis: ['admin', 'comum'], pasta: { caminhos: ['Compras-estoque'] }, papel: n => n === 'admin' ? 'write' : 'read',
    /* quem só lê a pasta pode gravar este arquivo: o primeiro que abrir depois da troca das planilhas atualiza o resumo */
    arquivosEditaveis: ['_resumo_processado.json.gz'],
    acesso: [['Compras-estoque', 'compras_pecas_acesso.json']] },
  { id: 'orcamento', dir: 'orcamento', sigla: 'OR', curto: 'Orçam.', nome: 'Orçamentação', grupo: 'Operação', desc: 'Orçamento de montagem de pivôs',
    niveis: ['admin', 'gestor', 'comum'], pasta: { caminhos: ['Orçamentos Montagem/Historico_montagem'] }, papel: () => 'write',
    acesso: [['Orçamentos Montagem/Historico_montagem', 'parametros.json']] },
  { id: 'obras', dir: 'obras', sigla: 'OB', curto: 'Obras', nome: 'Obras de montagem', grupo: 'Operação', desc: 'Clientes, obras, lançamentos e agenda',
    niveis: ['admin', 'comum', 'agenda'], rotulos: { agenda: 'Só agenda' }, pasta: { caminhos: ['Execução de Montagem'] },
    papel: n => n === 'admin' ? 'write' : 'read', le: n => n === 'admin' ? ['orcamento'] : [], escreve: n => n === 'admin' ? ['entregas'] : [],
    acesso: [['Execução de Montagem', 'mcpo_acesso.json']] },
  { id: 'entregas', dir: 'entregas', sigla: 'ET', curto: 'Entrega', nome: 'Entrega Técnica', grupo: 'Operação', desc: 'Entregas técnicas, aprovação da fábrica e pagamento',
    niveis: ['admin', 'consultor', 'comum'], rotulos: { consultor: 'Consultor' }, pasta: { caminhos: ['Entrega Técnica'], criar: true },
    papel: n => n === 'comum' ? 'read' : 'write' },
  { id: 'crmsam', dir: 'crmsam', sigla: 'CS', curto: 'CRM SAM', nome: 'CRM SAM', grupo: 'Comercial e pessoas', desc: 'Funil de Grandes Contas',
    niveis: ['admin', 'editor', 'comum'], rotulos: { editor: 'Editor (só os seus)' }, pasta: { link: LINK_CRMSAM, caminhos: ['CRM SAM'] },
    papel: n => n === 'comum' ? 'read' : 'write', acesso: [['@crmsam', 'crm_sam_acesso.json']] },
  { id: 'crmcomercial', dir: 'crmcomercial', sigla: 'CC', curto: 'CRM Com.', nome: 'CRM Comercial', grupo: 'Comercial e pessoas', desc: 'CRM MaqCampo (login próprio da planilha)',
    niveis: ['acesso'], rotulos: { acesso: 'Acesso' } },
  { id: 'followup', dir: 'followup', sigla: 'FU', curto: 'FollowUP', nome: 'FollowUP semanal', grupo: 'Comercial e pessoas', desc: 'Acompanhamento semanal por equipe',
    niveis: ['admin', 'equipe', 'visualizador'], rotulos: { equipe: 'Equipe (conforme FollowUP)', visualizador: 'Visualizador (todas as equipes; edita as que preenche)' },
    pasta: { caminhos: ['FollowUP'] }, papel: (n, em) => n === 'visualizador' && !FU_PREENCHE.has(em) ? 'read' : 'write',   /* visualizador que preenche alguma equipe continua gravando */ le: () => ['pecas', 'assistencia', 'treinamentos'],
    acesso: [['FollowUP', 'followup_acesso.json']] },
  { id: 'treinamentos', dir: 'treinamentos', sigla: 'TR', curto: 'Trein.', nome: 'Treinamentos', grupo: 'Comercial e pessoas', desc: 'Treinamentos de funcionários e clientes',
    niveis: ['admin', 'comum'], pasta: { caminhos: ['TreinamentosMCPO'] }, papel: () => 'write',
    acesso: [['TreinamentosMCPO', 'treinamentos_acesso.json']] },
];
const CENTRAL = { id: 'central', nome: 'Central de acessos', curto: 'Central', niveis: ['admin'] };
const MOD = Object.fromEntries(MODULOS.map(m => [m.id, m]));
const ROTULO = { admin: 'Admin', consultor: 'Consultor', comum: 'Comum', gestor: 'Gestor', editor: 'Editor', visualizador: 'Visualizador', equipe: 'Equipe', agenda: 'Só agenda', acesso: 'Acesso' };
const rotulo = (m, n) => (m.rotulos && m.rotulos[n]) || ROTULO[n] || n;

/* Funções do app no Azure → nível. MCPO.Admin = admin em tudo e na Central. */
const FUNCOES = {
  'Painel.Admin': ['painel', 'admin'], 'Painel.Comum': ['painel', 'comum'],
  'Assistencia.Admin': ['assistencia', 'admin'], 'Assistencia.Comum': ['assistencia', 'comum'],
  'Pecas.Admin': ['pecas', 'admin'], 'Pecas.Comum': ['pecas', 'comum'],
  'Compras.Admin': ['compras', 'admin'], 'Compras.Comum': ['compras', 'comum'],
  'Followup.Admin': ['followup', 'admin'], 'Followup.Leitor': ['followup', 'visualizador'],
  'Treinamentos.Admin': ['treinamentos', 'admin'], 'Treinamentos.Comum': ['treinamentos', 'comum'],
  'Obras.Admin': ['obras', 'admin'], 'Obras.Comum': ['obras', 'comum'],
  'Entregas.Admin': ['entregas', 'admin'], 'Entregas.Consultor': ['entregas', 'consultor'], 'Entregas.Comum': ['entregas', 'comum'],
  'Orcamento.Admin': ['orcamento', 'admin'], 'Orcamento.Gestor': ['orcamento', 'gestor'], 'Orcamento.Comum': ['orcamento', 'comum'],
  'CRMSAM.Admin': ['crmsam', 'admin'], 'CRMSAM.Editor': ['crmsam', 'editor'], 'CRMSAM.Comum': ['crmsam', 'comum'],
};

/* ---------------- regras puras (testáveis) ---------------- */
function melhorNivel(m, a, b) {
  if (!a) return b || null; if (!b) return a;
  const n = (m === 'central' ? CENTRAL : MOD[m]).niveis;
  const ia = n.indexOf(a), ib = n.indexOf(b);
  if (ia < 0) return b; if (ib < 0) return a;
  return ia <= ib ? a : b;
}
function niveisDe(acessos, email, funcoes) {
  const out = {};
  const u = acessos && acessos.usuarios && acessos.usuarios[email];
  const mods = (u && u.modulos) || {};
  for (const id of [...MODULOS.map(m => m.id), 'central']) {
    const m = id === 'central' ? CENTRAL : MOD[id];
    if (mods[id] && m.niveis.includes(mods[id])) out[id] = mods[id];
  }
  for (const f of funcoes || []) {
    if (f === 'MCPO.Admin') { MODULOS.forEach(m => out[m.id] = m.niveis[0]); out.central = 'admin'; continue; }
    const x = FUNCOES[f]; if (x) out[x[0]] = melhorNivel(x[0], out[x[0]], x[1]);
  }
  return out;
}
/* Pastas que cada pessoa precisa: { chave: { email: 'write'|'read' } }.
   Chaves: id do módulo dono da pasta, 'mcpo360' (acessos.json) e 'uso' (registro de uso). */
/* e-mails que preenchem alguma equipe no FollowUP (Equipes & Campos › Quem preenche) — lido na sincronização */
let FU_PREENCHE = new Set();
async function carregarPreenchedoresFU() {
  try {
    const u = await urlPasta('followup'); if (!u) return;
    const r = await lerJson(filho(u, 'followup_dados.json')); if (!r.ok || !r.data) return;
    const s = new Set();
    (((r.data.config || {}).equipes) || []).forEach(eq => { if ((eq.tipo || 'manual') === 'manual') (eq.membros || []).forEach(em => s.add(String(em).trim().toLowerCase())); });
    FU_PREENCHE = s;
  } catch (e) { console.warn('preenchedores FollowUP', e); }
}
function calcularDesejado(acessos) {
  const q = {};
  const por = (chave, email, papel) => {
    const x = (q[chave] = q[chave] || {});
    if (x[email] !== 'write') x[email] = papel;
  };
  for (const [email, u] of Object.entries((acessos && acessos.usuarios) || {})) {
    const mods = (u && u.modulos) || {};
    let algum = false;
    for (const [id, nivel] of Object.entries(mods)) {
      if (id === 'central') { if (nivel === 'admin') { por('mcpo360', email, 'write'); por('crmcfg', email, 'write'); algum = true; } continue; }
      const m = MOD[id]; if (!m || !m.niveis.includes(nivel)) continue;
      algum = true;
      if (id === 'crmcomercial') por('crmcfg', email, 'read');
      if (m.pasta && m.papel) por(id, email, m.papel(nivel, email));
      if (m.le) for (const outro of m.le(nivel)) if (MOD[outro] && MOD[outro].pasta) por(outro, email, 'read');
      if (m.escreve) for (const outro of m.escreve(nivel)) if (MOD[outro] && MOD[outro].pasta) por(outro, email, 'write');
    }
    if (algum) { por('mcpo360', email, 'read'); por('uso', email, 'write'); }
  }
  return q;
}
/* Diferença entre o que a pasta tem (registro) e o que deveria ter. */
function planejarPasta(tem, quer, forcar) {
  const liberar = [], retirar = [];
  for (const [em, papel] of Object.entries(quer || {})) if (forcar || tem[em] !== papel) liberar.push([em, papel, tem[em] || null]);
  for (const em of Object.keys(tem || {})) if (!(quer || {})[em]) retirar.push(em);
  return { liberar, retirar };
}
function mesclarImportado(acessos, email, mod, nivel) {
  email = String(email || '').trim().toLowerCase(); if (!email.includes('@')) return false;
  const m = mod === 'central' ? CENTRAL : MOD[mod]; if (!m || !m.niveis.includes(nivel)) return false;
  const u = (acessos.usuarios[email] = acessos.usuarios[email] || { modulos: {} });
  const novo = melhorNivel(mod, u.modulos[mod], nivel);
  if (novo === u.modulos[mod]) return false;
  u.modulos[mod] = novo; return true;
}
if (typeof module !== 'undefined') module.exports = { MODULOS, MOD, niveisDe, calcularDesejado, planejarPasta, mesclarImportado, melhorNivel };

/* ===================== daqui para baixo: só no navegador ===================== */
if (typeof window !== 'undefined') (function () {

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
function toast(m, ms) { const t = $('toast'); t.textContent = m; t.classList.add('show'); clearTimeout(toast._h); toast._h = setTimeout(() => t.classList.remove('show'), ms || 3200); }
function sync(txt) { const s = $('sync-status'); s.hidden = !txt; s.textContent = txt || ''; }
const hoje = () => new Date(Date.now() - 3 * 3600e3).toISOString().slice(0, 10);

let msal = null, conta = null;
let ACESSOS = null, ACESSOS_ETAG = null, ACESSOS_ERRO = '';
let NIVEIS = {};
const IFRAMES = {};

/* ---------------- login e Graph ---------------- */
async function iniciarMsal() {
  msal = msal_lib();
  await msal.initialize();
  try { await msal.handleRedirectPromise(); } catch (e) {}
  const contas = msal.getAllAccounts();
  conta = contas[0] || null;
  if (conta) msal.setActiveAccount(conta);
}
function msal_lib() {
  return new window.msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenant, redirectUri: REDIRECT },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false },
  });
}
async function token() {
  if (!conta) return null;
  try { return (await msal.acquireTokenSilent({ scopes: CFG.scopes, account: conta })).accessToken; }
  catch (e) {
    console.warn('MCPO 360: token silencioso falhou', e && (e.errorCode || e.message));
    pedirLoginDeNovo(); return null;   // popup só com clique (senão o navegador bloqueia)
  }
}
async function g(url, opt) {
  opt = opt || {};
  const tk = await token(); if (!tk) throw new Error('Sem login Microsoft.');
  const h = Object.assign({ Authorization: 'Bearer ' + tk }, opt.headers || {});
  return fetch(url.startsWith('http') ? url : GRAPH + url, Object.assign({}, opt, { headers: h, cache: 'no-store' }));
}
let _site = null;
async function siteId() {
  if (_site) return _site;
  const r = await g(`/sites/${CFG.host}:${CFG.site}`); if (!r.ok) throw new Error('Site do SharePoint: HTTP ' + r.status);
  _site = (await r.json()).id; return _site;
}
const enc = p => p.split('/').map(encodeURIComponent).join('/');
async function urlCaminho(caminho) { return `${GRAPH}/sites/${await siteId()}/drive/root:${enc('/' + CFG.bib + '/' + caminho)}`; }
const acao = (url, a) => url.includes('/root:') ? url + ':/' + a : url + '/' + a;
const filho = (url, nome) => url.includes('/root:') ? url + '/' + encodeURIComponent(nome) : url + ':/' + encodeURIComponent(nome);
function shareId(u) { return 'u!' + btoa(unescape(encodeURIComponent(u))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-'); }
const _pastaCache = {};
async function urlPasta(chave) {
  if (_pastaCache[chave] !== undefined) return _pastaCache[chave];
  let url = null;
  if (chave === 'mcpo360') url = await garantirPasta(CFG.pasta);
  else if (chave === 'uso') { await garantirPasta(CFG.pasta); url = await garantirPasta(CFG.pasta + '/uso'); }
  else if (chave === 'crmcfg') url = await garantirPasta(CFG.pastaCrm);
  else {
    const p = MOD[chave] && MOD[chave].pasta;
    if (p && p.link) {
      for (const l of [p.link, p.link.split('?')[0]]) {
        try { const r = await g(`/shares/${shareId(l)}/driveItem`); if (r.ok) { const it = await r.json(); if (it.folder && it.parentReference) { url = `${GRAPH}/drives/${it.parentReference.driveId}/items/${it.id}`; break; } } } catch (e) {}
      }
    }
    if (!url && p) for (const c of p.caminhos || []) {
      const u = await urlCaminho(c); const r = await g(u); if (r.ok) { url = u; break; }
    }
    if (!url && p && p.criar && p.caminhos && p.caminhos[0]) { try { url = await garantirPasta(p.caminhos[0]); } catch (e) { url = null; } }
  }
  _pastaCache[chave] = url; return url;
}
async function garantirPasta(caminho) {
  const u = await urlCaminho(caminho);
  const r = await g(u); if (r.ok) return u;
  const partes = caminho.split('/'); const nome = partes.pop();
  const pai = await urlCaminho(partes.join('/') || '');
  const paiUrl = partes.length ? pai : `${GRAPH}/sites/${await siteId()}/drive/root:${enc('/' + CFG.bib)}`;
  const c = await g(acao(paiUrl, 'children'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: nome, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) });
  if (c.ok || c.status === 409) return u;
  throw new Error('Não foi possível criar a pasta ' + caminho + ' (HTTP ' + c.status + ')');
}
async function lerJson(url) {   // url de arquivo → {ok, data, etag, status}
  const m = await g(url); if (m.status === 404) return { ok: true, missing: true, data: null, etag: null };
  if (!m.ok) return { ok: false, status: m.status };
  const etag = (await m.json()).eTag || null;
  const c = await g(acao(url, 'content')); if (!c.ok) return { ok: false, status: c.status };
  try { return { ok: true, data: await c.json(), etag }; } catch (e) { return { ok: false, status: 0 }; }
}
async function gravarJson(url, dados, etag) {
  const h = { 'Content-Type': 'application/json' };
  let u = acao(url, 'content');
  if (etag) h['If-Match'] = etag; else u += '?@microsoft.graph.conflictBehavior=fail';
  const r = await g(u, { method: 'PUT', headers: h, body: JSON.stringify(dados, null, 1) });
  let novo = null; if (r.ok) { try { novo = (await r.json()).eTag; } catch (e) {} }
  return { ok: r.ok, status: r.status, etag: novo };
}

/* ---------------- acessos.json ---------------- */
const vazio = () => ({ versao: 1, usuarios: {}, perfis: [], historico: [] });
async function urlAcessos() { return (await urlCaminho(CFG.pasta + '/acessos.json')); }
async function carregarAcessos() {
  ACESSOS_ERRO = '';
  try {
    const r = await lerJson(await urlAcessos());
    if (r.ok && r.missing) { ACESSOS = null; ACESSOS_ETAG = null; return; }
    if (!r.ok) { ACESSOS = null; ACESSOS_ERRO = r.status === 403 || r.status === 404 ? 'sem permissão para ler a Central' : 'erro ' + r.status; return; }
    ACESSOS = Object.assign(vazio(), r.data || {}); ACESSOS_ETAG = r.etag;
  } catch (e) { ACESSOS = null; ACESSOS_ERRO = e.message; }
}
/* Grava aplicando "mudar" sobre a versão mais nova (trava por ETag). */
async function salvarAcessos(mudar, descricao) {
  await garantirPasta(CFG.pasta);
  const url = await urlAcessos();
  for (let t = 0; t < 5; t++) {
    let base = ACESSOS ? JSON.parse(JSON.stringify(ACESSOS)) : vazio(), etag = ACESSOS_ETAG;
    if (t > 0) { const r = await lerJson(url); if (!r.ok) break; base = r.missing ? vazio() : Object.assign(vazio(), r.data); etag = r.etag; }
    const hist = mudar(base);
    if (hist === false) return true;
    (hist || []).forEach(h => base.historico.unshift(Object.assign({ em: new Date().toISOString(), por: emailAtual() }, h)));
    base.historico = base.historico.slice(0, 300);
    base.atualizadoEm = new Date().toISOString(); base.atualizadoPor = emailAtual();
    const w = await gravarJson(url, base, etag);
    if (w.ok) { ACESSOS = base; ACESSOS_ETAG = w.etag; recalcular(); return true; }
    if (w.status !== 412 && w.status !== 409) { toast('Não foi possível salvar a Central (HTTP ' + w.status + ')' + (w.status === 403 ? ': sem permissão de edição na pasta MCPO360.' : '.')); return false; }
  }
  toast('Não foi possível salvar a Central agora. Tente de novo.'); return false;
}

/* ---------------- níveis do usuário logado ---------------- */
const emailAtual = () => conta ? String(conta.username || '').toLowerCase() : '';
const funcoes = () => (conta && conta.idTokenClaims && Array.isArray(conta.idTokenClaims.roles)) ? conta.idTokenClaims.roles : [];
function recalcular() { NIVEIS = niveisDe(ACESSOS, emailAtual(), funcoes()); }
const ehAdminCentral = () => NIVEIS.central === 'admin';
const meusModulos = () => MODULOS.filter(m => NIVEIS[m.id]);

/* ---------------- telas ---------------- */
const TELAS = ['login', 'carregando', 'sem-acesso', 'inicio', 'central', 'uso', 'modulo'];
function mostrarTela(t) { TELAS.forEach(x => $('tela-' + x).hidden = x !== t); }
const menuItem = (href, nome, ativo, sg, nivel) =>
  `<a href="#${href}" class="${ativo ? 'ativo' : ''}" title="${esc(nome)}"><span class="ico" aria-hidden="true">${esc(sg)}</span><span class="txt">${esc(nome)}</span>${nivel ? `<span class="nivel">${esc(nivel)}</span>` : ''}</a>`;
function renderMenu(ativo) {
  const nav = $('menu'); if (!conta) { nav.innerHTML = ''; return; }
  const fechado = document.body.classList.contains('menu-fechado');
  let h = `<button type="button" class="menu-toggle" id="menu-toggle" aria-expanded="${!fechado}" aria-label="${fechado ? 'Abrir menu' : 'Fechar menu'}" title="${fechado ? 'Abrir menu' : 'Fechar menu'}"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="${fechado ? 'M9 6l6 6-6 6' : 'M15 6l-6 6 6 6'}"/></svg><span class="txt">Recolher menu</span></button>`;
  h += menuItem('inicio', 'Início', ativo === 'inicio', 'IN');
  let grupo = '';
  for (const m of meusModulos()) {
    if (m.grupo !== grupo) { grupo = m.grupo; h += `<div class="grupo">${esc(grupo)}</div>`; }
    h += menuItem(m.id, m.nome, ativo === m.id, m.sigla, rotulo(m, NIVEIS[m.id]));
  }
  if (ehAdminCentral()) {
    h += `<div class="grupo">Administração</div>`;
    h += menuItem('central', 'Central de acessos', ativo === 'central', 'CA');
    h += menuItem('uso', 'Uso dos módulos', ativo === 'uso', 'US');
  }
  h += `<div class="rodape">Só aparecem os módulos liberados para ${esc(emailAtual())}.</div>`;
  nav.innerHTML = h;
  $('menu-toggle').onclick = e => { e.stopPropagation(); alternarMenu(); };
}
function alternarMenu(forcar) {
  const fechar = typeof forcar === 'boolean' ? forcar : !document.body.classList.contains('menu-fechado');
  document.body.classList.toggle('menu-fechado', fechar);
  try { localStorage.setItem('mcpo360_menu', fechar ? 'fechado' : 'aberto'); } catch (e) {}
  const r = (location.hash || '#inicio').slice(1) || 'inicio'; renderMenu(r);
}
function renderInicio() {
  const mods = meusModulos();
  const nome = (conta && conta.name) ? conta.name.split(' ')[0] : emailAtual().split('@')[0];
  let h = `<div class="cab"><div><h1>Olá, ${esc(nome)}</h1><p>Você tem acesso a ${mods.length} módulo(s). Os dados continuam nas mesmas pastas do SharePoint MCPOIRRIGACAO.</p></div></div>`;
  if (ehAdminCentral() && !ACESSOS) h += `<div class="alerta" style="margin-bottom:16px"><b>A Central de acessos ainda não foi criada.</b> Abra a <a href="#central">Central de acessos</a> e use “Importar listas dos apps” para trazer quem já usa cada app.</div>`;
  h += `<div class="grade">` + mods.map(m => `<a class="card-mod" href="#${m.id}"><div class="linha"><strong>${esc(m.nome)}</strong><span class="selo">${esc(rotulo(m, NIVEIS[m.id]))}</span></div><span class="desc">${esc(m.desc)}</span></a>`).join('') + `</div>`;
  if (ehAdminCentral()) h += `<div class="grade" style="margin-top:14px"><a class="card-mod" href="#central"><div class="linha"><strong>Central de acessos</strong><span class="selo azure">admin</span></div><span class="desc">${ACESSOS ? Object.keys(ACESSOS.usuarios).length + ' pessoa(s) cadastradas' : 'ainda não criada'}</span></a><a class="card-mod" href="#uso"><div class="linha"><strong>Uso dos módulos</strong><span class="selo azure">admin</span></div><span class="desc">Quem abriu o quê nos últimos 30 dias</span></a></div>`;
  $('tela-inicio').innerHTML = h;
}
function abrirModulo(id) {
  const box = $('tela-modulo');
  if (!IFRAMES[id]) {
    const f = document.createElement('iframe');
    f.title = MOD[id].nome; f.src = 'modulos/' + MOD[id].dir + '/index.html?v=' + VERSAO;
    f.setAttribute('allow', 'clipboard-write');
    box.appendChild(f); IFRAMES[id] = f;
  }
  Object.entries(IFRAMES).forEach(([k, f]) => f.style.display = k === id ? 'block' : 'none');
  registrarUso(id);
}
function rota() {
  const r = (location.hash || '#inicio').slice(1) || 'inicio';
  $('menu').classList.remove('aberto');
  if (!conta) { mostrarTela('login'); renderMenu(); return; }
  /* sessão Microsoft expirada: não é falta de acesso, é só confirmar o login */
  if (_pedindo && !ACESSOS) { $('login-erro').textContent = 'Sua sessão expirou. Clique em Entrar para continuar.'; mostrarTela('login'); renderMenu(); return; }
  if (!meusModulos().length && !ehAdminCentral()) { $('sem-acesso-email').textContent = emailAtual(); $('sem-acesso-erro').textContent = ACESSOS_ERRO ? 'Detalhe: ' + ACESSOS_ERRO : ''; mostrarTela('sem-acesso'); renderMenu(); return; }
  if (MOD[r] && NIVEIS[r]) { renderMenu(r); mostrarTela('modulo'); abrirModulo(r); document.title = MOD[r].nome + ' · MCPO 360'; return; }
  if (r === 'central' && ehAdminCentral()) { renderMenu(r); renderCentral(); mostrarTela('central'); document.title = 'Central de acessos · MCPO 360'; return; }
  if (r === 'uso' && ehAdminCentral()) { renderMenu(r); mostrarTela('uso'); renderUso(); document.title = 'Uso · MCPO 360'; return; }
  if (MOD[r] && !NIVEIS[r]) toast('Você não tem acesso a ' + MOD[r].nome + '.');
  renderMenu('inicio'); renderInicio(); mostrarTela('inicio'); document.title = 'MCPO 360';
}

/* ---------------- Central de acessos ---------------- */
let FORM = { email: '', mods: {}, editando: false };
let FILTRO = '';
let PROGRESSO = '';

/* ---------------- armazenamento: tamanho dos arquivos + backup mensal ----------------
   - Medição: lista os arquivos de cada pasta (só leitura) e avisa quando um arquivo de dados passa de LIMITE_MB.
   - Backup mensal: uma vez por mês, um administrador da Central copia (cópia feita pelo próprio SharePoint)
     os arquivos de dados .json de cada pasta para <pasta>/backups/AAAA-MM/. Controle em MCPO360/backup_mensal.json. */
const LIMITE_MB = 3;
const SUB_IGNORAR = /^(backups?|uso|clientes|confirmacoes|convites)$/i;
let ARMAZ = null, ARMAZ_LENDO = false, BACKUP_INFO = null;
const ehDado = n => /\.json$/i.test(n);
const mesSP = () => new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit' }).format(new Date()).slice(0, 7);
const fmtTam = b => b >= 1048576 ? (b / 1048576).toFixed(1).replace('.', ',') + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
async function listarArquivos(url, pref, nivel) {
  const out = []; let next = acao(url, 'children') + '?$select=id,name,size,file,folder,lastModifiedDateTime,parentReference&$top=999';
  while (next) {
    const r = await g(next); if (!r.ok) throw new Error('HTTP ' + r.status);
    const j = await r.json();
    for (const it of j.value || []) {
      if (it.file) out.push({ nome: pref + it.name, tam: it.size || 0, mod: it.lastModifiedDateTime, id: it.id, driveId: it.parentReference && it.parentReference.driveId });
      else if (it.folder && nivel < 1 && !SUB_IGNORAR.test(it.name) && it.parentReference) out.push(...await listarArquivos(`${GRAPH}/drives/${it.parentReference.driveId}/items/${it.id}`, pref + it.name + '/', nivel + 1));
    }
    next = j['@odata.nextLink'] || null;
  }
  return out;
}
function pastasArmaz() { return [{ id: 'mcpo360', nome: 'MCPO 360 (Central)' }].concat(MODULOS.filter(m => m.pasta).map(m => ({ id: m.id, nome: m.nome }))); }
async function medirArmazenamento() {
  if (ARMAZ_LENDO) return; ARMAZ_LENDO = true;
  const res = [];
  for (const p of pastasArmaz()) {
    try { const u = await urlPasta(p.id); if (!u) { res.push({ ...p, erro: 'pasta não encontrada' }); continue; } res.push({ ...p, arquivos: await listarArquivos(u, '', 0) }); }
    catch (e) { res.push({ ...p, erro: e.message }); }
  }
  ARMAZ = { em: new Date().toISOString(), pastas: res }; ARMAZ_LENDO = false;
  return ARMAZ;
}
const acimaDoLimite = () => ARMAZ ? ARMAZ.pastas.flatMap(p => (p.arquivos || []).filter(a => ehDado(a.nome) && a.tam > LIMITE_MB * 1048576).map(a => p.nome + ' › ' + a.nome)) : [];
async function pastaFilha(paiUrl, nome) {
  const c = await g(acao(paiUrl, 'children'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: nome, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) });
  if (!c.ok && c.status !== 409) throw new Error('criar ' + nome + ': HTTP ' + c.status);
  const r = await g(filho(paiUrl, nome)); if (!r.ok) throw new Error('ler ' + nome + ': HTTP ' + r.status);
  const j = await r.json(); return { url: `${GRAPH}/drives/${j.parentReference.driveId}/items/${j.id}`, driveId: j.parentReference.driveId, id: j.id };
}
async function backupMensal(forcar) {
  if (!ehAdminCentral()) return null;
  const mes = mesSP(); const ctrlUrl = filho(await urlPasta('mcpo360'), 'backup_mensal.json');
  const r = await lerJson(ctrlUrl); if (!r.ok) return null;
  const ctrl = Object.assign({ meses: {} }, r.data || {}); BACKUP_INFO = ctrl;
  if (ctrl.meses[mes] && !forcar) return ctrl;
  const pastas = {};
  for (const p of pastasArmaz()) {
    let ok = 0, erro = 0;
    try {
      const u = await urlPasta(p.id); if (!u) continue;
      const arqs = (await listarArquivos(u, '', 0)).filter(a => ehDado(a.nome) && a.id && a.driveId);
      if (!arqs.length) continue;
      const dest = await pastaFilha((await pastaFilha(u, 'backups')).url, mes);
      for (const a of arqs) {
        const c = await g(`${GRAPH}/drives/${a.driveId}/items/${a.id}/copy?@microsoft.graph.conflictBehavior=replace`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ parentReference: { driveId: dest.driveId, id: dest.id }, name: a.nome.replace(/\//g, '__') }) });
        if (c.status === 202 || c.ok) ok++; else erro++;
      }
    } catch (e) { erro++; console.warn('backup', p.id, e); }
    pastas[p.id] = { arquivos: ok, erros: erro };
  }
  ctrl.meses[mes] = { em: new Date().toISOString(), por: emailAtual(), pastas };
  const w = await gravarJson(ctrlUrl, ctrl, r.etag);
  if (w.ok) BACKUP_INFO = ctrl;
  return ctrl;
}
async function rotinaArmazenamento() {
  if (!ehAdminCentral()) return;
  try { await backupMensal(); } catch (e) { console.warn('backup mensal', e); }
  try { await medirArmazenamento(); } catch (e) { console.warn('medição', e); }
  const al = acimaDoLimite();
  if (al.length) toast(`⚠ ${al.length} arquivo(s) de dados acima de ${LIMITE_MB} MB — veja "Armazenamento" na Central de acessos.`, 9000);
  if (!$('tela-central').hidden) renderCentral();
}
function htmlArmazenamento() {
  const bi = BACKUP_INFO && BACKUP_INFO.meses ? Object.keys(BACKUP_INFO.meses).sort().reverse() : [];
  const ult = bi.length ? BACKUP_INFO.meses[bi[0]] : null;
  const totBk = ult ? Object.values(ult.pastas || {}).reduce((a, x) => a + (x.arquivos || 0), 0) : 0;
  const errBk = ult ? Object.values(ult.pastas || {}).reduce((a, x) => a + (x.erros || 0), 0) : 0;
  let h = `<div class="secao"><div class="form-linha" style="justify-content:space-between"><h2>Armazenamento</h2><div class="acoes"><button class="btn" id="a-medir" type="button">${ARMAZ_LENDO ? 'Medindo…' : 'Medir agora'}</button><button class="btn" id="a-backup" type="button">Fazer backup agora</button></div></div><div class="cartao">`;
  h += `<p class="nota" style="margin:0 0 10px">Arquivos de dados (.json) acima de <b>${LIMITE_MB} MB</b> ficam em destaque: é o sinal para separar os movimentos por ano naquele módulo. Planilhas de origem (Excel/CSV) só são lidas quando os dados mudam e não entram no alerta. `
    + (ult ? `Último backup mensal: <b>${esc(bi[0])}</b> em ${esc(new Date(ult.em).toLocaleString('pt-BR'))} por ${esc(ult.por)} — ${totBk} arquivo(s) copiados para <i>backups/${esc(bi[0])}</i> de cada pasta${errBk ? `, <b style="color:var(--erro,#b3261e)">${errBk} erro(s)</b>` : ''}.` : 'Backup mensal: ainda não feito (roda sozinho quando um administrador da Central abre o MCPO 360).') + `</p>`;
  if (!ARMAZ) h += `<p class="nota">${ARMAZ_LENDO ? 'Medindo as pastas…' : 'Clique em <b>Medir agora</b> para ver o tamanho dos arquivos.'}</p>`;
  else {
    const lim = LIMITE_MB * 1048576;
    h += `<table class="t"><thead><tr><th>Pasta</th><th>Arquivos de dados</th><th style="text-align:right">Total dados</th><th>Maior arquivo de dados</th><th style="text-align:right">Planilhas / outros</th></tr></thead><tbody>`;
    ARMAZ.pastas.forEach(p => {
      if (p.erro) { h += `<tr><td><b>${esc(p.nome)}</b></td><td colspan="4" class="nota">${esc(p.erro)}</td></tr>`; return; }
      const d = p.arquivos.filter(a => ehDado(a.nome)), o = p.arquivos.filter(a => !ehDado(a.nome));
      const td = d.reduce((a, x) => a + x.tam, 0), to = o.reduce((a, x) => a + x.tam, 0);
      const mx = d.slice().sort((a, b) => b.tam - a.tam)[0];
      const alto = mx && mx.tam > lim;
      h += `<tr${alto ? ' style="background:#fdecea"' : ''}><td><b>${esc(p.nome)}</b></td><td>${d.length}</td><td style="text-align:right">${fmtTam(td)}</td><td>${mx ? `${alto ? '⚠ ' : ''}${esc(mx.nome)} · <b>${fmtTam(mx.tam)}</b>` : '—'}</td><td style="text-align:right">${o.length ? `${o.length} · ${fmtTam(to)}` : '—'}</td></tr>`;
    });
    h += `</tbody></table><p class="nota" style="margin:8px 0 0">Medido em ${esc(new Date(ARMAZ.em).toLocaleString('pt-BR'))}. Não entram as subpastas backups, uso, Clientes, confirmacoes e convites.</p>`;
  }
  return h + `</div></div>`;
}

function renderCentral() {
  const el = $('tela-central');
  const us = ACESSOS ? Object.entries(ACESSOS.usuarios).sort((a, b) => a[0].localeCompare(b[0])) : [];
  const vis = us.filter(([e]) => !FILTRO || e.includes(FILTRO.toLowerCase()));
  const todos = [...MODULOS, CENTRAL];
  let h = `<div class="cab"><div><h1>Central de acessos</h1><p>Um cadastro por pessoa: marque os módulos e o nível. Ao salvar, o menu e as pastas do SharePoint de cada módulo são liberados ou retirados.</p></div>
    <div class="acoes"><button class="btn" id="c-importar" type="button">Importar listas dos apps</button><button class="btn" id="c-sync" type="button" ${us.length ? '' : 'disabled'}>Sincronizar pastas do SharePoint</button><button class="btn" id="c-csv" type="button" ${us.length ? '' : 'disabled'}>Exportar (CSV)</button></div></div>`;
  if (!ACESSOS) h += `<div class="alerta" style="margin-bottom:16px"><b>Primeiro passo:</b> clique em <b>Importar listas dos apps</b>. Ele lê as listas de usuários dos apps (Assistência, Peças, Compras, Obras, Orçamentação, CRM SAM, FollowUP, Treinamentos e Painel) e cria a Central com todo mundo que já tinha acesso. Nenhuma pasta é alterada até você sincronizar.</div>`;
  if (PROGRESSO) h += `<div class="cartao progresso" style="margin-bottom:16px">${esc(PROGRESSO)}</div>`;
  const perfis = (ACESSOS && ACESSOS.perfis) || [];
  h += `<div class="cartao"><h2 style="font-size:21px">${FORM.editando ? 'Editar acesso' : 'Adicionar pessoa'}</h2>
    <div class="form-linha" style="margin-top:12px">
      <label class="campo" style="flex:2 1 320px">E-mail<input type="email" id="f-email" placeholder="nome.sobrenome@mcpoirrigacao.com.br" value="${esc(FORM.email)}" ${FORM.editando ? 'readonly' : ''}></label>
      <label class="campo">Aplicar perfil<select id="f-perfil"><option value="">—</option>${perfis.map((p, i) => `<option value="${i}">${esc(p.nome)}</option>`).join('')}</select></label>
      <label class="campo">Copiar acessos de<select id="f-copiar"><option value="">—</option>${us.map(([e]) => `<option value="${esc(e)}">${esc(e)}</option>`).join('')}</select></label>
    </div>
    <div class="mods">${todos.map(m => { const v = FORM.mods[m.id] || ''; return `<div class="mod-op ${v ? 'on' : ''}"><label><input type="checkbox" data-mod="${m.id}" ${v ? 'checked' : ''}>${esc(m.nome)}</label><select data-niv="${m.id}" aria-label="Nível em ${esc(m.nome)}" ${v ? '' : 'disabled'}>${m.niveis.map(n => `<option value="${n}" ${v === n ? 'selected' : ''}>${esc(rotulo(m, n))}</option>`).join('')}</select></div>`; }).join('')}</div>
    <div class="form-linha" style="margin-top:14px;justify-content:space-between"><span class="nota">FollowUP “Equipe”: o que a pessoa preenche ou vê continua definido nas equipes dentro do FollowUP.</span>
      <div class="acoes"><button class="btn" id="f-perfil-salvar" type="button">Salvar marcações como perfil</button>${FORM.editando ? '<button class="btn" id="f-cancelar" type="button">Cancelar</button>' : ''}<button class="btn primario" id="f-salvar" type="button">Salvar acesso</button></div></div></div>`;
  h += `<div class="secao"><h2>CRM Comercial · senha da planilha</h2><div class="cartao"><p class="nota" style="margin:0 0 10px">Fica guardada no SharePoint (pasta “${esc(CFG.pastaCrm)}”), com leitura só para quem tem o CRM Comercial liberado. O módulo usa essa senha para carregar a planilha sozinho.</p>
    <div class="form-linha"><label class="campo" style="flex:1 1 280px">Senha da planilha<input type="password" id="crm-senha" autocomplete="new-password" placeholder="${CRM_SENHA_OK ? '•••••••• (já cadastrada — digite para trocar)' : 'digite a senha da planilha'}"></label><button class="btn primario" id="crm-salvar" type="button">Salvar senha</button></div></div></div>`;
  h += `<div class="secao"><div class="form-linha" style="justify-content:space-between"><h2>Quem tem acesso (${us.length})</h2><label class="campo" style="flex:0 1 280px"><span class="sr">Filtrar</span><input type="search" id="c-filtro" placeholder="Filtrar por e-mail" value="${esc(FILTRO)}"></label></div>
    <div class="cartao tabela-box"><table class="t"><thead><tr><th>E-mail</th>${todos.map(m => `<th title="${esc(m.nome)}">${esc(m.curto || m.nome)}</th>`).join('')}<th><span class="sr">Ações</span></th></tr></thead><tbody>
    ${vis.length ? vis.map(([e, u]) => `<tr><td><b>${esc(e)}</b></td>${todos.map(m => { const n = (u.modulos || {})[m.id]; return n ? `<td class="${n === 'admin' ? 'adm' : ''}">${esc(ROTULO[n] || n)}</td>` : '<td class="vazio">—</td>'; }).join('')}<td style="white-space:nowrap"><button class="btn pequeno" data-editar="${esc(e)}" type="button">editar</button> <button class="btn pequeno perigo" data-remover="${esc(e)}" type="button">remover</button></td></tr>`).join('') : `<tr><td colspan="${todos.length + 2}" class="vazio">Ninguém cadastrado.</td></tr>`}
    </tbody></table></div></div>`;
  h += htmlArmazenamento();
  const hist = (ACESSOS && ACESSOS.historico) || [];
  h += `<div class="secao"><h2>Histórico</h2><div class="cartao">${hist.length ? hist.slice(0, 40).map(x => `<div class="hist">${esc(new Date(x.em).toLocaleString('pt-BR'))} · <b>${esc(x.por)}</b> ${esc(x.txt)}</div>`).join('') : '<span class="nota">Sem alterações ainda.</span>'}</div></div>`;
  el.innerHTML = h;

  el.querySelectorAll('[data-mod]').forEach(cb => cb.onchange = () => { const id = cb.dataset.mod; const sel = el.querySelector(`[data-niv="${id}"]`); const m = id === 'central' ? CENTRAL : MOD[id]; if (cb.checked) FORM.mods[id] = m.niveis[m.niveis.length - 1]; else delete FORM.mods[id]; FORM.email = $('f-email').value; renderCentral(); });
  el.querySelectorAll('[data-niv]').forEach(s => s.onchange = () => { FORM.mods[s.dataset.niv] = s.value; });
  $('f-perfil').onchange = e => { const p = perfis[+e.target.value]; if (p) { FORM.email = $('f-email').value; FORM.mods = Object.assign({}, p.modulos); renderCentral(); } };
  $('f-copiar').onchange = e => { const u = ACESSOS && ACESSOS.usuarios[e.target.value]; if (u) { FORM.email = $('f-email').value; FORM.mods = Object.assign({}, u.modulos); renderCentral(); } };
  $('f-salvar').onclick = salvarForm;
  if ($('f-cancelar')) $('f-cancelar').onclick = () => { FORM = { email: '', mods: {}, editando: false }; renderCentral(); };
  $('f-perfil-salvar').onclick = salvarPerfil;
  $('c-filtro').oninput = e => { FILTRO = e.target.value.trim(); const pos = e.target.selectionStart; renderCentral(); const f = $('c-filtro'); f.focus(); f.setSelectionRange(pos, pos); };
  el.querySelectorAll('[data-editar]').forEach(b => b.onclick = () => { const e = b.dataset.editar; FORM = { email: e, mods: Object.assign({}, ACESSOS.usuarios[e].modulos), editando: true }; renderCentral(); window.scrollTo({ top: 0, behavior: 'smooth' }); });
  el.querySelectorAll('[data-remover]').forEach(b => b.onclick = () => removerPessoa(b.dataset.remover));
  $('c-importar').onclick = importarListas;
  $('a-medir').onclick = async () => { const p = medirArmazenamento(); renderCentral(); await p; renderCentral(); };
  $('a-backup').onclick = async () => { const b = $('a-backup'); b.disabled = true; b.textContent = 'Copiando…'; try { await backupMensal(true); toast('Backup do mês feito.'); } catch (e) { toast('Backup: ' + e.message); } renderCentral(); };
  $('c-sync').onclick = () => sincronizar({ forcar: true, explicito: true });
  $('c-csv').onclick = exportarCsv;
  $('crm-salvar').onclick = salvarSenhaCrm;
  if (CRM_SENHA_OK === null) { CRM_SENHA_OK = false; (async () => { try { const r = await lerJson(await urlCaminho(CFG.pastaCrm + '/crm_config.json')); CRM_SENHA_OK = !!(r.ok && r.data && r.data.senhaPlanilha); if (CRM_SENHA_OK && !$('tela-central').hidden) renderCentral(); } catch (e) {} })(); }
}
let CRM_SENHA_OK = null;
async function salvarSenhaCrm() {
  const s = $('crm-senha').value; if (!s) { toast('Digite a senha.'); return; }
  try {
    await garantirPasta(CFG.pastaCrm);
    const url = await urlCaminho(CFG.pastaCrm + '/crm_config.json');
    const atual = await lerJson(url);
    const w = await gravarJson(url, { senhaPlanilha: s, atualizadoEm: new Date().toISOString(), atualizadoPor: emailAtual() }, atual.ok && !atual.missing ? atual.etag : null);
    if (!w.ok) { toast('Não foi possível salvar a senha (HTTP ' + w.status + ').'); return; }
    await salvarAcessos(() => [{ txt: 'atualizou a senha da planilha do CRM Comercial' }]);
    CRM_SENHA_OK = true; toast('Senha salva. Atualizando quem pode ler…'); renderCentral(); sincronizar({});
  } catch (e) { toast('Erro: ' + e.message); }
}
const resumo = mods => Object.entries(mods).map(([id, n]) => ((MOD[id] || CENTRAL).nome) + ': ' + rotulo(MOD[id] || CENTRAL, n)).join(', ') || 'nenhum módulo';
async function salvarForm() {
  const email = $('f-email').value.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { toast('Informe um e-mail válido.'); return; }
  const mods = Object.assign({}, FORM.mods);
  if (email === emailAtual() && !mods.central && !funcoes().includes('MCPO.Admin') && !confirm('Você vai tirar o seu próprio acesso à Central. Continuar?')) return;
  const ok = await salvarAcessos(a => {
    const antes = a.usuarios[email] ? resumo(a.usuarios[email].modulos) : null;
    if (!Object.keys(mods).length) { if (!a.usuarios[email]) return false; delete a.usuarios[email]; return [{ txt: `removeu ${email}` }]; }
    a.usuarios[email] = { modulos: mods, atualizadoEm: new Date().toISOString(), atualizadoPor: emailAtual() };
    return [{ txt: antes ? `alterou ${email}: ${resumo(mods)} (antes: ${antes})` : `adicionou ${email}: ${resumo(mods)}` }];
  });
  if (!ok) return;
  toast('Acesso salvo. Atualizando as pastas do SharePoint…');
  FORM = { email: '', mods: {}, editando: false }; renderCentral(); renderMenu('central');
  sincronizar({});
}
async function removerPessoa(email) {
  if (!confirm(`Remover todos os acessos de ${email}? As permissões que o MCPO 360 deu nas pastas do SharePoint também serão retiradas.`)) return;
  const ok = await salvarAcessos(a => { if (!a.usuarios[email]) return false; delete a.usuarios[email]; return [{ txt: `removeu ${email}` }]; });
  if (ok) { renderCentral(); sincronizar({}); }
}
async function salvarPerfil() {
  if (!Object.keys(FORM.mods).length) { toast('Marque pelo menos um módulo.'); return; }
  const nome = (prompt('Nome do perfil (ex.: Técnico de assistência):') || '').trim(); if (!nome) return;
  const mods = Object.assign({}, FORM.mods); FORM.email = $('f-email').value;
  const ok = await salvarAcessos(a => { a.perfis = (a.perfis || []).filter(p => p.nome !== nome); a.perfis.push({ nome, modulos: mods }); return [{ txt: `salvou o perfil "${nome}": ${resumo(mods)}` }]; });
  if (ok) { toast('Perfil salvo.'); renderCentral(); }
}
function exportarCsv() {
  const todos = [...MODULOS, CENTRAL];
  const linhas = [['E-mail', ...todos.map(m => m.nome)]].concat(Object.entries(ACESSOS.usuarios).sort().map(([e, u]) => [e, ...todos.map(m => (u.modulos || {})[m.id] ? rotulo(m, u.modulos[m.id]) : '')]));
  const csv = '\ufeff' + linhas.map(r => r.map(v => /[;"\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v).join(';')).join('\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })); a.download = `mcpo360_acessos_${hoje()}.csv`; a.click();
}

/* ---------------- importar as listas antigas dos apps ---------------- */
async function lerArquivoModulo(pastaRel, arquivo) {
  let url;
  if (pastaRel === '@crmsam') { const p = await urlPasta('crmsam'); if (!p) return null; url = filho(p, arquivo); }
  else url = await urlCaminho(pastaRel + '/' + arquivo);
  const r = await lerJson(url); return r.ok && !r.missing ? r.data : null;
}
async function importarListas() {
  if (!confirm('Ler as listas de usuários de cada app e juntar na Central? Quem já está na Central mantém o nível mais alto entre os dois. Nenhuma pasta do SharePoint é alterada agora.')) return;
  const novo = []; const log = [];
  const add = (e, mod, n) => novo.push([e, mod, n]);
  const L = a => (a || []).map(x => String(x).trim().toLowerCase()).filter(Boolean);
  PROGRESSO = 'Lendo as listas dos apps…'; renderCentral();
  for (const m of MODULOS) {
    if (!m.acesso) continue;
    let d = null;
    for (const [p, f] of m.acesso) { try { d = await lerArquivoModulo(p, f); } catch (e) {} if (d) break; }
    if (!d) { log.push(`${m.nome}: lista não encontrada`); continue; }
    const antes = novo.length;
    L(d.emailsAutorizados).forEach(e => add(e, m.id, 'admin'));
    if (m.id === 'orcamento') L(d.emailsGestores).forEach(e => add(e, m.id, 'gestor'));
    if (m.id === 'crmsam') L(d.emailsEditores).forEach(e => add(e, m.id, 'editor'));
    if (m.id === 'obras') L(d.emailsAgenda).forEach(e => add(e, m.id, 'agenda'));
    L(d.emailsComuns).forEach(e => add(e, m.id, m.id === 'followup' ? 'visualizador' : 'comum'));
    log.push(`${m.nome}: ${novo.length - antes} acesso(s)`);
  }
  try {   // FollowUP: quem preenche ou vê equipes
    const d = await lerArquivoModulo('FollowUP', 'followup_dados.json'); let n = 0;
    ((d && d.config && d.config.equipes) || []).forEach(eq => [...L(eq.membros), ...L(eq.visualizadores)].forEach(e => { add(e, 'followup', 'equipe'); n++; }));
    log.push(`FollowUP (equipes): ${n} acesso(s)`);
  } catch (e) {}
  try {   // Painel 360
    const p = await urlPasta('painel'); const r = p ? await lerJson(filho(p, 'painel_acesso.json')) : null; const d = r && r.ok && !r.missing ? r.data : null;
    if (d) { L(d.emailsAutorizados).forEach(e => add(e, 'painel', 'admin')); L(d.emailsComuns).forEach(e => add(e, 'painel', 'comum')); log.push(`Painel 360: ${L(d.emailsAutorizados).length + L(d.emailsComuns).length} acesso(s)`); }
    else log.push('Painel 360: lista não encontrada');
  } catch (e) {}
  const eu = emailAtual();
  const ok = await salvarAcessos(a => {
    let n = 0; novo.forEach(([e, m, niv]) => { if (mesclarImportado(a, e, m, niv)) n++; });
    if (funcoes().includes('MCPO.Admin') || !Object.values(a.usuarios).some(u => (u.modulos || {}).central === 'admin')) { if (mesclarImportado(a, eu, 'central', 'admin')) n++; }
    return [{ txt: `importou as listas dos apps (${n} acesso(s) novos ou elevados)` }];
  });
  PROGRESSO = (ok ? 'Importação concluída.\n' : 'A importação não foi salva.\n') + log.join('\n') + (ok ? '\n\nConfira a lista abaixo e depois clique em “Sincronizar pastas do SharePoint”.' : '');
  recalcular(); renderMenu('central'); renderCentral();
}

/* ---------------- permissões nas pastas ---------------- */
async function erroTxt(r) { let m = ''; try { m = (await r.json()).error.message; } catch (e) {} return r.status === 403 ? 'sem permissão para compartilhar a pasta' : (m || 'erro ' + r.status); }
async function conceder(url, email, papel) {
  try {
    const r = await g(acao(url, 'invite'), { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ recipients: [{ email }], requireSignIn: true, sendInvitation: false, roles: [papel === 'read' ? 'read' : 'write'] }) });
    return r.ok ? { ok: true } : { ok: false, msg: await erroTxt(r) };
  } catch (e) { return { ok: false, msg: 'falha de rede' }; }
}
function temEmail(p, em) {
  const ids = [p.grantedToV2, p.grantedTo, ...(p.grantedToIdentitiesV2 || []), ...(p.grantedToIdentities || [])].filter(Boolean);
  return ids.some(x => { const u = x.user || x.siteUser || {}; return [u.email, u.loginName, u.userPrincipalName].some(v => v && String(v).toLowerCase().includes(em)); });
}
async function retirar(url, email) {
  const em = email.toLowerCase();
  try {
    const r = await g(acao(url, 'permissions')); if (!r.ok) return { ok: false, msg: await erroTxt(r) };
    const perms = ((await r.json()).value || []).filter(p => !p.inheritedFrom && !(p.roles || []).includes('owner') && temEmail(p, em));
    for (const p of perms) {
      const ids = [...(p.grantedToIdentitiesV2 || []), ...(p.grantedToIdentities || [])];
      const d = (p.link && ids.length > 1)
        ? await g(acao(url, 'permissions/' + p.id + '/revokeGrants'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ grantees: [{ email }] }) })
        : await g(acao(url, 'permissions/' + p.id), { method: 'DELETE' });
      if (!d.ok) return { ok: false, msg: await erroTxt(d) };
    }
    return { ok: true };
  } catch (e) { return { ok: false, msg: 'falha de rede' }; }
}
let _fila = Promise.resolve();
function sincronizar(o) { _fila = _fila.then(() => _sincronizar(o || {})).catch(e => { toast('Sincronização: ' + e.message); sync(''); }); return _fila; }
async function _sincronizar({ forcar, explicito }) {
  if (!ehAdminCentral() || !ACESSOS || !Object.keys(ACESSOS.usuarios).length) return;
  await carregarPreenchedoresFU();
  const quer = calcularDesejado(ACESSOS);
  const chaves = [...MODULOS.filter(m => m.pasta).map(m => m.id), 'mcpo360', 'uso', 'crmcfg'];
  let lib = 0, ret = 0; const falhas = [], pulos = [], linhas = [];
  for (const chave of chaves) {
    const nome = chave === 'mcpo360' ? 'MCPO360' : chave === 'uso' ? 'MCPO360/uso' : chave === 'crmcfg' ? CFG.pastaCrm : MOD[chave].nome;
    sync('Sincronizando ' + nome + '…');
    let url; try { url = await urlPasta(chave); } catch (e) { url = null; }
    if (!url) { falhas.push(nome + ': pasta não encontrada'); continue; }
    const regUrl = filho(url, REG);
    let reg = { pastas: {} }; let regEtag = null;
    try { const r = await lerJson(regUrl); if (r.ok && r.data && r.data.pastas) { reg = r.data; regEtag = r.etag; } } catch (e) {}
    const tem = (reg.pastas.propria = reg.pastas.propria || {});
    const plano = planejarPasta(tem, quer[chave], forcar);
    let retirarLista = plano.retirar;
    const total = Object.keys(tem).length;
    if (retirarLista.length > 3 && retirarLista.length > total / 2) {
      if (!explicito || !confirm(`${nome}: a sincronização vai retirar o acesso de ${retirarLista.length} de ${total} pessoa(s) que o app tinha liberado:\n\n${retirarLista.slice(0, 25).join('\n')}${retirarLista.length > 25 ? '\n…' : ''}\n\nConfirmar a retirada?`)) { pulos.push(`${nome}: ${retirarLista.length} retirada(s) não feitas (confirme em Sincronizar)`); retirarLista = []; }
    }
    let mudou = false, l = 0, rt = 0;
    for (const [em, papel, antes] of plano.liberar) {
      if (antes && antes !== papel) await retirar(url, em);
      const r = await conceder(url, em, papel);
      if (r.ok) { if (tem[em] !== papel) { l++; } tem[em] = papel; mudou = true; } else falhas.push(`${nome} · ${em}: ${r.msg}`);
    }
    for (const em of retirarLista) {
      const r = await retirar(url, em);
      if (r.ok) { delete tem[em]; rt++; mudou = true; } else falhas.push(`${nome} · ${em}: ${r.msg}`);
    }
    /* arquivos que quem só lê a pasta também grava (ex.: resumo processado do Compras-estoque) */
    for (const arq of (MOD[chave] && MOD[chave].arquivosEditaveis) || []) {
      const fUrl = filho(url, arq);
      try { const ex = await g(fUrl); if (!ex.ok) { if (ex.status !== 404) falhas.push(`${nome} › ${arq}: HTTP ${ex.status}`); continue; } } catch (e) { continue; }
      const querArq = Object.fromEntries(Object.entries(quer[chave] || {}).filter(([, p]) => p === 'read').map(([em]) => [em, 'write']));
      const temArq = (reg.pastas['arq:' + arq] = reg.pastas['arq:' + arq] || {});
      const pa = planejarPasta(temArq, querArq, forcar);
      for (const [em, papel] of pa.liberar) { const r = await conceder(fUrl, em, papel); if (r.ok) { if (temArq[em] !== papel) l++; temArq[em] = papel; mudou = true; } else falhas.push(`${nome} › ${arq} · ${em}: ${r.msg}`); }
      for (const em of pa.retirar) { const r = await retirar(fUrl, em); if (r.ok) { delete temArq[em]; rt++; mudou = true; } else falhas.push(`${nome} › ${arq} · ${em}: ${r.msg}`); }
    }
    if (mudou) { reg.atualizadoEm = new Date().toISOString(); reg.atualizadoPor = emailAtual(); reg.origem = 'MCPO 360'; await gravarJson(regUrl, reg, regEtag); }
    lib += l; ret += rt; if (l || rt) linhas.push(`${nome}: ${l} liberado(s), ${rt} retirado(s)`);
  }
  sync('');
  const txt = (lib || ret ? `Pastas atualizadas: ${lib} liberado(s), ${ret} retirado(s).` : 'Pastas conferidas: nada a mudar.')
    + (linhas.length ? '\n' + linhas.join('\n') : '') + (pulos.length ? '\n\n' + pulos.join('\n') : '') + (falhas.length ? '\n\nNão foi possível:\n' + falhas.slice(0, 30).join('\n') : '');
  if (explicito || falhas.length || pulos.length) { PROGRESSO = txt; if (!$('tela-central').hidden) renderCentral(); }
  toast(txt.split('\n')[0] + (falhas.length ? ` (${falhas.length} falha(s) — veja a Central)` : ''), 5000);
}

/* ---------------- uso ---------------- */
let USO = null, _usoTimer = null;
const arqUso = () => CFG.pasta + '/uso/' + emailAtual().replace(/[^a-z0-9._-]/g, '_') + '.json';
async function registrarUso(id) {
  try {
    if (!USO) { const r = await lerJson(await urlCaminho(arqUso())); USO = (r.ok && r.data) ? r.data : { email: emailAtual(), dias: {} }; USO._etag = r.etag || null; }
    const d = hoje(); USO.dias[d] = USO.dias[d] || {}; USO.dias[d][id] = (USO.dias[d][id] || 0) + 1; USO.ultimo = new Date().toISOString();
    clearTimeout(_usoTimer); _usoTimer = setTimeout(gravarUso, 1200);
  } catch (e) {}
}
async function gravarUso() {
  try {
    const corte = new Date(Date.now() - 120 * 864e5).toISOString().slice(0, 10);
    Object.keys(USO.dias).forEach(k => { if (k < corte) delete USO.dias[k]; });
    const { _etag, ...dados } = USO;
    const u = await urlCaminho(arqUso());
    const h = { 'Content-Type': 'application/json' };
    const r = await g(acao(u, 'content'), { method: 'PUT', headers: h, body: JSON.stringify(dados) });
    if (r.ok) { try { USO._etag = (await r.json()).eTag; } catch (e) {} }
  } catch (e) {}
}
async function renderUso() {
  const el = $('tela-uso');
  el.innerHTML = `<div class="cab"><div><h1>Uso dos módulos</h1><p>Aberturas de cada módulo nos últimos 30 dias, a partir do registro que o MCPO 360 grava em MCPO360/uso.</p></div></div><div class="cartao">Lendo…</div>`;
  try {
    const pasta = await urlPasta('uso'); const r = await g(acao(pasta, 'children') + '?$top=999'); if (!r.ok) throw new Error('HTTP ' + r.status);
    const itens = ((await r.json()).value || []).filter(x => x.name.endsWith('.json') && x['@microsoft.graph.downloadUrl']);
    const dados = []; for (let i = 0; i < itens.length; i += 6) { const parte = await Promise.all(itens.slice(i, i + 6).map(x => fetch(x['@microsoft.graph.downloadUrl']).then(r => r.json()).catch(() => null))); dados.push(...parte.filter(Boolean)); }
    const corte = new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10);
    const porMod = {}, pessoas = [];
    dados.forEach(u => { let tot = 0; const mods = {}; Object.entries(u.dias || {}).forEach(([d, x]) => { if (d < corte) return; Object.entries(x).forEach(([m, n]) => { porMod[m] = porMod[m] || { n: 0, p: new Set() }; porMod[m].n += n; porMod[m].p.add(u.email); mods[m] = (mods[m] || 0) + n; tot += n; }); }); pessoas.push({ email: u.email, tot, ultimo: u.ultimo, mods }); });
    const max = Math.max(1, ...Object.values(porMod).map(x => x.n));
    const cad = ACESSOS ? Object.keys(ACESSOS.usuarios) : [];
    const semUso = cad.filter(e => !pessoas.some(p => p.email === e && p.tot));
    el.innerHTML = `<div class="cab"><div><h1>Uso dos módulos</h1><p>Últimos 30 dias · ${pessoas.filter(p => p.tot).length} pessoa(s) usaram o MCPO 360.</p></div></div>
      <div class="cartao tabela-box"><table class="t"><thead><tr><th>Módulo</th><th>Aberturas</th><th>Pessoas</th><th style="width:40%"><span class="sr">Gráfico</span></th></tr></thead><tbody>
      ${MODULOS.map(m => { const x = porMod[m.id] || { n: 0, p: new Set() }; return `<tr><td>${esc(m.nome)}</td><td>${x.n}</td><td>${x.p.size}</td><td><div class="barra" style="width:${(x.n / max * 100).toFixed(1)}%"></div></td></tr>`; }).join('')}
      </tbody></table></div>
      <div class="secao"><h2>Por pessoa</h2><div class="cartao tabela-box"><table class="t"><thead><tr><th>E-mail</th><th>Aberturas</th><th>Último acesso</th><th>Mais usados</th></tr></thead><tbody>
      ${pessoas.sort((a, b) => b.tot - a.tot).map(p => `<tr><td>${esc(p.email)}</td><td>${p.tot}</td><td>${p.ultimo ? esc(new Date(p.ultimo).toLocaleString('pt-BR')) : '—'}</td><td>${esc(Object.entries(p.mods).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([m, n]) => (MOD[m] ? MOD[m].nome : m) + ' (' + n + ')').join(', '))}</td></tr>`).join('') || '<tr><td colspan="4" class="vazio">Sem registros ainda.</td></tr>'}
      </tbody></table></div></div>
      ${semUso.length ? `<div class="secao"><h2>Cadastrados sem uso nos últimos 30 dias (${semUso.length})</h2><div class="cartao nota">${semUso.map(esc).join(', ')}</div></div>` : ''}`;
  } catch (e) { el.innerHTML = `<div class="cab"><div><h1>Uso dos módulos</h1></div></div><div class="alerta">Não foi possível ler os registros de uso: ${esc(e.message)}. Rode “Sincronizar pastas do SharePoint” na Central para criar a pasta MCPO360/uso.</div>`; }
}

/* ---------------- API para os módulos ---------------- */
window.MCPO360 = {
  versao: 1,
  nivel: id => NIVEIS[id] || null,
  email: () => emailAtual(),
  /* Entrega Técnica: e-mails liberados num módulo na Central ([{ email, nivel }]); null se a Central não pôde ser lida */
  usuarios: mod => { if (!ACESSOS || !ACESSOS.usuarios || !MOD[mod]) return null; return Object.entries(ACESSOS.usuarios).filter(([, u]) => u && u.modulos && u.modulos[mod]).map(([email, u]) => ({ email, nivel: u.modulos[mod] })); },
  abrir: id => { location.hash = '#' + id; },
  /* CRM Comercial: senha da planilha guardada no SharePoint (só quem tem o módulo lê) */
  senhaCrm: async () => { if (!NIVEIS.crmcomercial) return ''; try { const r = await lerJson(await urlCaminho(CFG.pastaCrm + '/crm_config.json')); return (r.ok && r.data && r.data.senhaPlanilha) || ''; } catch (e) { return ''; } },
  /* FollowUP: quem entra numa equipe precisa de acesso ao módulo. */
  garantirAcesso: async (email, mod, nivel) => {
    email = String(email || '').trim().toLowerCase();
    if (!email || !MOD[mod]) return;
    const atual = ACESSOS && ACESSOS.usuarios[email] && ACESSOS.usuarios[email].modulos[mod];
    if (atual) return;
    if (!ehAdminCentral()) { toast(`Peça a um admin da Central para liberar ${email} no ${MOD[mod].nome}.`, 5000); return; }
    const ok = await salvarAcessos(a => { if (!mesclarImportado(a, email, mod, nivel)) return false; return [{ txt: `liberou ${email} no ${MOD[mod].nome} (${rotulo(MOD[mod], nivel)}) pelo módulo` }]; });
    if (ok) { toast(`${email} liberado no ${MOD[mod].nome}.`); sincronizar({}); }
  },
};

function limparInteracao() {
  try { [localStorage, sessionStorage].forEach(s => Object.keys(s).filter(k => /interaction\.status|interaction_status/i.test(k)).forEach(k => s.removeItem(k))); } catch (e) {}
}
let _pedindo = false;
function pedirLoginDeNovo() {
  if (_pedindo) return; _pedindo = true;
  $('login-erro').textContent = 'Sua sessão precisa ser confirmada. Clique em Entrar.';
  mostrarTela('login');
}
/* ---------------- início ---------------- */
async function entrar() {
  $('login-erro').textContent = '';
  try {
    const r = await msal.loginPopup({ scopes: CFG.scopes, redirectUri: REDIRECT, prompt: 'select_account' });
    conta = r.account || (msal.getAllAccounts()[0] || null); msal.setActiveAccount(conta);
    _pedindo = false; $('login-erro').textContent = '';
    await depoisDoLogin();
  } catch (e) {
    const cod = (e && (e.errorCode || e.message)) || String(e);
    if (/interaction_in_progress/.test(cod)) { limparInteracao(); $('login-erro').textContent = 'Havia um login pela metade. Clique em Entrar de novo.'; return; }
    if (/popup_window_error|empty_window_error/.test(cod)) { $('login-erro').textContent = 'O navegador bloqueou a janela de login. Libere pop-ups para mcpoirrigacao.github.io e clique em Entrar de novo.'; return; }
    if (/user_cancelled/.test(cod)) { $('login-erro').textContent = 'Login cancelado.'; return; }
    $('login-erro').textContent = 'Não foi possível entrar: ' + (e.errorMessage || e.message || e);
  }
}
async function depoisDoLogin() {
  mostrarTela('carregando');
  $('quem').hidden = false; $('quem-email').textContent = emailAtual();
  await carregarAcessos(); recalcular();
  const f = funcoes(); $('quem-papel').textContent = f.includes('MCPO.Admin') ? 'Admin geral · Azure' : ehAdminCentral() ? 'Admin da Central' : meusModulos().length + ' módulo(s)';
  rota();
  if (ehAdminCentral()) setTimeout(() => rotinaArmazenamento(), 6000);
}
async function iniciar() {
  // janela de login (popup) voltando para cá: só deixa a janela principal ler a resposta
  if (window.opener && window.opener !== window && /[#&?](code|error|state)=/.test(location.hash + location.search)) { document.body.innerHTML = '<p style="padding:24px;color:#5f6e62;font-family:system-ui">Concluindo o login…</p>'; return; }
  $('btn-entrar').onclick = entrar;
  $('btn-reentrar').onclick = entrar;
  $('btn-sair').onclick = () => msal.logoutPopup({ account: conta, postLogoutRedirectUri: REDIRECT, mainWindowRedirectUri: BASE }).catch(() => { localStorage.clear(); location.reload(); });
  $('btn-menu').onclick = () => { if (window.matchMedia('(max-width:860px)').matches) $('menu').classList.toggle('aberto'); else alternarMenu(); };
  try { if (localStorage.getItem('mcpo360_menu') === 'fechado') document.body.classList.add('menu-fechado'); } catch (e) {}
  // menu recolhido: clicar numa área vazia da coluna abre de novo
  $('menu').addEventListener('click', e => { if (document.body.classList.contains('menu-fechado') && !e.target.closest('a,button')) alternarMenu(false); });
  window.addEventListener('hashchange', rota);
  limparInteracao();
  try { if (!window.msal) throw new Error('biblioteca de login não carregou'); await iniciarMsal(); }
  catch (e) { mostrarTela('login'); $('login-erro').textContent = 'Login indisponível: ' + e.message; return; }
  if (!conta) { mostrarTela('login'); return; }
  await depoisDoLogin();
}
iniciar();
})();
