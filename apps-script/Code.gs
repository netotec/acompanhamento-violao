/**
 * Violão – Acompanhamento de exercícios (Google Apps Script)
 * A aluna abre o link (sem login), marca os exercícios e envia.
 * Tudo é gravado na aba "Registros" desta planilha, com o nome dela.
 */

// Configuração opcional. Deixe em branco: o sistema descobre sozinho.
// 1) ID da planilha (o trecho entre /d/ e /edit). Em branco = usa a planilha onde o script está.
const PLANILHA_ID = '';

const ABA = { alunos: 'Alunos', aulas: 'Aulas', ex: 'Exercicios', envios: 'Envios', reg: 'Registros' };
const FUSO = 'America/Sao_Paulo';
// Quantos aparelhos cada aluno pode usar (ex.: celular + computador)
const LIMITE_APARELHOS = 2;

function onOpen() {
  try { // guarda o ID desta planilha para o web app encontrar
    const a = SpreadsheetApp.getActiveSpreadsheet();
    if (a) PropertiesService.getScriptProperties().setProperty('PLANILHA_ID', a.getId());
  } catch (e) {}
  SpreadsheetApp.getUi()
    .createMenu('Violão')
    .addItem('Abrir painel do professor', 'abrirPainel')
    .addItem('Painel no celular ou computador', 'linkPainelCelular')
    .addItem('Definir senha do painel', 'definirSenhaPainel')
    .addSeparator()
    .addItem('Gerar links pendentes', 'gerarLinks')
    .addItem('Ligar apostilas (PDF) às aulas', 'ligarApostilas')
    .addItem('Criar aulas 11 a 30 (apostilas)', 'criarAulas11a30')
    .addSeparator()
    .addItem('Registrar ausência do professor (repor aulas)', 'registrarAusencia')
    .addItem('Processar faltas agora', 'processarFaltas')
    .addItem('Ativar automação (AppSheet)', 'ativarAutomacao')
    .addSeparator()
    .addItem('Definir endereço do app (GitHub Pages)', 'definirEnderecoApp')
    .addItem('Desconectar todos os aparelhos do painel', 'desconectarAparelhosPainel')
    .addItem('Liberar tentativas de senha do painel', 'liberarTentativasPainel')
    .addToUi();
}

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = PLANILHA_ID || props.getProperty('PLANILHA_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) { props.setProperty('PLANILHA_ID', ativa.getId()); return ativa; }
  throw new Error('Planilha não encontrada. Abra a planilha e rode o Configurar uma vez.');
}

/** Executa fn com a trava do script (evita duas gravações ao mesmo tempo). */
function comTrava_(fn, esperaMs) {
  let trava = null;
  try { trava = LockService.getScriptLock(); } catch (e) { return fn(); } // gatilho simples sem acesso à trava
  trava.waitLock(esperaMs || 15000);
  try { return fn(); } finally { trava.releaseLock(); }
}

/** Endereço padrão dos links: as páginas do GitHub Pages (abrem em qualquer conta ou navegador).
 *  Pode ser trocado pelo painel (Enviar link) ou pelo menu Violão > Definir endereço do app. */
const ENDERECO_PADRAO = 'https://netotec.github.io/acompanhamento-violao/';

/** Endereço que vai nos links enviados aos alunos. Nunca depende de URL do Google. */
function urlPublica_() {
  const app = PropertiesService.getScriptProperties().getProperty('URL_PUBLICA') || '';
  return /^https:\/\//.test(app) ? app.replace(/\/?$/, '/') : ENDERECO_PADRAO;
}

function definirEnderecoApp() {
  soDono_();
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Endereço do app',
    'Cole o endereço do GitHub Pages (ex.: https://seuusuario.github.io/violao-acompanhamento/).\nDeixe em branco para voltar ao endereço padrão (' + ENDERECO_PADRAO + ').', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  const props = PropertiesService.getScriptProperties();
  if (!url) { props.deleteProperty('URL_PUBLICA'); ui.alert('Os links voltam ao endereço padrão: ' + ENDERECO_PADRAO); return; }
  if (!/^https:\/\/[^\s]+$/.test(url)) { ui.alert('Endereço inválido: ele precisa começar com https://'); return; }
  props.setProperty('URL_PUBLICA', url);
  ui.alert('Pronto! Os próximos links já vão abrir pelo app.');
}

// Memória de leitura de UMA execução: evita reler a mesma aba várias vezes (contexto_ é chamado em sequência).
// Só vale dentro de comMemo_(); limparMemo_() depois de qualquer gravação nas abas lidas.
let MEMO_ = null;
function comMemo_(fn) {
  const anterior = MEMO_;
  MEMO_ = anterior || {};
  try { return fn(); } finally { MEMO_ = anterior; }
}
function limparMemo_() { if (MEMO_) MEMO_ = {}; }

function tabela_(nome) {
  let valores = MEMO_ && MEMO_[nome];
  if (!valores) {
    valores = ss_().getSheetByName(nome).getDataRange().getValues();
    if (MEMO_) MEMO_[nome] = valores;
  }
  valores = valores.slice();
  const cab = valores.shift();
  return valores
    .filter(l => l[0] !== '')
    .map(l => Object.fromEntries(cab.map((c, i) => [c, l[i]])));
}

function data_(d) {
  return d ? Utilities.formatDate(new Date(d), FUSO, 'dd/MM/yyyy') : '';
}

/** Menu Violão > Ligar apostilas: preenche Link_Apostila (aba Aulas) com o PDF da aula, no endereço do app
 *  (apostilas/aula-NN.pdf, NN = número da aula). Só preenche células vazias: links que você colocou ficam. */
function ligarApostilas() {
  soDono_();
  const preenchidas = comTrava_(() => {
    const sh = ss_().getSheetByName(ABA.aulas);
    const v = sh.getDataRange().getValues();
    const cab = v[0];
    const cNum = cab.indexOf('Numero'), cLink = cab.indexOf('Link_Apostila');
    if (cNum < 0 || cLink < 0) throw new Error('A aba Aulas precisa das colunas Numero e Link_Apostila.');
    const base = urlPublica_() + 'apostilas/aula-';
    let n = 0;
    for (let i = 1; i < v.length; i++) {
      const num = Number(v[i][cNum]);
      if (!num || num < 1 || num > 99 || String(v[i][cLink] || '').trim()) continue;
      sh.getRange(i + 1, cLink + 1).setValue(base + String(num).padStart(2, '0') + '.pdf');
      n++;
    }
    return n;
  });
  SpreadsheetApp.getUi().alert(preenchidas + ' aula(s) ligada(s) à apostila em PDF. Aulas que já tinham link não foram alteradas.');
}

/** Menu Violão > Gerar links pendentes: cria token, link e mensagem de WhatsApp
 *  para cada linha da aba Envios que ainda não tem token. */
function gerarLinks() {
  soDono_();
  const criados = comTrava_(gerarLinks_);
  SpreadsheetApp.getUi().alert(criados + ' link(s) gerado(s). Clique na coluna WhatsApp para enviar.');
}

function gerarLinks_() {
  const alunos = tabela_(ABA.alunos);
  const aulas = tabela_(ABA.aulas);
  const sh = ss_().getSheetByName(ABA.envios);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const c = n => cab.indexOf(n);
  let criados = 0;
  for (let i = 1; i < v.length; i++) {
    const l = v[i];
    if (!l[c('ID_Aluno')] || !l[c('ID_Aula')] || l[c('Token')]) continue;
    if (!l[c('Data_Envio')]) l[c('Data_Envio')] = new Date();
    const token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    const link = urlPublica_() + '?t=' + token;
    const aluno = alunos.find(a => a.ID_Aluno === l[c('ID_Aluno')]) || {};
    const aula = aulas.find(a => a.ID_Aula === l[c('ID_Aula')]) || {};
    const primeiroNome = String(aluno.Nome || '').split(' ')[0];
    const msg = 'Olá, ' + primeiroNome + '! Seus exercícios da Aula ' + (aula.Numero || '') +
      ' (' + (aula.Titulo || '') + '). Marque cada dia que praticar. Link válido até ' +
      data_(l[c('Validade')]) + ': ' + link;
    const fone = String(aluno.Telefone || '').replace(/\D/g, '');
    l[c('Token')] = token;
    l[c('Link')] = link;
    l[c('WhatsApp')] = fone ? 'https://wa.me/55' + fone + '?text=' + encodeURIComponent(msg) : '';
    criados++;
  }
  // grava só as colunas que este processo preenche (preserva fórmulas e formatos das outras)
  if (criados > 0 && v.length > 1) {
    ['Data_Envio', 'Token', 'Link', 'WhatsApp'].forEach(nome => {
      const col = c(nome);
      if (col < 0) return;
      sh.getRange(2, col + 1, v.length - 1, 1).setValues(v.slice(1).map(l => [l[col]]));
    });
  }
  return criados;
}

/** Monta tudo o que a página precisa a partir do token. */
function contexto_(token) {
  if (!token) return { erro: 'Link inválido. Peça um novo link ao professor.' };
  const env = tabela_(ABA.envios).find(r => String(r.Token) === token);
  if (!env) return { erro: 'Link inválido. Peça um novo link ao professor.' };

  const validade = new Date(env.Validade);
  validade.setHours(23, 59, 59);
  if (isNaN(validade) || new Date() > validade) {
    return { erro: 'Este link venceu em ' + data_(env.Validade) + '. Na próxima aula você recebe o link novo.' };
  }

  const aluno = tabela_(ABA.alunos).find(r => r.ID_Aluno === env.ID_Aluno) || {};
  const aula = tabela_(ABA.aulas).find(r => r.ID_Aula === env.ID_Aula) || {};
  const regs = tabela_(ABA.reg).filter(r => String(r.Token) === token);

  const vezes = {};
  regs.forEach(r => String(r.IDs_Exercicios).split(',').map(s => s.trim()).filter(Boolean)
    .forEach(id => { vezes[id] = (vezes[id] || 0) + 1; }));
  const dias = new Set(regs.map(r => data_(r.Data_Hora))).size;
  const minutos = regs.reduce((s, r) => s + (Number(r.Minutos) || 0), 0);

  const exercicios = tabela_(ABA.ex)
    .filter(r => r.ID_Aula === env.ID_Aula)
    .sort((a, b) => a.Ordem - b.Ordem)
    .map(r => ({ id: r.ID_Exercicio, titulo: r.Titulo, descricao: r.Descricao, tempo: r.Tempo_min, vezes: vezes[r.ID_Exercicio] || 0 }));

  return {
    token: token,
    idAluno: env.ID_Aluno,
    idAula: env.ID_Aula,
    nome: String(aluno.Nome || ''),
    aula: { numero: aula.Numero, titulo: aula.Titulo, tarefa: aula.Tarefa, apostila: aula.Link_Apostila },
    validade: data_(env.Validade),
    exercicios: exercicios,
    dias: dias,
    minutos: minutos,
    hoje: regs.some(r => data_(r.Data_Hora) === data_(new Date()))
  };
}

/** Como o professor viu a reação do aluno a cada exercício (aba Avaliacoes), para o aluno acompanhar a evolução.
 *  Só reações e datas: as observações do professor não vão para a página do aluno. */
const NOTA_REACAO_ = { 'Dominou': 4, 'Foi bem': 3, 'Com dificuldade': 2, 'Não conseguiu': 1 };

function evolucaoAluno_(idAluno) {
  if (!ss_().getSheetByName('Avaliacoes')) return [];
  const titulos = {};
  tabela_(ABA.ex).forEach(x => { titulos[x.ID_Exercicio] = String(x.Titulo || ''); });
  const por = {};
  tabela_('Avaliacoes')
    .filter(r => r.ID_Aluno === idAluno && r.Data instanceof Date && NOTA_REACAO_[r.Reacao])
    .sort((a, b) => a.Data - b.Data)
    .forEach(r => { (por[r.ID_Exercicio] = por[r.ID_Exercicio] || []).push(r); });
  return Object.keys(por).map(id => {
    const seq = por[id];
    const ultima = seq[seq.length - 1];
    const dif = NOTA_REACAO_[ultima.Reacao] - NOTA_REACAO_[seq[0].Reacao];
    return {
      id: id, titulo: titulos[id] || id, ultima: ultima.Reacao, quando: ultima.Data.getTime(),
      tendencia: seq.length < 2 ? '' : dif > 0 ? 'sobe' : dif < 0 ? 'desce' : 'igual',
      historico: seq.slice(-6).map(r => ({ data: Utilities.formatDate(r.Data, FUSO, 'dd/MM'), reacao: r.Reacao }))
    };
  }).sort((a, b) => b.quando - a.quando).slice(0, 20).map(x => { delete x.quando; return x; });
}

/* ---------- Acesso: cadastro no 1º link, depois só o nome ---------- */

function alunoCadastrado_(aluno) {
  return String(aluno.Cadastro_Completo || '').toLowerCase() === 'sim';
}

/** O que a página recebe ANTES de a aluna se identificar: nada de dados pessoais. */
function publico_(token) { return comMemo_(() => publico__(token)); }

function publico__(token) {
  const ctx = contexto_(token);
  if (ctx.erro) return { erro: ctx.erro };
  const aluno = tabela_(ABA.alunos).find(r => r.ID_Aluno === ctx.idAluno) || {};
  return {
    etapa: alunoCadastrado_(aluno) ? 'nome' : 'cadastro',
    token: token,
    aula: { numero: ctx.aula.numero, titulo: ctx.aula.titulo },
    validade: ctx.validade
  };
}

/* ---------- API para a página do GitHub Pages ----------
 * A página fica em github.io e fala com este web app por requisições SEM cookies (fetch), então o Google não
 * troca de conta nem pede login ("/u/3/", "página não encontrada"). Respostas em JSON com CORS liberado. */

/** Só estas funções podem ser chamadas pela API (as mesmas que as páginas já usavam via google.script.run). */
function funcoesApi_() {
  return {
    entrar: entrar, cadastrar: cadastrar, registrar: registrar,
    painelLogin: painelLogin, painelSair: painelSair, painelDados: painelDados, painelSalvarAluno: painelSalvarAluno,
    painelResetar: painelResetar, painelCriarEnvio: painelCriarEnvio, painelPresenca: painelPresenca,
    painelAgendar: painelAgendar, painelSalvarDisp: painelSalvarDisp, painelSalvarAula: painelSalvarAula,
    painelNovaAula: painelNovaAula, painelSalvarAvaliacao: painelSalvarAvaliacao, painelRelatorio: painelRelatorio,
    painelDefinirEndereco: painelDefinirEndereco, painelSalvarPlanos: painelSalvarPlanos,
    painelSalvarPix: painelSalvarPix, painelMarcarPagamento: painelMarcarPagamento,
    painelResumoExclusao: painelResumoExclusao, painelExcluirAluno: painelExcluirAluno
  };
}

function saidaJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** POST { fn, args } -> { ok: true, r } ou { ok: false, erro }. */
function doPost(e) {
  try {
    const corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(corpo.fn || '');
    const tabela = funcoesApi_();
    if (!Object.prototype.hasOwnProperty.call(tabela, fn)) throw new Error('Função não permitida.');
    const args = Array.isArray(corpo.args) ? corpo.args : [];
    return saidaJson_({ ok: true, r: tabela[fn].apply(null, args) });
  } catch (err) {
    return saidaJson_({ ok: false, erro: String((err && err.message) || err).replace(/^Error:\s*/, '') });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.p === 'painel') return paginaPainel_();
  const ctx = publico_(String((e && e.parameter && e.parameter.t) || '').trim());
  if (e && e.parameter && e.parameter.modo === 'json') return saidaJson_(ctx); // página do GitHub Pages
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.dados = JSON.stringify(ctx).replace(/</g, '\\u003c');
  return tpl.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // permite abrir dentro do app (GitHub Pages)
    .setTitle('Violão – Meus exercícios')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function normaliza_(t) {
  return String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Aceita o nome completo ou o primeiro nome, sem diferenciar maiúsculas e acentos. */
function nomeConfere_(digitado, cadastrado) {
  const d = normaliza_(digitado), c = normaliza_(cadastrado);
  return d.length >= 2 && (d === c || d === c.split(' ')[0]);
}

/** Trava por aparelho: o link só abre nos aparelhos registrados do aluno.
 *  Para liberar um aparelho novo, apague a célula Aparelhos do aluno na aba Alunos. */
function verificarAparelho_(idAluno, aparelho, podeCadastrar) {
  aparelho = String(aparelho || '').replace(/[^a-z0-9]/gi, '').slice(0, 40);
  if (aparelho.length < 12) {
    throw new Error('Não foi possível reconhecer este aparelho. Abra o link no navegador do celular (Chrome ou Safari), fora da janela anônima.');
  }
  const sh = ss_().getSheetByName(ABA.alunos);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const cA = cab.indexOf('Aparelhos');
  const linha = v.findIndex((l, i) => i > 0 && l[cab.indexOf('ID_Aluno')] === idAluno);
  if (cA < 0 || linha < 0) return;
  const lista = String(v[linha][cA] || '').split(',').map(x => x.trim()).filter(Boolean);
  if (lista.indexOf(aparelho) >= 0) return;
  if (podeCadastrar && lista.length < LIMITE_APARELHOS) {
    lista.push(aparelho);
    sh.getRange(linha + 1, cA + 1).setValue(lista.join(', '));
    return;
  }
  throw new Error('Este link só abre nos aparelhos cadastrados da aluna. Se você trocou de celular ou de navegador, peça ao professor para liberar.');
}

/** 1º acesso: grava o cadastro na aba Alunos e libera os exercícios. */
function cadastrar(d) {
  return comTrava_(() => comMemo_(() => cadastrar_(d)));
}

function cadastrar_(d) {
  d = d || {};
  const token = String(d.token || '');
  const ctx = contexto_(token);
  if (ctx.erro) throw new Error(ctx.erro);
  const sh = ss_().getSheetByName(ABA.alunos);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const linha = v.findIndex((l, i) => i > 0 && l[cab.indexOf('ID_Aluno')] === ctx.idAluno);
  if (linha < 0) throw new Error('Aluno não encontrado. Fale com o professor.');
  if (String(v[linha][cab.indexOf('Cadastro_Completo')]).toLowerCase() === 'sim') {
    throw new Error('Seu cadastro já foi feito. Atualize a página e entre com o seu nome.');
  }

  const limpa = (t, max) => { let x = String(t || '').trim().slice(0, max); if (/^[=+\-@]/.test(x)) x = "'" + x; return x; };
  const nome = limpa(d.nome, 80);
  const fone = String(d.telefone || '').replace(/\D/g, '');
  if (normaliza_(nome).length < 3) throw new Error('Digite seu nome completo.');
  if (fone.length < 10 || fone.length > 11) throw new Error('Digite o WhatsApp com DDD (ex.: 94 99123-4567).');
  verificarAparelho_(ctx.idAluno, d.aparelho, true);

  limparMemo_(); // verificarAparelho_ pode ter gravado em Alunos
  const campos = {
    Nome: nome, Telefone: fone, Email: limpa(d.email, 80),
    Tem_Violao: limpa(d.temViolao, 10), Tipo_Violao: limpa(d.tipoViolao, 20),
    Experiencia: limpa(d.experiencia, 300), Estilos: limpa(d.estilos, 300), Objetivo: limpa(d.objetivo, 300),
    Data_Cadastro: new Date(), Cadastro_Completo: 'Sim'
  };
  Object.keys(campos).forEach(k => {
    const col = cab.indexOf(k);
    if (col >= 0) sh.getRange(linha + 1, col + 1).setValue(campos[k]);
  });
  SpreadsheetApp.flush();
  limparMemo_();
  const novo = contexto_(token);
  novo.evolucao = evolucaoAluno_(novo.idAluno);
  novo.nomeDigitado = nome;
  return novo;
}

/** Links seguintes: entra só com o nome. */
function entrar(token, nome, aparelho) { return comMemo_(() => entrar_(token, nome, aparelho)); }

function entrar_(token, nome, aparelho) {
  token = String(token || '');
  const ctx = contexto_(token);
  if (ctx.erro) throw new Error(ctx.erro);
  if (!nomeConfere_(nome, ctx.nome)) {
    falhaNome_(token); // só nomes errados contam para o limite
    throw new Error('Esse nome não confere com o cadastro deste link. Confira e tente de novo.');
  }
  comTrava_(() => verificarAparelho_(ctx.idAluno, aparelho, true));
  ctx.evolucao = evolucaoAluno_(ctx.idAluno);
  ctx.nomeDigitado = String(nome);
  return ctx;
}

/** Chamado pela página quando a aluna envia a prática do dia. */
function registrar(d) { return comMemo_(() => registrar_(d)); }

/** Conta só os nomes ERRADOS por link (10 a cada 15 minutos): uso normal nunca chega perto. */
function falhaNome_(token) {
  const cache = CacheService.getScriptCache();
  const k = 'falha_' + token;
  const n = Number(cache.get(k) || 0);
  if (n >= 10) throw new Error('Muitas tentativas. Espere alguns minutos e tente de novo.');
  cache.put(k, String(n + 1), 900);
}

function registrar_(d) {
  d = d || {};
  const ctx = contexto_(String(d.token || ''));
  if (ctx.erro) throw new Error(ctx.erro);
  if (!nomeConfere_(d.nome, ctx.nome)) {
    falhaNome_(ctx.token);
    throw new Error('Entre de novo com o seu nome.');
  }
  verificarAparelho_(ctx.idAluno, d.aparelho, false);
  const tipoAparelho = /Mobi|Android|iPhone|iPad/i.test(String(d.navegador || '')) ? 'Celular' : 'Computador';

  const validos = new Map(ctx.exercicios.map(x => [x.id, x.titulo]));
  const marcados = (d.exercicios || []).map(String).filter(id => validos.has(id));
  if (!marcados.length) throw new Error('Marque pelo menos um exercício.');

  const minutos = Math.max(0, Math.min(300, Math.round(Number(d.minutos) || 0)));
  const dificuldade = ['Fácil', 'Médio', 'Difícil'].indexOf(d.dificuldade) >= 0 ? d.dificuldade : '';
  let comentario = String(d.comentario || '').slice(0, 500);
  if (/^[=+\-@]/.test(comentario)) comentario = "'" + comentario; // evita fórmula na planilha

  const linha = [new Date(), ctx.token, ctx.idAluno, ctx.nome, ctx.idAula, marcados.join(', '),
    marcados.map(id => validos.get(id)).join(', '), minutos, dificuldade, comentario,
    tipoAparelho + ' ' + String(d.aparelho || '').slice(0, 6)];

  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    const sh = ss_().getSheetByName(ABA.reg);
    sh.getRange(sh.getLastRow() + 1, 1, 1, linha.length).setValues([linha]);
  } finally {
    trava.releaseLock();
  }
  limparMemo_();
  const novo = contexto_(ctx.token);
  novo.evolucao = evolucaoAluno_(novo.idAluno);
  novo.nomeDigitado = String(d.nome);
  return novo;
}

/* =====================================================================
 * FREQUÊNCIA: reposição automática pela sua DISPONIBILIDADE
 * Ao marcar "Falta" ou "Falta justificada" numa aula da aba Frequencia:
 *  - procura o primeiro horário livre seu (aba Disponibilidade) depois da falta,
 *    que não coincida com outra aula já agendada, e agenda a reposição;
 *  - estende a validade do link (aba Envios) até a reposição;
 *  - se ela faltar também na reposição, o link passa a valer até a próxima aula agendada.
 * Funciona sozinho ao editar a planilha (onEdit). Para edições feitas pelo AppSheet,
 * use o menu Violão > Ativar automação (verifica a cada hora).
 * ===================================================================== */
const DIAS_BUSCA = 14;          // até quantos dias à frente procurar horário
const DURACAO_AULA_MIN = 60;    // cada aula ocupa 1 hora
const DIAS_SEMANA = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Frequencia') return;
  const colPresenca = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('Presenca') + 1;
  if (e.range.getColumn() > colPresenca || e.range.getLastColumn() < colPresenca) return;
  try {
    const msgs = comTrava_(() => processarFaltas_(e.source), 10000);
    if (msgs.length) e.source.toast(msgs.join('\n'), 'Reposição', 10);
  } catch (err) {
    // sem trava a tempo: a verificação por hora (ou a próxima edição) refaz
  }
}

/** Menu e gatilho por hora. */
function processarFaltas() {
  const ss = ss_();
  const msgs = comTrava_(() => processarFaltas_(ss));
  try { if (msgs.length) ss.toast(msgs.join('\n'), 'Reposição', 10); } catch (err) {}
}

function ativarAutomacao() {
  soDono_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processarFaltas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processarFaltas').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Automação ativada: as faltas são verificadas a cada hora (inclusive as marcadas pelo AppSheet).');
}

// "18:00" ou um horário da planilha -> minutos desde 00:00
function minutos_(h) {
  if (h instanceof Date) return h.getHours() * 60 + h.getMinutes();
  const m = String(h).trim().match(/^(\d{1,2})[:h](\d{2})?/);
  return m ? Number(m[1]) * 60 + Number(m[2] || 0) : null;
}
function hhmm_(min) {
  return String(Math.floor(min / 60)).padStart(2, '0') + ':' + String(min % 60).padStart(2, '0');
}
function semAcento_(t) { return String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(); }

/** Primeiro horário livre depois de "depoisDe", ou null. Retorna {data, horario}. */
function acharHorario_(ss, depoisDe, ocupados) {
  const disp = ss.getSheetByName('Disponibilidade').getDataRange().getValues();
  const cab = disp[0];
  const c = n => cab.indexOf(n);
  const dia = d => Utilities.formatDate(new Date(d), FUSO, 'yyyy-MM-dd');

  // horários semanais (colunas A–E) e exceções por data (colunas G–J)
  const semanais = disp.slice(1).filter(l => l[c('Dia_Semana')] !== '' && semAcento_(l[c('Ativo')]) !== 'nao')
    .filter(l => ['reposicao', 'ambos'].indexOf(semAcento_(l[c('Uso')])) >= 0)
    .map(l => ({ dow: DIAS_SEMANA.map(semAcento_).indexOf(semAcento_(l[c('Dia_Semana')])),
                 ini: minutos_(l[c('Hora_Inicio')]), fim: minutos_(l[c('Hora_Fim')]) }));
  const excecoes = disp.slice(1).filter(l => l[c('Data')] instanceof Date)
    .map(l => ({ dia: dia(l[c('Data')]), tipo: semAcento_(l[c('Situacao')]),
                 ini: minutos_(l[c('Inicio_Extra')]), fim: minutos_(l[c('Fim_Extra')]) }));

  // nunca procura antes de hoje: falta marcada com atraso não gera reposição no passado
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const d = new Date(depoisDe);
  if (d < hoje) d.setTime(hoje.getTime());
  // dias de ausência não contam na janela de busca: a reposição vai para depois do retorno
  const ausentes = diasAusencia_(ss);
  for (let n = 1, uteis = 0; n <= DIAS_BUSCA + 90 && uteis < DIAS_BUSCA; n++) {
    const cand = new Date(d); cand.setDate(d.getDate() + n);
    const chave = dia(cand);
    if (ausentes.has(chave)) continue;
    uteis++;
    const exDia = excecoes.filter(x => x.dia === chave);
    if (exDia.some(x => x.tipo === 'indisponivel')) continue;
    const blocos = exDia.filter(x => x.tipo === 'disponivel extra' && x.ini != null && x.fim != null)
      .concat(semanais.filter(s => s.dow === cand.getDay() && s.ini != null && s.fim != null));
    blocos.sort((a, b) => a.ini - b.ini);
    for (const b of blocos) {
      for (let t = b.ini; t + DURACAO_AULA_MIN <= b.fim; t += DURACAO_AULA_MIN) {
        const horario = hhmm_(t) + '–' + hhmm_(t + DURACAO_AULA_MIN);
        if (!ocupados.has(chave + ' ' + hhmm_(t))) return { data: cand, horario: horario, chave: chave + ' ' + hhmm_(t) };
      }
    }
  }
  return null;
}

function processarFaltas_(ss) {
  const fr = ss.getSheetByName('Frequencia');
  const env = ss.getSheetByName('Envios');
  const v = fr.getDataRange().getValues();
  const cab = v[0];
  const c = n => cab.indexOf(n);
  const [cData, cAluno, cAula, cPres, cRep, cObs, cHora, cTipo, cNome] =
    ['Data', 'ID_Aluno', 'ID_Aula', 'Presenca', 'Reposicao_Data', 'Obs', 'Horario', 'Tipo', 'Nome'].map(c);
  const dia = d => Utilities.formatDate(new Date(d), FUSO, 'yyyy-MM-dd');
  const br = d => Utilities.formatDate(new Date(d), FUSO, 'dd/MM');
  const msgs = [];

  // horários já ocupados por aulas agendadas (para não marcar duas no mesmo horário)
  const ocupados = new Set(v.slice(1)
    .filter(x => x[cData] instanceof Date && x[cPres] === 'Agendada')
    .map(x => dia(x[cData]) + ' ' + hhmm_(minutos_(String(x[cHora]).split(/[–-]/)[0]) || 0)));

  for (let i = 1; i < v.length; i++) {
    const l = v[i];
    const pres = l[cPres];
    if ((pres !== 'Falta' && pres !== 'Falta justificada') || l[cRep] !== '' || !(l[cData] instanceof Date)) continue;
    const aluno = l[cAluno], aula = l[cAula], data = l[cData];
    const nome = l[cNome] || aluno;
    const ehReposicao = l[cTipo] === 'Reposição';
    const proximaAula = () => v.slice(1)
      .filter(x => x[cAluno] === aluno && x[cPres] === 'Agendada' && x[cData] instanceof Date && x[cData] > data)
      .map(x => x[cData]).sort((a, b) => a - b)[0] || null;
    let novaValidade = null;

    const slot = ehReposicao ? null : acharHorario_(ss, data, ocupados);
    if (slot) {
      ocupados.add(slot.chave);
      novaValidade = slot.data;
      fr.getRange(i + 1, cRep + 1).setValue(slot.data).setNumberFormat('dd/MM/yyyy');
      let alvo = v.findIndex((x, k) => k > 0 && x[cData] === '' && x[cAluno] === '');
      if (alvo === -1) { alvo = v.length; v.push(new Array(cab.length).fill('')); }
      const r = alvo + 1;
      fr.getRange(r, cData + 1).setValue(slot.data).setNumberFormat('dd/MM/yyyy');
      fr.getRange(r, cAluno + 1).setValue(aluno);
      fr.getRange(r, cAula + 1).setValue(aula);
      fr.getRange(r, cPres + 1).setValue('Agendada');
      fr.getRange(r, cObs + 1).setValue('Reposição da aula de ' + br(data));
      fr.getRange(r, cHora + 1).setValue(slot.horario);
      fr.getRange(r, cTipo + 1).setValue('Reposição');
      fr.getRange(r, cNome + 1).setFormula('=IF(B' + r + '="","",IFERROR(INDEX(Alunos!B:B,MATCH(B' + r + ',Alunos!A:A,0)),""))');
      fr.getRange(r, c('Aula') + 1).setFormula('=IF(D' + r + '="","",IFERROR("Aula "&INDEX(Aulas!B:B,MATCH(D' + r + ',Aulas!A:A,0))&" – "&INDEX(Aulas!C:C,MATCH(D' + r + ',Aulas!A:A,0)),""))');
      v[alvo][cData] = slot.data; v[alvo][cAluno] = aluno;
      msgs.push(nome + ': reposição agendada para ' + DIAS_SEMANA[slot.data.getDay()].toLowerCase() + ' ' +
        br(slot.data) + ', ' + slot.horario + '.');
    } else {
      novaValidade = proximaAula();
      fr.getRange(i + 1, cRep + 1).setValue(ehReposicao ? 'Sem nova reposição' : 'A definir');
      msgs.push(nome + (ehReposicao ? ': faltou também na reposição. ' :
        ': sem horário livre nos próximos ' + DIAS_BUSCA + ' dias, combine a reposição. ') +
        (novaValidade ? 'O link vale até a próxima aula (' + br(novaValidade) + ').' : 'Não há próxima aula agendada.'));
    }

    // estende a validade do link que vencia na data da falta
    if (novaValidade) {
      const ev = env.getDataRange().getValues();
      const eA = ev[0].indexOf('ID_Aluno'), eV = ev[0].indexOf('Validade');
      for (let k = 1; k < ev.length; k++) {
        if (ev[k][eA] === aluno && ev[k][eV] instanceof Date && dia(ev[k][eV]) === dia(data)) {
          env.getRange(k + 1, eV + 1).setValue(novaValidade).setNumberFormat('dd/MM/yyyy');
        }
      }
    }
  }
  return msgs;
}


/* =====================================================================
 * AUSÊNCIA DO PROFESSOR (férias, viagem...)
 * Os períodos ficam na aba Disponibilidade, colunas L (Ausencia_Inicio) e M (Ausencia_Fim).
 * Menu Violão > Registrar ausência: pede as datas, grava o período e, para cada aula
 * Agendada dentro dele, marca "Falta justificada" (motivo Outro, obs "Ausência do professor").
 * A reposição automática então agenda cada aula no primeiro horário livre DEPOIS do retorno.
 * ===================================================================== */
const COL_AUS_INI = 'Ausencia_Inicio';
const COL_AUS_FIM = 'Ausencia_Fim';
const OBS_AUSENCIA = 'Ausência do professor';

function paraData_(t) {
  const m = String(t || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return d.getMonth() === Number(m[2]) - 1 ? d : null;
}

/** Períodos de ausência cadastrados: [{ini: Date, fim: Date}] */
function periodosAusencia_(ss) {
  const sh = ss.getSheetByName('Disponibilidade');
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  const a = v[0].indexOf(COL_AUS_INI), b = v[0].indexOf(COL_AUS_FIM);
  if (a < 0 || b < 0) return [];
  return v.slice(1).filter(l => l[a] instanceof Date && l[b] instanceof Date && l[b] >= l[a])
    .map(l => ({ ini: l[a], fim: l[b] }));
}

/** Conjunto de dias 'yyyy-MM-dd' cobertos por alguma ausência. */
function diasAusencia_(ss) {
  const set = new Set();
  periodosAusencia_(ss).forEach(p => {
    const d = new Date(p.ini.getFullYear(), p.ini.getMonth(), p.ini.getDate(), 12);
    const fim = new Date(p.fim.getFullYear(), p.fim.getMonth(), p.fim.getDate(), 12);
    for (let k = 0; d <= fim && k < 400; k++, d.setDate(d.getDate() + 1)) {
      set.add(Utilities.formatDate(d, FUSO, 'yyyy-MM-dd'));
    }
  });
  return set;
}

function registrarAusencia() {
  soDono_();
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Ausência do professor', 'Primeiro dia da ausência (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Ausência do professor', 'Último dia da ausência (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const ini = paraData_(r1.getResponseText()), fim = paraData_(r2.getResponseText());
  if (!ini || !fim || fim < ini) { ui.alert('Datas inválidas. Use dd/mm/aaaa e o último dia não pode ser antes do primeiro.'); return; }
  if ((fim - ini) / 86400000 > 120) { ui.alert('Período longo demais (máximo 120 dias).'); return; }

  const msgs = comTrava_(() => {
    const ss = ss_();
    const sh = ss.getSheetByName('Disponibilidade');
    let cab = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 13)).getValues()[0];
    let a = cab.indexOf(COL_AUS_INI) + 1, b = cab.indexOf(COL_AUS_FIM) + 1;
    if (!a || !b) {
      a = 12; b = 13;
      sh.getRange(1, a, 1, 2).setValues([[COL_AUS_INI, COL_AUS_FIM]]).setFontWeight('bold');
    }
    const col = sh.getRange(1, a, Math.max(sh.getLastRow(), 2), 1).getValues();
    let linha = 2;
    while (linha <= col.length && col[linha - 1][0] !== '') linha++;
    sh.getRange(linha, a).setValue(ini).setNumberFormat('dd/MM/yyyy');
    sh.getRange(linha, b).setValue(fim).setNumberFormat('dd/MM/yyyy');
    SpreadsheetApp.flush();
    return aplicarAusencias_(ss);
  }, 30000);
  ui.alert('Ausência registrada.\n\n' + (msgs.length ? msgs.join('\n') : 'Nenhuma aula agendada caiu nesse período.'));
}

/** Para cada aula Agendada dentro de uma ausência: vira Falta justificada (a reposição automática agenda depois). */
function aplicarAusencias_(ss) {
  const ausentes = diasAusencia_(ss);
  if (!ausentes.size) return [];
  const fr = ss.getSheetByName('Frequencia');
  const v = fr.getDataRange().getValues();
  const cab = v[0];
  const c = n => cab.indexOf(n);
  const dia = d => Utilities.formatDate(new Date(d), FUSO, 'yyyy-MM-dd');
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  let marcadas = 0, refeitas = 0;

  for (let i = 1; i < v.length; i++) {
    const l = v[i];
    if (!(l[c('Data')] instanceof Date) || l[c('Presenca')] !== 'Agendada') continue;
    if (l[c('Data')] < hoje || !ausentes.has(dia(l[c('Data')]))) continue;
    const r = i + 1;

    if (l[c('Tipo')] === 'Reposição') {
      // reposição que caiu na ausência: apaga e devolve a falta original para ser reagendada
      for (let k = 1; k < v.length; k++) {
        const o = v[k];
        if (o[c('ID_Aluno')] === l[c('ID_Aluno')] && o[c('Reposicao_Data')] instanceof Date &&
            dia(o[c('Reposicao_Data')]) === dia(l[c('Data')]) && /^Falta/.test(o[c('Presenca')])) {
          fr.getRange(k + 1, c('Reposicao_Data') + 1).clearContent();
          v[k][c('Reposicao_Data')] = '';
        }
      }
      fr.getRange(r, 1, 1, cab.length).clearContent();
      v[i][c('Data')] = ''; v[i][c('ID_Aluno')] = '';
      refeitas++;
      continue;
    }
    fr.getRange(r, c('Presenca') + 1).setValue('Falta justificada');
    if (c('Motivo_Falta') >= 0) fr.getRange(r, c('Motivo_Falta') + 1).setValue('Outro');
    const obs = String(l[c('Obs')] || '');
    if (obs.indexOf(OBS_AUSENCIA) < 0) fr.getRange(r, c('Obs') + 1).setValue(obs ? obs + ' · ' + OBS_AUSENCIA : OBS_AUSENCIA);
    marcadas++;
  }
  SpreadsheetApp.flush();
  const msgs = processarFaltas_(ss);
  return ['Aulas marcadas como falta justificada: ' + marcadas + (refeitas ? ' (reposições refeitas: ' + refeitas + ')' : '') + '.']
    .concat(msgs);
}

/* =====================================================================
 * AULAS 11 A 30 (apostilas Intermediário e Avançado)
 * Menu Violão > Criar aulas 11 a 30. Só cria o que ainda não existe (por ID) e
 * já liga a apostila em PDF. Pode rodar mais de uma vez sem duplicar.
 * ===================================================================== */
const AULAS_11_30 = [
  [11,"Pestana: o acorde F","Entender a pestana (barra) e montar o Fá maior (F), primeiro na versão simplificada e depois completa.","Treinar o F todo dia, em sessões curtas (5 min, pausa, 5 min). A mão cansa: é normal. Se doer, pare e descanse.",[["Pestana solta","Pestana na casa 1 sem outros dedos; tocar corda por corda e checar todas",5],["F simplificado","Montar F de 4 cordas e trocar com C e G",8],["F completo","Montar F inteiro, segurar 10 segundos, soltar, repetir",7]]],
  [12,"Bm e acordes com pestana","Aprender Si menor (Bm) e o conceito de formas \"móveis\" para tocar qualquer acorde menor ou maior com pestana.","Tocar uma música que tenha Bm (ex.: sequência G - D - Em - Bm) com a batida da aula 5.",[["Bm","Montar Bm e trocar com G e D",8],["F#m","Montar F#m e trocar com A e E",7],["Forma móvel","Deslizar a forma tipo Em da casa 1 até a 5 dizendo o nome",5]]],
  [13,"Acordes com sétima","Conhecer os acordes com sétima (A7, D7, E7, G7, C7, B7) e perceber a sensação de \"movimento\" que eles criam.","Escolher uma música com ao menos 2 acordes com sétima e tocá-la completa.",[["Montar os 6 acordes","Montar cada um e tocar corda por corda",8],["Resoluções","Tocar G7-C, A7-D, E7-Am ouvindo a \"chegada\"",8],["Aplicar numa música","Tocar uma música que use sétimas",7]]],
  [14,"Ritmos: pop-rock e baião","Aprender duas batidas novas e perceber como o ritmo muda o \"estilo\" da mesma sequência de acordes.","Aplicar cada batida em uma música diferente do seu repertório.",[["Pop-rock","Batida de colcheias em G - D - Em - C, metrônomo a 70 bpm",8],["Baião","Padrão simplificado em G - C - D",8],["Comparar","Tocar a mesma sequência nos dois ritmos e perceber a diferença",4]]],
  [15,"Capotraste e transposição","Usar o capotraste para tocar músicas em tons mais altos com formas de acorde que você já domina, e transpor músicas para o tom da sua voz.","Escolher uma música do repertório e encontrar o tom ideal para cantar e tocar, anotando capotraste e acordes.",[["Capo nas casas 2, 3 e 5","Tocar G - C - D - Em com capotraste e dizer o tom que soa",8],["Transpor uma cifra","Subir ou descer uma música 2 meios-tons por escrito",10]]],
  [16,"Arpejos e dedilhado em 3/4","Aprender arpejos mais variados e tocar em compasso ternário (3/4), como valsas e baladas.","Dedilhar uma música inteira (ex.: uma valsa ou balada lenta) sem trocar para batida.",[["Arpejo 3/4","G - D - Em - C, um compasso por acorde",8],["Arpejo de 8 notas","Em - Am - C - G com o padrão completo",8],["Aplicar numa música","Tocar uma valsa ou balada dedilhando",9]]],
  [17,"Escala maior de Dó","Aprender a escala maior de Dó (C) no violão e entender sua fórmula, que vale para todas as escalas maiores.","Tocar a escala de C todo dia, com metrônomo a 60 bpm, uma nota por batida.",[["Subir e descer","Tocar a escala de C subindo e descendo devagar, falando o nome das notas",8],["Em terças","Tocar a escala de 3 em 3 notas (C-D-E, D-E-F...) como exercício de dedos",8],["Melodia","Tocar uma melodia simples (ex.: Ode à Alegria) usando as notas da escala",4]]],
  [18,"Campo harmônico maior","Entender quais acordes \"combinam\" entre si em cada tom, para tocar de ouvido e criar suas próprias sequências.","Escolher uma música nova e achar os acordes pelo ouvido usando o campo harmônico como guia.",[["Escrever o campo","Escrever o campo harmônico de G e de D de memória",6],["Tocar I - V - vi - IV","Tocar a sequência nos tons de C, G e D",10],["Ouvido","Pôr uma música e tentar achar o tom e os acordes",4]]],
  [19,"Pentatônica menor","Aprender a escala pentatônica menor (posição 1) e usá-la para improvisar sobre uma base de acordes.","Gravar 30 segundos de improviso e ouvir de volta. Anotar o que gostou.",[["Escala subindo e descendo","Pentatônica de Am, posição 1, metrônomo 60 bpm",8],["Improviso com base","Improvisar sobre Am - G - F - E",10],["Frase própria","Criar e repetir uma frase de 4 notas",4]]],
  [20,"Revisão e avaliação","Revisar o nível Intermediário e fazer a avaliação para liberar o nível Avançado.","Ensaiar para a avaliação",[["Preparar a avaliação","Ensaiar as músicas e o improviso que serão apresentados na avaliação do nível",20]]],
  [21,"Funções harmônicas e cadências","Entender para que serve cada acorde dentro do tom e usar isso para harmonizar melodias e substituir acordes.","Analisar a harmonia de 2 músicas e escrever as funções de cada compasso.",[["Ouvir cadências","Tocar as 4 cadências em C e descrever a sensação de cada uma",8],["Analisar música","Marcar T, SD e D numa música do repertório",10],["Substituir","Trocar 2 acordes por substitutos e comparar",7]]],
  [22,"Acordes com 7M, m7 e 9","Ampliar o vocabulário de acordes com sétima maior, menor com sétima e extensões (add9, sus2, sus4).","Trocar os acordes simples de uma música sua por versões com 7M, m7 ou add9 e ouvir a diferença.",[["Montar todos","Montar os 13 acordes e tocar corda por corda",10],["Sequência","Cmaj7 - Am7 - Dm7 - G7",8],["Arpejo com sus","Dsus4 - D - Dsus2 - D em arpejo",7]]],
  [23,"Inversões e acordes com baixo","Usar acordes com baixo diferente (slash chords) para criar linhas de baixo suaves entre os acordes.","Escolher uma música simples e acrescentar uma linha de baixo com slash chords.",[["Montar slash chords","Montar os 6 acordes da tabela",8],["Linha descendente","Tocar a sequência C - G/B - Am - Am/G - F - C/E - Dm - G",10],["Aplicar","Usar 2 slash chords numa música conhecida",7]]],
  [24,"Pentatônica nas 5 posições","Conhecer as 5 posições da pentatônica menor de Lá (Am) para improvisar em todo o braço, não só numa região.","Aprender uma posição nova por dia. No fim da semana, improvisar usando as cinco.",[["Posições 1 e 2","Tocar e unir as duas, subindo e descendo",8],["Posições 3, 4 e 5","Mesma coisa com as demais",10],["Improviso","Improvisar mudando de posição a cada 8 batidas",7]]],
  [25,"Escala maior no braço todo","Tocar a escala maior em formas móveis que funcionam em qualquer tom, usando a nota raiz como referência.","Escolher uma melodia simples e tocá-la em dois tons diferentes com a forma móvel.",[["Posição em G e C","Tocar a posição com raiz na casa 3 (G) e na casa 8 (C), subindo e descendo",10],["Tocar em tons","Falar o nome de cada nota ao tocar",10],["Melodia","Tocar uma melodia conhecida (ex.: Parabéns, Ode à Alegria) dentro da posição",5]]],
  [26,"Bossa nova e fingerstyle","Aprender a batida de bossa nova e os primeiros passos do fingerstyle, em que o violão toca baixo, harmonia e melodia ao mesmo tempo.","Ouvir 3 músicas de bossa nova prestando atenção só no violão. Tentar reproduzir a levada.",[["Polegar caminhante","Alternar raiz-quinta em Cmaj7 - Am7 - Dm7 - G7",8],["Bossa simplificada","Polegar + acorde i-m-a sobre a mesma sequência",10],["Melodia + baixo","Tocar uma melodia simples com o polegar fazendo o baixo",7]]],
  [27,"II-V-I e harmonia de jazz","Dominar a progressão II-V-I, o \"coração\" da harmonia da MPB, bossa nova e jazz, em vários tons.","Tocar o II-V-I nos 4 tons da tabela, sem olhar, e transpor para um quinto tom à sua escolha.",[["II-V-I em C e G","Tocar a progressão devagar, 2 batidas por acorde",10],["Em F e D","Mesma progressão nos tons de F e D",10],["Numa música","Encontrar II-V-I numa música do repertório e marcar",5]]],
  [28,"Técnicas: ligados, bend e slide","Aprender as técnicas de expressão que dão personalidade aos solos e melodias.","Escolher uma frase de uma música que use técnicas e reproduzir fielmente.",[["Cada técnica isolada","5 minutos de ligado, slide e bend em cordas soltas e na escala",10],["Frase com técnicas","Criar uma frase de 6 a 8 notas usando 2 técnicas",10],["Improviso","Improvisar 1 minuto usando as técnicas",5]]],
  [29,"Arranjo solo e repertório","Montar um arranjo próprio de uma música, combinando melodia, harmonia e baixo, e construir um repertório de apresentação.","Gravar o arranjo em áudio ou vídeo e ouvir de volta com a partitura de estrutura ao lado.",[["Estrutura","Escrever a estrutura da música escolhida (partes e compassos)",10],["Harmonia","Reharmonizar 2 trechos com acordes novos",10],["Tocar inteira","Tocar o arranjo completo sem parar",10]]],
  [30,"Projeto final e avaliação","Apresentar o projeto final do curso e receber a avaliação de conclusão.","Ensaiar para a avaliação",[["Preparar a avaliação","Ensaiar as músicas e o improviso que serão apresentados na avaliação do nível",20]]]
];

function criarAulas11a30() {
  soDono_();
  const ss = ss_();
  const r = comTrava_(() => {
    const aulas = ss.getSheetByName(ABA.aulas), ex = ss.getSheetByName(ABA.ex);
    const idsAulas = new Set(aulas.getRange(2, 1, Math.max(aulas.getLastRow() - 1, 1), 1).getValues().map(x => x[0]));
    const idsEx = new Set(ex.getRange(2, 1, Math.max(ex.getLastRow() - 1, 1), 1).getValues().map(x => x[0]));
    const base = urlPublica_() + 'apostilas/';
    const novasAulas = [], novosEx = [];
    AULAS_11_30.forEach(a => {
      const num = String(a[0]).padStart(2, '0'), id = 'AU' + num;
      if (idsAulas.has(id)) return;
      novasAulas.push([id, a[0], a[1], a[2], a[3], base + 'aula-' + num + '.pdf']);
      a[4].forEach((e, i) => {
        const idEx = 'EX' + num + '-' + (i + 1);
        if (!idsEx.has(idEx)) novosEx.push([idEx, id, i + 1, e[0], e[1], e[2]]);
      });
    });
    if (novasAulas.length) aulas.getRange(aulas.getLastRow() + 1, 1, novasAulas.length, 6).setValues(novasAulas);
    if (novosEx.length) ex.getRange(ex.getLastRow() + 1, 1, novosEx.length, 6).setValues(novosEx);
    limparMemo_();
    return { aulas: novasAulas.length, ex: novosEx.length };
  }, 30000);
  SpreadsheetApp.getUi().alert('Criadas ' + r.aulas + ' aulas e ' + r.ex + ' exercícios (aulas que já existiam foram mantidas).');
}
