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
    .addSeparator()
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
  for (let n = 1; n <= DIAS_BUSCA; n++) {
    const cand = new Date(d); cand.setDate(d.getDate() + n);
    const chave = dia(cand);
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
