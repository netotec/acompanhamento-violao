/**
 * PainelServidor.gs – Painel do professor (servidor)
 * Na planilha: menu Violão > Abrir painel do professor (não pede senha: só o dono abre).
 * No celular/computador: app do professor (GitHub Pages) ou URL do web app + ?p=painel, com senha.
 * A senha é guardada só como hash (SHA-256 com sal) nas propriedades do script.
 */

/* ---------- Segurança ---------- */

/** Chave interna usada pelo painel aberto dentro da planilha. */
function chavePainel_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('PAINEL_CHAVE');
  if (!k) {
    k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    p.setProperty('PAINEL_CHAVE', k);
  }
  return k;
}

/** Aceita a chave interna (planilha) ou uma sessão válida criada pelo login com senha. */
function checarChave_(k) {
  k = String(k || '');
  const props = PropertiesService.getScriptProperties();
  const certa = props.getProperty('PAINEL_CHAVE');
  if (k && certa && k === certa) return;
  if (/^[a-f0-9]{40,64}$/.test(k)) {
    const validade = Number(props.getProperty('SESS_' + k) || 0);
    if (validade > Date.now()) return;
    if (validade) props.deleteProperty('SESS_' + k);
  }
  throw new Error('SESSAO: Sua sessão terminou. Entre de novo com a senha.');
}

/** Bloqueia funções administrativas quando chamadas pela página pública. */
function soDono_() {
  const ativo = Session.getActiveUser().getEmail();
  const dono = Session.getEffectiveUser().getEmail();
  if (!ativo || ativo !== dono) throw new Error('Acesso negado.');
}

function hashSenha_(senha, sal) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + senha, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function definirSenhaPainel() {
  soDono_();
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Senha do painel do professor', 'Digite a nova senha (mínimo 6 caracteres):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const senha = r.getResponseText();
  if (senha.length < 6) { ui.alert('A senha precisa ter pelo menos 6 caracteres.'); return; }
  const r2 = ui.prompt('Confirme a senha', 'Digite a mesma senha de novo:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (r2.getResponseText() !== senha) { ui.alert('As senhas não são iguais. Nada foi alterado.'); return; }
  const sal = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ PAINEL_SAL: sal, PAINEL_SENHA: hashSenha_(senha, sal) });
  encerrarSessoes_();
  ui.alert('Senha salva. Por segurança, todos os aparelhos foram desconectados: entre de novo com a senha nova.');
}

/** Login do painel: devolve um código de sessão (12 horas, ou 30 dias se "lembrar"). */
function painelLogin(senha, lembrar) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('PAINEL_SENHA'), sal = props.getProperty('PAINEL_SAL');
  if (!hash || !sal) throw new Error('A senha do painel ainda não foi criada. Na planilha, use Violão > Definir senha do painel.');
  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get('login_falhas') || 0);
  if (falhas >= 5) throw new Error('Muitas tentativas erradas. Espere 15 minutos e tente de novo (ou, na planilha, use Violão > Liberar tentativas de senha do painel).');
  Utilities.sleep(400); // atrasa tentativas automáticas
  if (hashSenha_(String(senha || ''), sal) !== hash) {
    cache.put('login_falhas', String(falhas + 1), 900);
    throw new Error('Senha incorreta.' + (falhas + 1 >= 3 ? ' Restam ' + (5 - falhas - 1) + ' tentativas.' : ''));
  }
  cache.remove('login_falhas');
  limparSessoesVencidas_();
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const dias = lembrar ? 30 : 0.5;
  props.setProperty('SESS_' + token, String(Date.now() + dias * 864e5));
  return token;
}

function painelSair(k) {
  if (/^[a-f0-9]{40,64}$/.test(String(k || ''))) PropertiesService.getScriptProperties().deleteProperty('SESS_' + k);
  return true;
}

function limparSessoesVencidas_() {
  const props = PropertiesService.getScriptProperties();
  const todas = props.getProperties();
  Object.keys(todas).forEach(n => { if (n.indexOf('SESS_') === 0 && Number(todas[n]) < Date.now()) props.deleteProperty(n); });
}

function encerrarSessoes_() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(n => { if (n.indexOf('SESS_') === 0) props.deleteProperty(n); });
}

/** Menu: desconecta todos os aparelhos que entraram no painel com senha. */
function desconectarAparelhosPainel() {
  soDono_();
  encerrarSessoes_();
  SpreadsheetApp.getUi().alert('Pronto: todos os aparelhos saíram do painel. Para entrar de novo, é preciso digitar a senha.');
}

/** Menu: zera o bloqueio por tentativas erradas de senha (o painel da planilha nunca é bloqueado). */
function liberarTentativasPainel() {
  soDono_();
  CacheService.getScriptCache().remove('login_falhas');
  SpreadsheetApp.getUi().alert('Pronto: as tentativas de senha foram liberadas.');
}

/* ---------- Abrir o painel ---------- */

function abrirPainel() {
  soDono_();
  const t = HtmlService.createTemplateFromFile('Painel');
  t.chave = chavePainel_();
  t.modo = 'dialogo';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1180).setHeight(780), 'Painel do Professor');
}

function linkPainelCelular() {
  soDono_();
  const props = PropertiesService.getScriptProperties();
  const url = urlPublica_() + 'professor.html';
  const semSenha = !props.getProperty('PAINEL_SENHA');
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial;padding:8px">' +
    '<p><b>Endereço do seu painel</b> (celular ou computador):</p>' +
    '<textarea style="width:100%;height:60px" onclick="this.select()">' + url + '</textarea>' +
    (semSenha ? '<p style="color:#a13a26"><b>Crie a senha antes:</b> Violão &gt; Definir senha do painel.</p>'
              : '<p>Ao abrir, ele pede a sua senha. Marque "Lembrar neste aparelho" para não digitar toda vez.</p>') +
    '<p>No celular, abra o endereço e instale como app (Compartilhar &gt; Adicionar à Tela de Início, ou Instalar).</p></div>'
  ).setWidth(520).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Painel no celular');
}

/** Chamado pelo doGet quando a URL tem ?p=painel. A página pede a senha. */
function paginaPainel_() {
  const t = HtmlService.createTemplateFromFile('Painel');
  t.chave = '';
  t.modo = 'web';
  return t.evaluate().setTitle('Painel do Professor – Violão').addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Ajudantes ---------- */

function iso_(d) {
  return d instanceof Date && !isNaN(d) ? Utilities.formatDate(d, FUSO, 'yyyy-MM-dd') : '';
}
function deIso_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Data inválida.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function hoje_() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }

/** Lê uma aba com o número de cada linha. */
function linhas_(nome) {
  const sh = ss_().getSheetByName(nome);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '' || v[i][0] === null) continue;
    const o = { _r: i + 1 };
    cab.forEach((c, j) => { if (c) o[c] = v[i][j]; });
    rows.push(o);
  }
  return { sh: sh, cab: cab, rows: rows, total: v.length };
}

function primeiraLinhaVazia_(sh) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < col.length; i++) if (col[i][0] === '') return i + 1;
  return col.length + 1;
}

function gravar_(sh, cab, r, campos) {
  Object.keys(campos).forEach(k => {
    const c = cab.indexOf(k);
    if (c >= 0) sh.getRange(r, c + 1).setValue(campos[k]);
  });
}

function limpa_(t, max) {
  let x = String(t == null ? '' : t).trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(x)) x = "'" + x;
  return x;
}

/* ---------- Dados do painel ---------- */

function painelDados(k) {
  checarChave_(k);
  const H = hoje_();
  const alunos = linhas_(ABA.alunos).rows;
  const aulas = linhas_(ABA.aulas).rows;
  const envios = linhas_(ABA.envios).rows;
  const freq = linhas_('Frequencia').rows;
  const regs = linhas_(ABA.reg).rows;
  const exs = linhas_(ABA.ex).rows;
  const dispSh = ss_().getSheetByName('Disponibilidade').getRange(2, 1, 30, 10).getValues();
  const titulo = id => { const a = aulas.find(x => x.ID_Aula === id); return a ? 'Aula ' + a.Numero + ' – ' + a.Titulo : id; };
  const txt = v => (v instanceof Date ? Utilities.formatDate(v, FUSO, 'HH:mm') : String(v == null ? '' : v));

  const resumo = alunos.map(a => {
    const id = a.ID_Aluno;
    const meus = regs.filter(r => r.ID_Aluno === id);
    const ult7 = meus.filter(r => r.Data_Hora instanceof Date && (new Date() - r.Data_Hora) <= 7 * 864e5);
    const ultima = meus.map(r => r.Data_Hora).filter(d => d instanceof Date).sort((x, y) => y - x)[0];
    const minhas = freq.filter(f => f.ID_Aluno === id);
    const realizadas = minhas.filter(f => f.Tipo !== 'Reposição' && f.Presenca && f.Presenca !== 'Agendada').length;
    const presencas = minhas.filter(f => f.Presenca === 'Presente' || f.Presenca === 'Reposição').length;
    const prox = minhas.filter(f => f.Presenca === 'Agendada' && f.Data instanceof Date && f.Data >= H).sort((x, y) => x.Data - y.Data)[0];
    const env = envios.filter(e => e.ID_Aluno === id).sort((x, y) => (y.Data_Envio || 0) - (x.Data_Envio || 0))[0];
    const reenvio = env ? dadosEnvio_(a, aulas.find(x => x.ID_Aula === env.ID_Aula), env.Validade, String(env.Token || '')) : null;
    return {
      id: id,
      praticas7: new Set(ult7.map(r => iso_(r.Data_Hora))).size,
      minutos7: ult7.reduce((s, r) => s + (Number(r.Minutos) || 0), 0),
      minutosTotal: meus.reduce((s, r) => s + (Number(r.Minutos) || 0), 0),
      ultima: ultima ? Utilities.formatDate(ultima, FUSO, 'dd/MM HH:mm') : '',
      diasSemPraticar: ultima ? Math.floor((new Date() - ultima) / 864e5) : null,
      frequencia: realizadas ? Math.min(1, presencas / realizadas) : null,
      realizadas: realizadas, presencas: presencas,
      proxima: prox ? { data: iso_(prox.Data), horario: txt(prox.Horario), aula: titulo(prox.ID_Aula) } : null,
      envio: env ? { aula: titulo(env.ID_Aula), validade: iso_(env.Validade), ativo: env.Validade instanceof Date && env.Validade >= H, link: reenvio.link, whatsapp: reenvio.whatsapp } : null
    };
  });

  return {
    hoje: iso_(H),
    alunos: alunos.map(a => ({
      id: a.ID_Aluno, nome: String(a.Nome || ''), telefone: String(a.Telefone || ''), plano: String(a.Plano || ''),
      valor: a.Valor_Mensal === '' ? '' : Number(a.Valor_Mensal) || 0, status: String(a.Status || ''),
      cadastro: String(a.Cadastro_Completo || '').toLowerCase() === 'sim',
      aparelhos: String(a.Aparelhos || '').split(',').filter(x => x.trim()).length,
      estilos: String(a.Estilos || ''), objetivo: String(a.Objetivo || ''), violao: String(a.Tipo_Violao || '')
    })),
    aulas: aulas.slice().sort((x, y) => x.Numero - y.Numero).map(a => ({
      id: a.ID_Aula, numero: a.Numero, titulo: String(a.Titulo || ''), conteudo: String(a.Conteudo || ''),
      tarefa: String(a.Tarefa || ''), apostila: String(a.Link_Apostila || ''),
      exercicios: exs.filter(x => x.ID_Aula === a.ID_Aula).sort((x, y) => x.Ordem - y.Ordem).map(x => ({
        id: x.ID_Exercicio, titulo: String(x.Titulo || ''), descricao: String(x.Descricao || ''), tempo: Number(x.Tempo_min) || ''
      }))
    })),
    freq: freq.map(f => ({
      r: f._r, data: iso_(f.Data), aluno: f.ID_Aluno, aula: f.ID_Aula, aulaTitulo: titulo(f.ID_Aula),
      presenca: String(f.Presenca || ''), motivo: String(f.Motivo_Falta || ''),
      reposicao: f.Reposicao_Data instanceof Date ? iso_(f.Reposicao_Data) : String(f.Reposicao_Data || ''),
      obs: String(f.Obs || ''), horario: txt(f.Horario), tipo: String(f.Tipo || '')
    })).sort((x, y) => (x.data < y.data ? -1 : x.data > y.data ? 1 : 0)),
    registros: regs.filter(r => r.Data_Hora instanceof Date).sort((x, y) => y.Data_Hora - x.Data_Hora).slice(0, 300).map(r => ({
      quando: Utilities.formatDate(r.Data_Hora, FUSO, 'dd/MM HH:mm'), aluno: r.ID_Aluno, nome: String(r.Nome || ''),
      aula: titulo(r.ID_Aula), exercicios: String(r.Exercicios || ''), minutos: Number(r.Minutos) || 0,
      dificuldade: String(r.Dificuldade || ''), comentario: String(r.Comentario || ''), aparelho: String(r.Aparelho || '')
    })),
    disp: {
      semanais: dispSh.filter(l => l[0] !== '').map(l => ({ dia: String(l[0]), ini: txt(l[1]), fim: txt(l[2]), uso: String(l[3] || 'Ambos'), ativo: String(l[4] || 'Sim') })),
      excecoes: dispSh.filter(l => l[6] instanceof Date).map(l => ({ data: iso_(l[6]), situacao: String(l[7] || ''), ini: txt(l[8]), fim: txt(l[9]) }))
    },
    planos: lerPlanos_(),
    financeiro: lerFinanceiro_(),
    mensalidades: lerMensalidades_(),
    avaliacoes: (abaAvaliacoes_(false) ? linhas_(ABA_AVAL).rows : []).filter(r => r.Data instanceof Date)
      .map(r => ({ data: iso_(r.Data), aluno: r.ID_Aluno, aula: r.ID_Aula, ex: r.ID_Exercicio, reacao: String(r.Reacao || ''), obs: String(r.Obs || '') })),
    resumo: resumo,
    enderecoPublico: urlPublica_()
  };
}

/* ---------- Relatório de frequência ---------- */

/** Resumo por aluno entre duas datas (inclusive): aulas, presenças, faltas, reposições e prática.
 *  Mesma regra da visão geral: "realizadas" = aulas regulares já marcadas; presença conta Presente e Reposição. */
function painelRelatorio(k, deIso, ateIso) {
  checarChave_(k);
  const de = iso_(deIso_(deIso)), ate = iso_(deIso_(ateIso));
  if (de > ate) throw new Error('A data inicial precisa ser antes da final.');
  const alunos = linhas_(ABA.alunos).rows;
  const freq = linhas_('Frequencia').rows.filter(f => { const d = iso_(f.Data); return d && d >= de && d <= ate; });
  const regs = linhas_(ABA.reg).rows.filter(r => { const d = iso_(r.Data_Hora); return d && d >= de && d <= ate; });
  const linhas = alunos.map(a => {
    const f = freq.filter(x => x.ID_Aluno === a.ID_Aluno);
    const r = regs.filter(x => x.ID_Aluno === a.ID_Aluno);
    const conta = p => f.filter(x => x.Presenca === p).length;
    const realizadas = f.filter(x => x.Tipo !== 'Reposição' && x.Presenca && x.Presenca !== 'Agendada').length;
    const presencas = conta('Presente') + conta('Reposição');
    return {
      id: a.ID_Aluno, nome: String(a.Nome || ''), status: String(a.Status || ''),
      realizadas: realizadas, presentes: conta('Presente'), faltas: conta('Falta'), justificadas: conta('Falta justificada'),
      reposicoes: conta('Reposição'), agendadas: conta('Agendada'),
      frequencia: realizadas ? Math.min(1, presencas / realizadas) : null,
      diasPraticou: new Set(r.map(x => iso_(x.Data_Hora))).size,
      minutos: r.reduce((s, x) => s + (Number(x.Minutos) || 0), 0)
    };
  }).filter(l => l.status !== 'Inativo' || l.realizadas || l.minutos);
  return { de: de, ate: ate, linhas: linhas };
}

/* ---------- Reação do aluno a cada exercício (aula presencial) ---------- */

const ABA_AVAL = 'Avaliacoes';
const REACOES = ['Dominou', 'Foi bem', 'Com dificuldade', 'Não conseguiu'];

/** A aba é criada na primeira gravação (não exige rodar o Configurar, que apaga tudo). */
function abaAvaliacoes_(criar) {
  const ss = ss_();
  let sh = ss.getSheetByName(ABA_AVAL);
  if (!sh && criar) {
    sh = ss.insertSheet(ABA_AVAL);
    sh.getRange(1, 1, 1, 6).setValues([['Data', 'ID_Aluno', 'ID_Aula', 'ID_Exercicio', 'Reacao', 'Obs']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Grava (ou atualiza) como o aluno reagiu a cada exercício numa aula presencial.
 *  itens = [{ id: ID_Exercicio, reacao: 'Dominou' | ... | '' (apaga), obs: '' }] */
function painelSalvarAvaliacao(k, idAluno, idAula, dataIso, itens) {
  checarChave_(k);
  return comTrava_(() => painelSalvarAvaliacao_(idAluno, idAula, dataIso, itens));
}

function painelSalvarAvaliacao_(idAluno, idAula, dataIso, itens) {
  if (!tabela_(ABA.alunos).some(a => a.ID_Aluno === idAluno)) throw new Error('Escolha o aluno.');
  if (!tabela_(ABA.aulas).some(a => a.ID_Aula === idAula)) throw new Error('Escolha a aula.');
  const data = deIso_(dataIso);
  if (data > hoje_()) throw new Error('A data não pode ser no futuro.');
  const doDia = tabela_(ABA.ex).filter(x => x.ID_Aula === idAula).map(x => x.ID_Exercicio);
  const lista = (itens || []).map(i => ({ id: String(i.id || ''), reacao: String(i.reacao || ''), obs: limpa_(i.obs, 300) }));
  lista.forEach(i => {
    if (doDia.indexOf(i.id) < 0) throw new Error('Exercício que não pertence a esta aula.');
    if (i.reacao && REACOES.indexOf(i.reacao) < 0) throw new Error('Reação inválida.');
  });

  const sh = abaAvaliacoes_(true);
  const L = linhas_(ABA_AVAL);
  const dia = iso_(data);
  let salvos = 0, removidos = 0;
  lista.forEach(i => {
    const ja = L.rows.find(r => iso_(r.Data) === dia && r.ID_Aluno === idAluno && r.ID_Aula === idAula && r.ID_Exercicio === i.id);
    if (!i.reacao) {
      if (ja) { sh.getRange(ja._r, 1, 1, L.cab.length).clearContent(); ja.ID_Aluno = ''; removidos++; }
      return;
    }
    if (ja) { gravar_(sh, L.cab, ja._r, { Reacao: i.reacao, Obs: i.obs }); ja.Reacao = i.reacao; }
    else {
      const r = primeiraLinhaVazia_(sh);
      sh.getRange(r, 1).setNumberFormat('dd/MM/yyyy');
      gravar_(sh, L.cab, r, { Data: data, ID_Aluno: idAluno, ID_Aula: idAula, ID_Exercicio: i.id, Reacao: i.reacao, Obs: i.obs });
      L.rows.push({ _r: r, Data: data, ID_Aluno: idAluno, ID_Aula: idAula, ID_Exercicio: i.id, Reacao: i.reacao });
    }
    salvos++;
  });
  return { salvos: salvos, removidos: removidos };
}

/* ---------- Pix e mensalidades ---------- */

const ABA_MENS = 'Mensalidades';
const TIPOS_PIX = ['CPF', 'CNPJ', 'E-mail', 'Telefone', 'Aleatória'];

/** Mês da célula como 'aaaa-mm' (a planilha pode converter o texto em data). */
function mesDe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, FUSO, 'yyyy-MM');
  return String(v || '').trim();
}

function abaMensalidades_(criar) {
  const ss = ss_();
  let sh = ss.getSheetByName(ABA_MENS);
  if (!sh && criar) {
    sh = ss.insertSheet(ABA_MENS);
    sh.getRange(1, 1, 1, 5).setValues([['ID_Aluno', 'Mes', 'Valor', 'Pago', 'Data_Pagamento']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function lerMensalidades_() {
  if (!abaMensalidades_(false)) return [];
  return linhas_(ABA_MENS).rows
    .filter(r => r.ID_Aluno && mesDe_(r.Mes))
    .map(r => ({
      aluno: String(r.ID_Aluno), mes: mesDe_(r.Mes), valor: Number(r.Valor) || 0, pago: String(r.Pago || ''),
      data: r.Data_Pagamento instanceof Date ? iso_(r.Data_Pagamento) : ''
    }));
}

/** Dados do Pix e dia de vencimento (ficam nas propriedades privadas do script, não na planilha). */
function lerFinanceiro_() {
  const p = PropertiesService.getScriptProperties();
  return {
    tipo: p.getProperty('PIX_TIPO') || '', chave: p.getProperty('PIX_CHAVE') || '', titular: p.getProperty('PIX_TITULAR') || '',
    vencimento: Number(p.getProperty('DIA_VENCIMENTO')) || 10
  };
}

function painelSalvarPix(k, f) {
  checarChave_(k);
  f = f || {};
  const p = PropertiesService.getScriptProperties();
  const venc = Math.round(Number(f.vencimento) || 10);
  if (venc < 1 || venc > 28) throw new Error('O dia de vencimento precisa ficar entre 1 e 28.');
  const chaveBruta = String(f.chave || '').trim();
  if (!chaveBruta) {
    ['PIX_TIPO', 'PIX_CHAVE', 'PIX_TITULAR'].forEach(n => p.deleteProperty(n));
    p.setProperty('DIA_VENCIMENTO', String(venc));
    return lerFinanceiro_();
  }
  const tipo = String(f.tipo || '');
  if (TIPOS_PIX.indexOf(tipo) < 0) throw new Error('Escolha o tipo da chave Pix.');
  let chave = chaveBruta;
  const so = chaveBruta.replace(/\D/g, '');
  if (tipo === 'CPF') { if (so.length !== 11) throw new Error('CPF precisa ter 11 números.'); chave = so; }
  else if (tipo === 'CNPJ') { if (so.length !== 14) throw new Error('CNPJ precisa ter 14 números.'); chave = so; }
  else if (tipo === 'Telefone') { if (so.length < 10 || so.length > 13) throw new Error('Telefone com DDD, só números (10 a 13 dígitos).'); chave = so; }
  else if (tipo === 'E-mail') { if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(chaveBruta)) throw new Error('E-mail inválido.'); chave = chaveBruta.toLowerCase(); }
  else if (!/^[0-9a-f-]{32,36}$/i.test(chaveBruta)) throw new Error('Chave aleatória inválida: cole os 32 caracteres (letras e números) do seu banco.');
  p.setProperties({ PIX_TIPO: tipo, PIX_CHAVE: chave, PIX_TITULAR: limpa_(f.titular, 60), DIA_VENCIMENTO: String(venc) });
  return lerFinanceiro_();
}

/** Marca (pago = true) ou desmarca (false) a mensalidade de um aluno num mês ('aaaa-mm'). */
function painelMarcarPagamento(k, idAluno, mes, pago, dataIso) {
  checarChave_(k);
  return comTrava_(() => painelMarcarPagamento_(idAluno, mes, pago, dataIso));
}

function painelMarcarPagamento_(idAluno, mes, pago, dataIso) {
  mes = String(mes || '');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(mes)) throw new Error('Mês inválido.');
  const aluno = tabela_(ABA.alunos).find(a => a.ID_Aluno === idAluno);
  if (!aluno) throw new Error('Escolha o aluno.');
  let data = null;
  if (pago) {
    data = dataIso ? deIso_(dataIso) : hoje_();
    if (data > hoje_()) throw new Error('A data do pagamento não pode ser no futuro.');
  }
  const sh = abaMensalidades_(true);
  const L = linhas_(ABA_MENS);
  const ja = L.rows.find(r => String(r.ID_Aluno) === idAluno && mesDe_(r.Mes) === mes);
  if (!pago) {
    if (ja) sh.getRange(ja._r, 1, 1, L.cab.length).clearContent();
    return { pago: false };
  }
  const campos = { Valor: Number(aluno.Valor_Mensal) || 0, Pago: 'Sim', Data_Pagamento: data };
  if (ja) {
    gravar_(sh, L.cab, ja._r, campos);
  } else {
    const r = primeiraLinhaVazia_(sh);
    sh.getRange(r, L.cab.indexOf('Mes') + 1).setNumberFormat('@');
    sh.getRange(r, L.cab.indexOf('Data_Pagamento') + 1).setNumberFormat('dd/MM/yyyy');
    gravar_(sh, L.cab, r, Object.assign({ ID_Aluno: idAluno, Mes: mes }, campos));
  }
  return { pago: true, data: iso_(data) };
}

/* ---------- Planos e valores ---------- */

const ABA_PLANOS = 'Planos';

/** Planos iniciais (referência de mercado de Parauapebas/PA). Podem ser editados no painel. */
function planosPadrao_() {
  return [
    { nome: 'Básico', descricao: '4 aulas presenciais de 1 hora por mês, com o material da aula.', valor: 220, aulasMes: 4, ativo: 'Sim' },
    { nome: 'Com acompanhamento', descricao: '4 aulas de 1 hora por mês + link de exercícios depois de cada aula, registro das práticas, acompanhamento da evolução e reposição automática de faltas.', valor: 280, aulasMes: 4, ativo: 'Sim' },
    { nome: 'Intensivo', descricao: '8 aulas de 1 hora por mês (2 por semana) com acompanhamento completo.', valor: 480, aulasMes: 8, ativo: 'Sim' },
    { nome: 'Cortesia', descricao: 'Sem mensalidade: o aluno não entra no controle de pagamentos.', valor: 0, aulasMes: 4, ativo: 'Sim' }
  ];
}

/** Lê os planos da aba Planos; se ela ainda não existir, devolve os padrões. */
function lerPlanos_() {
  const sh = ss_().getSheetByName(ABA_PLANOS);
  if (!sh) return planosPadrao_();
  const v = sh.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '' || v[i][0] === null) continue;
    lista.push({
      nome: String(v[i][0]), descricao: String(v[i][1] || ''), valor: Number(v[i][2]) || 0,
      aulasMes: Number(v[i][3]) || 0, ativo: String(v[i][4] || 'Sim')
    });
  }
  return lista.length ? lista : planosPadrao_();
}

/** Salva a tabela de planos inteira. Cada item pode trazer `anterior` (nome antigo): se o nome mudou,
 *  os alunos desse plano passam para o nome novo. */
function painelSalvarPlanos(k, planos) {
  checarChave_(k);
  return comTrava_(() => painelSalvarPlanos_(planos));
}

function painelSalvarPlanos_(planos) {
  if (!Array.isArray(planos) || !planos.length) throw new Error('Mantenha pelo menos um plano.');
  const vistos = {};
  const lista = planos.map(p => {
    p = p || {};
    const nome = limpa_(p.nome, 40);
    if (normaliza_(nome).length < 2) throw new Error('Todo plano precisa de um nome.');
    const chave = normaliza_(nome);
    if (vistos[chave]) throw new Error('Há dois planos com o nome "' + nome + '".');
    vistos[chave] = true;
    const valor = Number(String(p.valor == null ? '' : p.valor).replace(',', '.'));
    if (!(valor >= 0) || valor > 100000) throw new Error('Valor inválido no plano "' + nome + '".');
    const aulas = Math.round(Number(p.aulasMes) || 0);
    if (aulas < 0 || aulas > 31) throw new Error('Aulas por mês inválidas no plano "' + nome + '".');
    return {
      nome: nome, anterior: String(p.anterior || ''), descricao: limpa_(p.descricao, 250),
      valor: Math.round(valor * 100) / 100, aulasMes: aulas, ativo: p.ativo === 'Não' ? 'Não' : 'Sim'
    };
  });
  if (!lista.some(p => p.ativo === 'Sim')) throw new Error('Deixe pelo menos um plano ativo.');

  const ss = ss_();
  let sh = ss.getSheetByName(ABA_PLANOS);
  if (!sh) sh = ss.insertSheet(ABA_PLANOS);
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 5).clearContent();
  const tabela = [['Nome', 'Descricao', 'Valor_Mensal', 'Aulas_Mes', 'Ativo']]
    .concat(lista.map(p => [p.nome, p.descricao, p.valor, p.aulasMes, p.ativo]));
  sh.getRange(1, 1, tabela.length, 5).setValues(tabela);

  // nome mudou: os alunos acompanham
  const renomeados = lista.filter(p => p.anterior && p.anterior !== p.nome);
  let alunosAtualizados = 0;
  if (renomeados.length) {
    const L = linhas_(ABA.alunos);
    L.rows.forEach(a => {
      const r = renomeados.find(p => p.anterior === String(a.Plano || ''));
      if (r) { gravar_(L.sh, L.cab, a._r, { Plano: r.nome }); alunosAtualizados++; }
    });
  }
  return { salvos: lista.length, alunosAtualizados: alunosAtualizados };
}

/* ---------- Alunos ---------- */

function painelSalvarAluno(k, a) {
  checarChave_(k);
  a = a || {};
  const nome = limpa_(a.nome, 80);
  if (normaliza_(nome).length < 2) throw new Error('Digite o nome do aluno.');
  const fone = String(a.telefone || '').replace(/\D/g, '');
  if (fone && (fone.length < 10 || fone.length > 11)) throw new Error('Telefone com DDD, só números (10 ou 11 dígitos).');
  const L = linhas_(ABA.alunos);
  const campos = { Nome: nome, Telefone: fone, Plano: limpa_(a.plano, 40), Status: limpa_(a.status || 'Ativo', 20) };
  if (a.valor !== '' && a.valor != null) campos.Valor_Mensal = Number(a.valor) || 0;
  const trava = LockService.getScriptLock(); trava.waitLock(10000);
  try {
    let existente = a.id ? L.rows.find(r => r.ID_Aluno === a.id) : null;
    if (existente) {
      gravar_(L.sh, L.cab, existente._r, campos);
      return { id: a.id, novo: false };
    }
    const nums = L.rows.map(r => Number(String(r.ID_Aluno).replace(/^\D+/, ''))).filter(n => !isNaN(n));
    const id = 'A' + String((nums.length ? Math.max.apply(null, nums) : 0) + 1).padStart(2, '0');
    const r = primeiraLinhaVazia_(L.sh);
    L.sh.getRange(r, L.cab.indexOf('Telefone') + 1).setNumberFormat('@');
    campos.ID_Aluno = id;
    campos.Data_Inicio = hoje_();
    gravar_(L.sh, L.cab, r, campos);
    return { id: id, novo: true };
  } finally { trava.releaseLock(); }
}

/* ---------- Excluir aluno ---------- */

/** Abas que guardam linhas por aluno (coluna ID_Aluno) e o nome do campo no resumo. */
function abasDoAluno_() {
  return [
    { aba: ABA.envios, campo: 'envios' }, { aba: ABA.reg, campo: 'praticas' }, { aba: 'Frequencia', campo: 'aulas' },
    { aba: ABA_AVAL, campo: 'avaliacoes' }, { aba: ABA_MENS, campo: 'mensalidades' }
  ];
}

/** Linhas de cada aba que pertencem ao aluno (com o número da linha). */
function linhasDoAluno_(idAluno) {
  const out = {};
  abasDoAluno_().forEach(x => {
    const sh = ss_().getSheetByName(x.aba);
    out[x.campo] = sh ? linhas_(x.aba).rows.filter(r => String(r.ID_Aluno) === idAluno) : [];
  });
  return out;
}

/** O que será apagado: serve para a tela de confirmação. */
function painelResumoExclusao(k, idAluno) {
  checarChave_(k);
  const aluno = tabela_(ABA.alunos).find(a => a.ID_Aluno === idAluno);
  if (!aluno) throw new Error('Aluno não encontrado.');
  const l = linhasDoAluno_(idAluno);
  const r = { nome: String(aluno.Nome || '') };
  Object.keys(l).forEach(c => { r[c] = l[c].length; });
  return r;
}

/** Apaga o aluno e TODO o histórico dele (links, práticas, frequência, avaliações, mensalidades).
 *  Exige digitar o nome do aluno. Deixa um registro (sem telefone) na aba Excluidos. */
function painelExcluirAluno(k, idAluno, confirmacao) {
  checarChave_(k);
  return comTrava_(() => painelExcluirAluno_(idAluno, confirmacao));
}

function painelExcluirAluno_(idAluno, confirmacao) {
  const L = linhas_(ABA.alunos);
  const aluno = L.rows.find(a => a.ID_Aluno === idAluno);
  if (!aluno) throw new Error('Aluno não encontrado.');
  if (!normaliza_(confirmacao) || normaliza_(confirmacao) !== normaliza_(aluno.Nome)) {
    throw new Error('Digite o nome do aluno exatamente como está cadastrado para confirmar.');
  }
  const ss = ss_();
  const l = linhasDoAluno_(idAluno);
  const contagem = {};
  Object.keys(l).forEach(c => { contagem[c] = l[c].length; });

  // registro mínimo do que foi apagado (sem telefone nem e-mail)
  let log = ss.getSheetByName('Excluidos');
  if (!log) {
    log = ss.insertSheet('Excluidos');
    log.getRange(1, 1, 1, 8).setValues([['Data', 'ID_Aluno', 'Nome', 'Links', 'Praticas', 'Aulas', 'Avaliacoes', 'Mensalidades']]);
  }
  const r = primeiraLinhaVazia_(log);
  log.getRange(r, 1, 1, 8).setValues([[new Date(), idAluno, String(aluno.Nome || ''), contagem.envios, contagem.praticas, contagem.aulas, contagem.avaliacoes, contagem.mensalidades]]);

  // apaga de baixo para cima, para os números das linhas não mudarem no meio do caminho
  abasDoAluno_().forEach(x => {
    const sh = ss.getSheetByName(x.aba);
    if (!sh) return;
    l[x.campo].map(row => row._r).sort((a, b) => b - a).forEach(n => sh.deleteRow(n));
  });
  L.sh.deleteRow(aluno._r);
  return contagem;
}

/** oque = 'aparelhos' (libera aparelho novo) ou 'cadastro' (pede o cadastro de novo). */
function painelResetar(k, id, oque) {
  checarChave_(k);
  const L = linhas_(ABA.alunos);
  const a = L.rows.find(r => r.ID_Aluno === id);
  if (!a) throw new Error('Aluno não encontrado.');
  const campos = { Aparelhos: '' };
  if (oque === 'cadastro') campos.Cadastro_Completo = '';
  gravar_(L.sh, L.cab, a._r, campos);
  return true;
}

/* ---------- Endereço dos links (GitHub Pages) ---------- */

/** Define (ou volta ao padrão, se vazio) o endereço que vai nos links dos alunos. Mesma regra do menu da planilha. */
function painelDefinirEndereco(k, url) {
  checarChave_(k);
  const props = PropertiesService.getScriptProperties();
  url = String(url || '').trim();
  if (!url) { props.deleteProperty('URL_PUBLICA'); return urlPublica_(); }
  if (/^https:\/\/(www\.)?github\.com\//i.test(url)) {
    throw new Error('Esse é o endereço do repositório. Use o do Pages: https://SEU-USUARIO.github.io/NOME-DO-REPOSITORIO/');
  }
  if (!/^https:\/\/[^\s?#]+$/.test(url) || /\.git$/i.test(url)) {
    throw new Error('Endereço inválido: precisa começar com https://, sem espaços e sem .git no final.');
  }
  const final = url.replace(/\/?$/, '/');
  props.setProperty('URL_PUBLICA', final);
  return final;
}

/* ---------- Enviar link ---------- */

/** Link, mensagem e WhatsApp de um envio, SEMPRE com o endereço atual (urlPublica_). Assim, links enviados no
 *  passado (com endereço do Google, que pode dar "Não foi possível abrir o arquivo") são reenviados no endereço certo.
 *  O código (token) é o mesmo: o link antigo e o novo abrem a mesma aula. */
function dadosEnvio_(aluno, aula, validade, token) {
  const link = token ? urlPublica_() + '?t=' + token : '';
  const primeiro = String((aluno && aluno.Nome) || '').split(' ')[0];
  const msg = 'Olá, ' + primeiro + '! Seus exercícios da Aula ' + ((aula && aula.Numero) || '') + ' (' + ((aula && aula.Titulo) || '') + '). ' +
    'Marque cada dia que praticar. Link válido até ' + data_(validade) + ': ' + link;
  const fone = String((aluno && aluno.Telefone) || '').replace(/\D/g, '');
  return { link: link, mensagem: msg, whatsapp: link && fone ? 'https://wa.me/55' + fone + '?text=' + encodeURIComponent(msg) : '' };
}

function painelCriarEnvio(k, idAluno, idAula, validadeIso) {
  checarChave_(k);
  const aluno = tabela_(ABA.alunos).find(a => a.ID_Aluno === idAluno);
  const aula = tabela_(ABA.aulas).find(a => a.ID_Aula === idAula);
  if (!aluno) throw new Error('Escolha o aluno.');
  if (!aula) throw new Error('Escolha a aula.');
  const validade = deIso_(validadeIso);
  if (validade < hoje_()) throw new Error('A validade precisa ser hoje ou depois.');

  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const e = dadosEnvio_(aluno, aula, validade, token);
  const link = e.link, msg = e.mensagem, whats = e.whatsapp;

  const sh = ss_().getSheetByName(ABA.envios);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const trava = LockService.getScriptLock(); trava.waitLock(10000);
  try {
    const r = primeiraLinhaVazia_(sh);
    gravar_(sh, cab, r, { ID_Aluno: idAluno, ID_Aula: idAula, Data_Envio: new Date(), Validade: validade, Token: token, Link: link, WhatsApp: whats });
  } finally { trava.releaseLock(); }
  return { link: link, whatsapp: whats, mensagem: msg };
}

/* ---------- Frequência ---------- */

function painelPresenca(k, linha, status, motivo) {
  checarChave_(k);
  return comTrava_(() => painelPresenca_(linha, status, motivo));
}

function painelPresenca_(linha, status, motivo) {
  const validos = ['Agendada', 'Presente', 'Falta', 'Falta justificada', 'Reposição'];
  if (validos.indexOf(status) < 0) throw new Error('Situação inválida.');
  const sh = ss_().getSheetByName('Frequencia');
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  linha = Number(linha);
  if (!(linha >= 2) || linha > sh.getLastRow()) throw new Error('Aula não encontrada.');
  const ehFalta = st => st === 'Falta' || st === 'Falta justificada';
  const atual = sh.getRange(linha, 1, 1, cab.length).getValues()[0];
  const antes = atual[cab.indexOf('Presenca')];
  const msgs = [];

  // desfazendo uma falta: apaga a reposição automática que ainda não aconteceu
  if (ehFalta(antes) && !ehFalta(status)) {
    const dataFalta = atual[cab.indexOf('Data')];
    const obsRep = 'Reposição da aula de ' + Utilities.formatDate(new Date(dataFalta), FUSO, 'dd/MM');
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const l = v[i];
      if (l[cab.indexOf('ID_Aluno')] === atual[cab.indexOf('ID_Aluno')] && l[cab.indexOf('Tipo')] === 'Reposição' &&
          l[cab.indexOf('Obs')] === obsRep && l[cab.indexOf('Presenca')] === 'Agendada') {
        sh.getRange(i + 1, 1, 1, cab.length).clearContent();
        msgs.push('A reposição que tinha sido agendada foi cancelada.');
      }
    }
  }

  const campos = { Presenca: status };
  campos.Motivo_Falta = ehFalta(status) ? limpa_(motivo, 40) : '';
  if (!ehFalta(status) || !ehFalta(antes)) campos.Reposicao_Data = '';
  gravar_(sh, cab, linha, campos);
  SpreadsheetApp.flush();
  return msgs.concat(processarFaltas_(ss_()));
}

/** Agenda as próximas N aulas do aluno nos seus horários de "Aula" (ou "Ambos"). */
function painelAgendar(k, idAluno, n) {
  checarChave_(k);
  return comTrava_(() => painelAgendar_(idAluno, n));
}

function painelAgendar_(idAluno, n) {
  n = Math.max(1, Math.min(12, Number(n) || 4));
  const ss = ss_();
  const aulas = tabela_(ABA.aulas).sort((a, b) => a.Numero - b.Numero);
  const F = linhas_('Frequencia');
  const minhas = F.rows.filter(f => f.ID_Aluno === idAluno && f.Data instanceof Date);
  if (!tabela_(ABA.alunos).some(a => a.ID_Aluno === idAluno)) throw new Error('Escolha o aluno.');

  // próxima aula do programa
  const numDe = id => { const a = aulas.find(x => x.ID_Aula === id); return a ? Number(a.Numero) : 0; };
  let ultimoNum = minhas.filter(f => f.Tipo !== 'Reposição').reduce((m, f) => Math.max(m, numDe(f.ID_Aula)), 0);
  let inicio = minhas.reduce((m, f) => (f.Data > m ? f.Data : m), hoje_());

  // horários semanais de aula e exceções
  const disp = ss.getSheetByName('Disponibilidade').getRange(2, 1, 30, 10).getValues();
  const semanais = disp.filter(l => l[0] !== '' && semAcento_(l[4]) !== 'nao' && ['aula', 'ambos'].indexOf(semAcento_(l[3])) >= 0)
    .map(l => ({ dow: DIAS_SEMANA.map(semAcento_).indexOf(semAcento_(l[0])), ini: minutos_(l[1]), fim: minutos_(l[2]) }))
    .filter(s => s.dow >= 0 && s.ini != null && s.fim != null);
  if (!semanais.length) throw new Error('Cadastre pelo menos um horário com uso "Aula" ou "Ambos" na Disponibilidade.');
  const bloqueados = new Set(disp.filter(l => l[6] instanceof Date && semAcento_(l[7]) === 'indisponivel').map(l => iso_(l[6])));
  diasAusencia_(ss).forEach(x => bloqueados.add(x));
  const ocupados = new Set(F.rows.filter(f => f.Data instanceof Date && f.Presenca === 'Agendada')
    .map(f => iso_(f.Data) + ' ' + String(f.Horario instanceof Date ? Utilities.formatDate(f.Horario, FUSO, 'HH:mm') : f.Horario).slice(0, 5)));

  const criadas = [];
  const d = new Date(inicio);
  for (let guarda = 0; criadas.length < n && guarda < 120; guarda++) {
    d.setDate(d.getDate() + 1);
    if (bloqueados.has(iso_(d))) continue;
    const doDia = semanais.filter(s => s.dow === d.getDay()).sort((a, b) => a.ini - b.ini);
    for (const s of doDia) {
      let achou = false;
      for (let t = s.ini; t + DURACAO_AULA_MIN <= s.fim; t += DURACAO_AULA_MIN) {
        const chave = iso_(d) + ' ' + hhmm_(t);
        if (ocupados.has(chave)) continue;
        const prox = aulas.find(a => Number(a.Numero) > ultimoNum);
        if (!prox) break;
        ultimoNum = Number(prox.Numero);
        ocupados.add(chave);
        criadas.push({ data: new Date(d), aula: prox.ID_Aula, horario: hhmm_(t) + '–' + hhmm_(t + DURACAO_AULA_MIN) });
        achou = true;
        break;
      }
      if (achou) break; // uma aula por dia
    }
    if (!aulas.some(a => Number(a.Numero) > ultimoNum)) break;
  }
  if (!criadas.length) throw new Error('Não encontrei horário livre, ou o aluno já está na última aula do programa.');

  const sh = F.sh, cab = F.cab;
  criadas.forEach(c => {
    const r = primeiraLinhaVazia_(sh);
    gravar_(sh, cab, r, { Data: c.data, ID_Aluno: idAluno, ID_Aula: c.aula, Presenca: 'Agendada', Horario: c.horario, Tipo: 'Regular' });
    sh.getRange(r, cab.indexOf('Data') + 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(r, cab.indexOf('Nome') + 1).setFormula('=IF(B' + r + '="","",IFERROR(INDEX(Alunos!B:B,MATCH(B' + r + ',Alunos!A:A,0)),""))');
    sh.getRange(r, cab.indexOf('Aula') + 1).setFormula('=IF(D' + r + '="","",IFERROR("Aula "&INDEX(Aulas!B:B,MATCH(D' + r + ',Aulas!A:A,0))&" – "&INDEX(Aulas!C:C,MATCH(D' + r + ',Aulas!A:A,0)),""))');
  });
  return criadas.map(c => ({ data: iso_(c.data), aula: c.aula, horario: c.horario }));
}

/* ---------- Disponibilidade ---------- */

function painelSalvarDisp(k, semanais, excecoes) {
  checarChave_(k);
  if ((semanais || []).length > 30 || (excecoes || []).length > 30) {
    throw new Error('A Disponibilidade aceita até 30 horários semanais e 30 exceções. Remova alguns antes de salvar.');
  }
  return comTrava_(() => painelSalvarDisp_(semanais, excecoes));
}

function painelSalvarDisp_(semanais, excecoes) {
  const dias = DIAS_SEMANA;
  const hora = h => { const m = minutos_(h); if (m == null) throw new Error('Horário inválido: ' + h); return hhmm_(m); };
  const S = (semanais || []).slice(0, 30).map(s => {
    const dia = dias.find(x => semAcento_(x) === semAcento_(s.dia));
    if (!dia) throw new Error('Dia inválido: ' + s.dia);
    const ini = hora(s.ini), fim = hora(s.fim);
    if (minutos_(fim) <= minutos_(ini)) throw new Error(dia + ': o fim precisa ser depois do início.');
    const uso = ['Aula', 'Reposição', 'Ambos'].indexOf(s.uso) >= 0 ? s.uso : 'Ambos';
    return [dia, ini, fim, uso, s.ativo === 'Não' ? 'Não' : 'Sim'];
  });
  const E = (excecoes || []).slice(0, 30).map(x => {
    const sit = x.situacao === 'Disponível extra' ? 'Disponível extra' : 'Indisponível';
    return [deIso_(x.data), sit, sit === 'Disponível extra' ? hora(x.ini) : '', sit === 'Disponível extra' ? hora(x.fim) : ''];
  });
  const sh = ss_().getSheetByName('Disponibilidade');
  sh.getRange(2, 1, 30, 5).clearContent();
  sh.getRange(2, 7, 30, 4).clearContent();
  sh.getRange(2, 2, 30, 2).setNumberFormat('@');
  sh.getRange(2, 9, 30, 2).setNumberFormat('@');
  if (S.length) sh.getRange(2, 1, S.length, 5).setValues(S);
  if (E.length) {
    sh.getRange(2, 7, E.length, 4).setValues(E);
    sh.getRange(2, 7, E.length, 1).setNumberFormat('dd/MM/yyyy');
  }
  return true;
}

/* ---------- Aulas e exercícios ---------- */

/** Salva uma aula e a lista de exercícios dela. Exercícios removidos saem da lista (os registros antigos continuam). */
function painelSalvarAula(k, aula, exercicios) {
  checarChave_(k);
  aula = aula || {};
  const titulo = limpa_(aula.titulo, 80);
  if (!titulo) throw new Error('Dê um título para a aula.');
  const apostila = String(aula.apostila || '').trim();
  if (apostila && !/^https:\/\/\S+$/.test(apostila)) throw new Error('O link da apostila precisa começar com https://');
  const lista = (exercicios || []).slice(0, 20).map((x, i) => ({
    id: String(x.id || ''), titulo: limpa_(x.titulo, 80), descricao: limpa_(x.descricao, 400),
    tempo: Math.max(0, Math.min(120, Math.round(Number(x.tempo) || 0))), ordem: i + 1
  })).filter(x => x.titulo);
  if (!lista.length) throw new Error('A aula precisa de pelo menos um exercício.');

  const trava = LockService.getScriptLock(); trava.waitLock(15000);
  try {
    const A = linhas_(ABA.aulas);
    const a = A.rows.find(r => r.ID_Aula === aula.id);
    if (!a) throw new Error('Aula não encontrada.');
    gravar_(A.sh, A.cab, a._r, { Titulo: titulo, Conteudo: limpa_(aula.conteudo, 500), Tarefa: limpa_(aula.tarefa, 300), Link_Apostila: apostila });

    const E = linhas_(ABA.ex);
    const daAula = E.rows.filter(r => r.ID_Aula === aula.id);
    const prefixo = 'EX' + String(aula.id).replace(/^AU/, '') + '-';
    let maior = daAula.reduce((m, r) => Math.max(m, Number(String(r.ID_Exercicio).split('-').pop()) || 0), 0);
    // não reaproveita IDs de exercícios removidos que ainda aparecem em Registros ou Avaliacoes
    const jaUsados = [];
    if (ss_().getSheetByName(ABA.reg)) linhas_(ABA.reg).rows.forEach(r => String(r.IDs_Exercicios || '').split(',').forEach(s => jaUsados.push(s.trim())));
    if (abaAvaliacoes_(false)) linhas_(ABA_AVAL).rows.forEach(r => jaUsados.push(String(r.ID_Exercicio || '').trim()));
    jaUsados.forEach(id => { if (id.indexOf(prefixo) === 0) maior = Math.max(maior, Number(id.slice(prefixo.length)) || 0); });
    const mantidos = new Set();
    lista.forEach(x => {
      const existente = x.id && daAula.find(r => r.ID_Exercicio === x.id);
      const campos = { Ordem: x.ordem, Titulo: x.titulo, Descricao: x.descricao, Tempo_min: x.tempo };
      if (existente) {
        gravar_(E.sh, E.cab, existente._r, campos);
        mantidos.add(existente.ID_Exercicio);
      } else {
        maior += 1;
        campos.ID_Exercicio = prefixo + maior;
        campos.ID_Aula = aula.id;
        gravar_(E.sh, E.cab, primeiraLinhaVazia_(E.sh), campos);
        mantidos.add(campos.ID_Exercicio);
      }
    });
    daAula.filter(r => !mantidos.has(r.ID_Exercicio)).forEach(r => E.sh.getRange(r._r, 1, 1, E.cab.length).clearContent());
  } finally { trava.releaseLock(); }
  return true;
}

/** Cria a próxima aula do programa (com um exercício para você editar). */
function painelNovaAula(k) {
  checarChave_(k);
  const trava = LockService.getScriptLock(); trava.waitLock(15000);
  try {
    const A = linhas_(ABA.aulas);
    const n = A.rows.reduce((m, r) => Math.max(m, Number(r.Numero) || 0), 0) + 1;
    const id = 'AU' + String(n).padStart(2, '0');
    gravar_(A.sh, A.cab, primeiraLinhaVazia_(A.sh), { ID_Aula: id, Numero: n, Titulo: 'Nova aula', Conteudo: '', Tarefa: '', Link_Apostila: '' });
    const E = linhas_(ABA.ex);
    gravar_(E.sh, E.cab, primeiraLinhaVazia_(E.sh), { ID_Exercicio: 'EX' + String(n).padStart(2, '0') + '-1', ID_Aula: id, Ordem: 1, Titulo: 'Primeiro exercício', Descricao: '', Tempo_min: 5 });
    return { id: id, numero: n };
  } finally { trava.releaseLock(); }
}
