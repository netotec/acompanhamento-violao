/**
 * Configurar.gs – monta a planilha do Violão (rode UMA vez).
 * No editor do Apps Script, escolha a função configurarPlanilha e clique em Executar.
 * Cuidado: apaga e recria as abas abaixo. Não rode de novo depois de começar a usar.
 */
function configurarPlanilha() {
  soDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet() || ss_();

  // Proteção: se a planilha já tem dados, apagar tudo exige digitar uma palavra de confirmação.
  const alunosAtual = ss.getSheetByName('Alunos');
  const jaTemDados = !!alunosAtual && alunosAtual.getLastRow() > 1;
  if (jaTemDados) {
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt('ATENÇÃO: isto APAGA todos os dados',
      'A planilha já está em uso. Rodar a configuração apaga e recria TODAS as abas ' +
      '(alunos, aulas, frequência, registros). Não dá para desfazer.\n\n' +
      'Para continuar mesmo assim, digite APAGAR TUDO. Para cancelar, feche esta janela.', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK || r.getResponseText().trim() !== 'APAGAR TUDO') {
      ui.alert('Nada foi alterado.');
      return;
    }
  }
  PropertiesService.getScriptProperties().setProperty('PLANILHA_ID', ss.getId());
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  const COR_CAB = '#9a4a17', COR_EDITA = '#fff2cc', COR_AUTO = '#ededed';
  console.log('Montando a planilha...');
  const abas = [
    { nome: "Como usar", linhas: 28, colunas: 1,
      texto: [], formatos: {},
      amarelo: [], cinza: [],
      larguras: {"A": 825}, notas: {},
      validacoes: [], ocultas: [],
      dados: [
      ["Violão – Acompanhamento de exercícios"],
      [''],
      ["Depois de cada aula presencial:"],
      ["1. Na aba Envios, preencha uma linha: ID_Aluno, ID_Aula e Validade (data da próxima aula)."],
      ["2. Menu Violão > Gerar links pendentes. O token, o link e a mensagem de WhatsApp são criados."],
      ["3. Clique no link da coluna WhatsApp e envie para a aluna."],
      [''],
      ["A aluna abre o link, marca os exercícios do dia e envia. Cada envio vira uma linha na aba Registros, com o nome dela."],
      ["Depois da Validade, o link para de funcionar."],
      [''],
      ["Células amarelas: você preenche. Abas Registros e colunas Token/Link/WhatsApp: preenchidas pelo sistema, não edite."],
      ["A aba Resumo mostra o desenvolvimento de cada aluno."],
      [''],
      ["Frequência (aba Frequencia):"],
      ["• As aulas do mês já estão agendadas (1x por semana, quarta das 18h às 19h). Depois de cada aula, troque Agendada por Presente, Falta, Falta justificada ou Reposição."],
      ["• Em caso de falta, escolha o motivo e, se for repor, preencha a data da reposição."],
      ["• Se ela faltar, não precisa gerar link novo: aumente a Validade na aba Envios e o mesmo link continua valendo até a nova aula."],
      ["• A aba Resumo mostra presenças, faltas, % de frequência e a próxima aula de cada aluno."],
      ["• Reposição automática: ao marcar Falta ou Falta justificada, o sistema agenda a reposição no primeiro horário livre da aba Disponibilidade e estende o link até lá."],
      [''],
      ["Disponibilidade (aba Disponibilidade):"],
      ["• Cadastre seus horários livres por dia da semana e marque o uso: Aula, Reposição ou Ambos."],
      ["• Para um dia em que você não pode, ou um horário extra só numa data, use a parte de Exceções (colunas G a J)."],
      [''],
      ["Acesso da aluna:"],
      ["• 1º link: ela faz o cadastro (nome, WhatsApp, violão, estilos e objetivo) e só depois vê os exercícios. Os dados vão para a aba Alunos."],
      ["• Links seguintes: ela entra só digitando o nome (primeiro nome ou completo), sem senha."],
      ["• Trava por aparelho: o link só abre no celular/computador que ela usou no cadastro (até 2 aparelhos). Se ela trocar de celular ou de navegador, apague a célula Aparelhos dela na aba Alunos para liberar."]
      ] },
    { nome: "Alunos", linhas: 2, colunas: 16,
      texto: [3], formatos: {},
      amarelo: ["B2:B10", "C2:C10", "D2:D10", "E2:E10", "F2:F10", "G2:G10"], cinza: ["H2:H10", "I2:I10", "J2:J10", "K2:K10", "L2:L10", "M2:M10", "N2:N10", "O2:O10", "P2:P10"],
      larguras: {"A": 75, "B": 225, "C": 112, "D": 97, "E": 165, "F": 97, "G": 75, "H": 120, "I": 97, "J": 97, "K": 97, "L": 97, "M": 97, "N": 97, "O": 97, "P": 225}, notas: {"B1": "Pode deixar só o primeiro nome. No cadastro ela confirma o nome completo.", "C1": "DDD + número, só dígitos (ex.: 94991234567). Usado no link do WhatsApp.", "F1": "Aluna atual: 0 (cortesia). Próximos: 240.", "L1": "Estilos e músicas que ela gosta (vem do cadastro).", "O1": "Fica 'Sim' quando ela faz o cadastro no 1º link. Apague para pedir o cadastro de novo.", "P1": "Aparelhos liberados (máx. 2: celular e computador). Preenchido sozinho no 1º acesso. Apague esta célula para liberar um aparelho novo."},
      validacoes: [], ocultas: [],
      dados: [
      ["ID_Aluno","Nome","Telefone","Data_Inicio","Plano","Valor_Mensal","Status","Email","Tem_Violao","Tipo_Violao","Experiencia","Estilos","Objetivo","Data_Cadastro","Cadastro_Completo","Aparelhos"],
      ["A01","Aluna 1 (trocar pelo nome)","94900000000","2026-10-06","Com acompanhamento",0,"Ativo",'','','','','','','','','']
      ] },
    { nome: "Aulas", linhas: 11, colunas: 6,
      texto: [], formatos: {},
      amarelo: ["F2:F11"], cinza: [],
      larguras: {"A": 67, "B": 60, "C": 195, "D": 375, "E": 300, "F": 225}, notas: {"F1": "Link da apostila (ex.: a da Aula 1). Aparece como botão na página da aluna."},
      validacoes: [], ocultas: [],
      dados: [
      ["ID_Aula","Numero","Titulo","Conteudo","Tarefa","Link_Apostila"],
      ["AU01",1,"Primeiros passos","Partes do violão, postura, nome das cordas, numeração dos dedos e casas, afinar com aplicativo","Exercício da aranha, 15 min por dia",''],
      ["AU02",2,"Em e Am","Primeiros acordes: Em e Am. Batida só para baixo, contando 1-2-3-4","Trocar Em ↔ Am sem parar o tempo",''],
      ["AU03",3,"G e D","Acordes G e D","Trocar G ↔ D e revisar Em/Am",''],
      ["AU04",4,"Acorde C","Acorde C e a sequência G – C – D","Tocar \"Parabéns pra Você\" (G, C, D)",''],
      ["AU05",5,"Batida para baixo e para cima","Batida ↓ ↓ ↑ ↑ ↓ ↑","Aplicar a batida em G – Em – C – D",''],
      ["AU06",6,"E e A","Acordes E e A","Sequência A – D – E",''],
      ["AU07",7,"Revisão e 1ª música","Revisão dos 7 acordes e trocas mais rápidas","Primeira música completa escolhida pela aluna",''],
      ["AU08",8,"Dedilhado","Polegar no baixo e dedos i-m-a nas cordas de cima","Dedilhar Em – Am – C – G",''],
      ["AU09",9,"Lendo cifra","Como ler cifra em sites e apps","Tirar uma música sozinha com a cifra",''],
      ["AU10",10,"Revisão e avaliação","Revisão geral e avaliação","Tocar 2 músicas do início ao fim",'']
      ] },
    { nome: "Exercicios", linhas: 18, colunas: 6,
      texto: [], formatos: {},
      amarelo: [], cinza: [],
      larguras: {"A": 90, "B": 67, "C": 52, "D": 195, "E": 450, "F": 75}, notas: {},
      validacoes: [], ocultas: [],
      dados: [
      ["ID_Exercicio","ID_Aula","Ordem","Titulo","Descricao","Tempo_min"],
      ["EX01-1","AU01",1,"Afinar o violão","Afinar as 6 cordas com o aplicativo (GuitarTuna ou Cifra Club)",2],
      ["EX01-2","AU01",2,"Cordas soltas","Tocar da 6ª à 1ª corda falando Mi – Lá – Ré – Sol – Si – Mi, e voltar",3],
      ["EX01-3","AU01",3,"Aranha 1-2-3-4","Dedos 1-2-3-4 nas casas 1 a 4, da 6ª à 1ª corda e de volta. Devagar, som limpo, 60 bpm",8],
      ["EX01-4","AU01",4,"Postura e descanso","Conferir a postura no espelho e relaxar as mãos",2],
      ["EX02-1","AU02",1,"Montar Em e Am","Montar cada acorde e tocar corda por corda, conferindo se todas soam",5],
      ["EX02-2","AU02",2,"Troca Em ↔ Am","Trocar entre os acordes contando 1-2-3-4, sem parar o tempo",8],
      ["EX03-1","AU03",1,"Troca G ↔ D","Trocar entre G e D contando 1-2-3-4",8],
      ["EX03-2","AU03",2,"Revisão Em/Am","Repetir a troca Em ↔ Am",4],
      ["EX04-1","AU04",1,"Sequência G – C – D","Tocar a sequência em loop, 4 batidas por acorde",6],
      ["EX04-2","AU04",2,"Parabéns pra Você","Tocar a música inteira com G, C e D",6],
      ["EX05-1","AU05",1,"Batida ↓ ↓ ↑ ↑ ↓ ↑","Fazer a batida com as cordas abafadas, depois num acorde",5],
      ["EX05-2","AU05",2,"G – Em – C – D","Aplicar a batida na sequência",8],
      ["EX06-1","AU06",1,"Sequência A – D – E","Trocar entre A, D e E com a batida",10],
      ["EX07-1","AU07",1,"Primeira música","Tocar a música escolhida do início ao fim",15],
      ["EX08-1","AU08",1,"Dedilhado Em – Am – C – G","Polegar no baixo, depois i-m-a",12],
      ["EX09-1","AU09",1,"Música pela cifra","Escolher uma música e tirar sozinha pela cifra",15],
      ["EX10-1","AU10",1,"Duas músicas completas","Tocar 2 músicas do início ao fim sem parar",15]
      ] },
    { nome: "Envios", linhas: 2, colunas: 7,
      texto: [], formatos: {"3": "dd/mm/yyyy", "4": "dd/mm/yyyy"},
      amarelo: ["A2:A2", "B2:B2", "C2:C2", "D2:D2"], cinza: [],
      larguras: {"A": 75, "B": 67, "C": 97, "D": 97, "E": 150, "F": 337, "G": 97}, notas: {"D1": "Data da próxima aula. Depois dela o link para de abrir.", "E1": "Preenchido pelo menu Violão > Gerar links pendentes."},
      validacoes: [], ocultas: [],
      dados: [
      ["ID_Aluno","ID_Aula","Data_Envio","Validade","Token","Link","WhatsApp"],
      ["A01","AU01",new Date(2026,9,10),new Date(2026,9,14),'','','']
      ] },
    { nome: "Disponibilidade", linhas: 35, colunas: 26,
      texto: [2, 3, 9, 10, 26], formatos: {"7": "dd/mm/yyyy"},
      amarelo: ["A2:A31", "B2:B31", "C2:C31", "D2:D31", "E2:E31", "G2:G31", "H2:H31", "I2:I31", "J2:J31"], cinza: [],
      larguras: {"A": 97, "B": 90, "C": 75, "D": 90, "E": 60, "F": 22, "G": 90, "H": 127, "I": 90, "J": 75, "K": 97, "L": 97, "M": 97, "N": 97, "O": 97, "P": 97, "Q": 97, "R": 97, "S": 97, "T": 97, "U": 97, "V": 97, "W": 97, "X": 97, "Y": 97, "Z": 97}, notas: {"A1": "Segunda a Domingo.", "B1": "Escolha na lista (HH:MM).", "D1": "Aula = horário das aulas fixas. Reposição = só para repor. Ambos = os dois.", "E1": "Não = ignora esta linha sem apagar.", "G1": "Uma data específica.", "H1": "Indisponível = não marcar nada neste dia. Disponível extra = horário a mais só neste dia.", "I1": "Só para 'Disponível extra'."},
      validacoes: [{"r": "A2:A31", "list": ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]}, {"r": "D2:D31", "list": ["Aula", "Reposição", "Ambos"]}, {"r": "E2:E31", "list": ["Sim", "Não"]}, {"r": "H2:H31", "list": ["Indisponível", "Disponível extra"]}, {"r": "B2:C31", "ref": "Z2:Z35"}, {"r": "I2:J31", "ref": "Z2:Z35"}], ocultas: ["Z"],
      dados: [
      ["Dia_Semana","Hora_Inicio","Hora_Fim","Uso","Ativo",'',"Data","Situacao","Inicio_Extra","Fim_Extra",'','','','','','','','','','','','','','','',"Horas"],
      ["Quarta","18:00","19:00","Aula","Sim",'','','','','','','','','','','','','','','','','','','','',"06:00"],
      ["Sexta","18:00","19:00","Reposição","Sim",'','','','','','','','','','','','','','','','','','','','',"06:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"07:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"07:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"08:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"08:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"09:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"09:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"10:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"10:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"11:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"11:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"12:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"12:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"13:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"13:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"14:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"14:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"15:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"15:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"16:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"16:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"17:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"17:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"18:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"18:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"19:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"19:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"20:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"20:30"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"21:00"],
      ['','','','','','','','','','','','','','','','','','','','','','','','','',"21:30"],
      ["Horários semanais (A–E): quando você pode dar aula. Exceções (G–J): datas em que você não pode, ou um horário extra só naquele dia.",'','','','','','','','','','','','','','','','','','','','','','','','',"22:00"],
      ["A reposição automática usa o primeiro horário livre de Uso 'Reposição' ou 'Ambos' depois da falta, sem bater com outra aula já agendada.",'','','','','','','','','','','','','','','','','','','','','','','','',"22:30"]
      ] },
    { nome: "Frequencia", linhas: 41, colunas: 11,
      texto: [10], formatos: {"1": "dd/mm/yyyy", "8": "dd/mm/yyyy"},
      amarelo: ["A2:A41", "B2:B41", "D2:D41", "F2:F41", "G2:G41", "H2:H41", "I2:I41", "J2:J41", "K2:K41"], cinza: [],
      larguras: {"A": 90, "B": 75, "C": 195, "D": 67, "E": 210, "F": 127, "G": 165, "H": 112, "I": 225, "J": 97, "K": 90}, notas: {"C1": "Automático a partir do ID_Aluno.", "E1": "Automático a partir do ID_Aula.", "F1": "Agendada (ainda não aconteceu), Presente, Falta, Falta justificada ou Reposição.", "G1": "Preencha só quando houver falta.", "H1": "Se for repor a aula, a data combinada.", "J1": "Aulas: quarta 18:00–19:00. Reposição: sexta 18:00–19:00.", "K1": "Regular = aula do calendário. Reposição = criada automaticamente quando há falta."},
      validacoes: [{"r": "F2:F41", "list": ["Agendada", "Presente", "Falta", "Falta justificada", "Reposição"]}, {"r": "G2:G41", "list": ["Doença", "Trabalho", "Estudos", "Viagem", "Compromisso familiar", "Transporte", "Esqueceu", "Outro"]}, {"r": "K2:K41", "list": ["Regular", "Reposição"]}], ocultas: [],
      dados: [
      ["Data","ID_Aluno","Nome","ID_Aula","Aula","Presenca","Motivo_Falta","Reposicao_Data","Obs","Horario","Tipo"],
      [new Date(2026,9,6),"A01","=IF(B2=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B2,Alunos!A:A,0)),\"\"))","AU01","=IF(D2=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D2,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D2,Aulas!A:A,0)),\"\"))","Presente",'','',"Aula-teste",'',"Regular"],
      [new Date(2026,9,14),"A01","=IF(B3=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B3,Alunos!A:A,0)),\"\"))","AU02","=IF(D3=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D3,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D3,Aulas!A:A,0)),\"\"))","Agendada",'','','',"18:00–19:00","Regular"],
      [new Date(2026,9,21),"A01","=IF(B4=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B4,Alunos!A:A,0)),\"\"))","AU03","=IF(D4=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D4,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D4,Aulas!A:A,0)),\"\"))","Agendada",'','','',"18:00–19:00","Regular"],
      [new Date(2026,9,28),"A01","=IF(B5=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B5,Alunos!A:A,0)),\"\"))","AU04","=IF(D5=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D5,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D5,Aulas!A:A,0)),\"\"))","Agendada",'','','',"18:00–19:00","Regular"],
      [new Date(2026,10,4),"A01","=IF(B6=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B6,Alunos!A:A,0)),\"\"))","AU05","=IF(D6=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D6,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D6,Aulas!A:A,0)),\"\"))","Agendada",'','',"Fechamento do mês: avaliação","18:00–19:00","Regular"],
      ['','',"=IF(B7=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B7,Alunos!A:A,0)),\"\"))",'',"=IF(D7=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D7,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D7,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B8=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B8,Alunos!A:A,0)),\"\"))",'',"=IF(D8=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D8,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D8,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B9=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B9,Alunos!A:A,0)),\"\"))",'',"=IF(D9=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D9,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D9,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B10=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B10,Alunos!A:A,0)),\"\"))",'',"=IF(D10=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D10,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D10,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B11=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B11,Alunos!A:A,0)),\"\"))",'',"=IF(D11=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D11,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D11,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B12=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B12,Alunos!A:A,0)),\"\"))",'',"=IF(D12=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D12,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D12,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B13=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B13,Alunos!A:A,0)),\"\"))",'',"=IF(D13=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D13,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D13,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B14=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B14,Alunos!A:A,0)),\"\"))",'',"=IF(D14=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D14,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D14,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B15=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B15,Alunos!A:A,0)),\"\"))",'',"=IF(D15=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D15,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D15,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B16=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B16,Alunos!A:A,0)),\"\"))",'',"=IF(D16=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D16,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D16,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B17=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B17,Alunos!A:A,0)),\"\"))",'',"=IF(D17=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D17,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D17,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B18=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B18,Alunos!A:A,0)),\"\"))",'',"=IF(D18=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D18,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D18,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B19=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B19,Alunos!A:A,0)),\"\"))",'',"=IF(D19=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D19,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D19,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B20=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B20,Alunos!A:A,0)),\"\"))",'',"=IF(D20=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D20,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D20,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B21=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B21,Alunos!A:A,0)),\"\"))",'',"=IF(D21=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D21,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D21,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B22=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B22,Alunos!A:A,0)),\"\"))",'',"=IF(D22=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D22,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D22,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B23=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B23,Alunos!A:A,0)),\"\"))",'',"=IF(D23=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D23,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D23,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B24=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B24,Alunos!A:A,0)),\"\"))",'',"=IF(D24=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D24,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D24,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B25=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B25,Alunos!A:A,0)),\"\"))",'',"=IF(D25=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D25,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D25,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B26=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B26,Alunos!A:A,0)),\"\"))",'',"=IF(D26=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D26,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D26,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B27=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B27,Alunos!A:A,0)),\"\"))",'',"=IF(D27=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D27,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D27,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B28=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B28,Alunos!A:A,0)),\"\"))",'',"=IF(D28=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D28,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D28,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B29=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B29,Alunos!A:A,0)),\"\"))",'',"=IF(D29=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D29,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D29,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B30=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B30,Alunos!A:A,0)),\"\"))",'',"=IF(D30=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D30,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D30,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B31=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B31,Alunos!A:A,0)),\"\"))",'',"=IF(D31=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D31,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D31,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B32=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B32,Alunos!A:A,0)),\"\"))",'',"=IF(D32=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D32,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D32,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B33=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B33,Alunos!A:A,0)),\"\"))",'',"=IF(D33=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D33,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D33,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B34=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B34,Alunos!A:A,0)),\"\"))",'',"=IF(D34=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D34,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D34,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B35=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B35,Alunos!A:A,0)),\"\"))",'',"=IF(D35=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D35,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D35,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B36=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B36,Alunos!A:A,0)),\"\"))",'',"=IF(D36=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D36,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D36,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B37=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B37,Alunos!A:A,0)),\"\"))",'',"=IF(D37=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D37,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D37,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B38=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B38,Alunos!A:A,0)),\"\"))",'',"=IF(D38=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D38,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D38,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B39=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B39,Alunos!A:A,0)),\"\"))",'',"=IF(D39=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D39,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D39,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B40=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B40,Alunos!A:A,0)),\"\"))",'',"=IF(D40=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D40,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D40,Aulas!A:A,0)),\"\"))",'','','','','',''],
      ['','',"=IF(B41=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(B41,Alunos!A:A,0)),\"\"))",'',"=IF(D41=\"\",\"\",IFERROR(\"Aula \"&INDEX(Aulas!B:B,MATCH(D41,Aulas!A:A,0))&\" – \"&INDEX(Aulas!C:C,MATCH(D41,Aulas!A:A,0)),\"\"))",'','','','','','']
      ] },
    { nome: "Registros", linhas: 1, colunas: 11,
      texto: [], formatos: {},
      amarelo: [], cinza: [],
      larguras: {"A": 135, "B": 97, "C": 75, "D": 180, "E": 67, "F": 165, "G": 255, "H": 67, "I": 82, "J": 300, "K": 135}, notas: {"A1": "Preenchida pela página da aluna. Não edite esta aba."},
      validacoes: [], ocultas: [],
      dados: [
      ["Data_Hora","Token","ID_Aluno","Nome","ID_Aula","IDs_Exercicios","Exercicios","Minutos","Dificuldade","Comentario","Aparelho"]
      ] },
    { nome: "Resumo", linhas: 5, colunas: 12,
      texto: [], formatos: {"5": "dd/mm/yyyy hh:mm", "11": "0%", "12": "dd/mm/yyyy"},
      amarelo: ["A2:A5"], cinza: [],
      larguras: {"A": 75, "B": 225, "C": 150, "D": 120, "E": 97, "F": 112, "G": 120, "H": 82, "I": 60, "J": 135, "K": 97, "L": 97}, notas: {"A1": "Digite o ID do aluno; o resto calcula sozinho.", "G1": "Aulas do calendário que já aconteceram (reposições não contam aqui, mas a presença nelas conta)."},
      validacoes: [], ocultas: [],
      dados: [
      ["ID_Aluno","Nome","Práticas registradas","Minutos no total","Última prática","Aulas enviadas","Aulas realizadas","Presenças","Faltas","Faltas justificadas","% Frequência","Próxima aula"],
      ["A01","=IF(A2=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(A2,Alunos!A:A,0)),\"\"))","=IF(A2=\"\",\"\",COUNTIF(Registros!C:C,A2))","=IF(A2=\"\",\"\",SUMIF(Registros!C:C,A2,Registros!H:H))","=IF(A2=\"\",\"\",IF(COUNTIF(Registros!C:C,A2)=0,\"-\",MAXIFS(Registros!A:A,Registros!C:C,A2)))","=IF(A2=\"\",\"\",COUNTIF(Envios!A:A,A2))","=IF(A2=\"\",\"\",COUNTIFS(Frequencia!B:B,A2,Frequencia!K:K,\"Regular\",Frequencia!F:F,\"<>Agendada\",Frequencia!F:F,\"<>\"))","=IF(A2=\"\",\"\",COUNTIFS(Frequencia!B:B,A2,Frequencia!F:F,\"Presente\")+COUNTIFS(Frequencia!B:B,A2,Frequencia!F:F,\"Reposição\"))","=IF(A2=\"\",\"\",COUNTIFS(Frequencia!B:B,A2,Frequencia!F:F,\"Falta\"))","=IF(A2=\"\",\"\",COUNTIFS(Frequencia!B:B,A2,Frequencia!F:F,\"Falta justificada\"))","=IF(A2=\"\",\"\",IF(G2=0,\"-\",MIN(1,H2/G2)))","=IF(A2=\"\",\"\",IF(COUNTIFS(Frequencia!B:B,A2,Frequencia!F:F,\"Agendada\")=0,\"-\",MINIFS(Frequencia!A:A,Frequencia!B:B,A2,Frequencia!F:F,\"Agendada\")))"],
      ['',"=IF(A3=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(A3,Alunos!A:A,0)),\"\"))","=IF(A3=\"\",\"\",COUNTIF(Registros!C:C,A3))","=IF(A3=\"\",\"\",SUMIF(Registros!C:C,A3,Registros!H:H))","=IF(A3=\"\",\"\",IF(COUNTIF(Registros!C:C,A3)=0,\"-\",MAXIFS(Registros!A:A,Registros!C:C,A3)))","=IF(A3=\"\",\"\",COUNTIF(Envios!A:A,A3))","=IF(A3=\"\",\"\",COUNTIFS(Frequencia!B:B,A3,Frequencia!K:K,\"Regular\",Frequencia!F:F,\"<>Agendada\",Frequencia!F:F,\"<>\"))","=IF(A3=\"\",\"\",COUNTIFS(Frequencia!B:B,A3,Frequencia!F:F,\"Presente\")+COUNTIFS(Frequencia!B:B,A3,Frequencia!F:F,\"Reposição\"))","=IF(A3=\"\",\"\",COUNTIFS(Frequencia!B:B,A3,Frequencia!F:F,\"Falta\"))","=IF(A3=\"\",\"\",COUNTIFS(Frequencia!B:B,A3,Frequencia!F:F,\"Falta justificada\"))","=IF(A3=\"\",\"\",IF(G3=0,\"-\",MIN(1,H3/G3)))","=IF(A3=\"\",\"\",IF(COUNTIFS(Frequencia!B:B,A3,Frequencia!F:F,\"Agendada\")=0,\"-\",MINIFS(Frequencia!A:A,Frequencia!B:B,A3,Frequencia!F:F,\"Agendada\")))"],
      ['',"=IF(A4=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(A4,Alunos!A:A,0)),\"\"))","=IF(A4=\"\",\"\",COUNTIF(Registros!C:C,A4))","=IF(A4=\"\",\"\",SUMIF(Registros!C:C,A4,Registros!H:H))","=IF(A4=\"\",\"\",IF(COUNTIF(Registros!C:C,A4)=0,\"-\",MAXIFS(Registros!A:A,Registros!C:C,A4)))","=IF(A4=\"\",\"\",COUNTIF(Envios!A:A,A4))","=IF(A4=\"\",\"\",COUNTIFS(Frequencia!B:B,A4,Frequencia!K:K,\"Regular\",Frequencia!F:F,\"<>Agendada\",Frequencia!F:F,\"<>\"))","=IF(A4=\"\",\"\",COUNTIFS(Frequencia!B:B,A4,Frequencia!F:F,\"Presente\")+COUNTIFS(Frequencia!B:B,A4,Frequencia!F:F,\"Reposição\"))","=IF(A4=\"\",\"\",COUNTIFS(Frequencia!B:B,A4,Frequencia!F:F,\"Falta\"))","=IF(A4=\"\",\"\",COUNTIFS(Frequencia!B:B,A4,Frequencia!F:F,\"Falta justificada\"))","=IF(A4=\"\",\"\",IF(G4=0,\"-\",MIN(1,H4/G4)))","=IF(A4=\"\",\"\",IF(COUNTIFS(Frequencia!B:B,A4,Frequencia!F:F,\"Agendada\")=0,\"-\",MINIFS(Frequencia!A:A,Frequencia!B:B,A4,Frequencia!F:F,\"Agendada\")))"],
      ['',"=IF(A5=\"\",\"\",IFERROR(INDEX(Alunos!B:B,MATCH(A5,Alunos!A:A,0)),\"\"))","=IF(A5=\"\",\"\",COUNTIF(Registros!C:C,A5))","=IF(A5=\"\",\"\",SUMIF(Registros!C:C,A5,Registros!H:H))","=IF(A5=\"\",\"\",IF(COUNTIF(Registros!C:C,A5)=0,\"-\",MAXIFS(Registros!A:A,Registros!C:C,A5)))","=IF(A5=\"\",\"\",COUNTIF(Envios!A:A,A5))","=IF(A5=\"\",\"\",COUNTIFS(Frequencia!B:B,A5,Frequencia!K:K,\"Regular\",Frequencia!F:F,\"<>Agendada\",Frequencia!F:F,\"<>\"))","=IF(A5=\"\",\"\",COUNTIFS(Frequencia!B:B,A5,Frequencia!F:F,\"Presente\")+COUNTIFS(Frequencia!B:B,A5,Frequencia!F:F,\"Reposição\"))","=IF(A5=\"\",\"\",COUNTIFS(Frequencia!B:B,A5,Frequencia!F:F,\"Falta\"))","=IF(A5=\"\",\"\",COUNTIFS(Frequencia!B:B,A5,Frequencia!F:F,\"Falta justificada\"))","=IF(A5=\"\",\"\",IF(G5=0,\"-\",MIN(1,H5/G5)))","=IF(A5=\"\",\"\",IF(COUNTIFS(Frequencia!B:B,A5,Frequencia!F:F,\"Agendada\")=0,\"-\",MINIFS(Frequencia!A:A,Frequencia!B:B,A5,Frequencia!F:F,\"Agendada\")))"]
      ] },
    { nome: "Mensalidades", linhas: 2, colunas: 5,
      texto: [2], formatos: {"5": "dd/MM/yyyy"},
      amarelo: ["A2:A2", "B2:B2", "C2:C2", "D2:D2", "E2:E2"], cinza: [],
      larguras: {"A": 75, "B": 97, "C": 97, "D": 97, "E": 112}, notas: {"D1": "Sim, Não ou Cortesia."},
      validacoes: [{"r": "D2:D200", "list": ["Sim", "Não", "Cortesia"]}], ocultas: [],
      dados: [
      ["ID_Aluno","Mes","Valor","Pago","Data_Pagamento"],
      ["A01","2026-10",0,"Cortesia",'']
      ] },
  ];

  // reaproveita as abas que já existem (não apaga nada no início)
  abas.forEach((a, i) => {
    console.log('Aba ' + (i + 1) + ' de ' + abas.length + ': ' + a.nome);
    let sh = ss.getSheetByName(a.nome);
    if (!sh) { sh = ss.insertSheet(a.nome); SpreadsheetApp.flush(); }
    sh.clear();
    sh.clearNotes();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
    sh.setConditionalFormatRules([]);
    if (sh.getMaxColumns() > 0) sh.showColumns(1, sh.getMaxColumns());
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
    if (sh.getMaxColumns() < a.colunas) sh.insertColumnsAfter(sh.getMaxColumns(), a.colunas - sh.getMaxColumns());
    if (sh.getMaxRows() < Math.max(a.linhas, 100)) sh.insertRowsAfter(sh.getMaxRows(), Math.max(a.linhas, 100) - sh.getMaxRows());
    a.texto.forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
    Object.keys(a.formatos).forEach(c => sh.getRange(2, Number(c), sh.getMaxRows() - 1, 1).setNumberFormat(a.formatos[c]));
    sh.getRange(1, 1, a.linhas, a.colunas).setValues(a.dados);
    sh.getRange(1, 1, Math.max(a.linhas, 50), a.colunas).setFontFamily('Arial');
    if (a.nome === 'Como usar') {
      sh.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor(COR_CAB);
      sh.setColumnWidth(1, 900);
      sh.getRange(1, 1, a.linhas, 1).setWrap(true);
    } else {
      sh.getRange(1, 1, 1, a.colunas).setBackground(COR_CAB).setFontColor('#ffffff').setFontWeight('bold');
      sh.setFrozenRows(1);
      if (a.amarelo.length) sh.getRangeList(a.amarelo).setBackground(COR_EDITA);
      if (a.cinza.length) sh.getRangeList(a.cinza).setBackground(COR_AUTO);
      Object.keys(a.larguras).forEach(col => {
        const n = col.charCodeAt(0) - 64;
        if (col.length === 1 && n <= 16 && !(a.nome === 'Disponibilidade' && n > 10)) sh.setColumnWidth(n, a.larguras[col]);
      });
    }
    Object.keys(a.notas).forEach(r => sh.getRange(r).setNote(a.notas[r]));
    a.validacoes.forEach(v => {
      const regra = v.list
        ? SpreadsheetApp.newDataValidation().requireValueInList(v.list, true)
        : SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange(v.ref), true);
      sh.getRange(v.r).setDataValidation(regra.setAllowInvalid(true).build());
    });
    a.ocultas.forEach(col => sh.hideColumns(col.charCodeAt(0) - 64));
  });
  // remove abas que sobraram (ex.: "Página1"), sem travar se não conseguir
  SpreadsheetApp.flush();
  const nomes = abas.map(a => a.nome);
  ss.getSheets().forEach(sh => {
    if (nomes.indexOf(sh.getName()) < 0) {
      try { ss.deleteSheet(sh); } catch (e) { console.log('Não consegui apagar a aba ' + sh.getName() + '. Pode apagar à mão.'); }
    }
  });

  // cores da presença
  const fr = ss.getSheetByName('Frequencia');
  const faixa = fr.getRange('F2:F' + fr.getMaxRows());
  const cores = { 'Presente': '#c6efce', 'Reposição': '#ddebf7', 'Falta': '#ffc7ce', 'Falta justificada': '#ffeb9c' };
  fr.setConditionalFormatRules(Object.keys(cores).map(t =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(cores[t]).setRanges([faixa]).build()));
  ss.getSheetByName('Resumo').getRange('K2:K30').setNumberFormat('0%');

  ss.setActiveSheet(ss.getSheetByName('Como usar'));
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert('Planilha configurada! Recarregue a página para ver o menu Violão.'); } catch (e) {}
}
