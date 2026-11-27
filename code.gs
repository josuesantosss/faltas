/**
 * SISTEMA DE BUSCA ATIVA ESCOLAR - BACKEND (GOOGLE APPS SCRIPT)
 * Integra as abas: "Banco de Dados", "FALTANTES" e "Historico_Envios"
 */

const CONFIG = {
  SPREADSHEET_ID: '', // Opcional: cole o ID da planilha se o script não estiver vinculado a ela
  ABA_BANCO: 'Banco de Dados',
  ABA_FALTANTES: 'FALTANTES',
  ABA_HISTORICO: 'Historico_Envios',
  DDD_PADRAO: '17'
};

/**
 * Obtém a planilha ativa ou abre pelo ID informado em CONFIG.SPREADSHEET_ID
 */
function obterPlanilha() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Planilha não vinculada. Preencha CONFIG.SPREADSHEET_ID com o ID da planilha.');
  }
  return ss;
}

/**
 * GET: Carrega turmas, alunos do Banco de Dados, lista de FALTANTES e histórico
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getDados';

    if (action === 'getDados') {
      const dados = carregarTudo();
      return respostaJSON({ status: 'ok', ...dados });
    }

    return respostaJSON({ status: 'erro', mensagem: 'Ação GET inválida.' });
  } catch (err) {
    return respostaJSON({ status: 'erro', mensagem: err.toString() });
  }
}

/**
 * POST: Salva faltas da semana, registra envio de WhatsApp, limpa semana ou atualiza telefone
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    switch (payload.action) {
      case 'salvarFaltas':
        const qtd = salvarFaltasNaAbaFaltantes(payload);
        return respostaJSON({ status: 'ok', mensagem: qtd + ' registro(s) salvo(s) na aba FALTANTES.' });

      case 'registrarEnvio':
        registrarEnvioWhatsApp(payload);
        return respostaJSON({ status: 'ok', mensagem: 'Envio registrado com sucesso.' });

      case 'atualizarTelefone':
        atualizarContatoBancoEFaltantes(payload);
        return respostaJSON({ status: 'ok', mensagem: 'Telefone atualizado com sucesso.' });

      case 'limparFaltantes':
        limparAbaFaltantes();
        return respostaJSON({ status: 'ok', mensagem: 'Lista da aba FALTANTES limpa para nova semana.' });

      default:
        return respostaJSON({ status: 'erro', mensagem: 'Ação POST não reconhecida.' });
    }
  } catch (err) {
    return respostaJSON({ status: 'erro', mensagem: err.toString() });
  }
}

/**
 * Carrega Banco de Dados, FALTANTES e Histórico
 */
function carregarTudo() {
  const ss = obterPlanilha();
  const alunos = lerBancoDeDados(ss);
  const faltantes = lerAbaFaltantes(ss, alunos);
  const historico = lerHistorico(ss);

  const turmasSet = new Set();
  alunos.forEach(a => turmasSet.add(a.turma));

  return {
    turmas: Array.from(turmasSet).sort(),
    alunos: alunos,
    faltantes: faltantes,
    historico: historico
  };
}

/**
 * Lê a aba "Banco de Dados" respeitando o cabeçalho de 25 colunas (A até Y)
 * 0:TURMA | 3:Nº | 4:Nome do Aluno | 5:RA | 12:Situação | 14:Condições especiais
 * 16:NÚMERO DE FALTAS | 17-18:Par1 | 19-20:Par2 | 21-22:Par3 | 23-24:Par4
 */
function lerBancoDeDados(ss) {
  ss = ss || obterPlanilha();
  const sheet = ss.getSheetByName(CONFIG.ABA_BANCO);
  if (!sheet) throw new Error('Aba "' + CONFIG.ABA_BANCO + '" não encontrada.');

  const valores = sheet.getDataRange().getDisplayValues();
  const alunos = [];
  const regexTurma = /^[1-9][A-Z]$/i;

  for (let i = 0; i < valores.length; i++) {
    const l = valores[i];
    const turma = limparTexto(l[0]).toUpperCase();
    const nome = limparTexto(l[4]);

    // Ignora linhas em branco e cabeçalhos repetidos
    if (!regexTurma.test(turma) || !nome || nome.toUpperCase() === 'NOME DO ALUNO') {
      continue;
    }

    const contatos = extrairContatos([
      { slot: 1, tipo: l[17], tel: l[18] },
      { slot: 2, tipo: l[19], tel: l[20] },
      { slot: 3, tipo: l[21], tel: l[22] },
      { slot: 4, tipo: l[23], tel: l[24] }
    ]);

    alunos.push({
      linhaBanco: i + 1,
      turma: turma,
      numero: limparTexto(l[3]),
      nome: nome,
      ra: limparTexto(l[5]),
      situacao: limparTexto(l[12]).toUpperCase() || 'ATIVO',
      especial: limparTexto(l[14]),
      contatos: contatos
    });
  }

  return alunos;
}

/**
 * Lê a aba "FALTANTES" (a partir da linha 3)
 * Colunas: A:Turma | B:Nome | C:Tipo1 | D:Tel1 | E:Faltas | F:Msg1
 *          G:Tipo2 | H:Tel2 | I:Msg2 | J:Tipo3 | K:Tel3 | L:Msg3
 *          M:Tipo4 | N:Tel4 | O:Msg4/BuscaAtiva | P:Data Registro
 */
function lerAbaFaltantes(ss, alunosBanco) {
  ss = ss || obterPlanilha();
  alunosBanco = alunosBanco || lerBancoDeDados(ss);

  const sheet = ss.getSheetByName(CONFIG.ABA_FALTANTES);
  if (!sheet || sheet.getLastRow() < 3) return [];

  const ultimaLinha = sheet.getLastRow();
  const valores = sheet.getRange(3, 1, ultimaLinha - 2, 16).getDisplayValues();

  // Mapa rápido para cruzar com o Banco de Dados (caso falte turma ou telefone na aba FALTANTES)
  const mapaBanco = {};
  alunosBanco.forEach(a => {
    mapaBanco[normalizarChave(a.nome)] = a;
  });

  const lista = [];

  for (let i = 0; i < valores.length; i++) {
    const l = valores[i];
    const nome = limparTexto(l[1]);
    if (!nome) continue;

    const alunoRef = mapaBanco[normalizarChave(nome)] || null;
    const turma = limparTexto(l[0]).toUpperCase() || (alunoRef ? alunoRef.turma : 'SEM TURMA');
    const faltas = parseInt(limparTexto(l[4]), 10) || 0;
    const dataRegistro = limparTexto(l[15]);

    const contatosFaltantes = extrairContatos([
      { slot: 1, tipo: l[2], tel: l[3], statusEnvio: limparTexto(l[5]) },
      { slot: 2, tipo: l[6], tel: l[7], statusEnvio: limparTexto(l[8]) },
      { slot: 3, tipo: l[9], tel: l[10], statusEnvio: limparTexto(l[11]) },
      { slot: 4, tipo: l[12], tel: l[13], statusEnvio: limparTexto(l[14]) }
    ]);

    // Se não houver contatos na linha de FALTANTES, tenta aproveitar do Banco de Dados
    const contatosFinais = contatosFaltantes.length > 0
      ? contatosFaltantes
      : (alunoRef ? alunoRef.contatos : []);

    lista.push({
      linhaFaltantes: i + 3,
      linhaBanco: alunoRef ? alunoRef.linhaBanco : null,
      turma: turma,
      nome: nome,
      ra: alunoRef ? alunoRef.ra : '',
      faltas: faltas,
      dataRegistro: dataRegistro,
      contatos: contatosFinais
    });
  }

  return lista;
}

/**
 * Salva os registros enviados pelo lançamento de faltas na aba FALTANTES
 */
function salvarFaltasNaAbaFaltantes(payload) {
  const ss = obterPlanilha();
  let sheetFaltantes = ss.getSheetByName(CONFIG.ABA_FALTANTES);

  if (!sheetFaltantes) {
    sheetFaltantes = ss.insertSheet(CONFIG.ABA_FALTANTES);
    sheetFaltantes.appendRow(['FALTANTES', 'FALTAS DA SEMANA']);
    sheetFaltantes.appendRow([
      'TURMA', 'NOME DE ALUNO', 'TIPO', 'TELEFONE 1', 'Faltas', 'MENSAGEM',
      'TIPO', 'TELEFONE 2', 'MENSAGEM', 'TIPO', 'TELEFONE 3', 'MENSAGEM',
      'TIPO', 'TELEFONE 4', 'BUSCA ATIVA', 'DATA REGISTRO'
    ]);
  }

  const registros = payload.registros || [];
  const dataRegistro = payload.data || Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');

  if (registros.length === 0) return 0;

  // Busca dados completos (parentescos e telefones) no Banco de Dados
  const alunosBanco = lerBancoDeDados(ss);
  const mapaBanco = {};
  alunosBanco.forEach(a => {
    mapaBanco[normalizarChave(a.turma + '|' + a.nome)] = a;
    mapaBanco[normalizarChave(a.nome)] = a;
  });

  // Mapeia alunos já existentes na aba FALTANTES para atualizar sem duplicar
  const existentes = {};
  const ultimaLinha = sheetFaltantes.getLastRow();
  if (ultimaLinha >= 3) {
    const dadosAtuais = sheetFaltantes.getRange(3, 1, ultimaLinha - 2, 2).getDisplayValues();
    for (let i = 0; i < dadosAtuais.length; i++) {
      const nomeExist = limparTexto(dadosAtuais[i][1]);
      if (nomeExist) {
        existentes[normalizarChave(nomeExist)] = i + 3;
      }
    }
  }

  let salvos = 0;

  registros.forEach(reg => {
    const nomeCompleto = limparTexto(reg.nome);
    const turma = limparTexto(reg.turma).toUpperCase();
    const faltas = parseInt(reg.faltas, 10);

    if (!nomeCompleto || ![3, 4, 5].includes(faltas)) return;

    const info = mapaBanco[normalizarChave(turma + '|' + nomeCompleto)] || mapaBanco[normalizarChave(nomeCompleto)];
    const c = info ? info.contatos : [];

    const getC = (idx) => c[idx] || { tipo: '', telefoneOriginal: '' };
    const c1 = getC(0), c2 = getC(1), c3 = getC(2), c4 = getC(3);

    const novaLinha = [
      turma,
      nomeCompleto,
      c1.tipo,
      c1.telefoneOriginal,
      faltas,
      c1.telefoneOriginal ? 'ENVIAR' : '',
      c2.tipo,
      c2.telefoneOriginal,
      c2.telefoneOriginal ? 'ENVIAR' : '',
      c3.tipo,
      c3.telefoneOriginal,
      c3.telefoneOriginal ? 'ENVIAR' : '',
      c4.tipo,
      c4.telefoneOriginal,
      c4.telefoneOriginal ? 'ENVIAR' : '',
      dataRegistro
    ];

    const chaveNome = normalizarChave(nomeCompleto);
    if (existentes[chaveNome]) {
      sheetFaltantes.getRange(existentes[chaveNome], 1, 1, 16).setValues([novaLinha]);
    } else {
      const proxLinha = Math.max(3, sheetFaltantes.getLastRow() + 1);
      sheetFaltantes.getRange(proxLinha, 1, 1, 16).setValues([novaLinha]);
      existentes[chaveNome] = proxLinha;
    }

    salvos++;
  });

  return salvos;
}

/**
 * Marca o status de envio na aba FALTANTES e grava na aba Historico_Envios
 */
function registrarEnvioWhatsApp(dados) {
  const ss = obterPlanilha();
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const agoraCurto = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm');

  // 1. Atualiza status na coluna correspondente da aba FALTANTES
  const sheetFaltantes = ss.getSheetByName(CONFIG.ABA_FALTANTES);
  const linha = parseInt(dados.linhaFaltantes, 10);
  const slot = parseInt(dados.slot, 10);

  if (sheetFaltantes && linha >= 3 && slot >= 1 && slot <= 4) {
    const colStatusPorSlot = { 1: 6, 2: 9, 3: 12, 4: 15 }; // Cols F, I, L, O
    sheetFaltantes.getRange(linha, colStatusPorSlot[slot]).setValue('ENVIADO (' + agoraCurto + ')');
  }

  // 2. Registra no Histórico
  let sheetHist = ss.getSheetByName(CONFIG.ABA_HISTORICO);
  if (!sheetHist) {
    sheetHist = ss.insertSheet(CONFIG.ABA_HISTORICO);
    sheetHist.appendRow([
      'Data/Hora', 'Turma', 'Nome do Aluno', 'Faltas',
      'Parentesco', 'Telefone', 'Mensagem'
    ]);
    sheetHist.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e2e8f0');
  }

  sheetHist.appendRow([
    agora,
    dados.turma || '',
    dados.aluno || '',
    dados.faltas || '',
    dados.contatoTipo || '',
    dados.telefone || '',
    dados.mensagem || ''
  ]);
}

/**
 * Atualiza o Telefone 1 no Banco de Dados e na linha correspondente de FALTANTES
 */
function atualizarContatoBancoEFaltantes(dados) {
  const ss = obterPlanilha();
  const tipo = limparTexto(dados.tipoContato) || 'Responsável';
  const tel = String(dados.telefone || '').replace(/\D/g, '');

  const linhaBanco = parseInt(dados.linhaBanco, 10);
  if (linhaBanco >= 2) {
    const sheetBanco = ss.getSheetByName(CONFIG.ABA_BANCO);
    sheetBanco.getRange(linhaBanco, 18).setValue(tipo); // Col R
    sheetBanco.getRange(linhaBanco, 19).setValue(tel);  // Col S
  }

  const linhaFaltantes = parseInt(dados.linhaFaltantes, 10);
  if (linhaFaltantes >= 3) {
    const sheetFaltantes = ss.getSheetByName(CONFIG.ABA_FALTANTES);
    sheetFaltantes.getRange(linhaFaltantes, 3).setValue(tipo);     // Col C
    sheetFaltantes.getRange(linhaFaltantes, 4).setValue(tel);      // Col D
    sheetFaltantes.getRange(linhaFaltantes, 6).setValue('ENVIAR'); // Col F
  }
}

/**
 * Limpa a aba FALTANTES a partir da linha 3 (mantém os cabeçalhos das linhas 1 e 2)
 */
function limparAbaFaltantes() {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(CONFIG.ABA_FALTANTES);
  if (!sheet || sheet.getLastRow() < 3) return;
  sheet.getRange(3, 1, sheet.getLastRow() - 2, 16).clearContent();
}

/**
 * Lê últimos 100 envios do histórico
 */
function lerHistorico(ss) {
  ss = ss || obterPlanilha();
  const sheet = ss.getSheetByName(CONFIG.ABA_HISTORICO);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const inicio = Math.max(2, sheet.getLastRow() - 99);
  const qtd = sheet.getLastRow() - inicio + 1;
  const dados = sheet.getRange(inicio, 1, qtd, 7).getDisplayValues();

  return dados.reverse().map(l => ({
    dataHora: l[0],
    turma: l[1],
    aluno: l[2],
    faltas: l[3],
    contatoTipo: l[4],
    telefone: l[5],
    mensagem: l[6]
  }));
}

/**
 * Helpers de limpeza e normalização
 */
function limparTexto(val) {
  const s = String(val || '').trim();
  if (!s || s.startsWith('#')) return '';
  return s;
}

function normalizarChave(str) {
  return String(str || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function extrairContatos(listaBruta) {
  const resultado = [];
  listaBruta.forEach(item => {
    const telNorm = normalizarTelefone(item.tel);
    if (!telNorm) return;

    resultado.push({
      slot: item.slot,
      tipo: limparTexto(item.tipo) || 'Responsável',
      telefoneOriginal: limparTexto(item.tel),
      telefoneWhats: telNorm.whatsapp,
      ehFixo: telNorm.ehFixo,
      statusEnvio: item.statusEnvio || 'ENVIAR'
    });
  });
  return resultado;
}

function normalizarTelefone(val) {
  const s = limparTexto(val);
  let digitos = s.replace(/\D/g, '');
  if (!digitos || digitos.length < 8) return null;

  if (digitos.length >= 12 && digitos.startsWith('55')) {
    digitos = digitos.substring(2);
  }
  if (digitos.length === 8 || digitos.length === 9) {
    digitos = CONFIG.DDD_PADRAO + digitos;
  }

  const ehFixo = digitos.length === 10 && /^[2-5]/.test(digitos.substring(2, 3));
  return { whatsapp: '55' + digitos, ehFixo: ehFixo };
}

function respostaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
