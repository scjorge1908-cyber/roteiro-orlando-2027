/**
 * ROTEIRO ORLANDO 2027 — Codigo.gs
 * Sistema de votação em grupo e montagem automática de roteiro
 * (Tampa → Orlando → Miami → Tampa, 08/02 a 19/02/2027)
 *
 * Banco de dados: a própria planilha onde este script está vinculado.
 * Primeiro uso: rode a função setup() uma vez (menu Executar).
 */

const ABA = {
  CONFIG: 'Config',
  DIAS: 'Dias',
  LOCAIS: 'Locais',
  ATRACOES: 'Atracoes',
  PESSOAS: 'Pessoas',
  VOTOS: 'Votos',
  ROTEIRO: 'Roteiro'
};

const CABECALHOS = {
  Config: ['chave', 'valor'],
  Dias: ['data', 'cidade', 'tipo', 'local_id', 'travado', 'observacao'],
  Locais: ['id', 'nome', 'grupo', 'regiao', 'tipo', 'dia_inteiro', 'ordem', 'descricao', 'site', 'foto_busca', 'foto_url', 'foto_credito'],
  Atracoes: ['id', 'local_id', 'area', 'ordem_area', 'nome', 'categoria', 'altura_min_cm', 'descricao', 'foto_busca', 'foto_url', 'foto_credito', 'ativo'],
  Pessoas: ['id', 'nome', 'pin', 'altura_cm', 'criado_em'],
  Votos: ['pessoa_id', 'tipo_item', 'item_id', 'valor', 'atualizado_em'],
  Roteiro: ['data', 'ordem', 'cidade', 'tipo', 'local_id', 'area', 'item_id', 'nome', 'pontos', 'alerta', 'gerado_em']
};

/* Cidades atendidas por dia: em um dia "Orlando" o sistema pode escolher parques de Orlando e a NASA (Space Coast). */
const REGIOES_POR_CIDADE = {
  'Orlando': ['Orlando', 'Space Coast'],
  'Tampa': ['Tampa'],
  'Miami': ['Miami']
};

const TIPOS_AUTOMATICOS = ['PARQUE', 'RETORNO', 'MIAMI'];

/* =========================================================
 *  WEB APP
 * ========================================================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Roteiro Orlando 2027')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎢 Roteiro')
    .addItem('1. Configurar planilha (setup)', 'setup')
    .addItem('2. Gerar roteiro agora', 'menuGerarRoteiro')
    .addItem('Mostrar PIN do organizador', 'menuMostrarPinAdmin')
    .addToUi();
}

function menuGerarRoteiro() {
  const r = gerarRoteiro_();
  SpreadsheetApp.getUi().alert('Roteiro gerado: ' + r.length + ' linhas na aba Roteiro.');
}

function menuMostrarPinAdmin() {
  SpreadsheetApp.getUi().alert('PIN do organizador: ' + getConfig_('pin_admin'));
}

/* =========================================================
 *  SETUP — cria abas e carrega o catálogo real
 * ========================================================= */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || ss_();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  Object.keys(CABECALHOS).forEach(function (nome) {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    sh.getRange(1, 1, sh.getMaxRows(), CABECALHOS[nome].length).setNumberFormat('@');
    const atual = sh.getRange(1, 1, 1, CABECALHOS[nome].length).getDisplayValues()[0];
    if (atual.join('') === '') {
      sh.getRange(1, 1, 1, CABECALHOS[nome].length).setValues([CABECALHOS[nome]]).setFontWeight('bold').setBackground('#0b1f4b').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });

  if (contarLinhas_(ABA.CONFIG) === 0) {
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    escreverLinhas_(ABA.CONFIG, [
      ['nome_viagem', 'Roteiro Orlando 2027'],
      ['data_chegada', '2027-02-08'],
      ['data_ultimo_dia', '2027-02-18'],
      ['voo_volta', 'Madrugada de 19/02/2027 saindo de Tampa (TPA)'],
      ['pin_admin', pin]
    ]);
  }
  if (contarLinhas_(ABA.DIAS) === 0) escreverLinhas_(ABA.DIAS, seedDias_());
  if (contarLinhas_(ABA.LOCAIS) === 0) escreverLinhas_(ABA.LOCAIS, seedLocais_());
  if (contarLinhas_(ABA.ATRACOES) === 0) escreverLinhas_(ABA.ATRACOES, seedAtracoes_());

  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);

  try {
    SpreadsheetApp.getUi().alert('Planilha configurada!\n\nPIN do organizador: ' + getConfig_('pin_admin') + '\n\nAgora publique: Implantar → Nova implantação → App da Web (Executar como: Eu / Acesso: Qualquer pessoa).');
  } catch (e) {
    Logger.log('PIN do organizador: ' + getConfig_('pin_admin'));
  }
}

/* =========================================================
 *  API PÚBLICA (chamada pelo Index.html)
 * ========================================================= */

/** Cria as abas automaticamente no primeiro acesso, caso o setup() ainda não tenha sido rodado. */
function garantirSetup_() {
  const planilha = ss_();
  if (planilha.getSheetByName(ABA.CONFIG) && planilha.getSheetByName(ABA.ATRACOES)) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!planilha.getSheetByName(ABA.CONFIG) || !planilha.getSheetByName(ABA.ATRACOES)) setup();
  } finally {
    lock.releaseLock();
  }
}

function api_dados() {
  try {
    garantirSetup_();
    const config = {};
    ler_(ABA.CONFIG).forEach(function (c) { if (c.chave !== 'pin_admin') config[c.chave] = c.valor; });
    const pessoas = ler_(ABA.PESSOAS).map(function (p) {
      return { id: p.id, nome: p.nome, altura_cm: p.altura_cm, tem_pin: p.pin !== '' };
    });
    return {
      ok: true,
      dados: {
        config: config,
        dias: ler_(ABA.DIAS),
        locais: ler_(ABA.LOCAIS),
        atracoes: ler_(ABA.ATRACOES).filter(function (a) { return a.ativo !== 'NAO'; }),
        pessoas: pessoas,
        votos: ler_(ABA.VOTOS),
        roteiro: ler_(ABA.ROTEIRO)
      }
    };
  } catch (e) {
    return { ok: false, erro: String(e.message || e) };
  }
}

function api_cadastrar(nome, pin, altura) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    nome = String(nome || '').trim();
    pin = String(pin || '').trim();
    altura = String(altura || '').replace(/\D/g, '');
    if (nome.length < 2) return { ok: false, erro: 'Digite o nome.' };
    if (!/^\d{4}$/.test(pin)) return { ok: false, erro: 'O PIN deve ter 4 números.' };
    const pessoas = ler_(ABA.PESSOAS);
    const existe = pessoas.some(function (p) { return p.nome.toLowerCase() === nome.toLowerCase(); });
    if (existe) return { ok: false, erro: 'Esse nome já está cadastrado. Volte e toque no seu nome.' };
    const id = 'P' + Utilities.getUuid().slice(0, 8).toUpperCase();
    escreverLinhas_(ABA.PESSOAS, [[id, nome, pin, altura, agora_()]]);
    return { ok: true, pessoa: { id: id, nome: nome, altura_cm: altura, tem_pin: true } };
  } finally {
    lock.releaseLock();
  }
}

function api_login(id, pin) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    pin = String(pin || '').trim();
    if (!/^\d{4}$/.test(pin)) return { ok: false, erro: 'O PIN deve ter 4 números.' };
    const sh = aba_(ABA.PESSOAS);
    const linhas = ler_(ABA.PESSOAS);
    const i = linhas.findIndex(function (p) { return p.id === id; });
    if (i < 0) return { ok: false, erro: 'Pessoa não encontrada.' };
    const p = linhas[i];
    if (p.pin === '') {
      sh.getRange(i + 2, CABECALHOS.Pessoas.indexOf('pin') + 1).setValue(pin);
    } else if (p.pin !== pin) {
      return { ok: false, erro: 'PIN incorreto.' };
    }
    return { ok: true, pessoa: { id: p.id, nome: p.nome, altura_cm: p.altura_cm, tem_pin: true } };
  } finally {
    lock.releaseLock();
  }
}

function api_atualizarAltura(id, pin, altura) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const i = validarPessoa_(id, pin);
    if (i < 0) return { ok: false, erro: 'Sessão inválida. Entre de novo.' };
    altura = String(altura || '').replace(/\D/g, '');
    aba_(ABA.PESSOAS).getRange(i + 2, CABECALHOS.Pessoas.indexOf('altura_cm') + 1).setValue(altura);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * valor: 2 = Quero muito | 1 = Tanto faz | -1 = Não quero | 0 = apagar voto
 */
function api_votar(id, pin, tipoItem, itemId, valor) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (validarPessoa_(id, pin) < 0) return { ok: false, erro: 'Sessão inválida. Entre de novo.' };
    if (['LOCAL', 'ATRACAO'].indexOf(tipoItem) < 0) return { ok: false, erro: 'Tipo inválido.' };
    valor = Number(valor);
    if ([2, 1, -1, 0].indexOf(valor) < 0) return { ok: false, erro: 'Voto inválido.' };

    const sh = aba_(ABA.VOTOS);
    const votos = ler_(ABA.VOTOS);
    const i = votos.findIndex(function (v) { return v.pessoa_id === id && v.tipo_item === tipoItem && v.item_id === itemId; });
    if (valor === 0) {
      if (i >= 0) sh.deleteRow(i + 2);
    } else if (i >= 0) {
      sh.getRange(i + 2, 4, 1, 2).setValues([[String(valor), agora_()]]);
    } else {
      escreverLinhas_(ABA.VOTOS, [[id, tipoItem, itemId, String(valor), agora_()]]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** Guarda a foto encontrada no Wikimedia Commons para não buscar de novo. */
function api_salvarFoto(tipoItem, itemId, url, credito) {
  if (!/^https:\/\/upload\.wikimedia\.org\//.test(String(url))) return { ok: false };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false };
  try {
    const nomeAba = tipoItem === 'LOCAL' ? ABA.LOCAIS : ABA.ATRACOES;
    const cab = CABECALHOS[nomeAba];
    const linhas = ler_(nomeAba);
    const i = linhas.findIndex(function (r) { return r.id === itemId; });
    if (i < 0 || linhas[i].foto_url !== '') return { ok: true };
    aba_(nomeAba).getRange(i + 2, cab.indexOf('foto_url') + 1, 1, 2).setValues([[url, String(credito || '')]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** Ações do organizador (exigem o PIN de 6 dígitos da aba Config). */
function api_admin(pinAdmin, acao, payload) {
  if (String(pinAdmin) !== getConfig_('pin_admin')) return { ok: false, erro: 'PIN do organizador incorreto.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    payload = payload || {};
    if (acao === 'checar') return { ok: true };

    if (acao === 'salvarDia') {
      const sh = aba_(ABA.DIAS);
      const dias = ler_(ABA.DIAS);
      const i = dias.findIndex(function (d) { return d.data === payload.data; });
      if (i < 0) return { ok: false, erro: 'Dia não encontrado.' };
      sh.getRange(i + 2, 2, 1, 5).setValues([[
        String(payload.cidade || ''),
        String(payload.tipo || ''),
        String(payload.local_id || ''),
        payload.travado ? 'SIM' : 'NAO',
        String(payload.observacao || '')
      ]]);
      return { ok: true, dias: ler_(ABA.DIAS) };
    }

    if (acao === 'gerarRoteiro') {
      return { ok: true, roteiro: gerarRoteiro_() };
    }

    if (acao === 'salvarPessoa') {
      const sh = aba_(ABA.PESSOAS);
      const pessoas = ler_(ABA.PESSOAS);
      const i = pessoas.findIndex(function (p) { return p.id === payload.id; });
      if (i < 0) return { ok: false, erro: 'Pessoa não encontrada.' };
      const nome = String(payload.nome || pessoas[i].nome).trim();
      const altura = String(payload.altura_cm || '').replace(/\D/g, '');
      const pin = payload.resetPin ? '' : pessoas[i].pin;
      sh.getRange(i + 2, 2, 1, 3).setValues([[nome, pin, altura]]);
      return { ok: true };
    }

    if (acao === 'excluirPessoa') {
      const pessoas = ler_(ABA.PESSOAS);
      const i = pessoas.findIndex(function (p) { return p.id === payload.id; });
      if (i < 0) return { ok: false, erro: 'Pessoa não encontrada.' };
      aba_(ABA.PESSOAS).deleteRow(i + 2);
      const shV = aba_(ABA.VOTOS);
      const votos = ler_(ABA.VOTOS);
      for (let k = votos.length - 1; k >= 0; k--) {
        if (votos[k].pessoa_id === payload.id) shV.deleteRow(k + 2);
      }
      return { ok: true };
    }

    return { ok: false, erro: 'Ação desconhecida.' };
  } catch (e) {
    return { ok: false, erro: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 *  MOTOR DO ROTEIRO
 * ========================================================= */

function gerarRoteiro_() {
  const dias = ler_(ABA.DIAS).slice().sort(function (a, b) { return a.data < b.data ? -1 : 1; });
  const locais = ler_(ABA.LOCAIS);
  const atracoes = ler_(ABA.ATRACOES).filter(function (a) { return a.ativo !== 'NAO'; });
  const pessoas = ler_(ABA.PESSOAS);
  const votos = ler_(ABA.VOTOS);
  const linhas = montarRoteiro_(dias, locais, atracoes, pessoas, votos, agora_());

  const sh = aba_(ABA.ROTEIRO);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CABECALHOS.Roteiro.length).clearContent();
  if (linhas.length) escreverLinhas_(ABA.ROTEIRO, linhas);
  return ler_(ABA.ROTEIRO);
}

/** Lógica pura (sem planilha) — também usada nos testes. Retorna linhas no formato da aba Roteiro. */
function montarRoteiro_(dias, locais, atracoes, pessoas, votos, carimbo) {
  const pontos = {};
  const temVoto = {};
  votos.forEach(function (v) {
    const k = v.tipo_item + ':' + v.item_id;
    pontos[k] = (pontos[k] || 0) + Number(v.valor || 0);
    temVoto[k] = true;
  });
  const ptsLocal = function (id) { return pontos['LOCAL:' + id] || 0; };
  const ptsAtr = function (id) { return pontos['ATRACAO:' + id] || 0; };
  const localPorId = {};
  locais.forEach(function (l) { localPorId[l.id] = l; });

  const usados = {};
  dias.forEach(function (d) {
    if (d.travado === 'SIM' && d.local_id) usados[d.local_id] = (usados[d.local_id] || 0) + 1;
  });

  const saida = [];
  dias.forEach(function (d) {
    let localId = d.local_id;
    const travado = d.travado === 'SIM' && localId;

    if (!travado && TIPOS_AUTOMATICOS.indexOf(d.tipo) >= 0) {
      const regioes = REGIOES_POR_CIDADE[d.cidade] || [d.cidade];
      const candidatos = locais.filter(function (l) {
        return l.dia_inteiro === 'SIM' && regioes.indexOf(l.regiao) >= 0 && ptsLocal(l.id) >= 0;
      }).sort(function (a, b) {
        const ua = usados[a.id] || 0, ub = usados[b.id] || 0;
        if (ua !== ub) return ua - ub;
        const pa = ptsLocal(a.id), pb = ptsLocal(b.id);
        if (pa !== pb) return pb - pa;
        return Number(a.ordem || 99) - Number(b.ordem || 99);
      });
      localId = candidatos.length ? candidatos[0].id : '';
      if (localId) usados[localId] = (usados[localId] || 0) + 1;
    }

    const local = localPorId[localId];
    saida.push([d.data, '0', d.cidade, d.tipo, localId || '', '', '', local ? local.nome : 'Dia livre', local ? String(ptsLocal(localId)) : '', d.observacao || '', carimbo]);
    if (!local) return;

    const lista = atracoes.filter(function (a) { return a.local_id === localId; });
    const alguemVotou = lista.some(function (a) { return temVoto['ATRACAO:' + a.id]; });
    let escolhidas = lista.filter(function (a) {
      const p = ptsAtr(a.id);
      return alguemVotou ? p > 0 : true;
    });

    const areas = [];
    lista.slice().sort(function (a, b) { return Number(a.ordem_area) - Number(b.ordem_area); }).forEach(function (a) {
      if (areas.indexOf(a.area) < 0) areas.push(a.area);
    });
    let inicio = 0;
    if (alguemVotou && escolhidas.length) {
      const top = escolhidas.slice().sort(function (a, b) { return ptsAtr(b.id) - ptsAtr(a.id); })[0];
      inicio = Math.max(0, areas.indexOf(top.area));
    }
    const ordemAreas = areas.slice(inicio).concat(areas.slice(0, inicio));

    escolhidas = escolhidas.sort(function (a, b) {
      const ia = ordemAreas.indexOf(a.area), ib = ordemAreas.indexOf(b.area);
      if (ia !== ib) return ia - ib;
      const pa = ptsAtr(a.id), pb = ptsAtr(b.id);
      if (pa !== pb) return pb - pa;
      return a.nome < b.nome ? -1 : 1;
    });

    escolhidas.forEach(function (a, idx) {
      const min = Number(a.altura_min_cm || 0);
      let alerta = '';
      if (min > 0) {
        const baixos = pessoas.filter(function (p) { return p.altura_cm !== '' && Number(p.altura_cm) < min; }).map(function (p) { return p.nome; });
        if (baixos.length) alerta = 'Altura mínima ' + min + ' cm — não podem ir: ' + baixos.join(', ');
      }
      saida.push([d.data, String(idx + 1), d.cidade, d.tipo, localId, a.area, a.id, a.nome, String(ptsAtr(a.id)), alerta, carimbo]);
    });
  });
  return saida;
}

/* =========================================================
 *  HELPERS DE PLANILHA
 * ========================================================= */

function ss_() {
  let ativa = null;
  try { ativa = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ativa = null; }
  if (ativa && ativa.getSheetByName(ABA.CONFIG)) return ativa;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) {
    try {
      const salva = SpreadsheetApp.openById(id);
      if (salva.getSheetByName(ABA.CONFIG) || !ativa) return salva;
    } catch (e) { /* ID antigo ou sem acesso: segue com a planilha ativa */ }
  }
  if (ativa) return ativa;
  throw new Error('Este script não está vinculado a uma planilha. Crie o projeto pela planilha (Extensões → Apps Script).');
}

function aba_(nome) {
  const planilha = ss_();
  const sh = planilha.getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não existe na planilha "' + planilha.getName() + '". Rode setup() no editor e depois crie uma nova versão da implantação.');
  return sh;
}

/** Diagnóstico: rode no editor e veja o resultado em "Registro de execução". */
function diagnostico() {
  let ativa = null;
  try { ativa = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ativa = null; }
  Logger.log('Planilha ativa: ' + (ativa ? ativa.getName() + ' (' + ativa.getId() + ')' : 'nenhuma'));
  Logger.log('SS_ID salvo: ' + PropertiesService.getScriptProperties().getProperty('SS_ID'));
  const usada = ss_();
  Logger.log('Planilha usada pelo app: ' + usada.getName() + ' (' + usada.getId() + ')');
  Logger.log('Abas: ' + usada.getSheets().map(function (s) { return s.getName(); }).join(', '));
}

function ler_(nome) {
  const sh = aba_(nome);
  const cab = CABECALHOS[nome];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const valores = sh.getRange(2, 1, n, cab.length).getDisplayValues();
  return valores
    .filter(function (r) { return r.join('') !== ''; })
    .map(function (r) {
      const o = {};
      cab.forEach(function (c, i) { o[c] = normalizar_(c, r[i]); });
      return o;
    });
}

function normalizar_(coluna, v) {
  v = String(v == null ? '' : v).trim();
  if (coluna === 'data') {
    const m = v.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return m[3] + '-' + m[2] + '-' + m[1];
  }
  return v;
}

function escreverLinhas_(nome, linhas) {
  if (!linhas.length) return;
  const sh = aba_(nome);
  const ini = sh.getLastRow() + 1;
  sh.getRange(ini, 1, linhas.length, linhas[0].length).setNumberFormat('@').setValues(linhas.map(function (r) {
    return r.map(function (c) { return c == null ? '' : String(c); });
  }));
}

function contarLinhas_(nome) {
  return Math.max(0, aba_(nome).getLastRow() - 1);
}

function getConfig_(chave) {
  const c = ler_(ABA.CONFIG).find(function (x) { return x.chave === chave; });
  return c ? c.valor : '';
}

function validarPessoa_(id, pin) {
  const pessoas = ler_(ABA.PESSOAS);
  return pessoas.findIndex(function (p) { return p.id === id && p.pin !== '' && p.pin === String(pin); });
}

function agora_() {
  return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss');
}

/* =========================================================
 *  CATÁLOGO REAL (situação para fevereiro/2027)
 *  - DinoLand U.S.A. (Animal Kingdom) fechou em fev/2026 → fora da lista
 *  - Rivers of America / Tom Sawyer Island (Magic Kingdom) fecharam em 2025 → fora
 *  - Rock 'n' Roller Coaster agora "Starring The Muppets" (2026)
 *  - Epic Universe (Universal) incluído com alturas oficiais
 *  Alturas em cm conforme regras oficiais; sempre confira no app do parque antes de ir.
 * ========================================================= */

function seedDias_() {
  return [
    ['2027-02-08', 'Tampa', 'CHEGADA', 'TPA', 'SIM', 'Chegada em Tampa: mercado, passeio leve e depois hospedagem.'],
    ['2027-02-09', 'Orlando', 'PARQUE', '', 'NAO', 'Início dos parques.'],
    ['2027-02-10', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-11', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-12', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-13', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-14', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-15', 'Miami', 'MIAMI', 'MIA', 'SIM', 'Bate-volta: cerca de 3h30 de carro em cada trecho saindo de Orlando.'],
    ['2027-02-16', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-17', 'Orlando', 'PARQUE', '', 'NAO', ''],
    ['2027-02-18', 'Tampa', 'RETORNO', '', 'NAO', 'Volta para Tampa (~1h30), último passeio, devolução do carro. Voo na madrugada de 19/02.']
  ];
}

function seedLocais_() {
  // id, nome, grupo, regiao, tipo, dia_inteiro, ordem, descricao, site, foto_busca, foto_url, foto_credito
  return [
    ['MK', 'Magic Kingdom', 'Disney', 'Orlando', 'PARQUE', 'SIM', '1', 'O parque do Castelo da Cinderela: clássicos da Disney, TRON, Space Mountain e fogos à noite.', 'https://disneyworld.disney.go.com/destinations/magic-kingdom/', 'Cinderella Castle Magic Kingdom', '', ''],
    ['EPCOT', 'EPCOT', 'Disney', 'Orlando', 'PARQUE', 'SIM', '2', 'Futuro e países do mundo: Guardiões da Galáxia, Test Track, Frozen, Ratatouille e o World Showcase. Em jan–fev acontece o Festival of the Arts.', 'https://disneyworld.disney.go.com/destinations/epcot/', 'Spaceship Earth Epcot', '', ''],
    ['HS', "Disney's Hollywood Studios", 'Disney', 'Orlando', 'PARQUE', 'SIM', '3', 'Star Wars Galaxy\'s Edge, Toy Story Land, Torre do Terror e a montanha-russa dos Muppets. Fantasmic! à noite.', 'https://disneyworld.disney.go.com/destinations/hollywood-studios/', 'Tower of Terror Hollywood Studios', '', ''],
    ['AK', "Disney's Animal Kingdom", 'Disney', 'Orlando', 'PARQUE', 'SIM', '4', 'Pandora (Avatar), safári africano e Expedition Everest. Abre cedo e é mais calmo à tarde.', 'https://disneyworld.disney.go.com/destinations/animal-kingdom/', 'Tree of Life Animal Kingdom', '', ''],
    ['EPIC', 'Universal Epic Universe', 'Universal', 'Orlando', 'PARQUE', 'SIM', '5', 'Parque mais novo de Orlando (2025): Super Nintendo World, Como Treinar seu Dragão, Ministério da Magia e Monstros Clássicos.', 'https://www.universalorlando.com/web/en/us/theme-parks/epic-universe', 'Epic Universe Universal Orlando', '', ''],
    ['IOA', 'Universal Islands of Adventure', 'Universal', 'Orlando', 'PARQUE', 'SIM', '6', 'Hogsmeade (Harry Potter), VelociCoaster, Hulk, Homem-Aranha e Jurassic Park.', 'https://www.universalorlando.com/web/en/us/theme-parks/islands-of-adventure', 'Hogsmeade Islands of Adventure', '', ''],
    ['USF', 'Universal Studios Florida', 'Universal', 'Orlando', 'PARQUE', 'SIM', '7', 'Beco Diagonal (Gringotts), Velozes e Furiosos, Minions, Simpsons e Múmia. O Hogwarts Express liga ao Islands (exige ingresso park-to-park).', 'https://www.universalorlando.com/web/en/us/theme-parks/universal-studios-florida', 'Universal Studios Florida', '', ''],
    ['KSC', 'Kennedy Space Center (NASA)', 'NASA', 'Space Coast', 'PASSEIO', 'SIM', '8', 'Centro espacial da NASA (~1h de Orlando): ônibus Atlantis, foguete Saturn V, simulador de lançamento. Confira se há lançamento de foguete na data.', 'https://www.kennedyspacecenter.com/', 'Kennedy Space Center Visitor Complex', '', ''],
    ['SWO', 'SeaWorld Orlando', 'SeaWorld', 'Orlando', 'PARQUE', 'SIM', '9', 'Montanhas-russas fortes (Mako, Kraken, Manta, Pipeline), pinguins e show das orcas.', 'https://seaworld.com/orlando/', 'SeaWorld Orlando Mako', '', ''],
    ['BGT', 'Busch Gardens Tampa Bay', 'Busch Gardens', 'Tampa', 'PARQUE', 'SIM', '10', 'Parque de Tampa com animais africanos e montanhas-russas recordistas (Iron Gwazi, SheiKra, Montu). Ótimo para o dia de volta a Tampa.', 'https://buschgardens.com/tampa/', 'Busch Gardens Tampa Bay', '', ''],
    ['LEGO', 'LEGOLAND Florida', 'LEGOLAND', 'Orlando', 'PARQUE', 'SIM', '11', 'Em Winter Haven (~50 min de Orlando). Ideal para crianças de até ~12 anos.', 'https://www.legoland.com/florida/', 'Legoland Florida Miniland', '', ''],
    ['VB', 'Universal Volcano Bay (aquático)', 'Universal', 'Orlando', 'PARQUE_AQUATICO', 'SIM', '12', 'Parque aquático da Universal com fila virtual. Em fevereiro a água é aquecida, mas o dia pode estar fresco.', 'https://www.universalorlando.com/web/en/us/theme-parks/volcano-bay', 'Volcano Bay Krakatau', '', ''],
    ['TL', "Disney's Typhoon Lagoon (aquático)", 'Disney', 'Orlando', 'PARQUE_AQUATICO', 'SIM', '13', 'Aquático da Disney com piscina de ondas gigante. No inverno normalmente só um dos dois aquáticos Disney fica aberto — confira o calendário.', 'https://disneyworld.disney.go.com/destinations/typhoon-lagoon/', 'Typhoon Lagoon Disney', '', ''],
    ['BB', "Disney's Blizzard Beach (aquático)", 'Disney', 'Orlando', 'PARQUE_AQUATICO', 'SIM', '14', 'Aquático com tema de estação de esqui e o tobogã Summit Plummet. Costuma fechar parte do inverno — confira.', 'https://disneyworld.disney.go.com/destinations/blizzard-beach/', 'Blizzard Beach Summit Plummet', '', ''],
    ['MIA', 'Dia em Miami', 'Passeio', 'Miami', 'PASSEIO', 'SIM', '15', 'Bate-volta para Miami: praia, arte de rua, Little Havana e compras.', 'https://www.miamiandbeaches.com/', 'South Beach Miami Ocean Drive', '', ''],
    ['TPA', 'Tampa — cidade, mercado e compras', 'Passeio', 'Tampa', 'PASSEIO', 'SIM', '16', 'Mercado na chegada, Riverwalk, Ybor City, aquário e outlets.', 'https://www.visittampabay.com/', 'Tampa Riverwalk', '', ''],
    ['CLW', 'Clearwater Beach', 'Passeio', 'Tampa', 'PASSEIO', 'SIM', '17', 'Praia de areia branca a ~40 min de Tampa, eleita várias vezes uma das melhores dos EUA.', 'https://www.visitstpeteclearwater.com/', 'Clearwater Beach Pier 60', '', ''],
    ['ORLX', 'Orlando — compras e passeios curtos', 'Passeio', 'Orlando', 'COMPRAS', 'NAO', '18', 'Para noites ou meio período livre: outlets, Disney Springs, roda-gigante e Gatorland.', 'https://www.visitorlando.com/', 'Disney Springs Orlando', '', '']
  ];
}

function seedAtracoes_() {
  // [area, ordem_area, nome, categoria, altura_min_cm, descricao]
  const C = {
    MK: [
      ['Main Street, U.S.A.', 1, 'Walt Disney World Railroad', 'Passeio', 0, 'Trem a vapor que dá a volta no parque.'],
      ['Main Street, U.S.A.', 1, 'Happily Ever After (fogos no castelo)', 'Show', 0, 'Show de fogos e projeções no castelo, à noite.'],
      ['Adventureland', 2, 'Pirates of the Caribbean', 'Família', 0, 'Barco pelos piratas — clássico.'],
      ['Adventureland', 2, 'Jungle Cruise', 'Família', 0, 'Passeio de barco cheio de piadas pela "selva".'],
      ['Adventureland', 2, 'The Magic Carpets of Aladdin', 'Infantil', 0, 'Tapetes voadores giratórios.'],
      ['Adventureland', 2, "Walt Disney's Enchanted Tiki Room", 'Show', 0, 'Show de pássaros animatrônicos.'],
      ['Frontierland', 3, 'Big Thunder Mountain Railroad', 'Radical', 102, 'Montanha-russa do trem desgovernado na mina.'],
      ['Frontierland', 3, "Tiana's Bayou Adventure", 'Aquática', 102, 'Descida de tronco com a Tiana — molha!'],
      ['Frontierland', 3, 'Country Bear Musical Jamboree', 'Show', 0, 'Show musical dos ursos.'],
      ['Liberty Square', 4, 'Haunted Mansion', 'Família', 0, 'Mansão mal-assombrada divertida, sem sustos fortes.'],
      ['Liberty Square', 4, 'The Hall of Presidents', 'Show', 0, 'Show com os presidentes americanos.'],
      ['Fantasyland', 5, 'Seven Dwarfs Mine Train', 'Radical', 97, 'Montanha-russa familiar dos Sete Anões — fila longa, vá cedo.'],
      ['Fantasyland', 5, "Peter Pan's Flight", 'Família', 0, 'Voo sobre Londres e a Terra do Nunca.'],
      ['Fantasyland', 5, "it's a small world", 'Família', 0, 'Barquinho pelas crianças do mundo.'],
      ['Fantasyland', 5, 'Under the Sea – Journey of The Little Mermaid', 'Família', 0, 'Passeio pela história da Ariel.'],
      ['Fantasyland', 5, "Mickey's PhilharMagic", 'Show', 0, 'Filme 3D com músicas Disney.'],
      ['Fantasyland', 5, 'Dumbo the Flying Elephant', 'Infantil', 0, 'Elefantes voadores.'],
      ['Fantasyland', 5, 'The Barnstormer', 'Infantil', 89, 'Mini montanha-russa do Pateta.'],
      ['Fantasyland', 5, 'Mad Tea Party', 'Infantil', 0, 'Xícaras giratórias.'],
      ['Fantasyland', 5, 'The Many Adventures of Winnie the Pooh', 'Família', 0, 'Passeio com o Ursinho Pooh.'],
      ['Fantasyland', 5, 'Enchanted Tales with Belle', 'Show', 0, 'Encontro interativo com a Bela.'],
      ['Tomorrowland', 6, 'TRON Lightcycle / Run', 'Radical', 122, 'Montanha-russa de moto lançada — a mais rápida do parque.'],
      ['Tomorrowland', 6, 'Space Mountain', 'Radical', 112, 'Montanha-russa no escuro, no espaço.'],
      ['Tomorrowland', 6, "Buzz Lightyear's Space Ranger Spin", 'Família', 0, 'Jogo de tiro ao alvo com o Buzz.'],
      ['Tomorrowland', 6, 'Tomorrowland Speedway', 'Família', 82, 'Carrinhos para dirigir (para dirigir sozinho: 137 cm).'],
      ['Tomorrowland', 6, 'Astro Orbiter', 'Infantil', 0, 'Foguetes giratórios no alto.'],
      ['Tomorrowland', 6, 'Tomorrowland Transit Authority PeopleMover', 'Passeio', 0, 'Passeio tranquilo pela Tomorrowland.'],
      ['Tomorrowland', 6, 'Monsters, Inc. Laugh Floor', 'Show', 0, 'Show de comédia interativo.'],
      ['Tomorrowland', 6, "Walt Disney's Carousel of Progress", 'Show', 0, 'Clássico teatro giratório.']
    ],
    EPCOT: [
      ['World Celebration', 1, 'Spaceship Earth', 'Família', 0, 'Dentro da bola gigante: história da comunicação.'],
      ['World Celebration', 1, 'Journey into Imagination with Figment', 'Família', 0, 'Passeio divertido com o dragãozinho Figment.'],
      ['World Discovery', 2, 'Guardians of the Galaxy: Cosmic Rewind', 'Radical', 107, 'Montanha-russa giratória com músicas dos anos 80 — reserva por fila virtual.'],
      ['World Discovery', 2, 'Test Track', 'Radical', 102, 'Projete seu carro e teste a até ~100 km/h.'],
      ['World Discovery', 2, 'Mission: SPACE', 'Simulador', 102, 'Simulador de missão a Marte (versão laranja mais intensa).'],
      ['World Nature', 3, "Soarin' Around the World", 'Simulador', 102, 'Voo de asa-delta sobre paisagens do mundo.'],
      ['World Nature', 3, 'Living with the Land', 'Passeio', 0, 'Barco pelas estufas.'],
      ['World Nature', 3, 'The Seas with Nemo & Friends', 'Família', 0, 'Passeio do Nemo e aquário enorme.'],
      ['World Nature', 3, 'Turtle Talk with Crush', 'Show', 0, 'Conversa ao vivo com a tartaruga Crush.'],
      ['World Nature', 3, 'Journey of Water, Inspired by Moana', 'Passeio', 0, 'Trilha interativa com água.'],
      ['World Showcase', 4, 'Frozen Ever After', 'Família', 0, 'Barco por Arendelle (Noruega).'],
      ['World Showcase', 4, "Remy's Ratatouille Adventure", 'Família', 0, 'Encolha ao tamanho do Remy (França).'],
      ['World Showcase', 4, 'Gran Fiesta Tour Starring The Three Caballeros', 'Família', 0, 'Barco pelo pavilhão do México.'],
      ['World Showcase', 4, 'Reflections of China', 'Show', 0, 'Filme 360° sobre a China.'],
      ['World Showcase', 4, 'Canada Far and Wide', 'Show', 0, 'Filme 360° sobre o Canadá.'],
      ['World Showcase', 4, 'Show noturno de fogos no lago', 'Show', 0, 'Show de fogos sobre a lagoa do World Showcase.']
    ],
    HS: [
      ['Hollywood Boulevard', 1, "Mickey & Minnie's Runaway Railway", 'Família', 0, 'Entre no desenho do Mickey.'],
      ['Echo Lake', 2, 'Star Tours – The Adventures Continue', 'Simulador', 102, 'Simulador de voo Star Wars.'],
      ['Echo Lake', 2, 'Indiana Jones Epic Stunt Spectacular!', 'Show', 0, 'Show de dublês.'],
      ['Echo Lake', 2, 'For the First Time in Forever: A Frozen Sing-Along', 'Show', 0, 'Cante junto com Frozen.'],
      ["Star Wars: Galaxy's Edge", 3, 'Star Wars: Rise of the Resistance', 'Família', 102, 'A atração mais completa da Disney — vá cedo ou use Lightning Lane.'],
      ["Star Wars: Galaxy's Edge", 3, 'Millennium Falcon: Smugglers Run', 'Simulador', 97, 'Pilote a Millennium Falcon em equipe.'],
      ['Toy Story Land', 4, 'Slinky Dog Dash', 'Radical', 97, 'Montanha-russa familiar do Slinky.'],
      ['Toy Story Land', 4, 'Toy Story Mania!', 'Família', 0, 'Jogo 3D de tiro ao alvo.'],
      ['Toy Story Land', 4, 'Alien Swirling Saucers', 'Infantil', 81, 'Discos voadores dos aliens.'],
      ['Animation Courtyard', 5, 'The Little Mermaid – A Musical Adventure', 'Show', 0, 'Musical da Pequena Sereia.'],
      ['Sunset Boulevard', 6, 'The Twilight Zone Tower of Terror', 'Radical', 102, 'Elevador em queda livre.'],
      ['Sunset Boulevard', 6, "Rock 'n' Roller Coaster Starring The Muppets", 'Radical', 122, 'Montanha-russa lançada no escuro, agora com os Muppets.'],
      ['Sunset Boulevard', 6, 'Beauty and the Beast – Live on Stage', 'Show', 0, 'Musical da Bela e a Fera.'],
      ['Sunset Boulevard', 6, 'Fantasmic!', 'Show', 0, 'Show noturno com água, fogo e projeções.']
    ],
    AK: [
      ['Discovery Island', 1, 'Zootopia: Better Zoogether!', 'Show', 0, 'Show 4D de Zootopia dentro da Árvore da Vida.'],
      ['Discovery Island', 1, 'Discovery Island Trails', 'Passeio', 0, 'Trilhas ao redor da Árvore da Vida.'],
      ['Pandora – The World of Avatar', 2, 'Avatar Flight of Passage', 'Simulador', 112, 'Voo em um banshee — uma das melhores da Disney.'],
      ['Pandora – The World of Avatar', 2, "Na'vi River Journey", 'Família', 0, 'Barco tranquilo pela floresta bioluminescente.'],
      ['Africa', 3, 'Kilimanjaro Safaris', 'Passeio', 0, 'Safári de caminhão com animais reais — melhor de manhã.'],
      ['Africa', 3, 'Gorilla Falls Exploration Trail', 'Passeio', 0, 'Trilha com gorilas e hipopótamos.'],
      ['Africa', 3, 'Festival of the Lion King', 'Show', 0, 'Musical do Rei Leão.'],
      ['Asia', 4, 'Expedition Everest – Legend of the Forbidden Mountain', 'Radical', 112, 'Montanha-russa com o Yeti, inclusive de ré.'],
      ['Asia', 4, 'Kali River Rapids', 'Aquática', 97, 'Bote nas corredeiras — molha muito.'],
      ['Asia', 4, 'Maharajah Jungle Trek', 'Passeio', 0, 'Trilha com tigres.'],
      ['Asia', 4, 'Feathered Friends in Flight!', 'Show', 0, 'Show de aves.']
    ],
    EPIC: [
      ['Celestial Park', 1, 'Stardust Racers', 'Radical', 122, 'Duas montanhas-russas competindo lado a lado.'],
      ['Celestial Park', 1, 'Constellation Carousel', 'Infantil', 0, 'Carrossel giratório das constelações.'],
      ['Celestial Park', 1, 'Astronomica', 'Aquática', 0, 'Área de jogos de água.'],
      ['Super Nintendo World', 2, "Mario Kart: Bowser's Challenge", 'Família', 102, 'Corrida de Mario Kart com realidade aumentada.'],
      ['Super Nintendo World', 2, 'Mine-Cart Madness', 'Radical', 102, 'Montanha-russa do Donkey Kong.'],
      ['Super Nintendo World', 2, "Yoshi's Adventure", 'Infantil', 87, 'Passeio tranquilo com o Yoshi.'],
      ['How to Train Your Dragon – Isle of Berk', 3, "Hiccup's Wing Gliders", 'Radical', 102, 'Montanha-russa lançada familiar.'],
      ['How to Train Your Dragon – Isle of Berk', 3, "Dragon Racer's Rally", 'Radical', 122, 'Giro de "asa" controlado por você.'],
      ['How to Train Your Dragon – Isle of Berk', 3, 'Fyre Drill', 'Aquática', 0, 'Barcos com jatos d\'água.'],
      ['How to Train Your Dragon – Isle of Berk', 3, 'The Untrainable Dragon', 'Show', 0, 'Show ao vivo com dragão gigante.'],
      ['How to Train Your Dragon – Isle of Berk', 3, 'Viking Training Camp', 'Infantil', 0, 'Playground viking.'],
      ['The Wizarding World of Harry Potter – Ministry of Magic', 4, 'Harry Potter and the Battle at the Ministry', 'Família', 102, 'Atração principal do Ministério da Magia.'],
      ['The Wizarding World of Harry Potter – Ministry of Magic', 4, 'Le Cirque Arcanus', 'Show', 0, 'Show com criaturas mágicas.'],
      ['Dark Universe', 5, 'Monsters Unchained: The Frankenstein Experiment', 'Radical', 122, 'Atração escura e intensa com os monstros clássicos.'],
      ['Dark Universe', 5, 'Curse of the Werewolf', 'Radical', 102, 'Montanha-russa giratória do lobisomem.']
    ],
    IOA: [
      ['Marvel Super Hero Island', 1, 'The Incredible Hulk Coaster', 'Radical', 137, 'Montanha-russa lançada com loopings.'],
      ['Marvel Super Hero Island', 1, 'The Amazing Adventures of Spider-Man', 'Simulador', 102, 'Clássico 3D do Homem-Aranha.'],
      ['Marvel Super Hero Island', 1, "Doctor Doom's Fearfall", 'Radical', 132, 'Torre de queda/lançamento.'],
      ['Marvel Super Hero Island', 1, 'Storm Force Accelatron', 'Infantil', 0, 'Xícaras giratórias da Tempestade.'],
      ['Toon Lagoon', 2, "Dudley Do-Right's Ripsaw Falls", 'Aquática', 112, 'Descida de tronco — molha muito.'],
      ['Toon Lagoon', 2, "Popeye & Bluto's Bilge-Rat Barges", 'Aquática', 107, 'Bote — sai encharcado.'],
      ['Skull Island', 3, 'Skull Island: Reign of Kong', 'Família', 91, 'Expedição de caminhão com o King Kong.'],
      ['Jurassic Park', 4, 'Jurassic World VelociCoaster', 'Radical', 130, 'Uma das melhores montanhas-russas do mundo.'],
      ['Jurassic Park', 4, 'Jurassic Park River Adventure', 'Aquática', 107, 'Barco com o T-Rex e queda grande.'],
      ['Jurassic Park', 4, 'Pteranodon Flyers', 'Infantil', 91, 'Só para crianças de 91 a 142 cm (adulto só acompanhando).'],
      ['Jurassic Park', 4, 'Camp Jurassic', 'Infantil', 0, 'Playground temático.'],
      ['The Wizarding World of Harry Potter – Hogsmeade', 5, "Hagrid's Magical Creatures Motorbike Adventure", 'Radical', 122, 'Montanha-russa de moto do Hagrid — muito popular.'],
      ['The Wizarding World of Harry Potter – Hogsmeade', 5, 'Harry Potter and the Forbidden Journey', 'Radical', 122, 'Dentro do castelo de Hogwarts.'],
      ['The Wizarding World of Harry Potter – Hogsmeade', 5, 'Flight of the Hippogriff', 'Família', 91, 'Montanha-russa familiar.'],
      ['The Wizarding World of Harry Potter – Hogsmeade', 5, 'Hogwarts Express – Hogsmeade Station', 'Passeio', 0, 'Trem até o Universal Studios (exige park-to-park).'],
      ['Seuss Landing', 6, 'The Cat in the Hat', 'Infantil', 91, 'Passeio pelo livro do Gato de Chapéu.'],
      ['Seuss Landing', 6, 'One Fish, Two Fish, Red Fish, Blue Fish', 'Infantil', 0, 'Peixinhos giratórios com água.'],
      ['Seuss Landing', 6, 'The High in the Sky Seuss Trolley Train Ride!', 'Infantil', 86, 'Trenzinho elevado.'],
      ['Seuss Landing', 6, 'Caro-Seuss-el', 'Infantil', 0, 'Carrossel do Dr. Seuss.']
    ],
    USF: [
      ['Production Central', 1, 'Hollywood Rip Ride Rockit', 'Radical', 130, 'Montanha-russa em que você escolhe a música.'],
      ['Production Central', 1, 'Transformers: The Ride-3D', 'Simulador', 102, 'Batalha 3D dos Transformers.'],
      ['Production Central', 1, 'Despicable Me Minion Mayhem', 'Simulador', 102, 'Simulador dos Minions.'],
      ['Production Central', 1, "Illumination's Villain-Con Minion Blast", 'Família', 0, 'Jogo interativo de tiro dos Minions.'],
      ['New York', 2, 'Revenge of the Mummy', 'Radical', 122, 'Montanha-russa no escuro com a Múmia.'],
      ['New York', 2, 'Race Through New York Starring Jimmy Fallon', 'Simulador', 102, 'Corrida 3D por Nova York.'],
      ['San Francisco', 3, 'Fast & Furious – Supercharged', 'Simulador', 102, 'Simulador de Velozes e Furiosos.'],
      ['The Wizarding World of Harry Potter – Diagon Alley', 4, 'Harry Potter and the Escape from Gringotts', 'Radical', 107, 'Montanha-russa dentro do banco Gringotts.'],
      ['The Wizarding World of Harry Potter – Diagon Alley', 4, "Hogwarts Express – King's Cross Station", 'Passeio', 0, 'Trem até o Islands (exige park-to-park).'],
      ['World Expo', 5, 'Men in Black: Alien Attack', 'Família', 107, 'Tiro ao alvo nos alienígenas, com placar.'],
      ['Springfield', 6, 'The Simpsons Ride', 'Simulador', 102, 'Simulador com a família Simpson.'],
      ["Springfield", 6, "Kang & Kodos' Twirl 'n' Hurl", 'Infantil', 0, 'Discos giratórios.'],
      ['DreamWorks Land', 7, 'Trolls Trollercoaster', 'Infantil', 91, 'Mini montanha-russa dos Trolls.'],
      ['DreamWorks Land', 7, 'E.T. Adventure', 'Família', 86, 'Bicicleta voadora com o E.T.'],
      ['Hollywood', 8, 'The Bourne Stuntacular', 'Show', 0, 'Show de ação com dublês.']
    ],
    KSC: [
      ['Entrada e Heroes & Legends', 1, 'Rocket Garden', 'Passeio', 0, 'Jardim com foguetes reais.'],
      ['Entrada e Heroes & Legends', 1, 'Heroes & Legends (U.S. Astronaut Hall of Fame)', 'Passeio', 0, 'Exposição dos astronautas americanos.'],
      ['Space Shuttle Atlantis', 2, 'Space Shuttle Atlantis', 'Passeio', 0, 'O ônibus espacial verdadeiro em exposição.'],
      ['Space Shuttle Atlantis', 2, 'Shuttle Launch Experience', 'Simulador', 112, 'Simulador de lançamento do ônibus espacial.'],
      ['Gateway: The Deep Space Launch Complex', 3, 'Spaceport KSC (simuladores de viagem)', 'Simulador', 0, 'Simuladores de viagens a Marte e outros destinos.'],
      ['Gateway: The Deep Space Launch Complex', 3, 'Planet Play', 'Infantil', 0, 'Área de brincar com tema espacial.'],
      ['Tour de ônibus', 4, 'KSC Bus Tour + Apollo/Saturn V Center', 'Passeio', 0, 'Ônibus pelas áreas de lançamento até o foguete Saturn V.'],
      ['Extras', 5, 'Assistir a um lançamento de foguete', 'Show', 0, 'Depende do calendário de lançamentos — conferir perto da data.'],
      ['Extras', 5, 'Chat with an Astronaut', 'Show', 0, 'Conversa com um astronauta (pago à parte).']
    ],
    SWO: [
      ['Sea of Legends', 1, 'Mako', 'Radical', 137, 'Hipercoaster mais alta de Orlando.'],
      ['Sea of Legends', 1, 'Kraken', 'Radical', 137, 'Montanha-russa sem chão.'],
      ['Sea of Legends', 1, 'Infinity Falls', 'Aquática', 107, 'Bote com queda recorde.'],
      ['Sea of Mystery', 2, 'Manta', 'Radical', 137, 'Montanha-russa "voando" como uma arraia.'],
      ['Sea of Delight', 3, 'Pipeline: The Surf Coaster', 'Radical', 137, 'Montanha-russa em pé com tema de surfe.'],
      ['Sea of Delight', 3, 'Journey to Atlantis', 'Aquática', 107, 'Barco com queda — molha.'],
      ['Sea of Ice', 4, 'Ice Breaker', 'Radical', 122, 'Montanha-russa lançada para frente e para trás.'],
      ['Sea of Ice', 4, 'Penguin Trek', 'Família', 99, 'Montanha-russa familiar que termina no pinguinário.'],
      ['Sea of Ice', 4, 'Antarctica: Empire of the Penguin', 'Família', 0, 'Passeio até a colônia de pinguins.'],
      ['Sesame Street Land', 5, 'Sesame Street Land (brinquedos infantis)', 'Infantil', 0, 'Área da Vila Sésamo.'],
      ['Shows', 6, 'Orca Encounter', 'Show', 0, 'Apresentação das orcas.'],
      ['Shows', 6, 'Dolphin Adventures', 'Show', 0, 'Apresentação dos golfinhos.']
    ],
    BGT: [
      ['Pantopia', 1, 'Iron Gwazi', 'Radical', 122, 'Montanha-russa híbrida — uma das mais bem avaliadas dos EUA.'],
      ['Pantopia', 1, 'Tigris', 'Radical', 137, 'Montanha-russa lançada.'],
      ['Pantopia', 1, "Falcon's Fury", 'Radical', 137, 'Torre de queda livre de cara para o chão.'],
      ['Cheetah Hunt', 2, 'Cheetah Hunt', 'Radical', 122, 'Montanha-russa lançada passando pela savana.'],
      ['Egito', 3, 'Montu', 'Radical', 137, 'Montanha-russa invertida clássica.'],
      ['Egito', 3, "Cobra's Curse", 'Família', 107, 'Montanha-russa giratória familiar.'],
      ['Congo', 4, 'Kumba', 'Radical', 137, 'Montanha-russa com loopings.'],
      ['Congo', 4, 'Congo River Rapids', 'Aquática', 107, 'Bote nas corredeiras.'],
      ['Stanleyville', 5, 'SheiKra', 'Radical', 137, 'Queda vertical de 90°.'],
      ['Stanleyville', 5, 'Phoenix Rising', 'Família', 97, 'Montanha-russa suspensa familiar.'],
      ['Planícies do Serengeti', 6, 'Serengeti Express (trem)', 'Passeio', 0, 'Trem que passa pelos animais da savana.']
    ],
    LEGO: [
      ['Miniland USA', 1, 'Miniland USA', 'Passeio', 0, 'Cidades americanas montadas em LEGO.'],
      ['LEGO Movie World', 2, 'The LEGO Movie Masters of Flight', 'Simulador', 0, 'Simulador de voo do filme LEGO.'],
      ['Kingdoms', 3, 'The Dragon', 'Família', 0, 'Montanha-russa do castelo (conferir altura no site).'],
      ['LEGO Technic', 4, 'Project X', 'Radical', 0, 'Montanha-russa mais radical do parque (conferir altura no site).'],
      ['Imagination Zone', 5, 'Ninjago The Ride', 'Família', 0, 'Jogo interativo 4D.']
    ],
    VB: [
      ['Krakatau', 1, 'Krakatau Aqua Coaster', 'Aquática', 107, 'Montanha-russa aquática dentro do vulcão.'],
      ['Krakatau', 1, "Ko'okiri Body Plunge", 'Aquática', 122, 'Tobogã com queda quase vertical.'],
      ['Wave Village', 2, 'Waturi Beach', 'Aquática', 0, 'Piscina de ondas.'],
      ['River Village', 3, 'TeAwa The Fearless River', 'Aquática', 0, 'Rio com correnteza forte.'],
      ['Rainforest Village', 4, 'Honu ika Moana', 'Aquática', 107, 'Tobogã em bote para grupo.']
    ],
    TL: [
      ['Surf Pool', 1, 'Typhoon Lagoon Surf Pool', 'Aquática', 0, 'Piscina de ondas gigantes.'],
      ['Mount Mayday', 2, "Crush 'n' Gusher", 'Aquática', 122, 'Montanha-russa aquática.'],
      ['Mount Mayday', 2, 'Humunga Kowabunga', 'Aquática', 122, 'Tobogã de queda rápida.'],
      ['Mount Mayday', 2, 'Miss Adventure Falls', 'Aquática', 0, 'Boia em família.'],
      ['Castaway Creek', 3, 'Castaway Creek', 'Aquática', 0, 'Rio lento de boia.']
    ],
    BB: [
      ['Mount Gushmore', 1, 'Summit Plummet', 'Aquática', 122, 'Um dos tobogãs mais altos e rápidos do mundo.'],
      ['Mount Gushmore', 1, 'Downhill Double Dipper', 'Aquática', 122, 'Corrida de tobogã lado a lado.'],
      ['Mount Gushmore', 1, 'Teamboat Springs', 'Aquática', 0, 'Bote em família.'],
      ['Cross Country Creek', 2, 'Cross Country Creek', 'Aquática', 0, 'Rio lento de boia.']
    ],
    MIA: [
      ['No caminho (Everglades)', 1, 'Passeio de airboat nos Everglades', 'Passeio', 0, 'Barco-ventilador com jacarés, no caminho para Miami.'],
      ['Miami Beach', 2, 'South Beach e Ocean Drive', 'Passeio', 0, 'Praia famosa e calçadão.'],
      ['Miami Beach', 2, 'Art Deco Historic District', 'Passeio', 0, 'Prédios coloridos art déco.'],
      ['Wynwood / Little Havana', 3, 'Wynwood Walls', 'Passeio', 0, 'Museu de grafite a céu aberto.'],
      ['Wynwood / Little Havana', 3, 'Little Havana – Calle Ocho', 'Passeio', 0, 'Bairro cubano, café e cultura.'],
      ['Downtown', 4, 'Bayside Marketplace + passeio de barco', 'Passeio', 0, 'Shopping na baía e barco pelas casas das celebridades.'],
      ['Compras', 5, 'Dolphin Mall (outlet)', 'Compras', 0, 'Grande outlet de Miami.'],
      ['Compras', 5, 'Aventura Mall', 'Compras', 0, 'Shopping grande e completo.']
    ],
    TPA: [
      ['Chegada', 1, 'Mercado (Walmart Supercenter / Publix / Target)', 'Compras', 0, 'Compras de mercado e itens básicos para a viagem.'],
      ['Centro', 2, 'Tampa Riverwalk', 'Passeio', 0, 'Calçadão às margens do rio.'],
      ['Centro', 2, 'Armature Works', 'Passeio', 0, 'Mercado gastronômico à beira-rio.'],
      ['Centro', 2, 'The Florida Aquarium', 'Família', 0, 'Aquário grande no centro de Tampa.'],
      ['Ybor City', 3, 'Ybor City', 'Passeio', 0, 'Bairro histórico cubano.'],
      ['Compras', 4, 'International Plaza', 'Compras', 0, 'Shopping principal de Tampa.'],
      ['Compras', 4, 'Tampa Premium Outlets', 'Compras', 0, 'Outlet em Lutz.']
    ],
    CLW: [
      ['Praia', 1, 'Clearwater Beach', 'Passeio', 0, 'Areia branca e mar calmo.'],
      ['Praia', 1, 'Pier 60 (pôr do sol)', 'Passeio', 0, 'Pier com artistas de rua ao entardecer.'],
      ['Família', 2, 'Clearwater Marine Aquarium', 'Família', 0, 'Casa da golfinho Winter (filme "Winter, o Golfinho").']
    ],
    ORLX: [
      ['Disney', 1, 'Disney Springs', 'Compras', 0, 'Lojas e restaurantes da Disney, entrada gratuita.'],
      ['International Drive', 2, 'ICON Park – The Wheel', 'Passeio', 0, 'Roda-gigante com vista de Orlando.'],
      ['International Drive', 2, 'Orlando International Premium Outlets', 'Compras', 0, 'Outlet na International Drive.'],
      ['Lake Buena Vista', 3, 'Orlando Vineland Premium Outlets', 'Compras', 0, 'Outlet perto da Disney.'],
      ['Kissimmee', 4, 'Old Town Kissimmee', 'Passeio', 0, 'Rua de diversões e carros antigos.'],
      ['Kissimmee', 4, 'Gatorland', 'Família', 0, 'Parque dos jacarés (meio período).']
    ]
  };

  const nomeCurto = { MK: 'Magic Kingdom', EPCOT: 'Epcot', HS: 'Hollywood Studios', AK: 'Animal Kingdom', EPIC: 'Epic Universe', IOA: 'Islands of Adventure', USF: 'Universal Studios Florida', KSC: 'Kennedy Space Center', SWO: 'SeaWorld Orlando', BGT: 'Busch Gardens Tampa', LEGO: 'Legoland Florida', VB: 'Volcano Bay', TL: 'Typhoon Lagoon', BB: 'Blizzard Beach', MIA: 'Miami', TPA: 'Tampa', CLW: 'Clearwater', ORLX: 'Orlando' };

  const linhas = [];
  Object.keys(C).forEach(function (localId) {
    C[localId].forEach(function (a, i) {
      const id = localId + '-' + ('0' + (i + 1)).slice(-2);
      const busca = a[2].replace(/\(.*?\)/g, '').replace(/[–—]/g, ' ').trim() + ' ' + nomeCurto[localId];
      linhas.push([id, localId, a[0], String(a[1]), a[2], a[3], a[4] ? String(a[4]) : '', a[5], busca, '', '', 'SIM']);
    });
  });
  return linhas;
}
