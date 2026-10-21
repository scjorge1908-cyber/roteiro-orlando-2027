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
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.motor = motorJs_();
  return tpl.evaluate()
    .setTitle('Roteiro Orlando 2027')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Motor do roteiro enviado para a tela: o mesmo código do servidor monta o roteiro ao vivo no celular. */
function motorJs_() {
  return 'var REGIOES_POR_CIDADE = ' + JSON.stringify(REGIOES_POR_CIDADE) + ';\n' +
    'var TIPOS_AUTOMATICOS = ' + JSON.stringify(TIPOS_AUTOMATICOS) + ';\n' +
    montarRoteiro_.toString() + '\n';
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

/** Preenche foto_url/foto_credito vazios com as fotos fixas do catálogo (FOTOS_CATALOGO, no fim do arquivo). */
function garantirFotos_() {
  [ABA.LOCAIS, ABA.ATRACOES].forEach(function (nomeAba) {
    const sh = aba_(nomeAba);
    const cab = CABECALHOS[nomeAba];
    const n = sh.getLastRow() - 1;
    if (n < 1) return;
    const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
    const col = cab.indexOf('foto_url') + 1;
    const atuais = sh.getRange(2, col, n, 2).getDisplayValues();
    let mudou = false;
    const novos = atuais.map(function (r, i) {
      const f = FOTOS_CATALOGO[ids[i][0]];
      if (r[0] === '' && f) { mudou = true; return [f[0], f[1]]; }
      return r;
    });
    if (mudou) sh.getRange(2, col, n, 2).setNumberFormat('@').setValues(novos);
  });
}

function api_dados() {
  try {
    garantirSetup_();
    garantirFotos_();
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

/**
 * Busca fotos no Wikimedia Commons pelo servidor (o navegador do Apps Script bloqueia a busca direta).
 * itens: [{tipo:'LOCAL'|'ATRACAO', id}] — no máximo 12 por chamada.
 * Retorna { ok, fotos: { 'TIPO:ID': {url, cred} }, sem: ['TIPO:ID', ...] }
 */
function api_buscarFotos(itens) {
  try {
    itens = (itens || []).slice(0, 12);
    const tabelas = { LOCAL: ler_(ABA.LOCAIS), ATRACAO: ler_(ABA.ATRACOES) };
    const fotos = {};
    const sem = [];
    let pendentes = [];

    itens.forEach(function (it) {
      const lista = tabelas[it.tipo];
      if (!lista) return;
      const r = lista.find(function (x) { return x.id === it.id; });
      if (!r) return;
      const k = it.tipo + ':' + it.id;
      if (r.foto_url) { fotos[k] = { url: r.foto_url, cred: r.foto_credito }; return; }
      pendentes.push({ k: k, tipo: it.tipo, id: it.id, termos: [r.foto_busca || r.nome, r.nome] });
    });

    for (let rodada = 0; rodada < 2 && pendentes.length; rodada++) {
      const pedidos = pendentes.map(function (p) {
        return { url: urlCommons_(p.termos[rodada] || p.termos[0]), muteHttpExceptions: true };
      });
      const respostas = UrlFetchApp.fetchAll(pedidos);
      const faltam = [];
      pendentes.forEach(function (p, i) {
        const f = extrairFotoCommons_(respostas[i]);
        if (f) {
          fotos[p.k] = f;
          gravarFoto_(p.tipo, p.id, f.url, f.cred);
        } else {
          faltam.push(p);
        }
      });
      pendentes = faltam.filter(function (p) { return p.termos[1] && p.termos[1] !== p.termos[0]; });
    }
    itens.forEach(function (it) {
      const k = it.tipo + ':' + it.id;
      if (!fotos[k]) sem.push(k);
    });
    return { ok: true, fotos: fotos, sem: sem };
  } catch (e) {
    return { ok: false, erro: String(e.message || e) };
  }
}

function urlCommons_(termo) {
  return 'https://commons.wikimedia.org/w/api.php?action=query&format=json' +
    '&generator=search&gsrnamespace=6&gsrlimit=4&gsrsearch=' + encodeURIComponent(termo + ' filetype:bitmap') +
    '&prop=imageinfo&iiprop=url&iiurlwidth=640';
}

function extrairFotoCommons_(resp) {
  try {
    if (resp.getResponseCode() !== 200) return null;
    const j = JSON.parse(resp.getContentText());
    const pages = j && j.query && j.query.pages;
    if (!pages) return null;
    const lista = Object.keys(pages).map(function (k) { return pages[k]; })
      .sort(function (a, b) { return (a.index || 0) - (b.index || 0); });
    for (let i = 0; i < lista.length; i++) {
      const ii = lista[i].imageinfo && lista[i].imageinfo[0];
      if (ii && ii.thumburl && /\.(jpe?g|png|webp)$/i.test(ii.url)) return { url: ii.thumburl, cred: ii.descriptionurl };
    }
  } catch (e) { /* resposta inválida */ }
  return null;
}

function gravarFoto_(tipoItem, itemId, url, credito) {
  const nomeAba = tipoItem === 'LOCAL' ? ABA.LOCAIS : ABA.ATRACOES;
  const cab = CABECALHOS[nomeAba];
  const linhas = ler_(nomeAba);
  const i = linhas.findIndex(function (r) { return r.id === itemId; });
  if (i < 0 || linhas[i].foto_url !== '') return;
  aba_(nomeAba).getRange(i + 2, cab.indexOf('foto_url') + 1, 1, 2).setValues([[url, String(credito || '')]]);
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
  /*
   * Regra democrática: cada pessoa vale 1.
   *   a favor = 😍 Quero muito + 🙂 Tanto faz   |   contra = 🙅 Não quero
   *   saldo = a favor − contra  (é o número mostrado como "pontos")
   * Ordem: maior saldo → mais 😍 → menos 🙅.
   * Empate de votos: vai primeiro o que está mais perto de onde a família está (área vizinha no parque).
   */
  const cont = {};
  votos.forEach(function (v) {
    const k = v.tipo_item + ':' + v.item_id;
    const c = cont[k] || (cont[k] = { quero: 0, tanto: 0, nao: 0 });
    const val = Number(v.valor);
    if (val === 2) c.quero++; else if (val === 1) c.tanto++; else if (val === -1) c.nao++;
  });
  const placar = function (tipo, id) {
    const c = cont[tipo + ':' + id] || { quero: 0, tanto: 0, nao: 0 };
    const favor = c.quero + c.tanto;
    return { favor: favor, contra: c.nao, quero: c.quero, saldo: favor - c.nao };
  };
  const compararPlacar = function (a, b) {
    if (a.saldo !== b.saldo) return b.saldo - a.saldo;
    if (a.quero !== b.quero) return b.quero - a.quero;
    return a.contra - b.contra;
  };
  const mesmoNivel = function (a, b) { return a.saldo === b.saldo && a.quero === b.quero && a.contra === b.contra; };
  const entra = function (p) { return p.favor > 0 && p.saldo >= 0; };

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
        return l.dia_inteiro === 'SIM' && regioes.indexOf(l.regiao) >= 0 && placar('LOCAL', l.id).saldo >= 0;
      }).sort(function (a, b) {
        // Máximo de parques diferentes: 1º os votados ainda não usados, 2º os sem voto ainda não usados, por último repetir.
        const pa = placar('LOCAL', a.id), pb = placar('LOCAL', b.id);
        const nivel = function (l, p) { return (usados[l.id] || 0) > 0 ? 2 : (entra(p) ? 0 : 1); };
        const na = nivel(a, pa), nb = nivel(b, pb);
        if (na !== nb) return na - nb;
        const ua = usados[a.id] || 0, ub = usados[b.id] || 0;
        if (ua !== ub) return ua - ub;
        const c = compararPlacar(pa, pb);
        if (c !== 0) return c;
        return Number(a.ordem || 99) - Number(b.ordem || 99);
      });
      localId = candidatos.length ? candidatos[0].id : '';
      if (localId) usados[localId] = (usados[localId] || 0) + 1;
    }

    const local = localPorId[localId];
    saida.push([d.data, '0', d.cidade, d.tipo, localId || '', '', '', local ? local.nome : (TIPOS_AUTOMATICOS.indexOf(d.tipo) >= 0 ? 'A definir pela votação' : 'Dia livre'), local ? String(placar('LOCAL', localId).saldo) : '', d.observacao || '', carimbo]);
    if (!local) return;

    const lista = atracoes.filter(function (a) { return a.local_id === localId; });

    // Áreas do parque em ordem de caminhada (o catálogo segue a volta real do parque).
    const areas = [];
    lista.slice().sort(function (a, b) { return Number(a.ordem_area) - Number(b.ordem_area); }).forEach(function (a) {
      if (areas.indexOf(a.area) < 0) areas.push(a.area);
    });
    const distancia = function (areaA, areaB) {
      const ia = areas.indexOf(areaA), ib = areas.indexOf(areaB);
      if (ia < 0 || ib < 0) return 99;
      const dd = Math.abs(ia - ib);
      return Math.min(dd, areas.length - dd);
    };

    // Só entram atrações que a maioria de quem votou quer.
    const restantes = lista.filter(function (a) { return entra(placar('ATRACAO', a.id)); });
    const escolhidas = [];
    let areaAtual = areas[0];
    while (restantes.length) {
      restantes.sort(function (a, b) { return compararPlacar(placar('ATRACAO', a.id), placar('ATRACAO', b.id)); });
      const topo = placar('ATRACAO', restantes[0].id);
      const empatadas = restantes.filter(function (a) { return mesmoNivel(placar('ATRACAO', a.id), topo); });
      empatadas.sort(function (a, b) {
        const da = distancia(areaAtual, a.area), db = distancia(areaAtual, b.area);
        if (da !== db) return da - db;
        const oa = areas.indexOf(a.area), ob = areas.indexOf(b.area);
        if (oa !== ob) return oa - ob;
        return a.nome < b.nome ? -1 : 1;
      });
      const prox = empatadas[0];
      escolhidas.push(prox);
      areaAtual = prox.area;
      restantes.splice(restantes.indexOf(prox), 1);
    }

    escolhidas.forEach(function (a, idx) {
      const min = Number(a.altura_min_cm || 0);
      let alerta = '';
      if (min > 0) {
        const baixos = pessoas.filter(function (p) { return p.altura_cm !== '' && Number(p.altura_cm) < min; }).map(function (p) { return p.nome; });
        if (baixos.length) alerta = 'Altura mínima ' + min + ' cm — não podem ir: ' + baixos.join(', ');
      }
      saida.push([d.data, String(idx + 1), d.cidade, d.tipo, localId, a.area, a.id, a.nome, String(placar('ATRACAO', a.id).saldo), alerta, carimbo]);
    });
  });
  return saida;
}


/* =========================================================
 *  PDF DO ROTEIRO (ao vivo, com quem votou)
 * ========================================================= */

function api_pdf() {
  try {
    garantirSetup_();
    const dias = ler_(ABA.DIAS).slice().sort(function (a, b) { return a.data < b.data ? -1 : 1; });
    const locais = ler_(ABA.LOCAIS);
    const atracoes = ler_(ABA.ATRACOES).filter(function (a) { return a.ativo !== 'NAO'; });
    const pessoas = ler_(ABA.PESSOAS);
    const votos = ler_(ABA.VOTOS);
    const carimbo = agora_();
    const linhas = montarRoteiro_(dias, locais, atracoes, pessoas, votos, carimbo);
    const html = htmlPdf_(linhas, dias, locais, atracoes, pessoas, votos, carimbo);
    const nome = 'Roteiro-Orlando-2027.pdf';
    const pdf = Utilities.newBlob(html, 'text/html', 'roteiro.html').getAs('application/pdf').setName(nome);
    return { ok: true, nome: nome, base64: Utilities.base64Encode(pdf.getBytes()) };
  } catch (e) {
    return { ok: false, erro: String(e.message || e) };
  }
}

function escPdf_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function dataPdf_(iso) {
  const p = String(iso).split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2]);
  const sem = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'][d.getDay()];
  return sem + ', ' + ('0' + p[2]).slice(-2) + '/' + ('0' + p[1]).slice(-2) + '/' + p[0];
}

function quemVotou_(votos, pessoas, tipo, id) {
  const nome = {};
  pessoas.forEach(function (p) { nome[p.id] = p.nome; });
  const g = { '2': [], '1': [], '-1': [] };
  let pts = 0;
  votos.forEach(function (v) {
    if (v.tipo_item !== tipo || v.item_id !== id || !nome[v.pessoa_id]) return;
    const val = String(Number(v.valor));
    if (g[val]) g[val].push(nome[v.pessoa_id]);
    pts += Number(v.valor);
  });
  return { quero: g['2'], tanto: g['1'], nao: g['-1'], pts: (g['2'].length + g['1'].length) - g['-1'].length, pontosAntigos: pts };
}

function textoVotos_(q) {
  const partes = [];
  if (q.quero.length) partes.push('<b>Quero muito:</b> ' + escPdf_(q.quero.join(', ')));
  if (q.tanto.length) partes.push('<b>Tanto faz:</b> ' + escPdf_(q.tanto.join(', ')));
  if (q.nao.length) partes.push('<b>Não quero:</b> ' + escPdf_(q.nao.join(', ')));
  return partes.length ? partes.join('<br>') : '<i>ninguém votou</i>';
}

function htmlPdf_(linhas, dias, locais, atracoes, pessoas, votos, carimbo) {
  const nomesTipo = { CHEGADA: 'Chegada', PARQUE: 'Parque', MIAMI: 'Miami', RETORNO: 'Retorno', LIVRE: 'Livre' };
  const localPorId = {};
  locais.forEach(function (l) { localPorId[l.id] = l; });
  const porDia = {};
  const ordem = [];
  linhas.forEach(function (r) {
    if (!porDia[r[0]]) { porDia[r[0]] = []; ordem.push(r[0]); }
    porDia[r[0]].push(r);
  });
  ordem.sort();

  const css = 'body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#172036;}' +
    'h1{font-size:18pt;color:#0b1f4b;margin:0 0 2px 0;}' +
    '.sub{color:#5d6a85;font-size:9pt;margin-bottom:12px;}' +
    '.dia{border:1px solid #c9d3e6;border-left:6px solid #f5b82e;margin:0 0 12px 0;padding:8px 10px;page-break-inside:avoid;}' +
    '.dt{font-size:12pt;font-weight:bold;color:#0b1f4b;}' +
    '.lc{font-size:11pt;font-weight:bold;margin:2px 0 4px 0;}' +
    '.obs{color:#5d6a85;font-size:9pt;margin:2px 0 6px 0;}' +
    '.vp{background:#f3f6fb;padding:5px 7px;font-size:9pt;margin:4px 0 6px 0;}' +
    'table{width:100%;border-collapse:collapse;font-size:9pt;}' +
    'th{background:#0b1f4b;color:#ffffff;text-align:left;padding:4px;}' +
    'td{border-bottom:1px solid #dfe6f2;padding:4px;vertical-align:top;}' +
    '.n{width:18px;text-align:center;font-weight:bold;}' +
    '.p{width:40px;text-align:center;font-weight:bold;color:#0b1f4b;}' +
    '.al{color:#d93636;font-size:8.5pt;}' +
    'h2{font-size:13pt;color:#0b1f4b;margin:16px 0 6px 0;}';

  let h = '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>';
  h += '<h1>Roteiro Orlando 2027</h1>';
  h += '<div class="sub">Tampa → Orlando → Miami → Tampa · 08 a 19/02/2027 · montado pelos votos em ' + escPdf_(carimbo) +
    ' · ' + pessoas.length + ' pessoas · ' + votos.length + ' votos</div>' +
    '<div class="sub">Regra: cada pessoa vale 1 voto. Saldo = (Quero muito + Tanto faz) − Não quero. Empate: mais "Quero muito" primeiro e, entre iguais, a atração mais perto da anterior no parque.</div>';

  ordem.forEach(function (data) {
    const rs = porDia[data].sort(function (a, b) { return Number(a[1]) - Number(b[1]); });
    const cab = rs[0];
    const itens = rs.slice(1);
    const local = localPorId[cab[4]];
    h += '<div class="dia"><div class="dt">' + escPdf_(dataPdf_(data)) + ' · ' + escPdf_(nomesTipo[cab[3]] || cab[3]) + ' · ' + escPdf_(cab[2]) + '</div>';
    h += '<div class="lc">' + escPdf_(cab[7]) + '</div>';
    if (cab[9]) h += '<div class="obs">' + escPdf_(cab[9]) + '</div>';
    if (local) {
      const qp = quemVotou_(votos, pessoas, 'LOCAL', local.id);
      h += '<div class="vp"><b>Votos neste local (' + qp.pts + ' pontos):</b><br>' + textoVotos_(qp) + '</div>';
      if (itens.length) {
        h += '<table><tr><th class="n">#</th><th>Atração</th><th>Área</th><th class="p">Saldo</th><th>Quem votou</th></tr>';
        itens.forEach(function (it) {
          const qa = quemVotou_(votos, pessoas, 'ATRACAO', it[6]);
          h += '<tr><td class="n">' + escPdf_(it[1]) + '</td><td><b>' + escPdf_(it[7]) + '</b>' +
            (it[9] ? '<br><span class="al">' + escPdf_(it[9]) + '</span>' : '') + '</td><td>' + escPdf_(it[5]) +
            '</td><td class="p">' + escPdf_(it[8]) + '</td><td>' + textoVotos_(qa) + '</td></tr>';
        });
        h += '</table>';
      } else {
        h += '<div class="obs">Nenhuma atração com votos positivos ainda.</div>';
      }
    }
    h += '</div>';
  });

  // Ranking geral
  const rank = locais.filter(function (l) { return l.dia_inteiro === 'SIM'; }).map(function (l) {
    return { l: l, q: quemVotou_(votos, pessoas, 'LOCAL', l.id) };
  }).sort(function (a, b) { return b.q.pts - a.q.pts || b.q.quero.length - a.q.quero.length || a.q.nao.length - b.q.nao.length || Number(a.l.ordem) - Number(b.l.ordem); });
  h += '<h2>Ranking dos parques e passeios</h2><table><tr><th class="n">#</th><th>Local</th><th class="p">Saldo</th><th>Quem votou</th></tr>';
  rank.forEach(function (r, i) {
    h += '<tr><td class="n">' + (i + 1) + '</td><td>' + escPdf_(r.l.nome) + '</td><td class="p">' + r.q.pts + '</td><td>' + textoVotos_(r.q) + '</td></tr>';
  });
  h += '</table>';

  // Participação
  const contagem = {};
  votos.forEach(function (v) { contagem[v.pessoa_id] = (contagem[v.pessoa_id] || 0) + 1; });
  h += '<h2>Participação</h2><table><tr><th>Pessoa</th><th class="p">Votos</th></tr>';
  pessoas.forEach(function (p) {
    h += '<tr><td>' + escPdf_(p.nome) + (p.altura_cm ? ' (' + escPdf_(p.altura_cm) + ' cm)' : '') + '</td><td class="p">' + (contagem[p.id] || 0) + '</td></tr>';
  });
  h += '</table></body></html>';
  return h;
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

/* =========================================================
 *  FOTOS DO CATÁLOGO — Wikimedia Commons (licença livre; crédito no link de cada foto)
 *  [url da miniatura 500px, página da foto com autor e licença]
 * ========================================================= */
const FOTOS_CATALOGO = {
  "AK": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a4/Tree_of_Life,_Disney's_Animal_Kingdom.jpg/500px-Tree_of_Life,_Disney's_Animal_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Tree_of_Life,_Disney's_Animal_Kingdom.jpg"],
  "AK-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a4/Tree_of_Life,_Disney's_Animal_Kingdom.jpg/500px-Tree_of_Life,_Disney's_Animal_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Tree_of_Life,_Disney's_Animal_Kingdom.jpg"],
  "AK-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/5a/Tree_of_Life_at_Animal_Kingdom,_Discovery_Island,_Disney's_Animal_Kingdom,_June_2006_-_ThrillZing.jpg/500px-Tree_of_Life_at_Animal_Kingdom,_Discovery_Island,_Disney's_Animal_Kingdom,_June_2006_-_ThrillZing.jpg","https://commons.wikimedia.org/wiki/File:Tree_of_Life_at_Animal_Kingdom,_Discovery_Island,_Disney's_Animal_Kingdom,_June_2006_-_ThrillZing.jpg"],
  "AK-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/Boat_for_Na'vi_River_Journey_(34716522955).jpg/500px-Boat_for_Na'vi_River_Journey_(34716522955).jpg","https://commons.wikimedia.org/wiki/File:Boat_for_Na'vi_River_Journey_(34716522955).jpg"],
  "AK-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/Kilimanjaro_Safaris_lion.JPG/500px-Kilimanjaro_Safaris_lion.JPG","https://commons.wikimedia.org/wiki/File:Kilimanjaro_Safaris_lion.JPG"],
  "AK-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/02/Pangani_Forest_Exploration_Trail_(42377914135).jpg/500px-Pangani_Forest_Exploration_Trail_(42377914135).jpg","https://commons.wikimedia.org/wiki/File:Pangani_Forest_Exploration_Trail_(42377914135).jpg"],
  "AK-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/e/e7/Disney-Animal-Kingdom-Lion-King-8457.jpg/500px-Disney-Animal-Kingdom-Lion-King-8457.jpg","https://commons.wikimedia.org/wiki/File:Disney-Animal-Kingdom-Lion-King-8457.jpg"],
  "AK-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/8b/Expedition_Everest.jpg/500px-Expedition_Everest.jpg","https://commons.wikimedia.org/wiki/File:Expedition_Everest.jpg"],
  "AK-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/90/Kali_River_Rapids.jpg/500px-Kali_River_Rapids.jpg","https://commons.wikimedia.org/wiki/File:Kali_River_Rapids.jpg"],
  "AK-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/64/Maharajah_Jungle_Trek_(42377960765).jpg/500px-Maharajah_Jungle_Trek_(42377960765).jpg","https://commons.wikimedia.org/wiki/File:Maharajah_Jungle_Trek_(42377960765).jpg"],
  "AK-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/05/Caravan_Stage_(43281745001).jpg/500px-Caravan_Stage_(43281745001).jpg","https://commons.wikimedia.org/wiki/File:Caravan_Stage_(43281745001).jpg"],
  "BB": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/91/Summit_Plummet_from_Blizzard_Beach_Entrance_-_panoramio.jpg/500px-Summit_Plummet_from_Blizzard_Beach_Entrance_-_panoramio.jpg","https://commons.wikimedia.org/wiki/File:Summit_Plummet_from_Blizzard_Beach_Entrance_-_panoramio.jpg"],
  "BB-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/91/Summit_Plummet_from_Blizzard_Beach_Entrance_-_panoramio.jpg/500px-Summit_Plummet_from_Blizzard_Beach_Entrance_-_panoramio.jpg","https://commons.wikimedia.org/wiki/File:Summit_Plummet_from_Blizzard_Beach_Entrance_-_panoramio.jpg"],
  "BB-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/38/Team_Boat_Springs.jpg/500px-Team_Boat_Springs.jpg","https://commons.wikimedia.org/wiki/File:Team_Boat_Springs.jpg"],
  "BB-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/38/Team_Boat_Springs.jpg/500px-Team_Boat_Springs.jpg","https://commons.wikimedia.org/wiki/File:Team_Boat_Springs.jpg"],
  "BB-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/6c/Blizzard_Beach_Aerial_photo_(7426584544).jpg/500px-Blizzard_Beach_Aerial_photo_(7426584544).jpg","https://commons.wikimedia.org/wiki/File:Blizzard_Beach_Aerial_photo_(7426584544).jpg"],
  "BGT": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d8/Jungala_entrance_Busch_Gardens_Tampa_(1).jpg/500px-Jungala_entrance_Busch_Gardens_Tampa_(1).jpg","https://commons.wikimedia.org/wiki/File:Jungala_entrance_Busch_Gardens_Tampa_(1).jpg"],
  "BGT-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/73/Iron_Gwazi_(11).jpg/500px-Iron_Gwazi_(11).jpg","https://commons.wikimedia.org/wiki/File:Iron_Gwazi_(11).jpg"],
  "BGT-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/97/Bengal_tiger_(Panthera_tigris_tigris)_female_3.jpg/500px-Bengal_tiger_(Panthera_tigris_tigris)_female_3.jpg","https://commons.wikimedia.org/wiki/File:Bengal_tiger_(Panthera_tigris_tigris)_female_3.jpg"],
  "BGT-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/Falcon's_Fury_(August_17,_2014).jpg/500px-Falcon's_Fury_(August_17,_2014).jpg","https://commons.wikimedia.org/wiki/File:Falcon's_Fury_(August_17,_2014).jpg"],
  "BGT-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d5/Cheetah_Hunt_@_Busch_Gardens_2012_-_panoramio.jpg/500px-Cheetah_Hunt_@_Busch_Gardens_2012_-_panoramio.jpg","https://commons.wikimedia.org/wiki/File:Cheetah_Hunt_@_Busch_Gardens_2012_-_panoramio.jpg"],
  "BGT-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2d/Montu_(Busch_Gardens_Africa)_01.jpg/500px-Montu_(Busch_Gardens_Africa)_01.jpg","https://commons.wikimedia.org/wiki/File:Montu_(Busch_Gardens_Africa)_01.jpg"],
  "BGT-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2b/Cobra's_Curse_queue_BGT_(2).jpg/500px-Cobra's_Curse_queue_BGT_(2).jpg","https://commons.wikimedia.org/wiki/File:Cobra's_Curse_queue_BGT_(2).jpg"],
  "BGT-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/Kumba_(Busch_Gardens_Africa)_01.jpg/500px-Kumba_(Busch_Gardens_Africa)_01.jpg","https://commons.wikimedia.org/wiki/File:Kumba_(Busch_Gardens_Africa)_01.jpg"],
  "BGT-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0d/Congo_section_Busch_Gardens_Tampa_(1).jpg/500px-Congo_section_Busch_Gardens_Tampa_(1).jpg","https://commons.wikimedia.org/wiki/File:Congo_section_Busch_Gardens_Tampa_(1).jpg"],
  "BGT-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/82/SheiKra_(Busch_Gardens_Africa)_05.JPG/500px-SheiKra_(Busch_Gardens_Africa)_05.JPG","https://commons.wikimedia.org/wiki/File:SheiKra_(Busch_Gardens_Africa)_05.JPG"],
  "BGT-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/86/Phoenix_Rising_view.jpg/500px-Phoenix_Rising_view.jpg","https://commons.wikimedia.org/wiki/File:Phoenix_Rising_view.jpg"],
  "BGT-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/31/Miniature_train_at_Busch_Gardens,_Tampa,_Florida,_1972_(1_of_2).jpg/500px-Miniature_train_at_Busch_Gardens,_Tampa,_Florida,_1972_(1_of_2).jpg","https://commons.wikimedia.org/wiki/File:Miniature_train_at_Busch_Gardens,_Tampa,_Florida,_1972_(1_of_2).jpg"],
  "CLW": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/Clearwater-beach-florida-pier-60.jpg/500px-Clearwater-beach-florida-pier-60.jpg","https://commons.wikimedia.org/wiki/File:Clearwater-beach-florida-pier-60.jpg"],
  "CLW-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7e/Clearwater_Beach_pirate_ship_excursion.jpg/500px-Clearwater_Beach_pirate_ship_excursion.jpg","https://commons.wikimedia.org/wiki/File:Clearwater_Beach_pirate_ship_excursion.jpg"],
  "CLW-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/Clearwater-beach-florida-pier-60.jpg/500px-Clearwater-beach-florida-pier-60.jpg","https://commons.wikimedia.org/wiki/File:Clearwater-beach-florida-pier-60.jpg"],
  "CLW-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/View_of_Clearwater_Marine_Aquarium_-_panoramio.jpg/500px-View_of_Clearwater_Marine_Aquarium_-_panoramio.jpg","https://commons.wikimedia.org/wiki/File:View_of_Clearwater_Marine_Aquarium_-_panoramio.jpg"],
  "EPCOT": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/33/Spaceship_Earth_tiles_(close).jpg/500px-Spaceship_Earth_tiles_(close).jpg","https://commons.wikimedia.org/wiki/File:Spaceship_Earth_tiles_(close).jpg"],
  "EPCOT-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/33/Spaceship_Earth_tiles_(close).jpg/500px-Spaceship_Earth_tiles_(close).jpg","https://commons.wikimedia.org/wiki/File:Spaceship_Earth_tiles_(close).jpg"],
  "EPCOT-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d0/Journey_into_Imagination_ride_logo.jpg/500px-Journey_into_Imagination_ride_logo.jpg","https://commons.wikimedia.org/wiki/File:Journey_into_Imagination_ride_logo.jpg"],
  "EPCOT-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/6a/Guardians_of_the_Galaxy_Cosmic_Rewind_at_EPCOT.jpg/500px-Guardians_of_the_Galaxy_Cosmic_Rewind_at_EPCOT.jpg","https://commons.wikimedia.org/wiki/File:Guardians_of_the_Galaxy_Cosmic_Rewind_at_EPCOT.jpg"],
  "EPCOT-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/e/ed/TestTrackEpcot.JPG/500px-TestTrackEpcot.JPG","https://commons.wikimedia.org/wiki/File:TestTrackEpcot.JPG"],
  "EPCOT-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/59/WDW_EPCOT_Mission_Space.JPG/500px-WDW_EPCOT_Mission_Space.JPG","https://commons.wikimedia.org/wiki/File:WDW_EPCOT_Mission_Space.JPG"],
  "EPCOT-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b5/Soarin_Seating_Epcot_Center.jpg/500px-Soarin_Seating_Epcot_Center.jpg","https://commons.wikimedia.org/wiki/File:Soarin_Seating_Epcot_Center.jpg"],
  "EPCOT-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c3/Living_with_the_Land_entrance.jpg/500px-Living_with_the_Land_entrance.jpg","https://commons.wikimedia.org/wiki/File:Living_with_the_Land_entrance.jpg"],
  "EPCOT-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/05/The_entrance_to_'The_Seas_With_Nemo_%26_Friends'_at_Epcot.jpg/500px-The_entrance_to_'The_Seas_With_Nemo_%26_Friends'_at_Epcot.jpg","https://commons.wikimedia.org/wiki/File:The_entrance_to_'The_Seas_With_Nemo_%26_Friends'_at_Epcot.jpg"],
  "EPCOT-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b5/Turtle_Talk_with_Crush,_Epcot.jpg/500px-Turtle_Talk_with_Crush,_Epcot.jpg","https://commons.wikimedia.org/wiki/File:Turtle_Talk_with_Crush,_Epcot.jpg"],
  "EPCOT-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c7/Journey_of_Water_Cast_Preview_(53282976488).jpg/500px-Journey_of_Water_Cast_Preview_(53282976488).jpg","https://commons.wikimedia.org/wiki/File:Journey_of_Water_Cast_Preview_(53282976488).jpg"],
  "EPCOT-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bd/Frozen_Ever_After_Sign_(27803925036).jpg/500px-Frozen_Ever_After_Sign_(27803925036).jpg","https://commons.wikimedia.org/wiki/File:Frozen_Ever_After_Sign_(27803925036).jpg"],
  "EPCOT-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/68/Remy's_Ratatouille_Adventure.jpg/500px-Remy's_Ratatouille_Adventure.jpg","https://commons.wikimedia.org/wiki/File:Remy's_Ratatouille_Adventure.jpg"],
  "EPCOT-13": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/4d/Gran_Fiesta_Tour_Starring_The_Three_Caballeros.jpg/500px-Gran_Fiesta_Tour_Starring_The_Three_Caballeros.jpg","https://commons.wikimedia.org/wiki/File:Gran_Fiesta_Tour_Starring_The_Three_Caballeros.jpg"],
  "EPCOT-14": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/Reflection_of_China_China_Pavilion_Epcot_Center.jpg/500px-Reflection_of_China_China_Pavilion_Epcot_Center.jpg","https://commons.wikimedia.org/wiki/File:Reflection_of_China_China_Pavilion_Epcot_Center.jpg"],
  "EPCOT-15": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/32/Epcot_-_54163012842.jpg/500px-Epcot_-_54163012842.jpg","https://commons.wikimedia.org/wiki/File:Epcot_-_54163012842.jpg"],
  "EPCOT-16": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/23/Starburst_(49576259231).jpg/500px-Starburst_(49576259231).jpg","https://commons.wikimedia.org/wiki/File:Starburst_(49576259231).jpg"],
  "EPIC": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/5f/Chronos_Tower_Front_Epic_Universe.jpg/500px-Chronos_Tower_Front_Epic_Universe.jpg","https://commons.wikimedia.org/wiki/File:Chronos_Tower_Front_Epic_Universe.jpg"],
  "EPIC-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Stardust_Racers_entrance.jpg/500px-Stardust_Racers_entrance.jpg","https://commons.wikimedia.org/wiki/File:Stardust_Racers_entrance.jpg"],
  "EPIC-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/fc/Constellation_Carousel.jpg/500px-Constellation_Carousel.jpg","https://commons.wikimedia.org/wiki/File:Constellation_Carousel.jpg"],
  "EPIC-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/5f/Chronos_Tower_Front_Epic_Universe.jpg/500px-Chronos_Tower_Front_Epic_Universe.jpg","https://commons.wikimedia.org/wiki/File:Chronos_Tower_Front_Epic_Universe.jpg"],
  "EPIC-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/31/Mario_Kart_Bowser's_Challenge_(Universal_Epic_Universe)_1.jpg/500px-Mario_Kart_Bowser's_Challenge_(Universal_Epic_Universe)_1.jpg","https://commons.wikimedia.org/wiki/File:Mario_Kart_Bowser's_Challenge_(Universal_Epic_Universe)_1.jpg"],
  "EPIC-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/79/Mine-Cart_Madness_on-ride_view_-_Universal_Epic_Universe.png/500px-Mine-Cart_Madness_on-ride_view_-_Universal_Epic_Universe.png","https://commons.wikimedia.org/wiki/File:Mine-Cart_Madness_on-ride_view_-_Universal_Epic_Universe.png"],
  "EPIC-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/25/Super_Nintendo_World,_Epic_Universe,_Orlando,_Fl%C3%B3rida.jpg/500px-Super_Nintendo_World,_Epic_Universe,_Orlando,_Fl%C3%B3rida.jpg","https://commons.wikimedia.org/wiki/File:Super_Nintendo_World,_Epic_Universe,_Orlando,_Fl%C3%B3rida.jpg"],
  "EPIC-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/18/Hiccup's_Wing_Gliders_1.jpg/500px-Hiccup's_Wing_Gliders_1.jpg","https://commons.wikimedia.org/wiki/File:Hiccup's_Wing_Gliders_1.jpg"],
  "EPIC-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/88/How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg/500px-How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg","https://commons.wikimedia.org/wiki/File:How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg"],
  "EPIC-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/88/How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg/500px-How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg","https://commons.wikimedia.org/wiki/File:How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg"],
  "EPIC-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/88/How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg/500px-How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg","https://commons.wikimedia.org/wiki/File:How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg"],
  "EPIC-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/88/How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg/500px-How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg","https://commons.wikimedia.org/wiki/File:How_to_Train_Your_Dragon_-_Isle_of_Berk_-_54563945175.jpg"],
  "EPIC-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/09/Ministry_of_Magic_Metrofloo,_The_Wizarding_World_of_Harry_Potter.jpg/500px-Ministry_of_Magic_Metrofloo,_The_Wizarding_World_of_Harry_Potter.jpg","https://commons.wikimedia.org/wiki/File:Ministry_of_Magic_Metrofloo,_The_Wizarding_World_of_Harry_Potter.jpg"],
  "EPIC-13": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/86/Preshow_tent_Le_Cirque_Arcanus_Epic_Universe.jpg/500px-Preshow_tent_Le_Cirque_Arcanus_Epic_Universe.jpg","https://commons.wikimedia.org/wiki/File:Preshow_tent_Le_Cirque_Arcanus_Epic_Universe.jpg"],
  "EPIC-14": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2e/Monsters_Unchained_The_Frankenstein_Experiment.jpg/500px-Monsters_Unchained_The_Frankenstein_Experiment.jpg","https://commons.wikimedia.org/wiki/File:Monsters_Unchained_The_Frankenstein_Experiment.jpg"],
  "EPIC-15": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/49/Curse_of_the_Werewolf.jpg/500px-Curse_of_the_Werewolf.jpg","https://commons.wikimedia.org/wiki/File:Curse_of_the_Werewolf.jpg"],
  "HS": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bd/Hollywood_Tower_Hotel.jpg/500px-Hollywood_Tower_Hotel.jpg","https://commons.wikimedia.org/wiki/File:Hollywood_Tower_Hotel.jpg"],
  "HS-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1f/Chinese_Theater_(49560227308).jpg/500px-Chinese_Theater_(49560227308).jpg","https://commons.wikimedia.org/wiki/File:Chinese_Theater_(49560227308).jpg"],
  "HS-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2d/Star_Tours_%E2%80%93_The_Adventures_Continue_(Disney's_Hollywood_Studios)_2.jpg/500px-Star_Tours_%E2%80%93_The_Adventures_Continue_(Disney's_Hollywood_Studios)_2.jpg","https://commons.wikimedia.org/wiki/File:Star_Tours_%E2%80%93_The_Adventures_Continue_(Disney's_Hollywood_Studios)_2.jpg"],
  "HS-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/be/Indiana_Jones_Stunt_Spectacular.jpg/500px-Indiana_Jones_Stunt_Spectacular.jpg","https://commons.wikimedia.org/wiki/File:Indiana_Jones_Stunt_Spectacular.jpg"],
  "HS-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1c/For_the_First_Time_in_Forever-_A_Frozen_Sing-Along_Celebration_-_21834852103.jpg/500px-For_the_First_Time_in_Forever-_A_Frozen_Sing-Along_Celebration_-_21834852103.jpg","https://commons.wikimedia.org/wiki/File:For_the_First_Time_in_Forever-_A_Frozen_Sing-Along_Celebration_-_21834852103.jpg"],
  "HS-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0a/Rise_of_the_Resistance.jpg/500px-Rise_of_the_Resistance.jpg","https://commons.wikimedia.org/wiki/File:Rise_of_the_Resistance.jpg"],
  "HS-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2c/Millennium_Falcon_Star_Wars_Galaxy's_edge.jpg/500px-Millennium_Falcon_Star_Wars_Galaxy's_edge.jpg","https://commons.wikimedia.org/wiki/File:Millennium_Falcon_Star_Wars_Galaxy's_edge.jpg"],
  "HS-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/09/Slinky_Dog_Dash.jpg/500px-Slinky_Dog_Dash.jpg","https://commons.wikimedia.org/wiki/File:Slinky_Dog_Dash.jpg"],
  "HS-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/ca/Toy_Story_Mania_(23822754735).jpg/500px-Toy_Story_Mania_(23822754735).jpg","https://commons.wikimedia.org/wiki/File:Toy_Story_Mania_(23822754735).jpg"],
  "HS-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/e/e2/Alien_Swirling_Saucers_(43084274482).jpg/500px-Alien_Swirling_Saucers_(43084274482).jpg","https://commons.wikimedia.org/wiki/File:Alien_Swirling_Saucers_(43084274482).jpg"],
  "HS-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0d/The_Little_Mermaid_a_Musical_Adventure.png/500px-The_Little_Mermaid_a_Musical_Adventure.png","https://commons.wikimedia.org/wiki/File:The_Little_Mermaid_a_Musical_Adventure.png"],
  "HS-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bd/Hollywood_Tower_Hotel.jpg/500px-Hollywood_Tower_Hotel.jpg","https://commons.wikimedia.org/wiki/File:Hollywood_Tower_Hotel.jpg"],
  "HS-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b5/Rock_'n'_Roller_Coaster_(29380299068).jpg/500px-Rock_'n'_Roller_Coaster_(29380299068).jpg","https://commons.wikimedia.org/wiki/File:Rock_'n'_Roller_Coaster_(29380299068).jpg"],
  "HS-13": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b7/Beauty_and_the_Beast_Live_on_Stage_at_Disney's_Hollywood_Studios_in_Orlando,_FL.jpg/500px-Beauty_and_the_Beast_Live_on_Stage_at_Disney's_Hollywood_Studios_in_Orlando,_FL.jpg","https://commons.wikimedia.org/wiki/File:Beauty_and_the_Beast_Live_on_Stage_at_Disney's_Hollywood_Studios_in_Orlando,_FL.jpg"],
  "HS-14": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/cf/Fantasmic!_(Disney's_Hollywood_Studios)_2.jpg/500px-Fantasmic!_(Disney's_Hollywood_Studios)_2.jpg","https://commons.wikimedia.org/wiki/File:Fantasmic!_(Disney's_Hollywood_Studios)_2.jpg"],
  "IOA": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b6/Hogsmeade_(Islands_of_Adventure)_May_2023.JPG/500px-Hogsmeade_(Islands_of_Adventure)_May_2023.JPG","https://commons.wikimedia.org/wiki/File:Hogsmeade_(Islands_of_Adventure)_May_2023.JPG"],
  "IOA-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1f/Incredible_Hulk_Coaster.png/500px-Incredible_Hulk_Coaster.png","https://commons.wikimedia.org/wiki/File:Incredible_Hulk_Coaster.png"],
  "IOA-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/The_Amazing_Adventures_of_Spider-Man_entrance_1.jpg/500px-The_Amazing_Adventures_of_Spider-Man_entrance_1.jpg","https://commons.wikimedia.org/wiki/File:The_Amazing_Adventures_of_Spider-Man_entrance_1.jpg"],
  "IOA-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/4b/Doctor_Doom's_Fearfall_1.jpg/500px-Doctor_Doom's_Fearfall_1.jpg","https://commons.wikimedia.org/wiki/File:Doctor_Doom's_Fearfall_1.jpg"],
  "IOA-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/StormForce.JPG/500px-StormForce.JPG","https://commons.wikimedia.org/wiki/File:StormForce.JPG"],
  "IOA-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b5/Ripsaw_Falls_at_Islands_of_Adventure.jpg/500px-Ripsaw_Falls_at_Islands_of_Adventure.jpg","https://commons.wikimedia.org/wiki/File:Ripsaw_Falls_at_Islands_of_Adventure.jpg"],
  "IOA-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/24/Popeye_and_Blutos_Bilge-Rat_Barges_13.jpg/500px-Popeye_and_Blutos_Bilge-Rat_Barges_13.jpg","https://commons.wikimedia.org/wiki/File:Popeye_and_Blutos_Bilge-Rat_Barges_13.jpg"],
  "IOA-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0b/Skull_Island_Reign_of_Kong_01.jpg/500px-Skull_Island_Reign_of_Kong_01.jpg","https://commons.wikimedia.org/wiki/File:Skull_Island_Reign_of_Kong_01.jpg"],
  "IOA-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/18/Islands_of_Adventure_Orlando_(51149417278)_(cropped).jpg/500px-Islands_of_Adventure_Orlando_(51149417278)_(cropped).jpg","https://commons.wikimedia.org/wiki/File:Islands_of_Adventure_Orlando_(51149417278)_(cropped).jpg"],
  "IOA-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Universal-Islands-of-Adventure-Jurassic-Park-River-9210.jpg/500px-Universal-Islands-of-Adventure-Jurassic-Park-River-9210.jpg","https://commons.wikimedia.org/wiki/File:Universal-Islands-of-Adventure-Jurassic-Park-River-9210.jpg"],
  "IOA-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/Pteranodon_Flyers_(Islands_of_Adventure).jpg/500px-Pteranodon_Flyers_(Islands_of_Adventure).jpg","https://commons.wikimedia.org/wiki/File:Pteranodon_Flyers_(Islands_of_Adventure).jpg"],
  "IOA-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1e/Camp_Jurassic_at_Islands_of_Adventure.jpg/500px-Camp_Jurassic_at_Islands_of_Adventure.jpg","https://commons.wikimedia.org/wiki/File:Camp_Jurassic_at_Islands_of_Adventure.jpg"],
  "IOA-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/f4/Hagrid's_Magical_Creatures_Motorbike_Adventure_2.jpg/500px-Hagrid's_Magical_Creatures_Motorbike_Adventure_2.jpg","https://commons.wikimedia.org/wiki/File:Hagrid's_Magical_Creatures_Motorbike_Adventure_2.jpg"],
  "IOA-13": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c4/Wizarding_World_of_Harry_Potter_Castle.jpg/500px-Wizarding_World_of_Harry_Potter_Castle.jpg","https://commons.wikimedia.org/wiki/File:Wizarding_World_of_Harry_Potter_Castle.jpg"],
  "IOA-14": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b5/Flight_of_the_Hippogriff_at_Islands_of_Adventure.jpg/500px-Flight_of_the_Hippogriff_at_Islands_of_Adventure.jpg","https://commons.wikimedia.org/wiki/File:Flight_of_the_Hippogriff_at_Islands_of_Adventure.jpg"],
  "IOA-15": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/f0/Hogsmeade_Station_(Islands_of_Adventure)_Hogwarts_Express_-_Sept2018.jpg/500px-Hogsmeade_Station_(Islands_of_Adventure)_Hogwarts_Express_-_Sept2018.jpg","https://commons.wikimedia.org/wiki/File:Hogsmeade_Station_(Islands_of_Adventure)_Hogwarts_Express_-_Sept2018.jpg"],
  "IOA-16": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/6d/The_Cat_in_the_Hat_dark_ride_entrance.jpg/500px-The_Cat_in_the_Hat_dark_ride_entrance.jpg","https://commons.wikimedia.org/wiki/File:The_Cat_in_the_Hat_dark_ride_entrance.jpg"],
  "IOA-17": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/8e/OneFishTwoFishRedFishBlueFishEntrance.JPG/500px-OneFishTwoFishRedFishBlueFishEntrance.JPG","https://commons.wikimedia.org/wiki/File:OneFishTwoFishRedFishBlueFishEntrance.JPG"],
  "IOA-18": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/22/TheHighInTheSkySeussTrolleyTrainRide.JPG/500px-TheHighInTheSkySeussTrolleyTrainRide.JPG","https://commons.wikimedia.org/wiki/File:TheHighInTheSkySeussTrolleyTrainRide.JPG"],
  "IOA-19": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/Caro-Seuss-el.jpg/500px-Caro-Seuss-el.jpg","https://commons.wikimedia.org/wiki/File:Caro-Seuss-el.jpg"],
  "KSC": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/88/Atlantis_Kennedy_Space_Center_Visitor_Complex_4.jpg/500px-Atlantis_Kennedy_Space_Center_Visitor_Complex_4.jpg","https://commons.wikimedia.org/wiki/File:Atlantis_Kennedy_Space_Center_Visitor_Complex_4.jpg"],
  "KSC-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/da/Kennedy_Space_Center,_Rocket_Garden,_Power_of_Apollo.jpg/500px-Kennedy_Space_Center,_Rocket_Garden,_Power_of_Apollo.jpg","https://commons.wikimedia.org/wiki/File:Kennedy_Space_Center,_Rocket_Garden,_Power_of_Apollo.jpg"],
  "KSC-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/U.S._Astronaut_Hall_of_Fame_(KSC-2014-1768).jpg/500px-U.S._Astronaut_Hall_of_Fame_(KSC-2014-1768).jpg","https://commons.wikimedia.org/wiki/File:U.S._Astronaut_Hall_of_Fame_(KSC-2014-1768).jpg"],
  "KSC-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/9e/Space_Shuttle_Atlantis_exhibit,_Kennedy_Space_Center,_Cape_Canaveral,_FL,_KW_(15545743177).jpg/500px-Space_Shuttle_Atlantis_exhibit,_Kennedy_Space_Center,_Cape_Canaveral,_FL,_KW_(15545743177).jpg","https://commons.wikimedia.org/wiki/File:Space_Shuttle_Atlantis_exhibit,_Kennedy_Space_Center,_Cape_Canaveral,_FL,_KW_(15545743177).jpg"],
  "KSC-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/cb/Shuttle_launch_experience_artists_rendition_.jpg/500px-Shuttle_launch_experience_artists_rendition_.jpg","https://commons.wikimedia.org/wiki/File:Shuttle_launch_experience_artists_rendition_.jpg"],
  "KSC-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/91/Entrance_Kennedy_Space_Center_Visitor_Complex.jpg/500px-Entrance_Kennedy_Space_Center_Visitor_Complex.jpg","https://commons.wikimedia.org/wiki/File:Entrance_Kennedy_Space_Center_Visitor_Complex.jpg"],
  "KSC-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/91/Entrance_Kennedy_Space_Center_Visitor_Complex.jpg/500px-Entrance_Kennedy_Space_Center_Visitor_Complex.jpg","https://commons.wikimedia.org/wiki/File:Entrance_Kennedy_Space_Center_Visitor_Complex.jpg"],
  "KSC-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/Apollo_11_Launch2.jpg/500px-Apollo_11_Launch2.jpg","https://commons.wikimedia.org/wiki/File:Apollo_11_Launch2.jpg"],
  "KSC-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b8/SES-10_Launch_-_world's_first_reflight_of_an_orbital_class_rocket_(33361035200).jpg/500px-SES-10_Launch_-_world's_first_reflight_of_an_orbital_class_rocket_(33361035200).jpg","https://commons.wikimedia.org/wiki/File:SES-10_Launch_-_world's_first_reflight_of_an_orbital_class_rocket_(33361035200).jpg"],
  "KSC-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c8/Astronaut,_cosmonaut_space_suit_model_in_Kennedy_Space_Center_Visitor_Complex.jpg/500px-Astronaut,_cosmonaut_space_suit_model_in_Kennedy_Space_Center_Visitor_Complex.jpg","https://commons.wikimedia.org/wiki/File:Astronaut,_cosmonaut_space_suit_model_in_Kennedy_Space_Center_Visitor_Complex.jpg"],
  "LEGO": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d9/Miniland,_Legoland,_Florida._(6497802865).jpg/500px-Miniland,_Legoland,_Florida._(6497802865).jpg","https://commons.wikimedia.org/wiki/File:Miniland,_Legoland,_Florida._(6497802865).jpg"],
  "LEGO-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/5a/Winter_Haven_-_Legoland_Florida_-_Miniland_USA_-_Tampa_(9421378317).jpg/500px-Winter_Haven_-_Legoland_Florida_-_Miniland_USA_-_Tampa_(9421378317).jpg","https://commons.wikimedia.org/wiki/File:Winter_Haven_-_Legoland_Florida_-_Miniland_USA_-_Tampa_(9421378317).jpg"],
  "LEGO-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/22/Legoland_Florida_2.jpg/500px-Legoland_Florida_2.jpg","https://commons.wikimedia.org/wiki/File:Legoland_Florida_2.jpg"],
  "LEGO-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/f6/Legoland_FL_(8219689220).jpg/500px-Legoland_FL_(8219689220).jpg","https://commons.wikimedia.org/wiki/File:Legoland_FL_(8219689220).jpg"],
  "LEGO-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/22/Legoland_Florida_2.jpg/500px-Legoland_Florida_2.jpg","https://commons.wikimedia.org/wiki/File:Legoland_Florida_2.jpg"],
  "LEGO-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/22/Legoland_Florida_2.jpg/500px-Legoland_Florida_2.jpg","https://commons.wikimedia.org/wiki/File:Legoland_Florida_2.jpg"],
  "MIA": ["https://upload.wikimedia.org/wikipedia/commons/thumb/e/ed/Miami_Beach_-_Ocean_Drive_(2023).jpg/500px-Miami_Beach_-_Ocean_Drive_(2023).jpg","https://commons.wikimedia.org/wiki/File:Miami_Beach_-_Ocean_Drive_(2023).jpg"],
  "MIA-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/93/Everglades-airboat-tours_960.png/500px-Everglades-airboat-tours_960.png","https://commons.wikimedia.org/wiki/File:Everglades-airboat-tours_960.png"],
  "MIA-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2d/Boulevard_Hotel_(Neon_sign),_Miami_Beach.jpg/500px-Boulevard_Hotel_(Neon_sign),_Miami_Beach.jpg","https://commons.wikimedia.org/wiki/File:Boulevard_Hotel_(Neon_sign),_Miami_Beach.jpg"],
  "MIA-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2d/Boulevard_Hotel_(Neon_sign),_Miami_Beach.jpg/500px-Boulevard_Hotel_(Neon_sign),_Miami_Beach.jpg","https://commons.wikimedia.org/wiki/File:Boulevard_Hotel_(Neon_sign),_Miami_Beach.jpg"],
  "MIA-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/57/Wynwood_Walls_(10106520595).jpg/500px-Wynwood_Walls_(10106520595).jpg","https://commons.wikimedia.org/wiki/File:Wynwood_Walls_(10106520595).jpg"],
  "MIA-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/16/Calle-ocho-festival-2001.jpg/500px-Calle-ocho-festival-2001.jpg","https://commons.wikimedia.org/wiki/File:Calle-ocho-festival-2001.jpg"],
  "MIA-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/49/Miami-Bayside.jpg/500px-Miami-Bayside.jpg","https://commons.wikimedia.org/wiki/File:Miami-Bayside.jpg"],
  "MIA-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/e/e4/Dolphin_Mall_-_Miami_(4279541881).jpg/500px-Dolphin_Mall_-_Miami_(4279541881).jpg","https://commons.wikimedia.org/wiki/File:Dolphin_Mall_-_Miami_(4279541881).jpg"],
  "MIA-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d8/Aventura_Mall_exterior_2020.jpg/500px-Aventura_Mall_exterior_2020.jpg","https://commons.wikimedia.org/wiki/File:Aventura_Mall_exterior_2020.jpg"],
  "MK": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/84/Cinderella_Castle.jpg/500px-Cinderella_Castle.jpg","https://commons.wikimedia.org/wiki/File:Cinderella_Castle.jpg"],
  "MK-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/01/Walt_Disney_World_Railroad_Main_Street_station.jpg/500px-Walt_Disney_World_Railroad_Main_Street_station.jpg","https://commons.wikimedia.org/wiki/File:Walt_Disney_World_Railroad_Main_Street_station.jpg"],
  "MK-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/df/Fireworks_in_Magic_Kingdom.jpg/500px-Fireworks_in_Magic_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Fireworks_in_Magic_Kingdom.jpg"],
  "MK-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d5/Pirates_of_the_Caribbean_Magic_Kingdom.JPG/500px-Pirates_of_the_Caribbean_Magic_Kingdom.JPG","https://commons.wikimedia.org/wiki/File:Pirates_of_the_Caribbean_Magic_Kingdom.JPG"],
  "MK-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2e/Jungle_Cruise_at_Magic_Kingdom.jpg/500px-Jungle_Cruise_at_Magic_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Jungle_Cruise_at_Magic_Kingdom.jpg"],
  "MK-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/9d/Camel_statue_spraying_water_on_the_Magic_Carpets_of_Aladdin_-_Magic_Kingdom_-_Flickr_-_JeffChristiansen.jpg/500px-Camel_statue_spraying_water_on_the_Magic_Carpets_of_Aladdin_-_Magic_Kingdom_-_Flickr_-_JeffChristiansen.jpg","https://commons.wikimedia.org/wiki/File:Camel_statue_spraying_water_on_the_Magic_Carpets_of_Aladdin_-_Magic_Kingdom_-_Flickr_-_JeffChristiansen.jpg"],
  "MK-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/de/Walt_Disney's_Enchanted_Tiki_Room_at_Magic_Kingdom.jpg/500px-Walt_Disney's_Enchanted_Tiki_Room_at_Magic_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Walt_Disney's_Enchanted_Tiki_Room_at_Magic_Kingdom.jpg"],
  "MK-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1e/Big_Thunder_Mountain_Railroad_2.jpg/500px-Big_Thunder_Mountain_Railroad_2.jpg","https://commons.wikimedia.org/wiki/File:Big_Thunder_Mountain_Railroad_2.jpg"],
  "MK-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/db/Tiana's_Bayou_Adventure_(Magic_Kingdom)_1.jpg/500px-Tiana's_Bayou_Adventure_(Magic_Kingdom)_1.jpg","https://commons.wikimedia.org/wiki/File:Tiana's_Bayou_Adventure_(Magic_Kingdom)_1.jpg"],
  "MK-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/db/Country_Bear_Jamboree_(29420435133).jpg/500px-Country_Bear_Jamboree_(29420435133).jpg","https://commons.wikimedia.org/wiki/File:Country_Bear_Jamboree_(29420435133).jpg"],
  "MK-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/9a/The_Haunted_Mansion.jpg/500px-The_Haunted_Mansion.jpg","https://commons.wikimedia.org/wiki/File:The_Haunted_Mansion.jpg"],
  "MK-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2d/Hall_of_Presidents_on_Election_Day_(30230222304).jpg/500px-Hall_of_Presidents_on_Election_Day_(30230222304).jpg","https://commons.wikimedia.org/wiki/File:Hall_of_Presidents_on_Election_Day_(30230222304).jpg"],
  "MK-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/f0/Seven_Dwarfs_Mine_Train.jpg/500px-Seven_Dwarfs_Mine_Train.jpg","https://commons.wikimedia.org/wiki/File:Seven_Dwarfs_Mine_Train.jpg"],
  "MK-13": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/31/Peter_Pan's_Flight_at_Magic_Kingdom.jpg/500px-Peter_Pan's_Flight_at_Magic_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Peter_Pan's_Flight_at_Magic_Kingdom.jpg"],
  "MK-14": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/36/It's_A_Small_World_(Magic_Kingdom)_1998_1.jpg/500px-It's_A_Small_World_(Magic_Kingdom)_1998_1.jpg","https://commons.wikimedia.org/wiki/File:It's_A_Small_World_(Magic_Kingdom)_1998_1.jpg"],
  "MK-15": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/80/Under_the_Sea_Journey_of_the_Little_Mermaid_-_Magic_Kingdom.jpg/500px-Under_the_Sea_Journey_of_the_Little_Mermaid_-_Magic_Kingdom.jpg","https://commons.wikimedia.org/wiki/File:Under_the_Sea_Journey_of_the_Little_Mermaid_-_Magic_Kingdom.jpg"],
  "MK-16": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2a/Mickey's_PhilharMagic_TDL.JPG/500px-Mickey's_PhilharMagic_TDL.JPG","https://commons.wikimedia.org/wiki/File:Mickey's_PhilharMagic_TDL.JPG"],
  "MK-17": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a3/Dumbo_the_Flying_Elephant_Close-up.jpg/500px-Dumbo_the_Flying_Elephant_Close-up.jpg","https://commons.wikimedia.org/wiki/File:Dumbo_the_Flying_Elephant_Close-up.jpg"],
  "MK-18": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/9e/Barnstormer_at_Magic_Kingdom_-_turn.jpg/500px-Barnstormer_at_Magic_Kingdom_-_turn.jpg","https://commons.wikimedia.org/wiki/File:Barnstormer_at_Magic_Kingdom_-_turn.jpg"],
  "MK-19": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/9d/Mad_Tea_Part_WDW.jpg/500px-Mad_Tea_Part_WDW.jpg","https://commons.wikimedia.org/wiki/File:Mad_Tea_Part_WDW.jpg"],
  "MK-20": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d5/Many_Adventures_of_Winnie_the_Pooh_MK.jpg/500px-Many_Adventures_of_Winnie_the_Pooh_MK.jpg","https://commons.wikimedia.org/wiki/File:Many_Adventures_of_Winnie_the_Pooh_MK.jpg"],
  "MK-21": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/64/Enchanted_Tales_with_Belle.jpg/500px-Enchanted_Tales_with_Belle.jpg","https://commons.wikimedia.org/wiki/File:Enchanted_Tales_with_Belle.jpg"],
  "MK-22": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d6/Tron_Lightcycle_Run.jpg/500px-Tron_Lightcycle_Run.jpg","https://commons.wikimedia.org/wiki/File:Tron_Lightcycle_Run.jpg"],
  "MK-23": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/6b/Space_Mountain_in_the_Magic_Kingdom_in_2021.png/500px-Space_Mountain_in_the_Magic_Kingdom_in_2021.png","https://commons.wikimedia.org/wiki/File:Space_Mountain_in_the_Magic_Kingdom_in_2021.png"],
  "MK-24": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b7/Buzz_Lightyear's_Space_Ranger_Spin_Entrance_(4730826024).jpg/500px-Buzz_Lightyear's_Space_Ranger_Spin_Entrance_(4730826024).jpg","https://commons.wikimedia.org/wiki/File:Buzz_Lightyear's_Space_Ranger_Spin_Entrance_(4730826024).jpg"],
  "MK-25": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/fb/Tomorrowland_Speedway_(42591060204).jpg/500px-Tomorrowland_Speedway_(42591060204).jpg","https://commons.wikimedia.org/wiki/File:Tomorrowland_Speedway_(42591060204).jpg"],
  "MK-26": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/8b/Tomorrowland_sign_with_PeopleMover_track_and_Astro_Orbiter,_Magic_Kingdom,_June_2006_-_ThrillZing.jpg/500px-Tomorrowland_sign_with_PeopleMover_track_and_Astro_Orbiter,_Magic_Kingdom,_June_2006_-_ThrillZing.jpg","https://commons.wikimedia.org/wiki/File:Tomorrowland_sign_with_PeopleMover_track_and_Astro_Orbiter,_Magic_Kingdom,_June_2006_-_ThrillZing.jpg"],
  "MK-27": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/07/Wdw-peoplemover.jpg/500px-Wdw-peoplemover.jpg","https://commons.wikimedia.org/wiki/File:Wdw-peoplemover.jpg"],
  "MK-28": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/11/MonstersIncLaughFloorSign.jpg/500px-MonstersIncLaughFloorSign.jpg","https://commons.wikimedia.org/wiki/File:MonstersIncLaughFloorSign.jpg"],
  "MK-29": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/18/Carousel_of_Progress_1920.jpg/500px-Carousel_of_Progress_1920.jpg","https://commons.wikimedia.org/wiki/File:Carousel_of_Progress_1920.jpg"],
  "ORLX": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/28/Disney_Springs_Ariel_View.jpg/500px-Disney_Springs_Ariel_View.jpg","https://commons.wikimedia.org/wiki/File:Disney_Springs_Ariel_View.jpg"],
  "ORLX-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/28/Disney_Springs_Ariel_View.jpg/500px-Disney_Springs_Ariel_View.jpg","https://commons.wikimedia.org/wiki/File:Disney_Springs_Ariel_View.jpg"],
  "ORLX-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/de/The_Wheel_at_ICON_Park_(yellow).jpg/500px-The_Wheel_at_ICON_Park_(yellow).jpg","https://commons.wikimedia.org/wiki/File:The_Wheel_at_ICON_Park_(yellow).jpg"],
  "ORLX-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1d/Orlando_Premium_Outlets_08.JPG/500px-Orlando_Premium_Outlets_08.JPG","https://commons.wikimedia.org/wiki/File:Orlando_Premium_Outlets_08.JPG"],
  "ORLX-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/ab/Orlando_Premium_Outlets_-_Vineland_Ave.jpeg/500px-Orlando_Premium_Outlets_-_Vineland_Ave.jpeg","https://commons.wikimedia.org/wiki/File:Orlando_Premium_Outlets_-_Vineland_Ave.jpeg"],
  "ORLX-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/93/Gatorland_entrance_-Florida-23Feb2006.jpg/500px-Gatorland_entrance_-Florida-23Feb2006.jpg","https://commons.wikimedia.org/wiki/File:Gatorland_entrance_-Florida-23Feb2006.jpg"],
  "SWO": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7c/Mako_(SeaWorld_Orlando)_1.jpg/500px-Mako_(SeaWorld_Orlando)_1.jpg","https://commons.wikimedia.org/wiki/File:Mako_(SeaWorld_Orlando)_1.jpg"],
  "SWO-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7c/Mako_(SeaWorld_Orlando)_1.jpg/500px-Mako_(SeaWorld_Orlando)_1.jpg","https://commons.wikimedia.org/wiki/File:Mako_(SeaWorld_Orlando)_1.jpg"],
  "SWO-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/55/Seaworld-Orlando-Kraken-1629.jpg/500px-Seaworld-Orlando-Kraken-1629.jpg","https://commons.wikimedia.org/wiki/File:Seaworld-Orlando-Kraken-1629.jpg"],
  "SWO-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bb/Infinity_Falls_layout_drop.jpg/500px-Infinity_Falls_layout_drop.jpg","https://commons.wikimedia.org/wiki/File:Infinity_Falls_layout_drop.jpg"],
  "SWO-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a6/Manta_at_SeaWorld_Orlando_65.jpg/500px-Manta_at_SeaWorld_Orlando_65.jpg","https://commons.wikimedia.org/wiki/File:Manta_at_SeaWorld_Orlando_65.jpg"],
  "SWO-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d8/Pipeline_The_Surf_Coaster's_over-banked_turn.jpg/500px-Pipeline_The_Surf_Coaster's_over-banked_turn.jpg","https://commons.wikimedia.org/wiki/File:Pipeline_The_Surf_Coaster's_over-banked_turn.jpg"],
  "SWO-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/40/Journey_to_Atlantis_-_SeaWorld_Orlando.jpg/500px-Journey_to_Atlantis_-_SeaWorld_Orlando.jpg","https://commons.wikimedia.org/wiki/File:Journey_to_Atlantis_-_SeaWorld_Orlando.jpg"],
  "SWO-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/21/Ice_Breaker_2020.jpg/500px-Ice_Breaker_2020.jpg","https://commons.wikimedia.org/wiki/File:Ice_Breaker_2020.jpg"],
  "SWO-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/9/94/Penguin_Trek_and_the_Manta_March_2026.jpg/500px-Penguin_Trek_and_the_Manta_March_2026.jpg","https://commons.wikimedia.org/wiki/File:Penguin_Trek_and_the_Manta_March_2026.jpg"],
  "SWO-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/6e/Antarctica_Empire_of_the_Penguin_entrance_1.jpg/500px-Antarctica_Empire_of_the_Penguin_entrance_1.jpg","https://commons.wikimedia.org/wiki/File:Antarctica_Empire_of_the_Penguin_entrance_1.jpg"],
  "SWO-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0a/Sunny_Day_Carousel_Mar_2026.jpg/500px-Sunny_Day_Carousel_Mar_2026.jpg","https://commons.wikimedia.org/wiki/File:Sunny_Day_Carousel_Mar_2026.jpg"],
  "SWO-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/Shamu_Stadium_SeaWorld_Orlando_Florida.jpg/500px-Shamu_Stadium_SeaWorld_Orlando_Florida.jpg","https://commons.wikimedia.org/wiki/File:Shamu_Stadium_SeaWorld_Orlando_Florida.jpg"],
  "SWO-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/76/SeaWorld_Orlando_Dolphins.JPG/500px-SeaWorld_Orlando_Dolphins.JPG","https://commons.wikimedia.org/wiki/File:SeaWorld_Orlando_Dolphins.JPG"],
  "TL": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/35/Walt_Disney_World_Typhoon_Lagoon_Entrance.JPG/500px-Walt_Disney_World_Typhoon_Lagoon_Entrance.JPG","https://commons.wikimedia.org/wiki/File:Walt_Disney_World_Typhoon_Lagoon_Entrance.JPG"],
  "TL-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/5a/Typhoon_Lagoon_Surf-Wave_Pool_(34261997536).jpg/500px-Typhoon_Lagoon_Surf-Wave_Pool_(34261997536).jpg","https://commons.wikimedia.org/wiki/File:Typhoon_Lagoon_Surf-Wave_Pool_(34261997536).jpg"],
  "TL-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1f/WDW_Crush_'n'_Gusher.JPG/500px-WDW_Crush_'n'_Gusher.JPG","https://commons.wikimedia.org/wiki/File:WDW_Crush_'n'_Gusher.JPG"],
  "TL-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a6/WDW_Humunga_Kowabunga.JPG/500px-WDW_Humunga_Kowabunga.JPG","https://commons.wikimedia.org/wiki/File:WDW_Humunga_Kowabunga.JPG"],
  "TL-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/2c/Miss_Adventure_Falls_(34144803442).jpg/500px-Miss_Adventure_Falls_(34144803442).jpg","https://commons.wikimedia.org/wiki/File:Miss_Adventure_Falls_(34144803442).jpg"],
  "TPA": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bc/Tampa_architectural_photos_256.jpg/500px-Tampa_architectural_photos_256.jpg","https://commons.wikimedia.org/wiki/File:Tampa_architectural_photos_256.jpg"],
  "TPA-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0c/Publix_Supermarket,_Miami_FL.jpg/500px-Publix_Supermarket,_Miami_FL.jpg","https://commons.wikimedia.org/wiki/File:Publix_Supermarket,_Miami_FL.jpg"],
  "TPA-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bc/Tampa_architectural_photos_256.jpg/500px-Tampa_architectural_photos_256.jpg","https://commons.wikimedia.org/wiki/File:Tampa_architectural_photos_256.jpg"],
  "TPA-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0a/Florida_aquarium_exterior.jpg/500px-Florida_aquarium_exterior.jpg","https://commons.wikimedia.org/wiki/File:Florida_aquarium_exterior.jpg"],
  "TPA-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/f7/7th_Avenue_in_Ybor_City_-_Eric_Statzer.jpg/500px-7th_Avenue_in_Ybor_City_-_Eric_Statzer.jpg","https://commons.wikimedia.org/wiki/File:7th_Avenue_in_Ybor_City_-_Eric_Statzer.jpg"],
  "TPA-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/ba/Tampa_Florida_November_2013-3g.jpg/500px-Tampa_Florida_November_2013-3g.jpg","https://commons.wikimedia.org/wiki/File:Tampa_Florida_November_2013-3g.jpg"],
  "USF": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Entrance_of_Universal_Studios_Florida_Minions_film.JPG/500px-Entrance_of_Universal_Studios_Florida_Minions_film.JPG","https://commons.wikimedia.org/wiki/File:Entrance_of_Universal_Studios_Florida_Minions_film.JPG"],
  "USF-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/74/Hollywood_Rip_Ride_Rockit_(May_2023).jpg/500px-Hollywood_Rip_Ride_Rockit_(May_2023).jpg","https://commons.wikimedia.org/wiki/File:Hollywood_Rip_Ride_Rockit_(May_2023).jpg"],
  "USF-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/4d/Transformers_The_Ride_-_Universal_Studios_Florida.JPG/500px-Transformers_The_Ride_-_Universal_Studios_Florida.JPG","https://commons.wikimedia.org/wiki/File:Transformers_The_Ride_-_Universal_Studios_Florida.JPG"],
  "USF-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/dd/Despicable_Me-_Minion_Mayhem_(Orlando)_May_2023.jpg/500px-Despicable_Me-_Minion_Mayhem_(Orlando)_May_2023.jpg","https://commons.wikimedia.org/wiki/File:Despicable_Me-_Minion_Mayhem_(Orlando)_May_2023.jpg"],
  "USF-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/8/83/Villain-Con_Minion_Blast_2.jpg/500px-Villain-Con_Minion_Blast_2.jpg","https://commons.wikimedia.org/wiki/File:Villain-Con_Minion_Blast_2.jpg"],
  "USF-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/39/Revenge_of_the_Mummy_(Universal_Studios_Florida)_entrance.jpg/500px-Revenge_of_the_Mummy_(Universal_Studios_Florida)_entrance.jpg","https://commons.wikimedia.org/wiki/File:Revenge_of_the_Mummy_(Universal_Studios_Florida)_entrance.jpg"],
  "USF-06": ["https://upload.wikimedia.org/wikipedia/commons/thumb/5/5d/Race_Through_New_York_Starring_Jimmy_Fallon_(36916486110).jpg/500px-Race_Through_New_York_Starring_Jimmy_Fallon_(36916486110).jpg","https://commons.wikimedia.org/wiki/File:Race_Through_New_York_Starring_Jimmy_Fallon_(36916486110).jpg"],
  "USF-07": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/4d/Fast_%26_Furious_Supercharged_(Universal_Studios_Florida)_1.jpg/500px-Fast_%26_Furious_Supercharged_(Universal_Studios_Florida)_1.jpg","https://commons.wikimedia.org/wiki/File:Fast_%26_Furious_Supercharged_(Universal_Studios_Florida)_1.jpg"],
  "USF-08": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c2/Harry_Potter_and_the_Escape_from_Gringotts.JPG/500px-Harry_Potter_and_the_Escape_from_Gringotts.JPG","https://commons.wikimedia.org/wiki/File:Harry_Potter_and_the_Escape_from_Gringotts.JPG"],
  "USF-09": ["https://upload.wikimedia.org/wikipedia/commons/thumb/0/0e/Hogwarts_Express_Backlot.jpg/500px-Hogwarts_Express_Backlot.jpg","https://commons.wikimedia.org/wiki/File:Hogwarts_Express_Backlot.jpg"],
  "USF-10": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bb/Men_In_Black_Alien_Attack_at_Universal_Studios_Florida.JPG/500px-Men_In_Black_Alien_Attack_at_Universal_Studios_Florida.JPG","https://commons.wikimedia.org/wiki/File:Men_In_Black_Alien_Attack_at_Universal_Studios_Florida.JPG"],
  "USF-11": ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bc/The_Simpsons_Ride_at_Universal_Studios_Florida_(cropped).jpg/500px-The_Simpsons_Ride_at_Universal_Studios_Florida_(cropped).jpg","https://commons.wikimedia.org/wiki/File:The_Simpsons_Ride_at_Universal_Studios_Florida_(cropped).jpg"],
  "USF-12": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/4e/Kang_%26_Kodos_Twirl_n_Hurl_Springfield_U_S_A.JPG/500px-Kang_%26_Kodos_Twirl_n_Hurl_Springfield_U_S_A.JPG","https://commons.wikimedia.org/wiki/File:Kang_%26_Kodos_Twirl_n_Hurl_Springfield_U_S_A.JPG"],
  "USF-13": ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d9/Universal-Studios-Woody-Woodpecker-Nuthouse-Coaster-9277.jpg/500px-Universal-Studios-Woody-Woodpecker-Nuthouse-Coaster-9277.jpg","https://commons.wikimedia.org/wiki/File:Universal-Studios-Woody-Woodpecker-Nuthouse-Coaster-9277.jpg"],
  "USF-14": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/f6/E.T._Adventure_02.jpg/500px-E.T._Adventure_02.jpg","https://commons.wikimedia.org/wiki/File:E.T._Adventure_02.jpg"],
  "USF-15": ["https://upload.wikimedia.org/wikipedia/commons/thumb/f/ff/The_Bourne_Stuntacular_1.jpg/500px-The_Bourne_Stuntacular_1.jpg","https://commons.wikimedia.org/wiki/File:The_Bourne_Stuntacular_1.jpg"],
  "VB": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/69/Krakatau_at_Volcano_Bay.jpg/500px-Krakatau_at_Volcano_Bay.jpg","https://commons.wikimedia.org/wiki/File:Krakatau_at_Volcano_Bay.jpg"],
  "VB-01": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/47/Krakatau_Aqua_Coaster_(37170103801).jpg/500px-Krakatau_Aqua_Coaster_(37170103801).jpg","https://commons.wikimedia.org/wiki/File:Krakatau_Aqua_Coaster_(37170103801).jpg"],
  "VB-02": ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c2/Ohyah_and_Ohno_water_slides.jpg/500px-Ohyah_and_Ohno_water_slides.jpg","https://commons.wikimedia.org/wiki/File:Ohyah_and_Ohno_water_slides.jpg"],
  "VB-03": ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/69/Krakatau_at_Volcano_Bay.jpg/500px-Krakatau_at_Volcano_Bay.jpg","https://commons.wikimedia.org/wiki/File:Krakatau_at_Volcano_Bay.jpg"],
  "VB-04": ["https://upload.wikimedia.org/wikipedia/commons/thumb/2/24/Kopiko_Wai_Winding_River_at_Volcano_Bay.jpg/500px-Kopiko_Wai_Winding_River_at_Volcano_Bay.jpg","https://commons.wikimedia.org/wiki/File:Kopiko_Wai_Winding_River_at_Volcano_Bay.jpg"],
  "VB-05": ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/49/Honu_and_Ika_Moana_(37170120711).jpg/500px-Honu_and_Ika_Moana_(37170120711).jpg","https://commons.wikimedia.org/wiki/File:Honu_and_Ika_Moana_(37170120711).jpg"]
};
