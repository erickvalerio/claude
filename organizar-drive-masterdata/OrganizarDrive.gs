/**
 * ORGANIZADOR DO GOOGLE DRIVE - MASTER DATA AUTOMAÇÃO
 * ---------------------------------------------------
 * Google Apps Script (https://script.google.com)
 *
 * O que faz:
 *   1. Corrige a numeração das pastas principais (ex.: "02 - SOFTWARES" -> "08 - SOFTWARES").
 *   2. Cria as pastas padrão que estiverem faltando (inclusive "99 - A CLASSIFICAR").
 *   3. Move pastas avulsas para dentro da pasta certa (ex.: NOTAS FISCAIS -> 01 - ADMINISTRATIVO).
 *   4. Move os arquivos soltos na raiz por regras de nome/extensão.
 *   5. Registra tudo numa planilha de LOG (o que foi feito ou o que SERIA feito).
 *
 * Segurança:
 *   - NUNCA apaga nada. Só renomeia e move.
 *   - Começa em modo SIMULAÇÃO (DRY_RUN = true): só gera o log, não altera nada.
 *   - Mexe apenas no NÍVEL RAIZ da pasta principal (não reorganiza dentro das subpastas).
 *
 * Como usar:
 *   1. Entre com a conta dona da pasta (automacao@masterdatanet.com.br) em script.google.com
 *   2. Novo projeto > cole este código > salve.
 *   3. Execute "organizar" (com DRY_RUN = true) e autorize o acesso ao Drive.
 *   4. Abra a planilha "LOG - Organização Drive" (link no log de execução) e confira.
 *   5. Se estiver tudo certo, mude DRY_RUN para false e execute de novo.
 *   6. (Opcional) Execute "instalarGatilhoSemanal" para manter a raiz limpa toda segunda-feira.
 */

// ===================== CONFIGURAÇÃO =====================

const DRY_RUN = true; // true = só simula | false = executa de verdade

// ID da pasta "MASTER DATA AUTOMACAO"
const PASTA_RAIZ_ID = '1LiYLBM75PXYUXQfTux-_8VkiBDIEeyuU';

// Estrutura padrão (nível raiz)
const ESTRUTURA = [
  '00 - AUTOMAÇÃO',
  '01 - ADMINISTRATIVO E FINANCEIRO',
  '02 - PROPOSTAS E ORÇAMENTOS',
  '03 - CLIENTES E POSTOS',
  '04 - MATERIAIS ADAPTIVE',
  '05 - LAUDOS',
  '06 - VIAGENS PLANILHAS',
  '07 - SERVIÇOS EXECUTADOS',
  '08 - SOFTWARES',
  '09 - PROJETOS',
  '99 - A CLASSIFICAR',
];

// Pastas que mudam de nome (nome atual -> nome novo)
const RENOMEAR = {
  '02 - SOFTWARES': '08 - SOFTWARES',
};

// Pastas avulsas que vão para dentro de outra (nome atual -> [pasta destino, novo nome])
const MOVER_PASTAS = {
  'NOTAS FISCAIS MENSAL CLIENTE SISTEMA': ['01 - ADMINISTRATIVO E FINANCEIRO', 'NOTAS FISCAIS MENSAL CLIENTE SISTEMA'],
  'CHAMADOS WHATSAPP ATÉ 10/07':          ['07 - SERVIÇOS EXECUTADOS', 'CHAMADOS WHATSAPP (ATÉ 10-07)'],
};

// Regras para arquivos soltos. A primeira regra que bater vence.
// "nome" é testado no nome do arquivo (sem diferenciar maiúsculas/acentos).
const REGRAS_ARQUIVOS = [
  { destino: '08 - SOFTWARES',                 ext: ['exe', 'msi', 'apk', 'iso'] },
  { destino: '00 - AUTOMAÇÃO',                 ext: ['cmd', 'bat', 'ps1', 'py', 'gs', 'js', 'sh'] },
  { destino: '02 - PROPOSTAS E ORÇAMENTOS',    nome: /(orcamento|proposta|cotacao)/ },
  { destino: '05 - LAUDOS',                    nome: /(laudo|relatorio|teste de estanqueidade|estanqueidade)/ },
  { destino: '07 - SERVIÇOS EXECUTADOS',       nome: /(ordens? de servico|\bos\b|chamado|atendimento)/ },
  { destino: '06 - VIAGENS PLANILHAS',         nome: /(viagem|viagens|deslocamento|reembolso|km)/ },
  { destino: '01 - ADMINISTRATIVO E FINANCEIRO', nome: /(nota fiscal|\bnf\b|nfe|boleto|fatura|contrato|financeiro)/ },
  { destino: '09 - PROJETOS',                  nome: /(projeto|planta|drenagem|arqueamento|tanque)/ },
  { destino: '09 - PROJETOS',                  ext: ['dwg', 'dxf', 'skp', 'step', 'stp'] },
  { destino: '04 - MATERIAIS ADAPTIVE',        nome: /adaptive/ },
];
const DESTINO_PADRAO = '99 - A CLASSIFICAR';

// ===================== EXECUÇÃO =====================

function organizar() {
  const raiz = DriveApp.getFolderById(PASTA_RAIZ_ID);
  const log = criarLog_();
  log('INÍCIO', raiz.getName(), DRY_RUN ? 'MODO SIMULAÇÃO - nada será alterado' : 'MODO REAL');

  // 1) Renomear pastas com numeração errada
  Object.keys(RENOMEAR).forEach(function (atual) {
    const novo = RENOMEAR[atual];
    const p = pastaFilha_(raiz, atual);
    if (!p) return;
    if (pastaFilha_(raiz, novo)) {
      log('CONFLITO', atual, 'Já existe "' + novo + '". Junte manualmente.');
      return;
    }
    if (!DRY_RUN) p.setName(novo);
    log('RENOMEAR', atual, '-> ' + novo);
  });

  // 2) Garantir a estrutura padrão
  const pastas = {};
  ESTRUTURA.forEach(function (nome) {
    let p = pastaFilha_(raiz, nome);
    if (!p) {
      if (!DRY_RUN) p = raiz.createFolder(nome);
      log('CRIAR PASTA', nome, '');
    }
    pastas[nome] = p; // pode ser null em simulação
  });

  // 3) Mover pastas avulsas
  Object.keys(MOVER_PASTAS).forEach(function (atual) {
    const p = pastaFilha_(raiz, atual);
    if (!p) return;
    const destino = MOVER_PASTAS[atual][0];
    const novoNome = MOVER_PASTAS[atual][1];
    if (!DRY_RUN) {
      p.moveTo(pastas[destino]);
      if (novoNome !== atual) p.setName(novoNome);
    }
    log('MOVER PASTA', atual, '-> ' + destino + ' / ' + novoNome);
  });

  // 4) Pastas na raiz que não fazem parte da estrutura (só avisa)
  const nomesEsperados = ESTRUTURA.concat(Object.keys(RENOMEAR), Object.keys(MOVER_PASTAS));
  const itPastas = raiz.getFolders();
  while (itPastas.hasNext()) {
    const p = itPastas.next();
    if (nomesEsperados.indexOf(p.getName()) === -1) {
      log('AVISO', p.getName(), 'Pasta fora do padrão - revisar manualmente');
    }
  }

  // 5) Mover arquivos soltos na raiz
  const itArq = raiz.getFiles();
  while (itArq.hasNext()) {
    const f = itArq.next();
    const destino = classificar_(f.getName());
    if (!DRY_RUN) f.moveTo(pastas[destino]);
    log('MOVER ARQUIVO', f.getName(), '-> ' + destino);
  }

  log('FIM', '', DRY_RUN ? 'Confira o log e mude DRY_RUN para false.' : 'Organização concluída.');
}

/** Roda toda segunda às 7h, mantendo a raiz limpa (respeita o DRY_RUN). */
function instalarGatilhoSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'organizar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('organizar').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  Logger.log('Gatilho semanal instalado.');
}

// ===================== AUXILIARES =====================

function classificar_(nomeArquivo) {
  const nome = normalizar_(nomeArquivo);
  const ext = (nome.match(/\.([a-z0-9]+)$/) || [, ''])[1];
  for (let i = 0; i < REGRAS_ARQUIVOS.length; i++) {
    const r = REGRAS_ARQUIVOS[i];
    if (r.ext && r.ext.indexOf(ext) !== -1) return r.destino;
    if (r.nome && r.nome.test(nome)) return r.destino;
  }
  return DESTINO_PADRAO;
}

function normalizar_(s) {
  return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
}

function pastaFilha_(pai, nome) {
  const it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : null;
}

function criarLog_() {
  const ss = SpreadsheetApp.create('LOG - Organização Drive ' +
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm'));
  const sh = ss.getActiveSheet();
  sh.appendRow(['Data/Hora', 'Ação', 'Item', 'Detalhe']);
  sh.setFrozenRows(1);
  Logger.log('Planilha de log: ' + ss.getUrl());
  return function (acao, item, detalhe) {
    sh.appendRow([new Date(), (DRY_RUN ? '[SIMULAÇÃO] ' : '') + acao, item, detalhe]);
    Logger.log(acao + ' | ' + item + ' | ' + detalhe);
  };
}
