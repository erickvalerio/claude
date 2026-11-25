/**
 * ORGANIZAR "MEU DRIVE"  (I:\Meu Drive no computador)
 * ---------------------------------------------------
 * Google Apps Script (https://script.google.com) - rode logado como erickvalerio@gmail.com
 *
 * Move os itens soltos na raiz do Meu Drive para:
 *   01 - MASTER DATA / (Ordens de Serviço | Orçamentos e Propostas | Laudos e Relatórios |
 *                       Projetos | Financeiro | Clientes | IA)
 *   02 - IMPRESSÃO 3D
 *   03 - PESSOAL
 *   04 - ARQUIVO ANTIGO   (pastas antigas e arquivos sem alteração desde antes de 2024)
 *   99 - A CLASSIFICAR    (o que não se encaixar em nenhuma regra)
 *
 * Ficam na raiz: o atalho "MASTER DATA AUTOMACAO", "Google Meet" e "Salvo do Chrome"
 * (o Meet e o Chrome gravam nessas pastas sozinhos).
 *
 * Segurança: NUNCA apaga nada (nem duplicados). Começa em SIMULAÇÃO (DRY_RUN = true),
 * gera uma planilha de log, e guarda de onde cada item saiu para "desfazerMeuDrive".
 */

const DRY_RUN_MEUDRIVE = true; // true = só simula | false = executa de verdade

const MANTER_NA_RAIZ = ['MASTER DATA AUTOMACAO', 'Google Meet', 'Salvo do Chrome',
  '01 - MASTER DATA', '02 - IMPRESSÃO 3D', '03 - PESSOAL', '04 - ARQUIVO ANTIGO', '99 - A CLASSIFICAR'];

const MD = '01 - MASTER DATA/';
// Ordem importa: a primeira regra que bater vence. Teste sem acentos e em minúsculas.
const REGRAS_MEUDRIVE = [
  ['99 - A CLASSIFICAR',            /^(planilha sem titulo|gostaria de colocar)/],
  ['03 - PESSOAL',                  /(casamento|viagem|apartamento|planta_sala|despesas_jun|lista de tarefas)/],
  ['02 - IMPRESSÃO 3D',             /(^3d$|filamento|precificacao_3d|klipper)/],
  ['04 - ARQUIVO ANTIGO',           /^(diversos|software|adaptive|airtable|9 - mach-tec)$/],
  [MD + 'IA',                       /(base_conhecimento|contexto_claude|ia compartilhada|^chatgpt$)/],
  [MD + 'Clientes',                 /(clientes|postos mg|produtos ativos|planilha_erick_brasil)/],
  [MD + 'Ordens de Serviço',        /(ordem de servico|ordens de servico|precificacao os|servicos mensal|chamados)/],
  [MD + 'Laudos e Relatórios',      /(laudo|^atg |isl - qqp|relatorio)/],
  [MD + 'Orçamentos e Propostas',   /(orcamento|proposta|cotac|pedido|valfluid|acrescimo|coloca 45|coloque 35|listagem_cabos|wayne)/],
  [MD + 'Financeiro',               /(financeiro|pix|boleto|acerto|consolidado - master|notas_fiscais|nota fiscal|danfe|espelho de remessa|faturamento|despesas)/],
  [MD + 'Projetos',                 /(cmaa|mapa_|masterducha|d8_|projeto|levantamento|modelo-instalacao|piso_concreto|viabilidade|mao_de_obra|arqueamento|bases_santa_lucia)/],
];
const ANTIGO_ANTES_DE = new Date(2024, 0, 1);
const CHAVE_HIST_MEUDRIVE = 'historicoMeuDrive';

function organizarMeuDrive() {
  const raiz = DriveApp.getRootFolder();
  const log = criarLogMeuDrive_();
  const props = PropertiesService.getScriptProperties();
  const hist = JSON.parse(props.getProperty(CHAVE_HIST_MEUDRIVE) || '{}');
  const cache = {};
  let n = 0;
  log('INÍCIO', '', DRY_RUN_MEUDRIVE ? 'MODO SIMULAÇÃO - nada será alterado' : 'MODO REAL');

  const itens = [];
  const pf = raiz.getFolders(); while (pf.hasNext()) itens.push(pf.next());
  const ar = raiz.getFiles();   while (ar.hasNext()) itens.push(ar.next());

  itens.forEach(function (item) {
    const nome = item.getName();
    if (MANTER_NA_RAIZ.indexOf(nome) !== -1) return;
    const destino = classificarMeuDrive_(nome, item.getLastUpdated());
    if (!DRY_RUN_MEUDRIVE) {
      hist[item.getId()] = raiz.getId();
      item.moveTo(pastaPorCaminho_(raiz, destino, cache));
      props.setProperty(CHAVE_HIST_MEUDRIVE, JSON.stringify(hist));
    }
    n++;
    log('MOVER', nome, '-> ' + destino);
  });
  log('FIM', '', n + ' itens ' + (DRY_RUN_MEUDRIVE ? 'seriam movidos' : 'movidos'));
}

/** Devolve tudo para a raiz do Meu Drive. */
function desfazerMeuDrive() {
  const props = PropertiesService.getScriptProperties();
  const hist = JSON.parse(props.getProperty(CHAVE_HIST_MEUDRIVE) || '{}');
  Object.keys(hist).forEach(function (id) {
    let item;
    try { item = DriveApp.getFileById(id); } catch (e) { item = DriveApp.getFolderById(id); }
    item.moveTo(DriveApp.getFolderById(hist[id]));
    delete hist[id];
    props.setProperty(CHAVE_HIST_MEUDRIVE, JSON.stringify(hist));
  });
  Logger.log('Desfeito.');
}

function classificarMeuDrive_(nome, data) {
  const n = nome.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  for (let i = 0; i < REGRAS_MEUDRIVE.length; i++) {
    if (REGRAS_MEUDRIVE[i][1].test(n)) return REGRAS_MEUDRIVE[i][0];
  }
  if (data && data < ANTIGO_ANTES_DE) return '04 - ARQUIVO ANTIGO';
  return '99 - A CLASSIFICAR';
}

function pastaPorCaminho_(raiz, caminho, cache) {
  if (cache[caminho]) return cache[caminho];
  let atual = raiz;
  caminho.split('/').forEach(function (parte) {
    const it = atual.getFoldersByName(parte);
    atual = it.hasNext() ? it.next() : atual.createFolder(parte);
  });
  cache[caminho] = atual;
  return atual;
}

function criarLogMeuDrive_() {
  const ss = SpreadsheetApp.create('LOG - Organizar Meu Drive ' +
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm'));
  const sh = ss.getActiveSheet();
  sh.appendRow(['Data/Hora', 'Ação', 'Item', 'Destino']);
  sh.setFrozenRows(1);
  Logger.log('Planilha de log: ' + ss.getUrl());
  return function (a, i, d) { sh.appendRow([new Date(), (DRY_RUN_MEUDRIVE ? '[SIMULAÇÃO] ' : '') + a, i, d]); };
}
