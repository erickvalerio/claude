/**
 * ORGANIZAR "03 - CLIENTES E POSTOS" POR ANO / MÊS / DIA
 * ------------------------------------------------------
 * Google Apps Script (https://script.google.com)
 *
 * Dentro de CADA pasta de cliente, os arquivos soltos vão para:
 *     <pasta do cliente>/2026/03 - Março/17/arquivo.jpeg
 *
 * Qual data é usada:
 *   1. A data no NOME do arquivo, se houver (ex.: "WhatsApp Image 2026-03-17 at 09.57.30.jpeg",
 *      "RELATORIO 17-03-2026.pdf", "OS 17.03.2026.xlsx").
 *   2. Senão, a data da última modificação do arquivo.
 *   (A data de criação NÃO é usada: todos foram enviados ao Drive em 16/04/2026.)
 *
 * Segurança:
 *   - NUNCA apaga nada. Só cria pastas de data e move arquivos.
 *   - Começa em SIMULAÇÃO (DRY_RUN = true): só gera uma planilha de log.
 *   - Pastas que já são de data (2026, "03 - Março", "17") são puladas, então pode rodar
 *     várias vezes sem bagunçar.
 *   - O Google limita cada execução a 6 minutos. Se aparecer "PARADO POR TEMPO" no log,
 *     é só executar de novo: ele continua de onde parou.
 *
 * Como usar:
 *   1. Em script.google.com (conta automacao@masterdatanet.com.br), cole este arquivo
 *      no mesmo projeto do OrganizarDrive.gs ou num projeto novo.
 *   2. Execute "organizarClientesPorData" com DRY_RUN = true e confira o log.
 *   3. Mude para DRY_RUN = false e execute de novo.
 */

const DRY_RUN_CLIENTES = true; // true = só simula | false = executa de verdade

// ID da pasta "03 - CLIENTES E POSTOS"
const CLIENTES_ID = '14BjKZWu_c911Ki8lqkK12blv8TQGmWDc';

// Pastas que não devem ser reorganizadas (nome exato)
const IGNORAR_PASTAS = ['98 - A classificar'];

const MESES = ['01 - Janeiro', '02 - Fevereiro', '03 - Março', '04 - Abril', '05 - Maio', '06 - Junho',
  '07 - Julho', '08 - Agosto', '09 - Setembro', '10 - Outubro', '11 - Novembro', '12 - Dezembro'];

const LIMITE_MS = 5 * 60 * 1000; // para com folga antes dos 6 min do Google

function organizarClientesPorData() {
  const inicio = Date.now();
  const log = criarLogClientes_();
  const stats = { movidos: 0, pastasCriadas: 0, pelaData: 0, peloNome: 0 };
  log('INÍCIO', '03 - CLIENTES E POSTOS', DRY_RUN_CLIENTES ? 'MODO SIMULAÇÃO' : 'MODO REAL');

  const terminou = processarPasta_(DriveApp.getFolderById(CLIENTES_ID), '03 - CLIENTES E POSTOS',
    true, log, stats, inicio);

  log(terminou ? 'FIM' : 'PARADO POR TEMPO', '',
    stats.movidos + ' arquivos (' + stats.peloNome + ' pela data do nome, ' + stats.pelaData +
    ' pela data de modificação), ' + stats.pastasCriadas + ' pastas criadas' +
    (terminou ? '' : ' - EXECUTE DE NOVO para continuar'));
}

/** Percorre a pasta e as subpastas. Devolve false se parou por tempo. */
function processarPasta_(pasta, caminho, ehRaiz, log, stats, inicio) {
  // Arquivos soltos nesta pasta (na raiz "03 - CLIENTES" não mexe: não é pasta de cliente)
  if (!ehRaiz) {
    const cache = {};
    const arquivos = pasta.getFiles();
    while (arquivos.hasNext()) {
      if (Date.now() - inicio > LIMITE_MS) return false;
      const f = arquivos.next();
      const info = dataDoArquivo_(f);
      const d = info.data;
      const ano = String(d.getFullYear());
      const mes = MESES[d.getMonth()];
      const dia = ('0' + d.getDate()).slice(-2);
      const destino = garantirCaminho_(pasta, [ano, mes, dia], cache, stats);
      if (!DRY_RUN_CLIENTES) f.moveTo(destino);
      stats.movidos++;
      info.peloNome ? stats.peloNome++ : stats.pelaData++;
      log('MOVER', caminho + ' / ' + f.getName(),
        '-> ' + ano + ' / ' + mes + ' / ' + dia + (info.peloNome ? ' (data do nome)' : ' (data de modificação)'));
    }
  }

  // Subpastas (pula as que já são de data e as ignoradas)
  const subpastas = pasta.getFolders();
  while (subpastas.hasNext()) {
    const sp = subpastas.next();
    const nome = sp.getName();
    if (ehPastaDeData_(nome) || IGNORAR_PASTAS.indexOf(nome) !== -1) continue;
    if (!processarPasta_(sp, caminho + ' / ' + nome, false, log, stats, inicio)) return false;
  }
  return true;
}

function ehPastaDeData_(nome) {
  return /^(19|20)\d{2}$/.test(nome) || MESES.indexOf(nome) !== -1 || /^(0[1-9]|[12]\d|3[01])$/.test(nome);
}

/** Data pelo nome do arquivo; se não houver, data de modificação. */
function dataDoArquivo_(f) {
  const n = f.getName();
  let m = n.match(/(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})/);           // 2026-03-17
  if (m) { const d = montarData_(+m[1], +m[2], +m[3]); if (d) return { data: d, peloNome: true }; }
  m = n.match(/(?:^|\D)(\d{1,2})[-_.\/](\d{1,2})[-_.\/](20\d{2})(?!\d)/); // 17-03-2026
  if (m) { const d = montarData_(+m[3], +m[2], +m[1]); if (d) return { data: d, peloNome: true }; }
  return { data: f.getLastUpdated(), peloNome: false };
}

function montarData_(a, me, di) {
  if (me < 1 || me > 12 || di < 1 || di > 31) return null;
  const d = new Date(a, me - 1, di);
  return d.getMonth() === me - 1 ? d : null;
}

/** Cria (se preciso) pai/ano/mês/dia e devolve a pasta do dia. */
function garantirCaminho_(pai, partes, cache, stats) {
  let atual = pai;
  let chave = '';
  for (let i = 0; i < partes.length; i++) {
    chave += '/' + partes[i];
    if (cache[chave]) { atual = cache[chave]; continue; }
    let prox = null;
    if (atual && !atual.__simulada) {
      const it = atual.getFoldersByName(partes[i]);
      if (it.hasNext()) prox = it.next();
    }
    if (!prox) {
      stats.pastasCriadas++;
      prox = DRY_RUN_CLIENTES ? { __simulada: true } : atual.createFolder(partes[i]);
    }
    cache[chave] = prox;
    atual = prox;
  }
  return atual;
}

function criarLogClientes_() {
  const ss = SpreadsheetApp.create('LOG - Clientes por data ' +
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm'));
  const sh = ss.getActiveSheet();
  sh.appendRow(['Data/Hora', 'Ação', 'Arquivo', 'Destino']);
  sh.setFrozenRows(1);
  Logger.log('Planilha de log: ' + ss.getUrl());
  const linhas = [];
  const fn = function (acao, item, detalhe) {
    linhas.push([new Date(), (DRY_RUN_CLIENTES ? '[SIMULAÇÃO] ' : '') + acao, item, detalhe]);
    if (linhas.length >= 200 || acao === 'FIM' || acao === 'PARADO POR TEMPO') {
      sh.getRange(sh.getLastRow() + 1, 1, linhas.length, 4).setValues(linhas);
      linhas.length = 0;
    }
  };
  return fn;
}
