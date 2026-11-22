/**
 * ORGANIZAR "03 - CLIENTES E POSTOS" POR ANO / MÊS / DIA  (só PASTAS)
 * -------------------------------------------------------------------
 * Google Apps Script (https://script.google.com)
 *
 * Resultado:
 *     03 - CLIENTES E POSTOS / 2026 / 03 - Março / 17 / AUTO POSTO CARIBE
 *     03 - CLIENTES E POSTOS / 2026 / 02 - Fevereiro / 04 / Rede Décio
 *
 * - Move a PASTA do cliente inteira. NÃO mexe em nenhum arquivo dentro dela.
 * - "Pasta do cliente" = cada pasta dentro das categorias abaixo (CATEGORIAS).
 *   As redes (Rede Décio, Grupo GGX, Rede Lunasa...) vão inteiras, como um cliente só.
 * - Data usada = data do PRIMEIRO (mais antigo) arquivo dentro da pasta do cliente
 *   (procura também nas subpastas). De cada arquivo, vale a data escrita no nome
 *   (ex.: "WhatsApp Image 2026-03-17 ...", "SICOOB_2026_02_04...", "LAUDO 17-03-2026")
 *   e, se não houver, a data da última modificação.
 * - Pasta de cliente sem nenhum arquivo: não é movida (aparece no log como "SEM ARQUIVOS").
 * - As pastas de categoria ficam vazias, mas NÃO são apagadas.
 *
 * Segurança:
 *   - NUNCA apaga nada.
 *   - Começa em SIMULAÇÃO (DRY_RUN_PASTAS = true): só gera uma planilha de log.
 *   - Cada movimento é registrado; a função "desfazerClientesPorData" devolve
 *     todas as pastas para onde estavam.
 *   - Limite do Google: 6 min por execução. Se o log disser "PARADO POR TEMPO",
 *     execute de novo e ele continua.
 *
 * Como usar:
 *   1. script.google.com, logado como automacao@masterdatanet.com.br.
 *   2. Cole este arquivo, execute "organizarClientesPorData" (DRY_RUN_PASTAS = true) e confira o log.
 *   3. Mude DRY_RUN_PASTAS para false e execute de novo.
 */

const DRY_RUN_PASTAS = true; // true = só simula | false = executa de verdade

const CLIENTES_ID = '14BjKZWu_c911Ki8lqkK12blv8TQGmWDc'; // 03 - CLIENTES E POSTOS

// Categorias cujas pastas de cliente serão movidas
const CATEGORIAS = [
  '01 - Redes de postos',
  '02 - Postos independentes',
  '03 - Usinas e indústrias',
  '04 - Transportadoras e fazendas',
];
// Não mexe em: 05 - Clientes em prospecção, 98 - A classificar, 99 - Clientes inativos

const MESES = ['01 - Janeiro', '02 - Fevereiro', '03 - Março', '04 - Abril', '05 - Maio', '06 - Junho',
  '07 - Julho', '08 - Agosto', '09 - Setembro', '10 - Outubro', '11 - Novembro', '12 - Dezembro'];

const LIMITE_MS = 5 * 60 * 1000;
const CHAVE_HISTORICO = 'historicoMovimentos';

function organizarClientesPorData() {
  const inicio = Date.now();
  const raiz = DriveApp.getFolderById(CLIENTES_ID);
  const log = criarLogPastas_('LOG - Clientes por data');
  const props = PropertiesService.getScriptProperties();
  const historico = JSON.parse(props.getProperty(CHAVE_HISTORICO) || '{}');
  const cache = {};
  let movidas = 0, semArquivos = 0, terminou = true;

  log('INÍCIO', '', DRY_RUN_PASTAS ? 'MODO SIMULAÇÃO - nada será alterado' : 'MODO REAL');

  externo:
  for (let c = 0; c < CATEGORIAS.length; c++) {
    const it = raiz.getFoldersByName(CATEGORIAS[c]);
    if (!it.hasNext()) { log('AVISO', CATEGORIAS[c], 'Categoria não encontrada'); continue; }
    const categoria = it.next();
    const lista = [];
    const itClientes = categoria.getFolders();
    while (itClientes.hasNext()) lista.push(itClientes.next());
    for (let k = 0; k < lista.length; k++) {
      if (Date.now() - inicio > LIMITE_MS) { terminou = false; break externo; }
      const cliente = lista[k];
      const d = primeiraData_(cliente);
      if (!d) {
        semArquivos++;
        log('SEM ARQUIVOS', CATEGORIAS[c] + ' / ' + cliente.getName(), 'Não movida');
        continue;
      }
      const ano = String(d.getFullYear()), mes = MESES[d.getMonth()], dia = ('0' + d.getDate()).slice(-2);
      if (!DRY_RUN_PASTAS) {
        const destino = garantirCaminho_(raiz, [ano, mes, dia], cache);
        historico[cliente.getId()] = categoria.getId();
        cliente.moveTo(destino);
        props.setProperty(CHAVE_HISTORICO, JSON.stringify(historico));
      }
      movidas++;
      log('MOVER PASTA', CATEGORIAS[c] + ' / ' + cliente.getName(), '-> ' + ano + ' / ' + mes + ' / ' + dia);
    }
  }

  log(terminou ? 'FIM' : 'PARADO POR TEMPO', '',
    movidas + ' pastas movidas, ' + semArquivos + ' sem arquivos (não movidas)' +
    (terminou ? '' : ' - EXECUTE DE NOVO para continuar'));
}

/** Devolve todas as pastas movidas para a categoria de origem. */
function desfazerClientesPorData() {
  const props = PropertiesService.getScriptProperties();
  const historico = JSON.parse(props.getProperty(CHAVE_HISTORICO) || '{}');
  const log = criarLogPastas_('LOG - Desfazer clientes por data');
  Object.keys(historico).forEach(function (id) {
    const pasta = DriveApp.getFolderById(id);
    pasta.moveTo(DriveApp.getFolderById(historico[id]));
    log('DESFEITO', pasta.getName(), 'voltou para a categoria original');
    delete historico[id];
    props.setProperty(CHAVE_HISTORICO, JSON.stringify(historico));
  });
  log('FIM', '', 'Desfeito.');
}

// ===================== AUXILIARES =====================

/** Data mais antiga entre todos os arquivos da pasta (e subpastas). null se não houver arquivos. */
function primeiraData_(pasta) {
  let menor = null;
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const d = dataDoArquivo_(arquivos.next());
    if (!menor || d < menor) menor = d;
  }
  const subs = pasta.getFolders();
  while (subs.hasNext()) {
    const d = primeiraData_(subs.next());
    if (d && (!menor || d < menor)) menor = d;
  }
  return menor;
}

function dataDoArquivo_(f) {
  const n = f.getName();
  let m = n.match(/(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})/);                // 2026-03-17
  if (m) { const d = montarData_(+m[1], +m[2], +m[3]); if (d) return d; }
  m = n.match(/(?:^|\D)(\d{1,2})[-_.\/](\d{1,2})[-_.\/](20\d{2})(?!\d)/); // 17-03-2026
  if (m) { const d = montarData_(+m[3], +m[2], +m[1]); if (d) return d; }
  return f.getLastUpdated();
}

function montarData_(a, me, di) {
  if (me < 1 || me > 12 || di < 1 || di > 31) return null;
  const d = new Date(a, me - 1, di);
  return d.getMonth() === me - 1 ? d : null;
}

function garantirCaminho_(pai, partes, cache) {
  let atual = pai, chave = '';
  partes.forEach(function (nome) {
    chave += '/' + nome;
    if (!cache[chave]) {
      const it = atual.getFoldersByName(nome);
      cache[chave] = it.hasNext() ? it.next() : atual.createFolder(nome);
    }
    atual = cache[chave];
  });
  return atual;
}

function criarLogPastas_(titulo) {
  const ss = SpreadsheetApp.create(titulo + ' ' +
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm'));
  const sh = ss.getActiveSheet();
  sh.appendRow(['Data/Hora', 'Ação', 'Pasta', 'Destino']);
  sh.setFrozenRows(1);
  Logger.log('Planilha de log: ' + ss.getUrl());
  return function (acao, item, detalhe) {
    sh.appendRow([new Date(), (DRY_RUN_PASTAS ? '[SIMULAÇÃO] ' : '') + acao, item, detalhe]);
  };
}
