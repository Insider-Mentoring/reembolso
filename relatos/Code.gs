// Recebe os relatos (problemas e sugestões de melhoria) enviados pelo app
// e grava uma linha por relato na aba "Relatos" desta planilha. O print, quando
// vem, é salvo numa pasta do Drive de quem publicou o script e o link vai na linha.
//
// Como publicar (uma vez só):
//   1. Crie uma planilha no Google Drive (ex: "Relatos - Reembolso IGD").
//   2. Extensões → Apps Script → apague o conteúdo e cole este arquivo inteiro → Salvar.
//   3. Implantar → Nova implantação → tipo "App da Web":
//        Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
//   4. Autorize o acesso quando pedir e copie a URL do app da Web (termina em /exec).
//   5. Essa URL vai na constante RELATOS_URL do index.html.
//
// Ao alterar este arquivo depois: Implantar → Gerenciar implantações → editar →
// Versão: "Nova versão" (senão a URL continua rodando o código antigo).

const ABA = 'Relatos';
const PASTA_PRINTS = 'Relatos Reembolso IGD - Prints';
const CABECALHO = ['Data', 'Tipo', 'Nome', 'E-mail', 'Mensagem', 'Status', 'Print', 'Versão do app', 'Diagnóstico'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const mensagem = String(d.mensagem || '').trim();
    if (!mensagem) return resposta({ ok: false, erro: 'mensagem vazia' });

    // O app reenvia quando não consegue ler a resposta — o mesmo id em até
    // 10 min é o mesmo relato chegando de novo, não grava duplicado.
    const cache = CacheService.getScriptCache();
    if (d.id) {
      if (cache.get('relato_' + d.id)) return resposta({ ok: true, duplicado: true });
      cache.put('relato_' + d.id, '1', 600);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CABECALHO);
      sh.setFrozenRows(1);
    } else if (sh.getRange(1, 2).getValue() !== 'Tipo') {
      // Aba criada antes da coluna Tipo existir: abre a coluna no lugar certo
      // em vez de desalinhar as linhas que já estão lá.
      sh.insertColumnBefore(2);
      sh.getRange(1, 2).setValue('Tipo');
    }

    let linkPrint = '';
    if (typeof d.print === 'string' && /^data:image\/(png|jpeg|webp);base64,/.test(d.print)) {
      const [meta, b64] = d.print.split(',');
      const mime = meta.slice(5, meta.indexOf(';'));
      const nome = 'relato-' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd_HH-mm-ss')
        + '.' + mime.split('/')[1];
      const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, nome);
      linkPrint = pastaPrints().createFile(blob).getUrl();
    }

    sh.appendRow([
      new Date(),
      d.tipo === 'sugestao' ? 'Sugestão' : 'Problema',
      texto(d.nome, 200),
      texto(d.email, 200),
      texto(mensagem, 5000),
      'novo',
      linkPrint,
      texto(d.versao, 100),
      // Célula do Sheets aceita até 50 mil caracteres.
      texto(JSON.stringify(d.diag || {}, null, 1), 45000),
    ]);
    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}

function pastaPrints() {
  const it = DriveApp.getFoldersByName(PASTA_PRINTS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_PRINTS);
}

// Texto que começa com = + - @ vira fórmula no Sheets — o apóstrofo força texto puro.
function texto(v, max) {
  let s = String(v == null ? '' : v).slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
