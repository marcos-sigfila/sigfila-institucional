/**
 * Google Apps Script para gravar todos os campos dos formularios do SIGFILA.
 *
 * Como usar:
 * 1. Abra a planilha "E-mails cadastrados".
 * 2. Va em Extensoes > Apps Script.
 * 3. Substitua o codigo atual por este arquivo.
 * 4. Publique novamente como Web app mantendo acesso para "Qualquer pessoa".
 * 5. Se a URL do Web app mudar, atualize SHEETS_WEBAPP_URL nos HTMLs do site.
 */

const SHEET_NAME = 'E-mails cadastrados';

const HEADERS = [
  'Data',
  'E-mail',
  'Pagina',
  'User Agent',
  'Nome',
  'Telefone/WhatsApp',
  'Cidade ou prefeitura',
  'Mensagem',
  'Titulo da pagina',
  'Origem/Referenciador',
  'UTM Source',
  'UTM Medium',
  'UTM Campaign',
  'UTM Content',
  'UTM Term',
  'Idioma',
  'Resumo do lead'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = parsePayload_(e);
    const sheet = getLeadsSheet_();
    ensureHeaders_(sheet);

    sheet.appendRow([
      new Date(),
      pick_(data, ['email']),
      pick_(data, ['pagina', 'page', 'url']),
      pick_(data, ['user_agent', 'userAgent']),
      pick_(data, ['nome', 'name']),
      pick_(data, ['telefone', 'phone']),
      pick_(data, ['cidadePrefeitura', 'cidade', 'city', 'prefeitura']),
      pick_(data, ['mensagem', 'message']),
      pick_(data, ['pageTitle']),
      pick_(data, ['origem', 'referrer']),
      pick_(data, ['utmSource', 'utm_source']),
      pick_(data, ['utmMedium', 'utm_medium']),
      pick_(data, ['utmCampaign', 'utm_campaign']),
      pick_(data, ['utmContent', 'utm_content']),
      pick_(data, ['utmTerm', 'utm_term']),
      pick_(data, ['language']),
      pick_(data, ['leadResumo'])
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, service: 'SIGFILA leads' });
}

function parsePayload_(e) {
  const body = e && e.postData && e.postData.contents ? e.postData.contents : '';

  if (body) {
    try {
      return JSON.parse(body);
    } catch (err) {
      return parseQueryString_(body);
    }
  }

  return e && e.parameter ? e.parameter : {};
}

function parseQueryString_(body) {
  return body.split('&').reduce((acc, pair) => {
    const parts = pair.split('=');
    const key = decodeURIComponent((parts[0] || '').replace(/\+/g, ' '));
    const value = decodeURIComponent((parts.slice(1).join('=') || '').replace(/\+/g, ' '));
    if (key) {
      acc[key] = value;
    }
    return acc;
  }, {});
}

function getLeadsSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getActiveSheet() || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const firstCell = String(firstRow[0] || '').trim();
  const hasAnyValue = firstRow.some(value => String(value || '').trim() !== '');
  const alreadyHasHeaders = firstCell === HEADERS[0];

  if (!hasAnyValue) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  if (!alreadyHasHeaders) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  const currentHeaders = firstRow.map(value => String(value || '').trim());
  const needsUpdate = HEADERS.some((header, index) => currentHeaders[index] !== header);

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function pick_(data, keys) {
  for (const key of keys) {
    if (data[key] !== undefined && data[key] !== null && String(data[key]).trim() !== '') {
      return String(data[key]).trim();
    }
  }
  return '';
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
