// PricingIgoal.gs

const IGOAL_API_BASE_URL = 'https://my.spun.com.br';
const IGOAL_API_ENDPOINT = '/api/prices/update';
const IGOAL_API_TOKEN = '8jwl4v1ZmBYQlwFzPPEHNkYC8IOvRxB3ino1665b93f36cd228';

function buildIgoalApiUrl_() {
  return IGOAL_API_BASE_URL.replace(/\/$/, '') + IGOAL_API_ENDPOINT;
}

function runIgoalPricingUpdate() {
  try {
    var sheet = ensurePricingSheetIgoal_();
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { ok: true, message: 'Nenhuma regra para sincronizar na Igoal.' };
    }

    var headers = data[0] || [];
    var idx = getHeaderIndexMap_(headers);
    var syncIdx = headers.indexOf('Sincronizar');
    if (syncIdx === -1) {
      throw new Error("A coluna 'Sincronizar' não foi encontrada na aba de precificação da Igoal.");
    }

    var rowsToSync = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var shouldSync = String(row[syncIdx] || '').toLowerCase() === 'sim';
      if (!shouldSync) continue;
      rowsToSync.push({ row: row, rowNum: i + 1 });
    }

    if (!rowsToSync.length) {
      return { ok: true, message: 'Nenhuma regra marcada para sincronização na Igoal.' };
    }

    var logMessages = [];
    var url = buildIgoalApiUrl_();

    rowsToSync.forEach(function(item) {
      var row = item.row;
      var dominio = String(row[idx['dominio']] || '').trim();
      var urlPath = String(row[idx['url']] || '').trim();
      var companyIdRaw = String(row[idx['company_id']] || '').trim();
      var priceRule = String(row[idx['price_rule']] || '').trim();
      var utmSource = String(row[idx['utm_source']] || '').trim();
      var bloco = String(row[idx['bloco']] || '').trim();

      if (!dominio || !urlPath || !companyIdRaw || !priceRule) {
        logMessages.push('- ' + dominio + '/' + urlPath + ': ERRO! Campos obrigatórios ausentes.');
        return;
      }

      var companyIdValue = parseInt(companyIdRaw, 10);
      var payload = {
        dominio: dominio,
        url: urlPath,
        company_id: isNaN(companyIdValue) ? companyIdRaw : companyIdValue,
        price_rule: priceRule
      };
      if (utmSource) payload.utm_source = utmSource;
      if (bloco) payload.bloco = bloco;

      try {
        var response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          muteHttpExceptions: true,
          payload: JSON.stringify(payload),
          headers: { Authorization: IGOAL_API_TOKEN }
        });

        var statusCode = response.getResponseCode();
        var raw = response.getContentText() || '';
        var json;
        try { json = raw ? JSON.parse(raw) : null; } catch (err) { json = null; }

        var success = (statusCode === 200) && (!!json && (json.status === true || json.message === '200' || json.success === true));
        var message = '';
        if (success) {
          message = json && (json.data || json.message) ? (json.data || json.message) : 'Preço atualizado com sucesso.';
          sheet.getRange(item.rowNum, syncIdx + 1).clearContent();
          logMessages.push('- ' + dominio + '/' + urlPath + ': OK! ' + message);
        } else {
          message = json && (json.data || json.message) ? (json.data || json.message) : (raw || ('Erro HTTP ' + statusCode));
          logMessages.push('- ' + dominio + '/' + urlPath + ': ERRO! ' + message);
        }
      } catch (err) {
        logMessages.push('- ' + dominio + '/' + urlPath + ': ERRO! ' + err.message);
      }
    });

    return { ok: true, message: 'Resultado da Sincronização Igoal:\n\n' + logMessages.join('\n') };
  } catch (e) {
    Logger.log('[IGOAL][ERRO FATAL] ' + e.message);
    return { ok: false, message: 'Erro ao sincronizar com a Igoal: ' + e.message };
  }
}