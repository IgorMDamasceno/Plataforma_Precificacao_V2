// Pricing.gs

const PMD_API_URL = "https://pmd.cloud.adseleto.com/api/domains/";
const PMD_API_TOKEN = "6fae52ca5fb3d41eca803d903a73e2d011bb96b5415f101c563025a8c3470d85b249b2103a3e2a4dc073c3f62aebd438f0ddf2b72d6bc0c9156a8e5801d3d73023600b64232f6e407f4cea1f2c97e18c1471978083dd12ac164d76b3689729d9649e7e325b33bc06714842ea19e3160a5a4a4dda11b29bcdb1a4b31ca562a6f4";

/**
 * Helper para buscar as regras atuais diretamente da API PMD.
*/
function getPricingRuleUrl_(domainId) {
  try {
    const response = UrlFetchApp.fetch(PMD_API_URL + domainId, {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + PMD_API_TOKEN },
      muteHttpExceptions: true
    });
if (response.getResponseCode() !== 200) {
      Logger.log(`[ERRO GET] Falha ao buscar regras para domínio ${domainId}. Resposta: ${response.getContentText()}`);
return null; // Retorna null para indicar falha na busca
    }
    const data = JSON.parse(response.getContentText());
return data.data?.attributes?.pricing_rule_url || [];
  } catch(e) {
    Logger.log(`[ERRO GET] Falha ao conectar na API para o domínio ${domainId}: ${e.message}`);
return null; // Retorna null em caso de erro de conexão
  }
}

/**
 * Sincroniza de forma inteligente: busca, mescla e envia as regras.
*/
function runAdSeletoPricingUpdate() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Precificação - AdSeleto');
if (!sheet) throw new Error('A aba "Precificação - AdSeleto" não foi encontrada.');

    const range = sheet.getDataRange();
    const allData = range.getValues();
const headers = allData.shift();
    const syncColumnIndex = headers.indexOf('Sincronizar');
    
    if (syncColumnIndex === -1) throw new Error("A coluna 'Sincronizar' não foi encontrada.");
const rowsToSync = allData.map((row, index) => ({ data: row, rowNum: index + 2 }))
                              .filter(item => item.data[syncColumnIndex] && item.data[syncColumnIndex].toString().toLowerCase() === 'sim');
if (rowsToSync.length === 0) {
      return { ok: true, message: 'Nenhuma regra nova ou alterada para sincronizar.'
};
    }
    
    const rulesFromSheetByDomain = {};
rowsToSync.forEach(item => {
      const domainId = item.data[1];
      if (domainId) {
        if (!rulesFromSheetByDomain[domainId]) rulesFromSheetByDomain[domainId] = [];
        rulesFromSheetByDomain[domainId].push({
            spnprice_id: item.data[0], domain_id: domainId, price_rule: parseFloat(item.data[2]),
            utm_source: item.data[3], slot_id: item.data[4], url: item.data[5],
            _rowNum: item.rowNum // Guarda o número da linha para limpar depois
     
   });
      }
    });

    const domainsToUpdate = Object.keys(rulesFromSheetByDomain);
const logMessages = [];
    
    domainsToUpdate.forEach(domainId => {
      Logger.log(`[INFO] Sincronizando domínio ${domainId}...`);
      
      // 1. GET - Busca regras existentes na plataforma
      const currentRulesFromAPI = getPricingRuleUrl_(domainId);
      if (currentRulesFromAPI === null) {
        logMessages.push(`- Domínio ${domainId}: ERRO! Não foi possível buscar as regras atuais da API. O domínio foi ignorado.`);
        return; // Pula para o próximo domínio
      }

  
    const finalRules = currentRulesFromAPI.slice(); // Cria uma cópia para trabalhar
      const rulesFromSheet = rulesFromSheetByDomain[domainId];

      // 2. MERGE - Mescla as regras da planilha com as da API
      rulesFromSheet.forEach(ruleFromSheet => {
        let found = false;
        for (let i = 0; i < finalRules.length; i++) {
          // Tenta encontrar por ID primeiro
         
 if (finalRules[i].spnprice_id && String(finalRules[i].spnprice_id) == String(ruleFromSheet.spnprice_id)) {
            finalRules[i].price_rule = ruleFromSheet.price_rule;
found = true;
            break;
          }
        }
        if (!found) {
          for (let i = 0; i < finalRules.length; i++) {
            // Se não achou por ID, tenta pela combinação
            if (finalRules[i].url == ruleFromSheet.url && finalRules[i].slot_id == ruleFromSheet.slot_id && finalRules[i].utm_source == ruleFromSheet.utm_source) {
              finalRules[i].price_rule 
= ruleFromSheet.price_rule;
              found = true;
              break;
            }
          }
        }
        if (!found) { // Se não encontrou de nenhuma forma, é uma regra nova
          finalRules.push({
              spnprice_id: ruleFromSheet.spnprice_id, domain_id: ruleFromSheet.domain_id, price_rule: ruleFromSheet.price_rule,
              utm_source: ruleFromSheet.utm_source, slot_id: ruleFromSheet.slot_id, url: ruleFromSheet.url
       
   });
        }
      });
const payload = { data: { pricing_rule_url: finalRules } };
// 3. PUT - Envia a lista completa e mesclada de volta
      const response = UrlFetchApp.fetch(PMD_API_URL + domainId, {
        method: 'PUT', contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + PMD_API_TOKEN },
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
const responseCode = response.getResponseCode();
      if (responseCode >= 200 && responseCode < 300) {
        logMessages.push(`- Domínio ${domainId}: OK! Sincronizado com sucesso.`);
// Limpa o "Sim" da coluna
        rulesFromSheet.forEach(rule => {
          sheet.getRange(rule._rowNum, syncColumnIndex + 1).clearContent();
        });
} else {
        const errorText = response.getContentText();
        logMessages.push(`- Domínio ${domainId}: ERRO! (${responseCode}): ${errorText}`);
}
    });
    
    return { ok: true, message: 'Resultado da Sincronização:\n\n' + logMessages.join('\n') };
} catch (e) {
    Logger.log(`[ERRO FATAL] ${e.message}`);
return { ok: false, message: `Ocorreu um erro fatal no script: ${e.message}` };
  }
}