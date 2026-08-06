// ============================================================
// [UL] Dashboard_Taxonomias — 04_WebApp.gs
// Dados para o painel HTML e rotas do Web App
// ============================================================

function getDados() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PAINEL');
  if (!sh) return { erro: 'Aba PAINEL não encontrada. Execute a sincronização primeiro.' };
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return { erro: 'PAINEL vazio.' };
  const vals = sh.getRange(1,1,lastRow,lastCol).getValues();
  const hdr = vals[0].map(String);
  const rows = [];
  const iPaiKey = hdr.indexOf('PAI_KEY');
  const iFilhoKey = hdr.indexOf('FILHO_KEY');
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    // Descarta só a linha de timestamp (⏱) — mantém orphans (PAI_KEY vazio mas FILHO_KEY preenchido)
    const paiKey   = iPaiKey   >= 0 ? String(r[iPaiKey]   || '') : '';
    const filhoKey = iFilhoKey >= 0 ? String(r[iFilhoKey] || '') : '';
    if (!paiKey && !filhoKey) continue;          // linha vazia
    if (paiKey.startsWith('⏱')) continue;        // linha de timestamp
    const o = {};
    hdr.forEach((h,j) => { o[h] = r[j] !== null && r[j] !== undefined ? r[j] : ''; });
    rows.push(o);
  }
  // Conta incorretos para o label da topbar
  const shInc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚠️ INCORRETOS');
  const nIncorretos = shInc && shInc.getLastRow() > 1 ? shInc.getLastRow() - 1 : 0;

  // totalJira = todas as issues buscadas (lida da propriedade gravada pelo sync)
  const totalJiraStr = PropertiesService.getScriptProperties().getProperty('SYNC_TOTAL') || '';
  const totalJira = totalJiraStr ? parseInt(totalJiraStr, 10) : null;

  return {
    sync: new Date().toISOString(),
    total: rows.length,
    totalJira: totalJira,
    totalIncorretos: nIncorretos,
    avisoToken: getAvisoToken(),
    rows
  };
}

function getIncorretos() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚠️ INCORRETOS');
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  return vals.slice(1).map(r => {
    const o = {};
    hdr.forEach((h,j) => { o[h] = r[j] !== null && r[j] !== undefined ? r[j] : ''; });
    return o;
  });
}

function getDadosPeriodo(de, ate, campo) {
  const jqlParts = [];
  if (de)  jqlParts.push(campo + ' >= "' + de + '"');
  if (ate) jqlParts.push(campo + ' <= "' + ate + '"');
  const issues = _buscar(jqlParts.join(' AND '));
  if (!issues.length) return { sync: new Date().toISOString(), total: 0, rows: [] };
  const pais = {}, filhos = [];
  issues.forEach(i => {
    if ((i.fields.issuetype||{}).name==='Tarefa'||!i.fields.parent) pais[i.key]=i;
    else filhos.push(i);
  });
  filhos.forEach(f => {
    const pk=(f.fields.parent||{}).key;
    if(pk&&!pais[pk]) pais[pk]=_stubPai(pk,(f.fields.parent.fields||{}).summary||pk);
  });
  const dv=_carregarDePara('VEICULO');
  const H=['PAI_KEY','PAI_SUMMARY','BU','AGÊNCIA','MARCA (JIRA)','CAMP. MÍDIA','STATUS ADP (PAI)',
           'FILHO_KEY','SUMMARY (FILHO)','VEÍCULO','STATUS','RESPONSÁVEL','RELATOR',
           'QTD TAX.','PLATAFORMA','DATA LIMITE','ÚLT. ATU.'];
  const rows=[];
  const mapaPai={};
  Object.values(pais).forEach(p=>{
    const c=_cx(p.fields);
    mapaPai[p.key]={key:p.key,summary:p.fields.summary||'',bu:c.bu,agencia:c.agencia,
      marcaJira:c.marcaJira,campMidia:c.campMidia,statusAdp:c.statusAdp,
      assignee:(p.fields.assignee||{}).displayName||'',relator:c.relator,
      duedate:c.duedate,plataforma:c.plataforma,qtdTax:c.qtdTax,
      status:(p.fields.status||{}).name||'',updated:_fmt(p.fields.updated)};
  });
  const porPai={};
  filhos.forEach(f=>{const pk=(f.fields.parent||{}).key||'SEM_PAI';if(!porPai[pk])porPai[pk]=[];porPai[pk].push(f);});
  Object.keys(mapaPai).sort().forEach(pk=>{
    const pai=mapaPai[pk],fd=porPai[pk]||[];
    if(!fd.length){
      const o={};H.forEach((h,i)=>o[h]=[pai.key,pai.summary,pai.bu||'⚠️ vazio',pai.agencia||'⚠️ vazio',
        pai.marcaJira||'⚠️ vazio',pai.campMidia||'',pai.statusAdp||'⚠️ vazio',
        '','(sem subtarefas)','',pai.status,pai.assignee,pai.relator,
        pai.qtdTax||'',pai.plataforma||'',pai.duedate||'',pai.updated||''][i]);
      rows.push(o);
    } else {
      fd.forEach(f=>{
        const ff=f.fields,cf=_cx(ff),sum=ff.summary||'';
        const vv=_extrairTokens(sum,dv);if(!vv.length)vv.push('⚠️ Não mapeado');
        vv.forEach(vei=>{
          const bu=cf.bu||pai.bu||'⚠️ vazio',ag=cf.agencia||pai.agencia||'⚠️ vazio';
          const mk=cf.marcaJira||pai.marcaJira||'⚠️ vazio',cm=cf.campMidia||pai.campMidia||'';
          const sa=cf.statusAdp||pai.statusAdp||'⚠️ vazio',pl=cf.plataforma||pai.plataforma||'';
          const qt=cf.qtdTax!==''?cf.qtdTax:pai.qtdTax||'',dd=cf.duedate||pai.duedate||'';
          const o={};
          H.forEach((h,i)=>o[h]=[pai.key,pai.summary,bu,ag,mk,cm,sa,f.key,sum,vei,
            (ff.status||{}).name||'',(ff.assignee||{}).displayName||'',cf.relator||'',
            qt,pl,dd,_fmt(ff.updated)][i]);
          rows.push(o);
        });
      });
    }
  });
  return { sync: new Date().toISOString(), total: rows.length, rows };
}

// ── Formatação da planilha ────────────────────────────────────
function formatarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const rawLegada = ss.getSheetByName('RAW');
    if (rawLegada) { ss.deleteSheet(rawLegada); Logger.log('Aba RAW legada removida.'); }
  } catch(e) { Logger.log('Erro ao remover aba RAW legada: ' + e); }

  const shInicio = ss.getSheetByName('🏠 INICIO');
  if (shInicio) try { shInicio.setColumnWidth(1, 680); } catch(e) {}

  ss.getSheets().forEach(function(sh) {
    const nome = sh.getName();
    if (nome === '🏠 INICIO') return;
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;
    try {
      const maxR = sh.getMaxRows(), maxC = sh.getMaxColumns();
      if (maxR > lastRow + 1) sh.deleteRows(lastRow + 2, maxR - lastRow - 1);
      if (maxC > lastCol) sh.deleteColumns(lastCol + 1, maxC - lastCol);
    } catch(e) {}
    sh.getRange(1, 1, lastRow, lastCol).setVerticalAlignment('middle');
    if (lastRow >= 1) {
      sh.getRange(1, 1, 1, lastCol).setHorizontalAlignment('center').setFontWeight('bold');
      sh.setRowHeight(1, 28);
    }
    if (lastRow > 1) {
      for (var r = 2; r <= lastRow; r++) sh.setRowHeight(r, 22);
    }
    Logger.log('Formatando: ' + nome + ' (' + lastRow + ' linhas, ' + lastCol + ' cols)');
  });
}

// ── Web App ───────────────────────────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'status') {
    const prog = getProgresso();
    return ContentService.createTextOutput(JSON.stringify(prog)).setMimeType(ContentService.MimeType.JSON);
  }

  if (p.action === 'sync') {
    sincronizarJira();
    const t = HtmlService.createTemplateFromFile('Painel');
    const d = getDados();
    const sUrl = ScriptApp.getService().getUrl();
    t.dadosB64      = Utilities.base64Encode(JSON.stringify(d), Utilities.Charset.UTF_8);
    t.periodoB64    = Utilities.base64Encode(JSON.stringify(null), Utilities.Charset.UTF_8);
    t.incorretosB64 = Utilities.base64Encode(JSON.stringify(getIncorretos()), Utilities.Charset.UTF_8);
    t.serviceUrl    = sUrl;
    return t.evaluate().setTitle('Dashboard UL').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (p.acao === 'periodo') {
    const t = HtmlService.createTemplateFromFile('Painel');
    const dp = getDadosPeriodo(p.de||'', p.ate||'', p.campo||'updated');
    const per = { de: p.de||'', ate: p.ate||'', campo: p.campo||'updated' };
    const sUrl = ScriptApp.getService().getUrl();
    t.dadosB64      = Utilities.base64Encode(JSON.stringify(dp), Utilities.Charset.UTF_8);
    t.periodoB64    = Utilities.base64Encode(JSON.stringify(per), Utilities.Charset.UTF_8);
    t.incorretosB64 = Utilities.base64Encode(JSON.stringify(getIncorretos()), Utilities.Charset.UTF_8);
    t.serviceUrl    = sUrl;
    return t.evaluate().setTitle('Dashboard UL — Período').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const t = HtmlService.createTemplateFromFile('Painel');
  const d = getDados();
  const inc = getIncorretos();
  const sUrl = ScriptApp.getService().getUrl();
  t.dadosB64      = Utilities.base64Encode(JSON.stringify(d), Utilities.Charset.UTF_8);
  t.periodoB64    = Utilities.base64Encode(JSON.stringify(null), Utilities.Charset.UTF_8);
  t.incorretosB64 = Utilities.base64Encode(JSON.stringify(inc), Utilities.Charset.UTF_8);
  t.serviceUrl    = sUrl;
  return t.evaluate().setTitle('Dashboard UL').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
