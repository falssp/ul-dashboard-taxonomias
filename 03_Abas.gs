// ============================================================
// [UL] Dashboard_Taxonomias — 03_Abas.gs
// Gravação, formatação e proteção de todas as abas
// ============================================================

// ── Colunas que devem ter alinhamento CENTRAL (horizontal) ───
const COLS_CENTRO = ['KEY','TIPO','STATUS','CRIADO','ATUALIZADO','DATA LIMITE',
  'QTD TAXONOMIA','QTD TAX.','STATUS ADP','STATUS ADP (PAI)','BU','MARCA (JIRA)',
  'PLATAFORMA','ÚLT. ATU.','PARENT_KEY','FILHO_KEY','PAI_KEY','VEÍCULO','CAMP. MÍDIA'];

// Aplica alinhamento horizontal por coluna com base em COLS_CENTRO
function _alinharColunas(sh, hdr, nRows) {
  if (nRows < 2) return;
  hdr.forEach((h, i) => {
    const col = i + 1;
    const alinhamento = COLS_CENTRO.includes(h) ? 'center' : 'left';
    sh.getRange(2, col, nRows - 1, 1).setHorizontalAlignment(alinhamento);
  });
}

// Wrap nas colunas de texto livre, corte nas demais
function _wrapColunas(sh, hdr, nRows) {
  const WRAP_COLS = ['SUMMARY','PAI_SUMMARY','PARENT_SUMMARY','SUMMARY (FILHO)','ASSIGNEE','RELATOR','RESPONSÁVEL','AGÊNCIA_UL','AGÊNCIA'];
  hdr.forEach((h, i) => {
    const col = i + 1;
    const wrap = WRAP_COLS.includes(h) ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP;
    sh.getRange(1, col, Math.max(nRows, 1), 1).setWrapStrategy(wrap);
  });
}

// Protege o cabeçalho (linha 1) — só o dono pode editar
function _protegerCabecalho(sh) {
  try {
    // Remove proteções antigas na mesma aba
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
    const prot = sh.getRange(1, 1, 1, sh.getLastColumn()).protect();
    prot.setDescription('Cabeçalho — não editar');
    prot.setWarningOnly(true); // aviso ao tentar editar, mas não bloqueia completamente
  } catch(e) {}
}

// ── RAW PAI ───────────────────────────────────────────────────
function _gravarRAWPai(_ss, pais) {
  _ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = _ss.getSheetByName('RAW_PAI'); if (old) _ss.deleteSheet(old);
  const sh = _ss.insertSheet('RAW_PAI'); sh.setTabColor('#2E5BCD');
  const H = ['KEY','TIPO','SUMMARY','STATUS','ASSIGNEE','CRIADO','ATUALIZADO',
             'RELATOR','DATA LIMITE','QTD TAXONOMIA','STATUS ADP','BU','MARCA (JIRA)',
             'PLATAFORMA','AGÊNCIA_UL','CAMP. MÍDIA'];
  const rows = [H];
  pais.forEach(i => {
    const f = i.fields, c = _cx(f);
    rows.push([i.key,(f.issuetype||{}).name||'',f.summary||'',(f.status||{}).name||'',
      (f.assignee||{}).displayName||'',_fmt(f.created),_fmt(f.updated),
      c.relator,c.duedate,c.qtdTax,c.statusAdp,c.bu,c.marcaJira,c.plataforma,c.agencia,c.campMidia]);
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length);
  [80,75,360,110,160,90,90,180,90,60,110,120,110,100,110,120].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  _sinalizar(sh, rows.length, H, ['STATUS ADP','BU','MARCA (JIRA)','AGÊNCIA_UL']);
  _protegerCabecalho(sh);
  _trim(sh, rows.length, H.length);
}

// ── RAW FILHO ─────────────────────────────────────────────────
function _gravarRAWFilho(ss, filhos) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName('RAW_FILHO'); if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('RAW_FILHO'); sh.setTabColor('#7C3AED');
  const H = ['KEY','TIPO','SUMMARY','PARENT_KEY','PARENT_SUMMARY','STATUS','ASSIGNEE',
             'CRIADO','ATUALIZADO','RELATOR','DATA LIMITE','QTD TAXONOMIA','STATUS ADP',
             'BU','MARCA (JIRA)','PLATAFORMA','AGÊNCIA_UL','CAMP. MÍDIA'];
  const rows = [H];
  filhos.forEach(i => {
    const f = i.fields, pf = (f.parent||{}).fields||{}, c = _cx(f);
    rows.push([i.key,(f.issuetype||{}).name||'',f.summary||'',(f.parent||{}).key||'',
      pf.summary||'',(f.status||{}).name||'',(f.assignee||{}).displayName||'',
      _fmt(f.created),_fmt(f.updated),c.relator,c.duedate,c.qtdTax,
      c.statusAdp,c.bu,c.marcaJira,c.plataforma,c.agencia,c.campMidia]);
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length);
  [80,75,340,80,300,110,160,90,90,180,90,60,110,120,110,100,110,120].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  _sinalizar(sh, rows.length, H, ['STATUS ADP','BU','MARCA (JIRA)','AGÊNCIA_UL']);
  _protegerCabecalho(sh);
  _trim(sh, rows.length, H.length);
}

// ── DE_PARA ───────────────────────────────────────────────────
function inicializarDePara() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('DE_PARA')) return;
  _gravarDePara(ss.insertSheet('DE_PARA'));
}
function inicializarDeParaForcar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName('DE_PARA');
  if (s) ss.deleteSheet(s);
  _gravarDePara(ss.insertSheet('DE_PARA'));
  SpreadsheetApp.getUi().alert('DE_PARA recriada.');
}
function _gravarDePara(sh) {
  const d = [
    ['TIPO','TOKEN_JIRA','NOME_OFICIAL'],
    ['MARCA','Dove','Dove'],['MARCA','Dove Skincare','Dove'],['MARCA','Dove Diva','Dove'],
    ['MARCA','Dove MSTB','Dove'],['MARCA','Dove Deos','Dove'],['MARCA','Dove Men Care','Dove'],
    ['MARCA','Debeers','Dove'],['MARCA','Dove Debeers','Dove'],['MARCA','Dune','Dove'],['MARCA','DMC','Dove'],
    ['MARCA','Seda','Seda'],['MARCA','Seda Core','Seda'],['MARCA','Seda Steel','Seda'],['MARCA','RES','Seda'],
    ['MARCA','Closeup','Closeup'],['MARCA','Close Up','Closeup'],['MARCA','Close-Up','Closeup'],
    ['MARCA','Hellmanns',"Hellmann's"],['MARCA',"Hellmann's","Hellmann's"],['MARCA','HELLMANNS',"Hellmann's"],
    ['MARCA','Copa Supreme',"Hellmann's"],['MARCA','OMO','OMO'],['MARCA','Alchemist','OMO'],
    ['MARCA','CIF','CIF'],['MARCA','Rexona','Rexona'],['MARCA','Lux','Lux'],
    ['MARCA','Clear','Clear'],['MARCA','Crown','Clear'],['MARCA','Crown World Cup','Clear'],
    ['MARCA','Knorr','Knorr'],['MARCA','KNORR','Knorr'],
    ['VEICULO','YT','YouTube'],['VEICULO','Youtube','YouTube'],['VEICULO','YouTube','YouTube'],
    ['VEICULO','Meta','Meta'],['VEICULO','Tiktok','TikTok'],['VEICULO','TikTok','TikTok'],
    ['VEICULO','DV360','DV360'],['VEICULO','Pinterest','Pinterest'],['VEICULO','Social','Social'],
    ['VEICULO','Flashtalking','Flashtalking'],['VEICULO','Flashtalkng','Flashtalking'],
    ['VEICULO','MP','Mídia Programática'],['VEICULO','Programmatic','Mídia Programática'],
    ['VEICULO','Midia Programatica','Mídia Programática'],
    ['VEICULO','Amazon DSP','Amazon DSP'],['VEICULO','Amz Ads','Amazon DSP'],
    ['VEICULO','Google Search','Google Search'],['VEICULO','Netflix','Netflix'],
    ['VEICULO','Spotify','Spotify'],['VEICULO','Meliads','Meliads'],['VEICULO','UOL','UOL'],
    ['VEICULO','Futbol Sites','Futbol Sites'],['VEICULO','365 Scores','365 Scores'],
    ['VEICULO','GETV','GETV'],['VEICULO','Globoplay','Globoplay'],
    ['VEICULO','Tag Kantar','Tag Kantar'],['VEICULO','B&W','B&W'],
  ];
  sh.getRange(1,1,d.length,3).setValues(d);
  _estilizarSheet(sh, d.length, 3);
  [90,220,200].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  // Alinhamento: TIPO centralizado, TOKEN e NOME à esquerda
  sh.getRange(1,1,d.length,3).setVerticalAlignment('middle');
  sh.getRange(2,1,d.length-1,1).setHorizontalAlignment('center');
  sh.getRange(2,2,d.length-1,2).setHorizontalAlignment('left');
  sh.setFrozenRows(1);
  _protegerCabecalho(sh);
  _trim(sh, d.length, 3);
}

function _extrairTokens(texto, mapa) {
  const found = new Set();
  for (const token of Object.keys(mapa)) {
    const re = new RegExp('(?:^|[\\s|\\-/,])' + token.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '(?:$|[\\s|\\-/,])', 'i');
    if (re.test(' ' + texto + ' ')) found.add(mapa[token]);
  }
  return [...found];
}

function _carregarDePara(tipo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DE_PARA');
  if (!sh) return {};
  const map = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]===tipo && r[1] && r[2]) map[r[1].toString().toLowerCase().trim()] = r[2].toString().trim();
  });
  return map;
}

// ── PAINEL ────────────────────────────────────────────────────
function _gravarPainel(ss, pais, filhos) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName('PAINEL'); if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('PAINEL');
  const dv = _carregarDePara('VEICULO');
  const H = ['PAI_KEY','PAI_SUMMARY','BU','AGÊNCIA','MARCA (JIRA)','CAMP. MÍDIA','STATUS ADP (PAI)',
             'FILHO_KEY','SUMMARY (FILHO)','VEÍCULO','STATUS','RESPONSÁVEL','RELATOR',
             'QTD TAX.','PLATAFORMA','DATA LIMITE','ÚLT. ATU.'];
  const mapaPai = {};
  pais.forEach(p => {
    const c = _cx(p.fields);
    mapaPai[p.key] = { key:p.key, summary:p.fields.summary||'',
      bu:c.bu, agencia:c.agencia, marcaJira:c.marcaJira, campMidia:c.campMidia,
      statusAdp:c.statusAdp, assignee:(p.fields.assignee||{}).displayName||'',
      relator:c.relator, duedate:c.duedate, plataforma:c.plataforma, qtdTax:c.qtdTax,
      status:(p.fields.status||{}).name||'', updated:_fmt(p.fields.updated) };
  });
  const porPai = {};
  filhos.forEach(f => {
    const pk = (f.fields.parent||{}).key||'SEM_PAI';
    if (!porPai[pk]) porPai[pk] = [];
    porPai[pk].push(f);
  });
  const rows = [H];
  Object.keys(mapaPai).sort().forEach(pk => {
    const pai = mapaPai[pk], fd = porPai[pk]||[];
    if (!fd.length) {
      rows.push([pai.key,pai.summary,pai.bu||'⚠️ vazio',pai.agencia||'⚠️ vazio',
        pai.marcaJira||'⚠️ vazio',pai.campMidia||'',pai.statusAdp||'⚠️ vazio',
        '','(sem subtarefas)','',pai.status,pai.assignee,pai.relator,
        pai.qtdTax||'',pai.plataforma||'',pai.duedate||'',pai.updated||'']);
    } else {
      fd.forEach(f => {
        const ff = f.fields, cf = _cx(ff);
        const bu=cf.bu||pai.bu||'⚠️ vazio', ag=cf.agencia||pai.agencia||'⚠️ vazio';
        const mk=cf.marcaJira||pai.marcaJira||'⚠️ vazio', cm=cf.campMidia||pai.campMidia||'';
        const sa=cf.statusAdp||pai.statusAdp||'⚠️ vazio', pl=cf.plataforma||pai.plataforma||'';
        const qt=cf.qtdTax!==''?cf.qtdTax:pai.qtdTax||'', dd=cf.duedate||pai.duedate||'';
        const rl=cf.relator||pai.relator||'', upd=_fmt(ff.updated);
        const sum=ff.summary||'';
        const vv=_extrairTokens(sum,dv); if(!vv.length) vv.push('⚠️ Não mapeado');
        vv.forEach(vei => rows.push([pai.key,pai.summary,bu,ag,mk,cm,sa,
          f.key,sum,vei,(ff.status||{}).name||'',(ff.assignee||{}).displayName||'',rl,qt,pl,dd,upd]));
      });
    }
  });
  // Órfãos
  (porPai['SEM_PAI']||[]).forEach(f => {
    const ff=f.fields, cf=_cx(ff), sum=ff.summary||'';
    const vv=_extrairTokens(sum,dv); if(!vv.length) vv.push('⚠️ Não mapeado');
    vv.forEach(vei => rows.push(['','(sem pai)',cf.bu||'⚠️ vazio',cf.agencia||'⚠️ vazio',
      cf.marcaJira||'⚠️ vazio',cf.campMidia||'',cf.statusAdp||'⚠️ vazio',
      f.key,sum,vei,(ff.status||{}).name||'',(ff.assignee||{}).displayName||'',cf.relator||'',
      cf.qtdTax||'',cf.plataforma||'',cf.duedate||'',_fmt(ff.updated)]));
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length);
  [80,340,115,115,120,115,115,80,320,115,115,155,155,55,100,90,90].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  sh.getRange(rows.length+2,1).setValue('⏱ Sync: '+new Date().toLocaleString('pt-BR')).setFontColor('#888').setFontStyle('italic');
  _protegerCabecalho(sh);
  _trim(sh, rows.length+2, H.length);
}

// ── VISUAL ────────────────────────────────────────────────────
function atualizarAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { _gravarVisual(ss); } catch(e) { Logger.log('VISUAL: '+e); }
  try { _gravarTabela(ss); } catch(e) { Logger.log('TABELA: '+e); }
  SpreadsheetApp.getUi().alert('Abas VISUAL e TABELA atualizadas.');
}

function _gravarVisual(ss) {
  const painel = ss.getSheetByName('PAINEL'); if (!painel) return;
  const old = ss.getSheetByName('📊 VISUAL'); if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('📊 VISUAL');
  const data = painel.getDataRange().getValues(), hdr = data[0].map(String);
  const I = {};
  hdr.forEach((h,i) => I[h]=i);
  const rows = data.slice(1).filter(r => r[0] !== undefined && r[0] !== '' && !String(r[0]).startsWith('⏱'));
  const arvore = {};
  rows.forEach(r => {
    const bu=String(r[I['BU']]||'⚠️ Sem BU'), ag=String(r[I['AGÊNCIA']]||'⚠️ Sem agência');
    const mk=String(r[I['MARCA (JIRA)']]||'⚠️ Não mapeada'), camp=String(r[I['PAI_SUMMARY']]||'—');
    if(!arvore[bu]) arvore[bu]={};
    if(!arvore[bu][ag]) arvore[bu][ag]={};
    if(!arvore[bu][ag][mk]) arvore[bu][ag][mk]={};
    if(!arvore[bu][ag][mk][camp]) arvore[bu][ag][mk][camp]=[];
    arvore[bu][ag][mk][camp].push(r);
  });
  let estimativa = 5;
  for(const bu of Object.keys(arvore)){
    estimativa += 1;
    for(const ag of Object.keys(arvore[bu]||{})){
      estimativa += 1;
      for(const mk of Object.keys(arvore[bu][ag]||{})){
        estimativa += 1;
        for(const camp of Object.keys(arvore[bu][ag][mk]||{})){
          estimativa += 2 + arvore[bu][ag][mk][camp].length;
        }
      }
    }
    estimativa += 1;
  }
  const linhasNecessarias = estimativa + 20;
  try {
    const atualRows = sh.getMaxRows();
    if (linhasNecessarias > atualRows) sh.insertRowsAfter(atualRows, linhasNecessarias - atualRows);
  } catch(e) {}

  const CORES=['#2E5BCD','#7C3AED','#059669','#D97706','#DC2626','#0891B2','#7C2D12','#4F46E5'];
  const corBU={};let ci=0,ln=1;
  sh.getRange(ln,1,1,9).merge().setValue('📋 Dashboard Taxonomias UL')
    .setBackground('#1a1a2e').setFontColor('#fff').setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(ln++,24);
  const sync = painel.getRange(painel.getLastRow(),1).getValue();
  sh.getRange(ln,1,1,9).merge().setValue(String(sync))
    .setBackground('#252d42').setFontColor('#8899bb').setFontSize(8).setFontStyle('italic')
    .setHorizontalAlignment('left');
  sh.setRowHeight(ln++,14);
  for(const bu of Object.keys(arvore).sort()){
    if(!corBU[bu]) corBU[bu]=CORES[ci++%CORES.length];
    const cor=corBU[bu];
    const tot=Object.values(arvore[bu]).flatMap(a=>Object.values(a)).flatMap(m=>Object.values(m)).flat().length;
    if(ln > 3) {
      sh.getRange(ln,1,1,9).merge().setValue('').setBackground('#f0f0f0');
      sh.setRowHeight(ln++,5);
    }
    sh.getRange(ln,1,1,9).merge().setValue('  '+bu+'   ('+(tot===1?'1 ticket':tot+' tickets')+')')
      .setBackground(cor).setFontColor('#fff').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(ln++,22);
    for(const ag of Object.keys(arvore[bu]).sort()){
      const totAg=Object.values(arvore[bu][ag]).flatMap(m=>Object.values(m)).flat().length;
      sh.getRange(ln,1,1,9).merge().setValue('  🏢 '+ag+'   ('+(totAg===1?'1 ticket':totAg+' tickets')+')')
        .setBackground('#e8ecf4').setFontColor('#374151').setFontSize(10).setFontWeight('bold')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
      sh.setRowHeight(ln++,20);
      for(const mk of Object.keys(arvore[bu][ag]).sort()){
        const totMk=Object.values(arvore[bu][ag][mk]).flat().length;
        sh.getRange(ln,1,1,9).merge().setValue('    ◈ '+mk+'   ('+(totMk===1?'1 ticket':totMk+' tickets')+')')
          .setBackground('#f5f7ff').setFontColor(cor).setFontSize(9).setFontWeight('bold')
          .setHorizontalAlignment('left').setVerticalAlignment('middle');
        sh.setRowHeight(ln++,18);
        for(const camp of Object.keys(arvore[bu][ag][mk]).sort()){
          const cr=arvore[bu][ag][mk][camp], totC=cr.length;
          sh.getRange(ln,1,1,9).merge().setValue('      ◆ '+camp+'   ('+(totC===1?'1 ticket':totC+' tickets')+')')
            .setBackground('#dce3ff').setFontColor('#1e3a8a').setFontSize(9).setFontWeight('bold')
            .setHorizontalAlignment('left').setVerticalAlignment('middle');
          sh.setRowHeight(ln++,18);
          // Cabeçalho das colunas de ticket
          sh.getRange(ln,1,1,9).setValues([['Filho / Veículo','Status ADP','Status','Plataforma','Qtd Tax','Responsável','','','Últ. Atu.']])
            .setBackground('#374151').setFontColor('#fff').setFontSize(8).setFontWeight('bold')
            .setHorizontalAlignment('center').setVerticalAlignment('middle');
          sh.setRowHeight(ln++,16);
          cr.forEach(r=>{
            const fsum=String(r[I['SUMMARY (FILHO)']]||'—'), vei=String(r[I['VEÍCULO']]||'—');
            const label=fsum!=='(sem subtarefas)'?fsum+' / '+vei:fsum;
            const sts=String(r[I['STATUS']]||'—');
            const row=sh.getRange(ln,1,1,9);
            row.setValues([[label,String(r[I['STATUS ADP (PAI)']]||''),sts,String(r[I['PLATAFORMA']]||''),
              r[I['QTD TAX.']]||'',String(r[I['RESPONSÁVEL']]||''),'',' ',
              (function(v){
                if(!v) return '';
                var s=String(v);
                var m=s.match(/(\d{4})-(\d{2})-(\d{2})/);
                if(m) return m[3]+'/'+m[2]+'/'+m[1];
                if(s.includes('GMT')||s.includes('UTC')){var d=new Date(s);if(!isNaN(d))return ('0'+d.getDate()).slice(-2)+'/'+ ('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();}
                return s.substring(0,10);
              })(r[I['ÚLT. ATU.']])]])
              .setFontSize(9).setBackground('#fff').setFontColor('#1f2937')
              .setHorizontalAlignment('center').setVerticalAlignment('middle');
            sh.getRange(ln,1).setHorizontalAlignment('left')
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
            row.setBorder(true,true,true,true,false,false,'#d1d5db',SpreadsheetApp.BorderStyle.SOLID);
            sh.setRowHeight(ln++,18);
          });
        }
      }
    }
  }
  [310,95,105,100,65,140,15,95,95].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1);
  const visualLastRow = sh.getLastRow();
  _trim(sh, visualLastRow, 9);
}

// ── TABELA ────────────────────────────────────────────────────
function _gravarTabela(ss) {
  const painel=ss.getSheetByName('PAINEL'); if(!painel) return;
  const old=ss.getSheetByName('📋 TABELA'); if(old) ss.deleteSheet(old);
  const sh=ss.insertSheet('📋 TABELA');
  const data=painel.getDataRange().getValues();
  if(data.length<2) return;
  const hdr=data[0].map(String);
  // Cabeçalho
  sh.getRange(1,1,1,hdr.length).setValues([hdr])
    .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1,30); sh.setFrozenRows(1); sh.setFrozenColumns(2);
  // Dados
  const rows=data.slice(1).filter(r=>r[0]!==undefined&&r[0]!==''&&!String(r[0]).startsWith('⏱'));
  if(!rows.length) return;
  sh.getRange(2,1,rows.length,hdr.length).setValues(rows).setFontSize(10);
  // Zebra + altura + alinhamento vertical
  for(let i=0;i<rows.length;i++){
    sh.getRange(i+2,1,1,hdr.length)
      .setBackground(i%2===0?'#f0f4ff':'#fff')
      .setVerticalAlignment('middle');
    sh.setRowHeight(i+2,24);
  }
  // Alinhamento horizontal por coluna
  _alinharColunas(sh, hdr, rows.length+1);
  // Wrap nas colunas de texto
  _wrapColunas(sh, hdr, rows.length+1);
  [80,340,115,115,120,115,115,80,320,115,115,155,155,55,100,90,90].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length+1,hdr.length).createFilter(); } catch(e) {}
  _protegerCabecalho(sh);
  _trim(sh,rows.length+1,hdr.length);
}

// ── INCORRETOS ────────────────────────────────────────────────
function _gravarIncorretos(ss, incorretos) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const nome = '⚠️ INCORRETOS';
  const old = ss.getSheetByName(nome); if (old) ss.deleteSheet(old);
  if (!incorretos.length) return;
  const sh = ss.insertSheet(nome); sh.setTabColor('#DC2626');
  const H = ['KEY','TIPO','SUMMARY','STATUS','RESPONSÁVEL','ATUALIZADO','LINK JIRA'];
  const rows = [H];
  incorretos.forEach(i => {
    const f = i.fields;
    rows.push([i.key,(f.issuetype||{}).name||'',f.summary||'',(f.status||{}).name||'',
      (f.assignee||{}).displayName||'',_fmt(f.updated),JIRA_BASE + '/browse/' + i.key]);
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length);
  [80,100,340,120,155,90,300].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  _protegerCabecalho(sh);
  _trim(sh, rows.length, H.length);
  Logger.log('INCORRETOS: ' + incorretos.length + ' issues com tipo inválido.');
}

// ── Helpers de estilo ─────────────────────────────────────────
function _estilizarSheet(sh, nRows, nCols) {
  // Cabeçalho
  sh.getRange(1,1,1,nCols)
    .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);
  sh.setRowHeight(1,30);
  // Dados — zebra claro, todas as células: vertical=meio, fonte 10
  for(let i=2;i<=nRows;i++){
    sh.getRange(i,1,1,nCols)
      .setVerticalAlignment('middle').setFontSize(10)
      .setBackground(i%2===0?'#f0f4ff':'#fff');
    sh.setRowHeight(i,24);
  }
}

function _sinalizar(sh, nRows, hdr, colNames) {
  if (nRows < 2) return;
  colNames.forEach(nome => {
    const col = hdr.indexOf(nome);
    if (col < 0) return;
    const vals = sh.getRange(2, col+1, nRows-1, 1).getValues();
    vals.forEach((row, i) => {
      if (!row[0] || String(row[0]).trim() === '') {
        sh.getRange(i+2, col+1)
          .setValue('⚠️ vazio')
          .setBackground('#fff3cd')
          .setFontColor('#856404')
          .setHorizontalAlignment('center');
      }
    });
  });
}

function _trim(sh, usedRows, usedCols) {
  try {
    const mr=sh.getMaxRows(), mc=sh.getMaxColumns();
    if(mr>usedRows+1) sh.deleteRows(usedRows+2, mr-usedRows-1);
    if(mc>usedCols)   sh.deleteColumns(usedCols+1, mc-usedCols);
  } catch(e) {}
}
