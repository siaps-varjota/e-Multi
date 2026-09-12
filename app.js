/* eMulti - Painel de Indicadores M1 e M2 */
'use strict';

var SPREADSHEET_ID = '1EeE2PqXCXn4sxQW9FkGTQ5Zx-7Fuh8gKpvHUhQ-Y_dA';
var EQUIPES = ['Equipe A', 'Equipe B', 'Equipe C'];
var JANELA_MESES = 2;
var REFERENCIA_SUFIXO = ' (ref)';
var NOTAS_METODOLOGICAS = '<ol class="notes-list"><li>M1: Atendimentos por pessoa (Atendimentos + Atividades Coletivas) ÷ Pessoas Atendidas</li><li>M2: Ações Compartilhadas ÷ Ações Realizadas</li></ol>';
var OV_STATUS = {Ótimo:'Ótimo desempenho',Bom:'Bom desempenho',Suficiente:'Desempenho suficiente',Regular:'Desempenho regular'};
var OV_STATUS_FALLBACK = 'Sem dados';
var PONTOS_POR_CLASSE = {Ótimo:4,Bom:3,Suficiente:2,Regular:1};

  function classificarM1(val){
    if(val===null || val===undefined) return null;
    if(val >= 4) return 'Ótimo';
    if(val >= 3) return 'Bom';
    if(val >= 2) return 'Suficiente';
    return 'Regular';
  }
  function classificarM2(val){
    if(val===null || val===undefined) return null;
    if(val >= 80) return 'Ótimo';
    if(val >= 60) return 'Bom';
    if(val >= 40) return 'Suficiente';
    return 'Regular';
  }
  function classificarDesempenho(notaFinal){
    if(notaFinal===null || notaFinal===undefined) return null;
    if(notaFinal >= 32) return 'Ótimo';
    if(notaFinal >= 24) return 'Bom';
    if(notaFinal >= 16) return 'Suficiente';
    return 'Regular';
  }
  function escapeHtml(s){
    if(!s) return '';
    return String(s).replace(/[&<>"']/g, function(c){
      var m = {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'};
      return m[c];
    });
  }
  function requiredSheetNames(){
    return ['M1_' + EQUIPES[0], 'M2_' + EQUIPES[0], 'M1_' + EQUIPES[1], 'M2_' + EQUIPES[1], 'M1_' + EQUIPES[2], 'M2_' + EQUIPES[2]];
  }
  function startOfMonth(d){ return new Date(d.getFullYear(), d.getMonth(), 1); }
  function addMonths(d, n){ return new Date(d.getFullYear(), d.getMonth()+n, 1); }
  function monthOptionValue(d){ return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0'); }

  function processDashboard(record, serie){
    var resultadosMensais = [];
    var resultadosJanela = [];
    var HOJE = new Date();
    var HOJE_YYYY_MM = monthOptionValue(HOJE);

    record.data.forEach(function(mes, mesIdx){
      var mês = mes.mesReferencia.split('-');
      var d = new Date(parseInt(mês[0], 10), parseInt(mês[1], 10)-1, 1);
      var anoMês = monthOptionValue(d);
      var ativo = anoMês <= HOJE_YYYY_MM;

      var numerador = parseInt(mes.atendimentos, 10) + parseInt(mes.atividadesColetivas, 10);
      var denominador = parseInt(mes.pessoasAtendidas, 10);
      var m1 = denominador ? numerador / denominador : null;
      var m2 = parseInt(mes.acoesCompartilhadas, 10);
      var m2_denom = parseInt(mes.acoesRealizadas, 10);
      var m2_pct = m2_denom ? (m2 / m2_denom * 100) : null;

      resultadosMensais.push({
        mesReferencia: anoMês,
        ativo: ativo,
        data: {
          m1: m1,
          m2: m2_pct,
          numeradorM1: numerador,
          denominadorM1: denominador,
          numeradorM2: m2,
          denominadorM2: m2_denom,
          atendimentos: numerador,
          pessoasAtendidas: denominador,
          acoesCompartilhadas: m2,
          acoesRealizadas: m2_denom
        },
        pessoasAtendidas: {headers: ["Nome","Atendimentos","Participantes Ativ. Coletiva","Total"], rows: []}
      });
    });

    function media(campo){
      var vals = resultadosJanela.map(function(r){ return r.data[campo]; }).filter(function(v){ return v!=null; });
      if(!vals.length) return null;
      return vals.reduce(function(a,b){ return a+b; }, 0) / vals.length;
    }
    function soma(campo){
      return resultadosMensais.reduce(function(a,r){ return a + (r.data[campo]||0); }, 0);
    }
    function mediaJanela(campo){
      var vals = resultadosJanela.map(function(r){ return r.data[campo]; }).filter(function(v){ return v!=null; });
      if(!vals.length) return null;
      var media = vals.reduce(function(a,b){ return a+b; }, 0) / vals.length;
      return Math.round(media);
    }
    var m1 = media('m1');
    var m2 = media('m2');
    var classificacaoM1 = classificarM1(m1);
    var classificacaoM2 = classificarM2(m2);
    var pontosM1 = PONTOS_POR_CLASSE[classificacaoM1];
    var pontosM2 = PONTOS_POR_CLASSE[classificacaoM2];
    var pontosM1Pesados = pontosM1!==undefined ? pontosM1*6 : null;
    var pontosM2Pesados = pontosM2!==undefined ? pontosM2*4 : null;
    var notaFinal = (pontosM1Pesados!=null && pontosM2Pesados!=null) ? (pontosM1Pesados+pontosM2Pesados) : null;
    var desempenho = classificarDesempenho(notaFinal);

    var pessoasMap = {};
    resultadosMensais.forEach(function(r){
      r.pessoasAtendidas.rows.forEach(function(row){
        var chave = String(row[0]).trim().toUpperCase();
        if(!pessoasMap[chave]) pessoasMap[chave] = {nome:row[0], at:0, part:0};
        pessoasMap[chave].at += row[1];
        pessoasMap[chave].part += row[2];
      });
    });
    var pessoasLista = Object.keys(pessoasMap).map(function(k){ return pessoasMap[k]; })
      .sort(function(a,b){ return b.at+b.part - (a.at+a.part); });
    var pessoasAtendidasRows = pessoasLista.map(function(p){ return [p.nome, p.at, p.part, p.at+p.part]; });

    return {
      periodo: record.periodo,
      m1: m1,
      m2: m2,
      classificacaoM1: classificacaoM1,
      classificacaoM2: classificacaoM2,
      desempenho: desempenho,
      contexto: {
        numeradorM1: mediaJanela('numeradorM1'),
        denominadorM1: mediaJanela('denominadorM1'),
        numeradorM2: mediaJanela('numeradorM2'),
        denominadorM2: mediaJanela('denominadorM2'),
        numeradorM1Janela: mediaJanela('numeradorM1'),
        denominadorM1Janela: mediaJanela('denominadorM1'),
        numeradorM2Janela: mediaJanela('numeradorM2'),
        denominadorM2Janela: mediaJanela('denominadorM2'),
        m2: m2,
        classificacaoM2: classificacaoM2,
        pontosM1: pontosM1Pesados,
        pontosM2: pontosM2Pesados,
        notaFinal: notaFinal,
        desempenho: desempenho
      },
      notes: NOTAS_METODOLOGICAS,
      pessoasAtendidas: {headers: ["Nome","Atendimentos","Participantes Ativ. Coletiva","Total"], rows: pessoasAtendidasRows}
    };
  }
  function sheetCsvUrl(sheetName){
    return "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID
      + "/gviz/tq?tqx=out:csv&sheet=" + encodeURIComponent(sheetName);
  }
  function fetchAllSheets(){
    return Promise.all(requiredSheetNames().map(function(name){
      return fetch(sheetCsvUrl(name), {cache:'no-store'})
        .then(function(res){
          if(!res.ok) throw new Error('HTTP ' + res.status);
          return res.text();
        })
        .then(function(csvText){ return {name:name, csvText:csvText, ok:true}; })
        .catch(function(err){ return {name:name, error:err, ok:false}; });
    }));
  }

  var CLASS_PILL_HEX = {"Ótimo":"#2F6F5E","Bom":"#6B8F71","Suficiente":"#C68A3D","Regular":"#B5474B"};
  var CLASS_ARC_HEX = {"Regular":"#DC4C4C","Suficiente":"#F2A93B","Bom":"#4CAF6D","Ótimo":"#3B7DDD"};
  var CLASS_ARC_HEX_OV = {"Regular":"#BF2929","Suficiente":"#CF7E09","Bom":"#2A894A","Ótimo":"#1B59B5"};

  function m1ListNames(){ return ["Atendimentos", "Participantes Ativ. Coletiva", "Pessoas atendidas"].map(suffixedName); }
  function m2ListNames(){ return ["Atendimentos", "Resumo Reuniões", "Resumo Atividade Coletiva"].map(suffixedName); }
  var latestSheets = {}; 
  var listDateColIdx = {};
  var listMonthFilters = {};
  
  function pillHex(c){ return CLASS_PILL_HEX[c] || "#9AA69E"; }
  function arcHex(c){ return CLASS_ARC_HEX[c] || "#9AA69E"; }
  function arcHexOv(c){ return CLASS_ARC_HEX_OV[c] || "#9AA69E"; }
  function suffixedName(name){ return name + REFERENCIA_SUFIXO; }

  function displayListName(name){
    return name.replace(REFERENCIA_SUFIXO, '').replace(/_/g, ' ');
  }

  function parseBRDate(s){
    if(!s) return null;
    var m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if(!m) return null;
    return new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10));
  }
  function fmtInt(n){ return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

  function renderListCard(name){
    var cached = latestSheets[name] || {error:'Sem dados'};
    var hasTable = cached.rows && cached.rows.length > 0;
    var body = '';
    if(cached.error){
      body = '<div class="list-placeholder">'+escapeHtml(cached.error)+'</div>';
    } else if(hasTable){
      var dateColIdx = -1;
      cached.headers.forEach(function(h, i){
        if(String(h).toLowerCase().indexOf('data')>=0 || String(h).toLowerCase().indexOf('date')>=0){
          dateColIdx = i;
        }
      });
      listDateColIdx[name] = dateColIdx;
      
      var monthFilterHtml = '';
      if(dateColIdx >= 0){
        var optsCalc = [];
        cached.rows.forEach(function(row){
          var raw = row[dateColIdx];
          var d = parseBRDate(raw);
          var mv = d ? monthOptionValue(d) : null;
          if(mv && !optsCalc.find(function(o){ return o.value===mv; })){
            optsCalc.push({label:d.toLocaleDateString('pt-BR',{month:'2-digit',year:'numeric'}),value:mv});
          }
        });
        if(optsCalc.length){
          monthFilterHtml = '<div class="list-month-filter">'
            + '<span class="list-month-filter-label">Mês:</span>'
            + '<div class="list-month-filter-label list-month-filter-label" data-month-filter="'+escapeHtml(name)+'" data-computed-months="0"></div>'
            + '</div>';
        }
      }
      
      var filterPairsHtml = '<div class="filter-pair"></div>';
      var theadHtml = cached.headers.map(function(h){ return '<th>'+escapeHtml(h)+'</th>'; }).join('');
      var bodyHtml = cached.rows.map(function(row){
        return '<tr>' + row.map(function(cell){ return '<td>'+escapeHtml(cell)+'</td>'; }).join('') + '</tr>';
      }).join('');
      
      body = '<p class="list-meta">'+fmtInt(cached.rows.length)+' linhas</p>'
        + (monthFilterHtml||'') 
        + '<div class="list-filters" data-list-filters="'+escapeHtml(name)+'">'+filterPairsHtml+'</div>'
        + '<input class="list-search" type="text" placeholder="Filtrar nesta lista…" data-filter-key="'+escapeHtml(name)+'">'
        + '<div class="table-wrap"><table class="data-table"><thead>'+theadHtml+'</thead><tbody>'+bodyHtml+'</tbody></table></div>';
    }
    var pdfBtnHtml = hasTable
      ? '<button type="button" class="pdf-btn" data-pdf-btn="'+escapeHtml(name)+'">'
        + '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M9 15h1a1.5 1.5 0 0 0 0-3H9v5"/><path d="M13 12v5h1a2 2 0 0 0 0-5z"/><path d="M18.5 12H17v5"/><path d="M17 14.5h1.3"/></svg>'
        + '<span>Gerar PDF</span></button>'
      : '';
    return '<div class="card list-card" data-list-card="'+escapeHtml(name)+'">'
      + '<div class="list-card-head"><h4>'+escapeHtml(displayListName(name))+'</h4>'+pdfBtnHtml+'</div>'
      + body + '</div>';
  }

  function renderListsSection(containerId, names){
    var el = document.getElementById(containerId);
    if(!el) return;
    
    // Render as horizontal pills container
    var pillsHtml = '<div class="sheets-pills-container">'
      + names.map(function(name){
          return '<button class="sheet-pill" data-sheet-name="'+escapeHtml(name)+'">'
            + escapeHtml(displayListName(name))
            + '</button>';
        }).join('')
      + '</div>';
    
    var cardsHtml = names.map(renderListCard).join('');
    el.innerHTML = pillsHtml + cardsHtml;
    
    // Add event listeners to pills
    el.querySelectorAll('.sheet-pill').forEach(function(pill, idx){
      var sheetName = pill.getAttribute('data-sheet-name');
      if(idx === 0) pill.classList.add('active');
      
      pill.addEventListener('click', function(){
        el.querySelectorAll('.sheet-pill').forEach(function(p){ p.classList.remove('active'); });
        pill.classList.add('active');
        
        el.querySelectorAll('.list-card').forEach(function(card){ card.style.display = 'none'; });
        var targetCard = el.querySelector('[data-list-card="'+escapeHtml(sheetName)+'"]');
        if(targetCard) targetCard.style.display = 'block';
      });
    });
    
    // Hide all cards except first
    el.querySelectorAll('.list-card').forEach(function(card, idx){
      card.style.display = idx === 0 ? 'block' : 'none';
    });

    function applyFilters(card){
      var listName = card.getAttribute('data-list-card');
      var cached = latestSheets[listName];
      var dateColIdx = listDateColIdx[listName];
      var selectedMonths = listMonthFilters[listName] || [];
      var textInput = card.querySelector('.list-search');
      var term = textInput ? textInput.value.trim().toLowerCase() : '';
      var activeFilters = [];
      card.querySelectorAll('.filter-pair').forEach(function(pair){
        var colSelect = pair.querySelector('.filter-col');
        var valWrap = pair.querySelector('.filter-val-ms');
        var colIdx = colSelect && colSelect.value !== '' ? parseInt(colSelect.value, 10) : null;
        var vals = (valWrap && valWrap._msInstance) ? valWrap._msInstance.getSelected() : [];
        if(colIdx !== null && vals.length){ activeFilters.push({colIdx:colIdx, vals:vals}); }
      });
      var visibleCount = 0;
      card.querySelectorAll('tbody tr').forEach(function(tr, rowIdx){
        var matchesText = !term || tr.textContent.toLowerCase().indexOf(term) !== -1;
        var matchesCols = activeFilters.every(function(f){
          var cell = tr.children[f.colIdx];
          return cell && f.vals.indexOf(cell.textContent.trim()) >= 0;
        });
        var matchesMonth = true;
        if(selectedMonths.length && dateColIdx != null && dateColIdx >= 0){
          var raw = cached && cached.rows[rowIdx] ? cached.rows[rowIdx][dateColIdx] : null;
          var d = parseBRDate(raw);
          var mv = d ? monthOptionValue(d) : null;
          matchesMonth = !!mv && selectedMonths.indexOf(mv) >= 0;
        }
        var visible = matchesText && matchesCols && matchesMonth;
        tr.style.display = visible ? '' : 'none';
        if(visible) visibleCount++;
      });
      var metaEl = card.querySelector('.list-meta');
      if(metaEl) metaEl.textContent = fmtInt(visibleCount) + (visibleCount === 1 ? ' linha' : ' linhas');
    }

    el.querySelectorAll('[data-month-filter]').forEach(function(container){
      var name = container.getAttribute('data-month-filter');
      var optsCalc = [];
      if(latestSheets[name] && latestSheets[name].rows){
        latestSheets[name].rows.forEach(function(row){
          var dateColIdx = listDateColIdx[name];
          if(dateColIdx >= 0){
            var raw = row[dateColIdx];
            var d = parseBRDate(raw);
            var mv = d ? monthOptionValue(d) : null;
            if(mv && !optsCalc.find(function(o){ return o.value===mv; })){
              optsCalc.push({label:d.toLocaleDateString('pt-BR',{month:'2-digit',year:'numeric'}),value:mv});
            }
          }
        });
      }
      listMonthFilters[name] = [];
    });

    el.querySelectorAll('.list-search').forEach(function(input){
      input.addEventListener('keyup', function(){
        var card = input.closest('.list-card');
        applyFilters(card);
      });
    });
  }

  function renderDashboard(record, serie){
    document.getElementById('topEquipe').textContent = record.equipe || '—';
    document.getElementById('topPeriodo').textContent = record.periodo || '—';
    document.getElementById('topUpdated').textContent = 'Atualizado ' + new Date(record.timestamp).toLocaleDateString('pt-BR');

    var dash = processDashboard(record, serie);
    
    renderListsSection('listsM1', m1ListNames());
    renderListsSection('listsM2', m2ListNames());
  }

  var currentRecordId = null;
  var currentEquipes = EQUIPES;
  var latestWb = null;
  var refreshBtn = document.getElementById('refreshBtn');
  var refreshLabel = document.getElementById('refreshLabel');
  var fetchStatusEl = document.getElementById('statusState') || {};

  function saveHistoryArray(arr){
    try {
      localStorage.setItem('eMulti_history', JSON.stringify(arr));
      return Promise.resolve();
    } catch(e){
      return Promise.reject(e);
    }
  }

  function loadHistoryArray(){
    try {
      var s = localStorage.getItem('eMulti_history') || '[]';
      return JSON.parse(s);
    } catch(e){
      return [];
    }
  }

  function addToHistory(equipe, data, notes, periodo){
    var now = Date.now();
    var record = {
      id: 'u'+now+Math.random().toString(36).slice(2,7),
      timestamp: now,
      equipe: equipe,
      data: data,
      notes: notes,
      periodo: periodo
    };
    var arr = loadHistoryArray();
    arr.push(record);
    saveHistoryArray(arr).then(function(){
      currentRecordId = record.id;
      fetchStatusEl.textContent = 'Planilha lida e calculada com sucesso.';
      renderDashboard(record, null);
    });
  }

  function fetchAndLoad(){
    refreshBtn.classList.add('loading');
    refreshBtn.disabled = true;
    refreshLabel.textContent = 'Atualizando…';
    fetchStatusEl.textContent = 'Buscando dados…';
    fetchStatusEl.className = 'fetch-status';

    Promise.all([fetchAllSheets()])
      .then(function(arr){
        var results = arr[0];
        var faltando = results.filter(function(r){ return !r.ok; });
        if(faltando.length){
          throw new Error('Não foi possível ler a(s) aba(s) "' + faltando.map(function(r){return r.name;}).join('", "')
            + '" (verifique se elas ainda existem com esse nome e se a planilha está com acesso "qualquer pessoa com o link pode visualizar").');
        }

        var wb = {SheetNames:[], Sheets:{}};
        results.forEach(function(r){
          var parsedRows = parseCsv(r.csvText);
          if(!parsedRows.length) return;
          var key = suffixedName(r.name);
          wb.Sheets[key] = {headers: parsedRows[0]||[], rows: parsedRows.slice(1)};
          wb.SheetNames.push(key);
          latestSheets[key] = wb.Sheets[key];
        });

        latestWb = wb;
        var record = {
          equipe: currentEquipes[0],
          data: [],
          notes: '',
          periodo: 'Q4/2024'
        };
        renderDashboard(record, null);
      })
      .catch(function(err){
        var msg = (err && err.message) ? err.message : 'verifique sua conexão e o link publicado.';
        fetchStatusEl.textContent = 'Não foi possível ler a planilha: ' + msg;
        fetchStatusEl.className = 'fetch-status err';
      })
      .finally(function(){
        refreshBtn.classList.remove('loading');
        refreshBtn.disabled = false;
        refreshLabel.textContent = 'Atualizar agora';
      });
  }

  function parseCsv(text){
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    for(var i=0; i<text.length; i++){
      var c = text[i];
      if(c === '"'){
        if(inQuotes && text[i+1] === '"'){ field += '"'; i++; }
        else { inQuotes = !inQuotes; }
      } else if(c === ',' && !inQuotes){
        row.push(field);
        field = '';
      } else if((c === '\n' || c === '\r') && !inQuotes){
        if(field.length || row.length){ row.push(field); rows.push(row); row=[]; field=''; }
        if(c === '\r' && text[i+1] === '\n') i++;
      } else {
        field += c;
      }
    }
    if(field.length || row.length){ row.push(field); rows.push(row); }
    return rows;
  }

  document.addEventListener('DOMContentLoaded', function(){
    var tabs = document.querySelectorAll('.tab');
    tabs.forEach(function(tab){
      tab.addEventListener('click', function(){
        tabs.forEach(function(t){ t.classList.remove('active'); });
        tab.classList.add('active');
        
        var tabName = tab.getAttribute('data-tab');
        var panels = document.querySelectorAll('.tab-panel');
        panels.forEach(function(p){ p.classList.remove('active'); });
        var panel = document.getElementById('tab'+tabName.charAt(0).toUpperCase()+tabName.slice(1));
        if(panel) panel.classList.add('active');
      });
    });

    if(refreshBtn){
      refreshBtn.addEventListener('click', fetchAndLoad);
    }
    
    fetchAndLoad();
  });
