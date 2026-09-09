(function(){
  "use strict";

  // ID real da planilha (do link "Compartilhar", não do "Publicar na web").
  // O link antigo de "Publicar na web" só expunha 1 aba mesmo pedindo xlsx
  // (o Google ignorava o parâmetro), o que quebrava o fetch no navegador.
  // Agora buscamos cada aba separadamente via endpoint gviz/tq (CSV com
  // suporte a CORS de verdade), que funciona para qualquer aba por nome.
  var SPREADSHEET_ID = "1ujHEI_pERAcKmxQRuF9AmgU0bN22w0A9nnpakCVjY18";
  // Janela móvel usada SÓ pela aba "Tendência" (mês a mês): cada ponto do
  // gráfico é o M1/M2 calculado com uma janela de JANELA_MESES meses
  // terminando naquele mês. Mude só este número se quiser 3, 4 ou 6 meses
  // de janela.
  var JANELA_MESES = 4;
  // Quantos pontos (meses) mostrar nos gráficos de tendência — cada ponto
  // é o M1/M2 daquele mês, já calculado com sua própria janela de
  // JANELA_MESES meses terminando naquele mês.
  var TREND_MESES = 8;
  // ---- Filtro principal da Visão geral: Quadrimestre + Mês (opcional) ----
  // Quadrimestres fixos do ano civil: Q1 Jan–Abr, Q2 Mai–Ago, Q3 Set–Dez.
  var QUAD_LABELS = ['Jan–Abr (Q1)', 'Mai–Ago (Q2)', 'Set–Dez (Q3)'];
  // Quadrimestre selecionado (ano + índice 0/1/2). Começa no quadrimestre
  // que contém o mês atual; muda quando o usuário mexe no filtro de Quadrimestre.
  var quadSelecionado = quadrimestreDoMes(new Date());
  // Data(s) escolhida(s) no filtro de Mês. Array vazio = padrão => usa a
  // MÉDIA dos 4 meses do quadrimestre selecionado. Um ou mais meses
  // marcados: cada mês entra com o SEU PRÓPRIO resultado (já calculado
  // com a janela móvel de JANELA_MESES meses terminando nele — ver
  // calcularJanelaPeriodo) e, havendo mais de um, os resultados são
  // combinados pela média (mesma lógica já usada pra média do
  // quadrimestre, ver mediaDeMeses).
  var refMonthDates = [];
  function quadrimestreDoMes(d){
    return {ano: d.getFullYear(), qIndex: Math.floor(d.getMonth()/4)};
  }
  // Chave/rótulo do quadrimestre de um mês, usados pra agrupar os pontos
  // do gráfico de Tendência e desenhar a linha de média de cada
  // quadrimestre (ver sparkline).
  function quadKeyOfDate(d){
    var q = quadrimestreDoMes(d);
    return q.ano + '-' + q.qIndex;
  }
  function quadShortLabel(d){
    var q = quadrimestreDoMes(d);
    var base = QUAD_LABELS[q.qIndex].replace(/\s*\(Q\d\)/, '');
    return base + '/' + String(q.ano).slice(2);
  }
  // Código curto (Q1/Q2/Q3) usado só dentro do gráfico de tendência, onde
  // o espaço é pequeno — o rótulo completo (quadShortLabel) fica só como
  // referência textual fora do SVG.
  function quadCode(d){
    return 'Q' + (quadrimestreDoMes(d).qIndex + 1);
  }
  // Os 4 meses (dia 1 de cada) que compõem um quadrimestre.
  function mesesDoQuadrimestre(ano, qIndex){
    var meses = [];
    for(var i=0; i<4; i++){ meses.push(new Date(ano, qIndex*4+i, 1)); }
    return meses;
  }
  // Período de um único mês (do dia 1 ao último dia do mesmo mês).
  function periodoMesUnico(d){
    var inicio = new Date(d.getFullYear(), d.getMonth(), 1, 0,0,0,0);
    var fim = new Date(d.getFullYear(), d.getMonth()+1, 0, 23,59,59,999);
    return {inicio: inicio, fim: fim};
  }
  // Mês "âncora" usado pela aba Tendência: o mais recente dos meses
  // escolhidos, se houver algum, ou o último mês do quadrimestre
  // selecionado.
  function anchorMonthDate(){
    if(refMonthDates.length) return refMonthDates[refMonthDates.length-1];
    return new Date(quadSelecionado.ano, quadSelecionado.qIndex*4+3, 1);
  }
  // Só as abas de dados BRUTOS — o painel calcula M1/M2 sozinho a partir
  // delas (não lê mais nenhum valor pronto da aba "Indicadores M1 e M2").
  // IMPORTANTE: não existem abas separadas por equipe na planilha — os
  // nomes abaixo são os nomes REAIS das abas (conferidos direto no rodapé
  // do Google Sheets). O filtro por equipe acontece linha a linha, pela
  // coluna "equipe_unidade" de cada aba (ver filtrarLinhasPorEquipe).
  var BASE_SHEET_NAMES = [
    "Atendimentos",
    "Participantes Ativ. Coletiva",
    "Resumo Atividade Coletiva",
    "Resumo Reuniões"
  ];
  // matchKeyword: trecho (sem acento, maiúsculo) que precisa aparecer no
  // valor da coluna "equipe_unidade" pra a linha pertencer a esta equipe.
  // Ex.: "EMULTI CROATA DOS MARTINS - Croata" e "EMULTI CROATA DOS
  // MARTINS" (formatos variam entre abas) casam com "CROATA".
  var EQUIPES = [
    {key:"centro", label:"EMULTI Centro", suffix:"Centro", matchKeyword:"CENTRO"},
    {key:"croata", label:"EMULTI Croatá", suffix:"Croatá", matchKeyword:"CROATA"}
  ];
  // Agora suporta seleção múltipla: quando mais de uma equipe está
  // marcada, as linhas de AMBAS entram no cálculo (resultado combinado/
  // somado das equipes selecionadas). Sempre fica pelo menos 1 marcada.
  var currentEquipes = [EQUIPES[0]];

  function suffixedName(baseName){
    return baseName + " — " + currentEquipes.map(function(e){ return e.suffix; }).join('+');
  }
  function displayListName(name){
    // remove o sufixo " — Centro"/" — Croatá" só pra exibição (o título da
    // equipe já aparece no topo da página). Esse sufixo agora é só uma
    // CHAVE INTERNA de cache (ver wb.Sheets) — não é mais o nome real da
    // aba buscada no Google.
    return name.replace(/ — .+$/, '');
  }
  function requiredSheetNames(){
    // Nomes REAIS das abas — sem sufixo de equipe (ver comentário acima
    // de BASE_SHEET_NAMES).
    return BASE_SHEET_NAMES.slice();
  }
  function startOfMonth(d){ return new Date(d.getFullYear(), d.getMonth(), 1); }
  function addMonths(d, n){ return new Date(d.getFullYear(), d.getMonth()+n, 1); }
  function monthOptionValue(d){ return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0'); }
  function monthOptionLabel(d){
    var s = d.toLocaleDateString('pt-BR', {month:'long', year:'numeric'});
    return s.charAt(0).toUpperCase() + s.slice(1);
  }
  // refMonthDates vazio: usuário está vendo a MÉDIA do quadrimestre
  // (nenhum mês específico marcado). Com 1+ meses marcados, mostra o(s)
  // mês(es) escolhido(s) — este helper monta o rótulo certo pros dois
  // casos (usado nos lugares que exibem "Mês de referência (...)").
  function refMonthLabel(){
    if(!refMonthDates.length) return 'Média do quadrimestre';
    if(refMonthDates.length === 1) return monthOptionLabel(refMonthDates[0]);
    return refMonthDates.map(monthShortLabel).join(' + ');
  }
  function monthShortLabel(d){
    var s = d.toLocaleDateString('pt-BR', {month:'short'}).replace('.', '');
    return s.charAt(0).toUpperCase() + s.slice(1) + '/' + String(d.getFullYear()).slice(2);
  }
  // Janela móvel de 4 meses (JANELA_MESES) TERMINANDO no mês de referência
  // informado (ex.: referência = maio → janela = fev a maio, incluindo os
  // dois extremos). "refMonth" é sempre o dia 1 do mês.
  function calcularJanelaPeriodo(refMonth){
    var fim = new Date(refMonth.getFullYear(), refMonth.getMonth()+1, 0, 23,59,59,999); // último dia do mês de referência
    var inicioMes = addMonths(refMonth, -(JANELA_MESES-1));
    var inicio = new Date(inicioMes.getFullYear(), inicioMes.getMonth(), 1, 0,0,0,0);
    return {inicio: inicio, fim: fim};
  }
  // Série pra tendência: um ponto por mês (os últimos TREND_MESES meses,
  // terminando no mês de referência selecionado), cada um com sua PRÓPRIA
  // janela móvel de JANELA_MESES meses (não é o mesmo período repetido).
  function calcularSerieTendencia(wb, refMonth, n){
    var pontos = [];
    for(var i=n-1; i>=0; i--){
      var mes = addMonths(refMonth, -i);
      var janela = calcularJanelaPeriodo(mes);
      var res = calcularIndicadoresDoPeriodo(wb, janela);
      pontos.push({
        mes:mes, m1:res.data.m1, m2:res.data.m2,
        numeradorM1: res.data.numeradorM1, denominadorM1: res.data.denominadorM1,
        numeradorM2: res.data.numeradorM2, denominadorM2: res.data.denominadorM2,
        notaFinal: res.data.notaFinal,
        janela:janela
      });
    }
    return pontos;
  }
  // Popula o #quadMs com quadrimestres do ano atual e dos 2 anteriores
  // (mais recente primeiro), e o #mesMs com os 4 meses do quadrimestre
  // atualmente selecionado + uma opção vazia ("média"). Os dois são
  // widgets de valor único (multi:false) com o mesmo visual arredondado
  // do seletor de Equipe.
  var quadMs = null, mesMs = null;
  function populateQuadSelect(){
    var container = document.getElementById('quadMs');
    if(!container || quadMs) return; // já populado (não recria a cada render)
    quadMs = createMultiSelect(container, {
      placeholder: 'Selecione', multi: false, search: false,
      onChange: function(keys){
        var parts = keys[0].split('-');
        quadSelecionado = {ano: +parts[0], qIndex: +parts[1]};
        refMonthDates = []; // volta a mostrar a média do quadrimestre escolhido
        populateMonthSelectForQuad();
        aplicarMesReferencia(false);
      }
    });
    var anoAtual = new Date().getFullYear();
    var opts = [];
    for(var ano=anoAtual; ano>=anoAtual-2; ano--){
      for(var q=2; q>=0; q--){
        if(ano===anoAtual && q > quadSelecionado.qIndex) continue; // não mostra quadrimestre futuro do ano atual
        opts.push({value: ano+'-'+q, label: QUAD_LABELS[q]+'/'+ano});
      }
    }
    quadMs.setOptions(opts);
    quadMs.setSelected([quadSelecionado.ano+'-'+quadSelecionado.qIndex]);
    populateMonthSelectForQuad();
  }
  // Preenche #mesMs com os 4 meses do quadrimestre selecionado — agora em
  // multisseleção: marcar 1+ meses troca o resultado pro(s) mês(es)
  // escolhido(s) (cada um com sua janela móvel própria, combinados pela
  // média quando há mais de um); nenhum marcado = média do quadrimestre
  // inteiro. As opções são refeitas toda vez que o quadrimestre muda; o
  // widget em si (mesMs) é criado uma única vez.
  function populateMonthSelectForQuad(){
    var container = document.getElementById('mesMs');
    if(!container) return;
    if(!mesMs){
      mesMs = createMultiSelect(container, {
        placeholder: 'Média do quadrimestre', multi: true, search: false, showTags: true,
        onChange: function(keys){
          refMonthDates = keys.map(function(v){
            var parts = v.split('-');
            return new Date(+parts[0], +parts[1]-1, 1);
          }).sort(function(a,b){ return a-b; });
          aplicarMesReferencia(false);
        }
      });
    }
    var meses = mesesDoQuadrimestre(quadSelecionado.ano, quadSelecionado.qIndex);
    var mesesValidos = meses.map(monthOptionValue);
    var opts = meses.map(function(d){
      return {value: monthOptionValue(d), label: monthOptionLabel(d)};
    });
    // Ao trocar de quadrimestre, mantém só a seleção que ainda faz parte
    // do novo quadrimestre (evita "mês fantasma" de outro período).
    refMonthDates = refMonthDates.filter(function(d){ return mesesValidos.indexOf(monthOptionValue(d)) >= 0; });
    mesMs.setOptions(opts);
    mesMs.setSelected(refMonthDates.map(monthOptionValue));
  }
  // Combina os indicadores de vários meses fazendo a MÉDIA de M1 e M2 —
  // é assim que o quadrimestre vira "a média dos meses do quadrimestre".
  // IMPORTANTE: o M1/M2 de CADA mês que entra nessa média já é o valor
  // "oficial" daquele mês/competência, ou seja, calculado com a janela
  // móvel de JANELA_MESES meses terminando naquele mês (ex.: M1 de maio =
  // fev+mar+abr+maio) — é por isso que a função recebe DOIS conjuntos de
  // resultados: `resultadosJanela` (M1/M2 de cada mês já com a janela
  // móvel, usados para a média) e `resultadosMensais` (dados BRUTOS só
  // daquele mês isolado, sem janela, usados apenas pra somar contagens de
  // contexto — atendimentos, participações etc. — e montar a lista de
  // "Pessoas atendidas" sem contar o mesmo atendimento mais de uma vez).
  function mediaDeMeses(resultadosMensais, resultadosJanela){
    function media(campo){
      var vals = resultadosJanela.map(function(r){ return r.data[campo]; }).filter(function(v){ return v!=null; });
      if(!vals.length) return null;
      return vals.reduce(function(a,b){ return a+b; }, 0) / vals.length;
    }
    function soma(campo){
      return resultadosMensais.reduce(function(a,r){ return a + (r.data[campo]||0); }, 0);
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

    // "Pessoas atendidas": une as listas dos 4 meses, somando atendimentos
    // e participações de quem aparece em mais de um mês.
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
      .sort(function(a,b){ return a.nome.localeCompare(b.nome,'pt-BR'); });
    var pessoasAtendidasRows = pessoasLista.map(function(p){ return [p.nome, p.at, p.part, p.at+p.part]; });

    return {
      equipe: currentEquipes.map(function(e){ return e.label; }).join(' + '),
      data: {
        atendimentosIndividuais: soma('atendimentosIndividuais'),
        participacoesColetivas: soma('participacoesColetivas'),
        numeradorM1: soma('numeradorM1'),
        denominadorM1: soma('denominadorM1'),
        m1: m1,
        classificacaoM1: classificacaoM1,
        atividadesTotais: soma('atividadesTotais'),
        atividadesCompartilhadas: soma('atividadesCompartilhadas'),
        reunioesTotais: soma('reunioesTotais'),
        reunioesCompartilhadas: soma('reunioesCompartilhadas'),
        denominadorM2: soma('denominadorM2'),
        numeradorM2: soma('numeradorM2'),
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

  // ---------- Listas complementares ----------
  function m1ListNames(){ return ["Atendimentos", "Participantes Ativ. Coletiva", "Pessoas atendidas"].map(suffixedName); }
  function m2ListNames(){ return ["Atendimentos", "Resumo Reuniões", "Resumo Atividade Coletiva"].map(suffixedName); }
  var latestSheets = {}; // nome da aba -> {headers, rows} | {error}
  // Filtro de mês (multisseleção) das listas das abas M1/M2: por lista
  // (chave = nome sufixado da aba), guarda o índice da coluna de data
  // encontrada e os meses atualmente marcados (["" ] vazio = todos os
  // meses). Persistem entre re-renders pra não perder a seleção do
  // usuário a cada atualização dos dados.
  var listDateColIdx = {};
  var listMonthFilters = {};
  // Acha a coluna de data de uma lista bruta, testando os nomes usados
  // nas abas de origem ("data" na maioria, "data_hora" em Atendimentos).
  function dateColIndexForList(headers){
    var idx = colIndex(headers, "data_hora");
    if(idx >= 0) return idx;
    return colIndex(headers, "data");
  }
  // Monta as opções de mês (mais recente primeiro) a partir dos valores
  // de data realmente presentes nas linhas da lista.
  function monthOptionsForList(cached, dateColIdx){
    var seen = {}, months = [];
    cached.rows.forEach(function(r){
      var d = parseBRDate(r[dateColIdx]);
      if(!d) return;
      var v = monthOptionValue(d);
      if(!seen[v]){ seen[v] = true; months.push(new Date(d.getFullYear(), d.getMonth(), 1)); }
    });
    months.sort(function(a,b){ return b-a; });
    return months.map(function(d){ return {value: monthOptionValue(d), label: monthOptionLabel(d)}; });
  }

  var STORAGE_KEY = "uploads";
  var currentRecordId = null;
  var STORAGE_AVAILABLE = !!(window.storage && typeof window.storage.get === 'function'
    && typeof window.storage.set === 'function');
  var memoryHistory = [];

  // ---------- Helpers ----------
  function fmtInt(v){
    if(v===null||v===undefined||isNaN(v)) return "—";
    return Number(v).toLocaleString('pt-BR');
  }
  function fmtDec(v,d){
    if(v===null||v===undefined||isNaN(v)) return "—";
    return Number(v).toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d});
  }
  function fmtDate(ts){
    var d = new Date(ts);
    return d.toLocaleDateString('pt-BR') + " às " + d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
  }
  function shortDate(ts){
    var d = new Date(ts);
    return d.toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit'});
  }
  function escapeHtml(s){
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }
  function pillHex(c){ return CLASS_PILL_HEX[c] || "#9AA69E"; }
  function arcHex(c){ return CLASS_ARC_HEX[c] || "#9AA69E"; }

  // ---------- Multi-select arredondado (Equipe / Quadrimestre / Mês) ----------
  // Componente genérico: em modo multi:true permite marcar vários valores
  // (com "Selecionar tudo"/"Limpar" e tags abaixo do botão); em modo
  // multi:false funciona como um "select" de valor único, mas com o
  // mesmo visual arredondado — clicar numa opção troca a seleção e fecha.
  function createMultiSelect(container, cfg){
    cfg = cfg || {};
    var state = {options: [], selected: [], isOpen: false, searchTerm: ''};
    container.innerHTML =
        '<button type="button" class="ms-btn">'
      +   '<span class="ms-btn-text"></span>'
      +   '<svg class="ms-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>'
      + '</button>'
      + '<div class="ms-panel" style="display:none;"></div>'
      + (cfg.showTags ? '<div class="ms-tags"></div>' : '');
    var btn = container.querySelector('.ms-btn');
    var btnText = container.querySelector('.ms-btn-text');
    var panel = container.querySelector('.ms-panel');
    var tagsBox = container.querySelector('.ms-tags');

    function labelFor(value){
      var found = state.options.filter(function(o){ return o.value===value; })[0];
      return found ? found.label : value;
    }

    function renderTags(){
      if(!tagsBox) return;
      if(!cfg.multi || state.selected.length<2){ tagsBox.innerHTML=''; return; }
      tagsBox.innerHTML = state.selected.map(function(v){
        return '<span class="ms-tag" data-value="'+escapeHtml(v)+'">'+escapeHtml(labelFor(v))
          + '<button type="button" data-remove="'+escapeHtml(v)+'">'
          +   '<svg viewBox="0 0 24 24" width="11" height="11" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M18 6L6 18M6 6l12 12"/></svg>'
          + '</button></span>';
      }).join('');
      tagsBox.querySelectorAll('[data-remove]').forEach(function(b){
        b.addEventListener('click', function(e){
          e.stopPropagation();
          setSelected(state.selected.filter(function(v){ return v!==b.getAttribute('data-remove'); }));
        });
      });
    }

    function render(){
      var selLabels = state.selected.map(labelFor);
      btnText.textContent = selLabels.length===0 ? (cfg.placeholder || 'Todos')
        : (cfg.multi && selLabels.length>1 ? selLabels.length+' selecionados' : selLabels.join(', '));
      container.classList.toggle('ms-has-value', selLabels.length>0);
      btn.classList.toggle('ms-open', state.isOpen);
      renderTags();

      if(!state.isOpen){ panel.style.display='none'; panel.innerHTML=''; return; }
      panel.style.display='block';

      var term = state.searchTerm.toLowerCase();
      var filtered = !term ? state.options : state.options.filter(function(o){
        return o.label.toLowerCase().indexOf(term) >= 0;
      });

      var html = '';
      if(cfg.search){
        html += '<div class="ms-search-wrap"><input type="text" class="ms-search" placeholder="Buscar…" value="'+escapeHtml(state.searchTerm)+'"></div>';
      }
      if(cfg.multi){
        html += state.selected.length>0
          ? '<button type="button" class="ms-action" data-action="clear">Limpar seleção</button>'
          : '<button type="button" class="ms-action" data-action="all">Selecionar tudo</button>';
      }
      html += '<div class="ms-list">';
      html += filtered.length===0
        ? '<div class="ms-empty">Nenhum resultado encontrado</div>'
        : filtered.map(function(o){
            var checked = state.selected.indexOf(o.value)>=0;
            return '<button type="button" class="ms-option'+(checked?' ms-option-checked':'')+'" data-value="'+escapeHtml(o.value)+'">'
              + '<span class="ms-option-label">'+escapeHtml(o.label)+'</span>'
              + (checked ? '<svg class="ms-check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>' : '')
              + '</button>';
          }).join('');
      html += '</div>';
      panel.innerHTML = html;

      var searchInput = panel.querySelector('.ms-search');
      if(searchInput){
        searchInput.focus();
        var pos = state.searchTerm.length;
        searchInput.setSelectionRange(pos,pos);
        searchInput.addEventListener('input', function(){ state.searchTerm = searchInput.value; render(); });
      }
      var actionBtn = panel.querySelector('.ms-action');
      if(actionBtn){
        actionBtn.addEventListener('click', function(){
          if(actionBtn.getAttribute('data-action')==='clear') setSelected([]);
          else setSelected(filtered.map(function(o){ return o.value; }));
        });
      }
      panel.querySelectorAll('.ms-option').forEach(function(elOpt){
        elOpt.addEventListener('click', function(){
          var v = elOpt.getAttribute('data-value');
          if(cfg.multi){
            var next = state.selected.indexOf(v)>=0
              ? state.selected.filter(function(x){ return x!==v; })
              : state.selected.concat([v]);
            setSelected(next);
          } else {
            state.isOpen = false;
            setSelected([v]);
          }
        });
      });
    }

    function setSelected(values, silent){
      state.selected = values;
      render();
      if(!silent && cfg.onChange) cfg.onChange(state.selected.slice());
    }

    btn.addEventListener('click', function(){
      state.isOpen = !state.isOpen;
      state.searchTerm = '';
      render();
    });
    document.addEventListener('mousedown', function(e){
      if(state.isOpen && !container.contains(e.target)){
        state.isOpen = false;
        state.searchTerm = '';
        render();
      }
    });

    return {
      setOptions: function(opts){ state.options = opts; render(); },
      setSelected: function(values){ setSelected(values, true); },
      getSelected: function(){ return state.selected.slice(); }
    };
  }

  // ---------- Parsing ----------
  function sheetToRows(ws){
    if(Array.isArray(ws)) return ws; // já é uma matriz de linhas (vindo do parseCsv)
    return XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
  }

  // ---------- Filtro por equipe (linha a linha) ----------
  // Não existem abas separadas por equipe — cada linha da aba tem uma
  // coluna "equipe_unidade" (ou similar) que identifica a equipe. Aqui a
  // gente acha essa coluna e mantém só as linhas da equipe selecionada.
  function normalizeText(s){
    return String(s||"").toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  }
  function equipeColIndex(headerRow){
    for(var i=0;i<headerRow.length;i++){
      var h = normalizeText(headerRow[i]).replace(/\s+/g,'_');
      if(h === "EQUIPE_UNIDADE") return i;
    }
    for(var j=0;j<headerRow.length;j++){
      if(normalizeText(headerRow[j]).indexOf("EQUIPE") !== -1) return j;
    }
    return -1;
  }
  function filtrarLinhasPorEquipe(matrix, equipes){
    if(!matrix || !matrix.length) return matrix || [];
    var header = matrix[0];
    var idx = equipeColIndex(header);
    if(idx < 0) return matrix; // aba sem coluna de equipe: não filtra
    var keywords = equipes.map(function(eq){ return normalizeText(eq.matchKeyword || eq.suffix); });
    var linhas = matrix.slice(1).filter(function(r){
      var valor = normalizeText(r[idx]);
      return keywords.some(function(kw){ return valor.indexOf(kw) !== -1; });
    });
    return [header].concat(linhas);
  }

  // Parser de CSV manual (RFC4180: respeita campos entre aspas, vírgulas e
  // quebras de linha dentro de campos). Usado em vez do XLSX.read(string)
  // porque a leitura automática de string do SheetJS não separava as
  // linhas corretamente para o CSV retornado pelo endpoint gviz.
  function parseCsv(text){
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    for(var i=0; i<text.length; i++){
      var c = text[i];
      if(inQuotes){
        if(c === '"'){
          if(text[i+1] === '"'){ field += '"'; i++; }
          else { inQuotes = false; }
        } else {
          field += c;
        }
      } else {
        if(c === '"'){ inQuotes = true; }
        else if(c === ','){ row.push(field); field=''; }
        else if(c === '\r'){ /* ignora, quebra tratada no \n */ }
        else if(c === '\n'){ row.push(field); field=''; rows.push(row); row=[]; }
        else { field += c; }
      }
    }
    if(field.length || row.length){ row.push(field); rows.push(row); }
    return rows;
  }

  // Datas nas abas brutas vêm como texto dd/mm/aaaa (é assim que o script
  // de extração grava). Também aceita aaaa-mm-dd como reforço, caso a
  // célula tenha sido digitada nesse formato.
  function parseBRDate(raw){
    var s = String(raw||"").trim();
    if(!s) return null;
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if(m) return new Date(+m[3], +m[2]-1, +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(m) return new Date(+m[1], +m[2]-1, +m[3]);
    return null;
  }
  function fmtBRDate(d){
    if(!d) return "—";
    return String(d.getDate()).padStart(2,'0') + "/" + String(d.getMonth()+1).padStart(2,'0') + "/" + d.getFullYear();
  }
  function withinPeriod(dateVal, inicio, fim){
    return dateVal && dateVal >= inicio && dateVal <= fim;
  }
  function colIndex(headerRow, name){
    for(var i=0;i<headerRow.length;i++){
      if(String(headerRow[i]||"").trim() === name) return i;
    }
    return -1;
  }
  function toInt(v){
    var n = parseInt(String(v===undefined||v===null?"":v).trim(), 10);
    return isNaN(n) ? 0 : n;
  }
  // Normaliza texto pra comparar tipo_atividade sem depender de acento,
  // maiúscula/minúscula ou espaço/barra diferente ("Avaliação/Procedimento
  // coletivo" vs "Avaliação / Procedimento Coletivo" etc.).
  function normalizarTexto(v){
    return String(v||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
      .toLowerCase().replace(/\s+/g," ").replace(/\s*\/\s*/g,"/").trim();
  }

  var PONTOS_POR_CLASSE = {"Regular":0.25, "Suficiente":0.5, "Bom":0.75, "Ótimo":1};
  // Mesmas cores dos "pills" de classificação (ver :root), usadas pra
  // colorir a linha/rótulo de média de cada quadrimestre no gráfico de
  // Tendência conforme a faixa em que a média cai.
  var CLASS_COLOR = {"Regular":"#B5474B", "Suficiente":"#C68A3D", "Bom":"#6B8F71", "Ótimo":"#2F6F5E"};

  function classificarM1(v){
    if(v===null) return "—";
    if(v>3) return "Ótimo";
    if(v>2) return "Bom";
    if(v>1) return "Suficiente";
    return "Regular";
  }
  function classificarM2(v){
    if(v===null) return "—";
    if(v>5) return "Ótimo";
    if(v>2.5) return "Bom";
    if(v>1) return "Suficiente";
    return "Regular";
  }
  function classificarDesempenho(nota){
    if(nota===null) return "—";
    if(nota>7.5) return "Ótimo";
    if(nota>=5) return "Bom";
    if(nota>=2.6) return "Suficiente";
    return "Regular";
  }

  var NOTAS_METODOLOGICAS = [
    "Cálculo feito pelo próprio painel, direto dos dados brutos extraídos do e-SUS PEC (Atendimentos + Registro Tardio + Atividade Coletiva + Reuniões) para esta equipe/EMULTI, seguindo as fórmulas das Notas Metodológicas M1 (NT 43/2026-CGIAD/DEAPS/SAPS/MS) e M2 (NT 44/2026-CGIAD/DEAPS/SAPS/MS), na janela dos últimos 4 meses (ver 'Período' no topo da página) — não um quadrimestre fixo do calendário.",
    "M1 usa NOME da pessoa (a nota oficial usa CPF/CNS) — pessoas diferentes com o mesmo nome seriam contadas como se fossem uma só.",
    "M2 oficial soma 3 componentes: atendimentos individuais compartilhados, atividades coletivas compartilhadas e compartilhamento de cuidado (PEC). Esta extração só consegue aproximar as parcelas de 'atividades coletivas' e 'reuniões', usando 'nº de profissionais envolvidos ≥ 2' como indício de ação compartilhada — não há como checar CBO/CNS de cada profissional (principal/secundário) pra aplicar a regra oficial à risca.",
    "Atendimentos individuais compartilhados e compartilhamento de cuidado (PEC) NÃO entram no numerador do M2 aqui (a Lista de Atendimentos do e-SUS não indica se um atendimento individual teve mais de um profissional) — por isso o M2 calculado aqui tende a ficar ABAIXO do valor oficial do indicador.",
    "Atividade Coletiva só conta como 'compartilhada' aqui quando o tipo_atividade é Educação em saúde, Atendimento em grupo, Avaliação/Procedimento coletivo ou Mobilização social (códigos 04-07) E tem 2+ profissionais envolvidos — sem CBO/CNS de cada um, não dá pra confirmar que um deles é de fato cadastrado em eMulti, então ainda é uma aproximação.",
    "Reuniões (Resumo Reuniões) só contam oficialmente pra M2 quando são dos tipos 'Reunião de equipe', 'Reunião com outras equipes de saúde' ou 'Reunião intersetorial' (códigos 01-03) E registradas com o tema 'Discussão de Caso/Projeto Terapêutico Singular' — como a aba de reuniões não tem uma coluna de tema, esta extração conta qualquer reunião com 2+ profissionais, o que pode puxar o M2 um pouco PRA CIMA nesse componente específico.",
    "'Desempenho quadrimestral' NÃO é uma fórmula oficial do Ministério da Saúde — é uma síntese própria: Nota final = pontos M1 × 6 + pontos M2 × 4 (pontos por classificação: Regular=0,25, Suficiente=0,5, Bom=0,75, Ótimo=1), classificada como Regular < 2,6, Suficiente 2,6 a 4,9, Bom 5 a 7,5, Ótimo > 7,5 — pra dar uma visão geral rápida; os indicadores oficiais continuam sendo M1 e M2 separados."
  ];

  // Motor de cálculo: recebe o "workbook" (abas já em formato de matriz de
  // linhas) e o período {inicio, fim} (objetos Date) e calcula M1, M2 e o
  // Desempenho quadrimestral direto dos dados brutos — replica a lógica do
  // extrair_esus_unificado.py (_calcular_indicadores_m1_m2), mas já
  // filtrando pela janela móvel de 4 meses.
  function calcularIndicadoresDoPeriodo(wb, periodo){
    function rowsOf(baseName){
      var ws = wb.Sheets[suffixedName(baseName)];
      return ws ? sheetToRows(ws) : [];
    }

    // ---------- Atendimentos ----------
    var atRows = rowsOf("Atendimentos");
    var atHeader = atRows[0] || [];
    var iData = colIndex(atHeader, "data_hora");
    var iNome = colIndex(atHeader, "nome");
    var atFiltradas = atRows.slice(1).filter(function(r){
      var nome = String(r[iNome]||"").trim();
      return nome && withinPeriod(parseBRDate(r[iData]), periodo.inicio, periodo.fim);
    });
    var atendimentosIndividuais = atFiltradas.length;

    // ---------- Participantes Ativ. Coletiva ----------
    var partRows = rowsOf("Participantes Ativ. Coletiva");
    var partHeader = partRows[0] || [];
    var iPData = colIndex(partHeader, "data");
    var iPNome = colIndex(partHeader, "participante");
    var partFiltradas = partRows.slice(1).filter(function(r){
      var nome = String(r[iPNome]||"").trim();
      return nome && nome.indexOf("(sem lista nominal") !== 0
        && withinPeriod(parseBRDate(r[iPData]), periodo.inicio, periodo.fim);
    });
    var participacoesColetivas = partFiltradas.length;

    // ---------- M1: numerador/denominador ----------
    var numeradorM1 = atendimentosIndividuais + participacoesColetivas;
    var pessoasSet = {}; // nome em maiúsculas -> {at, part}
    atFiltradas.forEach(function(r){
      var chave = String(r[iNome]).trim().toUpperCase();
      if(!pessoasSet[chave]) pessoasSet[chave] = {nome:String(r[iNome]).trim(), at:0, part:0};
      pessoasSet[chave].at++;
    });
    partFiltradas.forEach(function(r){
      var chave = String(r[iPNome]).trim().toUpperCase();
      if(!pessoasSet[chave]) pessoasSet[chave] = {nome:String(r[iPNome]).trim(), at:0, part:0};
      pessoasSet[chave].part++;
    });
    var pessoasLista = Object.keys(pessoasSet).map(function(k){ return pessoasSet[k]; })
      .sort(function(a,b){ return a.nome.localeCompare(b.nome,'pt-BR'); });
    var denominadorM1 = pessoasLista.length;
    var m1 = denominadorM1 ? (numeradorM1/denominadorM1) : null;
    var classificacaoM1 = classificarM1(m1);

    // ---------- Resumo Atividade Coletiva ----------
    var racRows = rowsOf("Resumo Atividade Coletiva");
    var racHeader = racRows[0] || [];
    var iRacData = colIndex(racHeader, "data");
    var iRacTipo = colIndex(racHeader, "tipo_atividade");
    var iRacTotalProf = colIndex(racHeader, "qtd_total_profissionais");
    var iRacProfEnv = colIndex(racHeader, "qtd_profissionais_envolvidos");
    // Só estes 4 tipos (códigos 04-07 da Atividade Coletiva) contam como
    // "Atividade Coletiva Compartilhada" pra M2 — reuniões (códigos 01-03)
    // vêm de outra aba (Resumo Reuniões) e têm regra própria.
    var TIPOS_ATIV_COLETIVA_COMPARTILHADA = [
      "educacao em saude", "atendimento em grupo",
      "avaliacao/procedimento coletivo", "mobilizacao social"
    ];
    var racFiltradas = racRows.slice(1).filter(function(r){
      return withinPeriod(parseBRDate(r[iRacData]), periodo.inicio, periodo.fim);
    });
    var atividadesTotais = racFiltradas.length;
    var atividadesCompartilhadas = racFiltradas.filter(function(r){
      var totalProf = iRacTotalProf>=0 && r[iRacTotalProf]!=="" ? toInt(r[iRacTotalProf]) : 1+toInt(r[iRacProfEnv]);
      var tipoOk = iRacTipo<0 || TIPOS_ATIV_COLETIVA_COMPARTILHADA.indexOf(normalizarTexto(r[iRacTipo])) >= 0;
      return totalProf >= 2 && tipoOk;
    }).length;

    // ---------- Resumo Reuniões ----------
    var rrRows = rowsOf("Resumo Reuniões");
    var rrHeader = rrRows[0] || [];
    var iRrData = colIndex(rrHeader, "data");
    var iRrQtd = colIndex(rrHeader, "qtd_participantes");
    var rrFiltradas = rrRows.slice(1).filter(function(r){
      return withinPeriod(parseBRDate(r[iRrData]), periodo.inicio, periodo.fim);
    });
    var reunioesTotais = rrFiltradas.length;
    var reunioesCompartilhadas = rrFiltradas.filter(function(r){ return toInt(r[iRrQtd]) >= 2; }).length;

    // ---------- M2 ----------
    var numeradorM2 = atividadesCompartilhadas + reunioesCompartilhadas;
    var denominadorM2 = atendimentosIndividuais + numeradorM2;
    var m2 = denominadorM2 ? (numeradorM2/denominadorM2*100) : null;
    var classificacaoM2 = classificarM2(m2);

    // ---------- Desempenho quadrimestral (síntese própria) ----------
    var pontosM1 = PONTOS_POR_CLASSE[classificacaoM1];
    var pontosM2 = PONTOS_POR_CLASSE[classificacaoM2];
    var pontosM1Pesados = pontosM1!==undefined ? pontosM1*6 : null;
    var pontosM2Pesados = pontosM2!==undefined ? pontosM2*4 : null;
    var notaFinal = (pontosM1Pesados!==null && pontosM2Pesados!==null) ? (pontosM1Pesados+pontosM2Pesados) : null;
    var desempenho = classificarDesempenho(notaFinal);

    // ---------- "Pessoas atendidas" (lista dinâmica, só do período) ----------
    var pessoasAtendidasHeaders = ["Nome","Atendimentos","Participantes Ativ. Coletiva","Total"];
    var pessoasAtendidasRows = pessoasLista.map(function(p){
      return [p.nome, p.at, p.part, p.at+p.part];
    });

    return {
      equipe: currentEquipes.map(function(e){ return e.label; }).join(' + '),
      data: {
        atendimentosIndividuais: atendimentosIndividuais,
        participacoesColetivas: participacoesColetivas,
        numeradorM1: numeradorM1,
        denominadorM1: denominadorM1,
        m1: m1,
        classificacaoM1: classificacaoM1,
        atividadesTotais: atividadesTotais,
        atividadesCompartilhadas: atividadesCompartilhadas,
        reunioesTotais: reunioesTotais,
        reunioesCompartilhadas: reunioesCompartilhadas,
        denominadorM2: denominadorM2,
        numeradorM2: numeradorM2,
        m2: m2,
        classificacaoM2: classificacaoM2,
        pontosM1: pontosM1Pesados,
        pontosM2: pontosM2Pesados,
        notaFinal: notaFinal,
        desempenho: desempenho
      },
      notes: NOTAS_METODOLOGICAS,
      pessoasAtendidas: {headers: pessoasAtendidasHeaders, rows: pessoasAtendidasRows}
    };
  }

  // ---------- Listas ----------
  // "Pessoas atendidas" com filtro de mês PRÓPRIO (independente do filtro
  // de Mês do topo): deduplica direto de Atendimentos + Participantes
  // Ativ. Coletiva (já filtradas por equipe no fetch), sem depender de
  // nenhum link/aba externa. monthValues vazio = todos os meses
  // disponíveis (sem filtro); com meses marcados, só entram atendimentos/
  // participações daqueles meses.
  function pessoasAtendidasParaMeses(monthValues){
    var pessoasSet = {}; // nome em maiúsculas -> {nome, at, part, datas:[Date,...]}
    function dentroDoFiltro(d){
      if(!monthValues || !monthValues.length) return true;
      return !!d && monthValues.indexOf(monthOptionValue(d)) >= 0;
    }
    var atCached = latestSheets[suffixedName("Atendimentos")];
    if(atCached){
      var iData = colIndex(atCached.headers, "data_hora");
      var iNome = colIndex(atCached.headers, "nome");
      if(iData >= 0 && iNome >= 0){
        atCached.rows.forEach(function(r){
          var nome = String(r[iNome]||"").trim();
          var d = parseBRDate(r[iData]);
          if(!nome || !dentroDoFiltro(d)) return;
          var chave = nome.toUpperCase();
          if(!pessoasSet[chave]) pessoasSet[chave] = {nome:nome, at:0, part:0, datas:[]};
          pessoasSet[chave].at++;
          if(d) pessoasSet[chave].datas.push(d);
        });
      }
    }
    var partCached = latestSheets[suffixedName("Participantes Ativ. Coletiva")];
    if(partCached){
      var iPData = colIndex(partCached.headers, "data");
      var iPNome = colIndex(partCached.headers, "participante");
      if(iPData >= 0 && iPNome >= 0){
        partCached.rows.forEach(function(r){
          var nome = String(r[iPNome]||"").trim();
          var d = parseBRDate(r[iPData]);
          if(!nome || nome.indexOf("(sem lista nominal") === 0 || !dentroDoFiltro(d)) return;
          var chave = nome.toUpperCase();
          if(!pessoasSet[chave]) pessoasSet[chave] = {nome:nome, at:0, part:0, datas:[]};
          pessoasSet[chave].part++;
          if(d) pessoasSet[chave].datas.push(d);
        });
      }
    }
    var pessoasLista = Object.keys(pessoasSet).map(function(k){ return pessoasSet[k]; })
      .sort(function(a,b){ return a.nome.localeCompare(b.nome,'pt-BR'); });
    // Ordena as datas de cada pessoa em ordem cronológica e descobre o
    // maior número de datas entre todas as pessoas, pra saber quantas
    // colunas "Data N" a tabela precisa ter (colunas sobrando ficam "—"),
    // limitado a no máximo 10 colunas (MAX_DATAS_PESSOA_ATENDIDA) — quem
    // tiver mais de 10 eventos no período só mostra os 10 primeiros.
    var MAX_DATAS_PESSOA_ATENDIDA = 10;
    var maxDatas = 0;
    pessoasLista.forEach(function(p){
      p.datas.sort(function(a,b){ return a-b; });
      if(p.datas.length > maxDatas) maxDatas = p.datas.length;
    });
    maxDatas = Math.min(maxDatas, MAX_DATAS_PESSOA_ATENDIDA);
    var dataHeaders = [];
    for(var i=1;i<=maxDatas;i++){ dataHeaders.push("Data "+i); }
    return {
      headers: ["Nome","Atendimentos","Participantes Ativ. Coletiva","Total"].concat(dataHeaders),
      rows: pessoasLista.map(function(p){
        var row = [p.nome, p.at, p.part, p.at+p.part];
        for(var i=0;i<maxDatas;i++){
          row.push(p.datas[i] ? fmtBRDate(p.datas[i]) : "—");
        }
        return row;
      })
    };
  }
  // Meses disponíveis pro filtro de "Pessoas atendidas": união dos meses
  // com dado em Atendimentos e em Participantes Ativ. Coletiva (mais
  // recente primeiro).
  function monthOptionsParaPessoasAtendidas(){
    var seen = {}, months = [];
    function coletar(name, dateHeader){
      var cached = latestSheets[name];
      if(!cached) return;
      var idx = colIndex(cached.headers, dateHeader);
      if(idx < 0) return;
      cached.rows.forEach(function(r){
        var d = parseBRDate(r[idx]);
        if(!d) return;
        var v = monthOptionValue(d);
        if(!seen[v]){ seen[v] = true; months.push(new Date(d.getFullYear(), d.getMonth(), 1)); }
      });
    }
    coletar(suffixedName("Atendimentos"), "data_hora");
    coletar(suffixedName("Participantes Ativ. Coletiva"), "data");
    months.sort(function(a,b){ return b-a; });
    return months.map(function(d){ return {value: monthOptionValue(d), label: monthOptionLabel(d)}; });
  }
  function populateSheetsCache(wb){
    latestSheets = {};
    wb.SheetNames.forEach(function(name){
      var rows = sheetToRows(wb.Sheets[name]).filter(function(r){
        return r.some(function(c){ return String(c).trim() !== ""; });
      });
      if(!rows.length) return;
      var headers = rows[0].map(function(h){ return String(h||"").trim() || "—"; });
      latestSheets[name] = {headers: headers, rows: rows.slice(1)};
    });
  }

  function renderListCard(name){
    // "Pessoas atendidas" é uma lista calculada aqui mesmo no navegador
    // (dedup de Atendimentos + Participantes Ativ. Coletiva) — ver
    // pessoasAtendidasParaMeses. Tem filtro de mês PRÓPRIO, independente
    // do filtro de Mês do topo da página.
    var isPessoasAtendidas = (name === suffixedName("Pessoas atendidas"));
    var cached = isPessoasAtendidas
      ? pessoasAtendidasParaMeses(listMonthFilters[name] || [])
      : latestSheets[name];
    if(isPessoasAtendidas) latestSheets[name] = cached;
    var body;
    var hasTable = false;
    if(!cached){
      body = '<div class="list-placeholder">Não encontramos uma aba chamada "'+escapeHtml(name)+'" na planilha publicada.</div>';
    } else if(!cached.rows.length && !isPessoasAtendidas){
      body = '<div class="list-placeholder">Esta lista está vazia.</div>';
    } else {
      hasTable = true;
      var dateColIdx = dateColIndexForList(cached.headers);
      listDateColIdx[name] = dateColIdx;
      var theadHtml = '<tr>'+cached.headers.map(function(h){ return '<th>'+escapeHtml(h)+'</th>'; }).join('')+'</tr>';
      var bodyHtml = cached.rows.map(function(r){
        return '<tr>'+cached.headers.map(function(h,i){
          var v = r[i];
          return '<td>'+escapeHtml(v===undefined||v===null?'':v)+'</td>';
        }).join('')+'</tr>';
      }).join('');
      var colOptionsHtml = '<option value="">Filtrar por coluna…</option>'
        + cached.headers.map(function(h,i){ return '<option value="'+i+'">'+escapeHtml(h)+'</option>'; }).join('');
      var filterPairsHtml = [0,1,2].map(function(idx){
        return '<div class="filter-pair">'
          + '<select class="filter-col">'+colOptionsHtml+'</select>'
          + '<div class="ms-wrap filter-val-ms ms-disabled" data-pair-idx="'+idx+'"></div>'
          + '</div>';
      }).join('');
      // Filtro de mês (multisseleção) — aparece quando a lista tem uma
      // coluna de data reconhecível ("data" ou "data_hora"), ou é a
      // "Pessoas atendidas" calculada (filtro próprio, ver acima). Fica
      // na MESMA linha dos filtros de coluna (dentro de .list-filters),
      // como o primeiro item da fileira.
      var monthFilterHtml = (dateColIdx >= 0 || isPessoasAtendidas)
        ? '<div class="list-month-filter"><label class="list-month-filter-label">Mês</label>'
          + '<div class="ms-wrap" data-month-filter="'+escapeHtml(name)+'"'+(isPessoasAtendidas ? ' data-computed-months="1"' : '')+'></div></div>'
        : '';
      body = '<p class="list-meta">'+fmtInt(cached.rows.length)+(cached.rows.length===1?' linha':' linhas')+'</p>'
        + '<div class="list-filters" data-list-filters="'+escapeHtml(name)+'">'+monthFilterHtml+filterPairsHtml+'</div>'
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
    el.innerHTML = names.map(renderListCard).join('');

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
      // Contagem de linhas mostrada acima da lista: reflete o resultado
      // depois de aplicar TODOS os filtros ativos (mês, colunas e busca),
      // não o total bruto da lista.
      var metaEl = card.querySelector('.list-meta');
      if(metaEl) metaEl.textContent = fmtInt(visibleCount) + (visibleCount === 1 ? ' linha' : ' linhas');
    }

    el.querySelectorAll('[data-month-filter]').forEach(function(container){
      var name = container.getAttribute('data-month-filter');
      if(container.getAttribute('data-computed-months') === '1'){
        // "Pessoas atendidas": filtro de mês próprio — recalcula a
        // dedup (Atendimentos + Participantes Ativ. Coletiva) na hora,
        // em vez de só esconder/mostrar linhas de uma tabela fixa.
        var optsCalc = monthOptionsParaPessoasAtendidas();
        var validCalc = optsCalc.map(function(o){ return o.value; });
        listMonthFilters[name] = (listMonthFilters[name] || []).filter(function(v){
          return validCalc.indexOf(v) >= 0;
        });
        var calcMs = createMultiSelect(container, {
          placeholder: 'Todos os meses', multi: true, search: optsCalc.length > 8, showTags: true,
          onChange: function(keys){
            listMonthFilters[name] = keys;
            var card = container.closest('.list-card');
            var novoCached = pessoasAtendidasParaMeses(keys);
            latestSheets[name] = novoCached;
            var tbody = card.querySelector('tbody');
            if(tbody){
              tbody.innerHTML = novoCached.rows.map(function(r){
                return '<tr>'+novoCached.headers.map(function(h,i){
                  var v = r[i];
                  return '<td>'+escapeHtml(v===undefined||v===null?'':v)+'</td>';
                }).join('')+'</tr>';
              }).join('');
            }
            applyFilters(card);
          }
        });
        calcMs.setOptions(optsCalc);
        calcMs.setSelected(listMonthFilters[name]);
        applyFilters(container.closest('.list-card'));
        return;
      }
      var cached = latestSheets[name];
      var dateColIdx = listDateColIdx[name];
      if(!cached || dateColIdx == null || dateColIdx < 0) return;
      var opts = monthOptionsForList(cached, dateColIdx);
      var validValues = opts.map(function(o){ return o.value; });
      // Mantém só a seleção anterior que ainda faz sentido (evita "mês
      // fantasma" depois que os dados são atualizados).
      listMonthFilters[name] = (listMonthFilters[name] || []).filter(function(v){
        return validValues.indexOf(v) >= 0;
      });
      var monthMs = createMultiSelect(container, {
        placeholder: 'Todos os meses', multi: true, search: opts.length > 8, showTags: true,
        onChange: function(keys){
          listMonthFilters[name] = keys;
          applyFilters(container.closest('.list-card'));
        }
      });
      monthMs.setOptions(opts);
      monthMs.setSelected(listMonthFilters[name]);
      applyFilters(container.closest('.list-card'));
    });

    el.querySelectorAll('[data-filter-key]').forEach(function(input){
      input.addEventListener('input', function(){
        applyFilters(input.closest('.list-card'));
      });
    });

    el.querySelectorAll('.filter-pair').forEach(function(pair){
      var colSelect = pair.querySelector('.filter-col');
      var valWrap = pair.querySelector('.filter-val-ms');
      // Multisseleção de valores ("Todos os valores"): fica desabilitada
      // (opacidade + sem clique, via .ms-disabled) até uma coluna ser
      // escolhida no select ao lado. A instância fica pendurada no
      // próprio elemento (._msInstance) pra applyFilters conseguir ler
      // os valores marcados sem precisar de um estado global à parte.
      var msInst = createMultiSelect(valWrap, {
        placeholder: 'Todos os valores', multi: true, search: true, showTags: true,
        onChange: function(){ applyFilters(pair.closest('.list-card')); }
      });
      valWrap._msInstance = msInst;

      colSelect.addEventListener('change', function(){
        var card = colSelect.closest('.list-card');
        var listName = card.querySelector('[data-list-filters]').getAttribute('data-list-filters');
        var colIdx = colSelect.value !== '' ? parseInt(colSelect.value, 10) : null;
        if(colIdx === null){
          msInst.setOptions([]);
          msInst.setSelected([]);
          valWrap.classList.add('ms-disabled');
        } else {
          var cached = latestSheets[listName];
          var seen = {};
          var values = [];
          (cached ? cached.rows : []).forEach(function(r){
            var v = r[colIdx];
            v = (v===undefined||v===null) ? '' : String(v).trim();
            if(v && !seen[v]){ seen[v] = true; values.push(v); }
          });
          values.sort(function(a,b){ return a.localeCompare(b, 'pt-BR'); });
          msInst.setOptions(values.map(function(v){ return {value:v, label:v}; }));
          msInst.setSelected([]);
          valWrap.classList.remove('ms-disabled');
        }
        applyFilters(card);
      });
    });

    el.querySelectorAll('[data-pdf-btn]').forEach(function(btn){
      btn.addEventListener('click', function(){
        gerarPdfLista(btn.getAttribute('data-pdf-btn'), btn.closest('.list-card'), btn);
      });
    });
  }

  // ---------- Exportar lista em PDF ----------
  // Gera um PDF "elegante" (faixa de cabeçalho colorida + tabela) a partir
  // do que está REALMENTE visível na tela: lê o <thead>/<tbody> do próprio
  // card já filtrado (busca + filtros de coluna + filtro de mês), em vez
  // de reconstruir a partir de latestSheets — assim o PDF bate 100% com o
  // que os filtros ativos estão mostrando, sem duplicar a lógica deles.
  function monthValueToLabel(v){
    var parts = String(v).split('-');
    return monthOptionLabel(new Date(+parts[0], +parts[1]-1, 1));
  }
  function slugifyFileName(s){
    return normalizeText(s).replace(/[^A-Z0-9]+/g,'_').replace(/^_+|_+$/g,'');
  }
  function gerarPdfLista(listName, card, btn){
    if(!card) return;
    var jspdfNs = window.jspdf;
    if(!jspdfNs || !jspdfNs.jsPDF){
      alert('Não foi possível carregar a biblioteca de geração de PDF (verifique a conexão com a internet) — tente novamente.');
      return;
    }
    var headers = Array.prototype.map.call(card.querySelectorAll('thead th'), function(th){ return th.textContent.trim(); });
    var todasLinhas = card.querySelectorAll('tbody tr');
    var linhasVisiveis = Array.prototype.filter.call(todasLinhas, function(tr){ return tr.style.display !== 'none'; })
      .map(function(tr){ return Array.prototype.map.call(tr.children, function(td){ return td.textContent.trim(); }); });
    if(!linhasVisiveis.length){
      alert('Nenhuma linha visível com os filtros atuais dessa lista — ajuste os filtros antes de gerar o PDF.');
      return;
    }

    // Monta o resumo dos filtros ativos nesta lista, pra registrar no
    // cabeçalho do PDF exatamente o que foi aplicado.
    var filtrosAtivos = [];
    var searchInput = card.querySelector('.list-search');
    if(searchInput && searchInput.value.trim()) filtrosAtivos.push('Busca: "'+searchInput.value.trim()+'"');
    var mesesSelecionados = listMonthFilters[listName] || [];
    if(mesesSelecionados.length){
      filtrosAtivos.push('Mês: '+mesesSelecionados.map(monthValueToLabel).join(', '));
    }
    card.querySelectorAll('.filter-pair').forEach(function(pair){
      var colSelect = pair.querySelector('.filter-col');
      var valWrap = pair.querySelector('.filter-val-ms');
      var colIdx = colSelect && colSelect.value !== '' ? parseInt(colSelect.value, 10) : null;
      var vals = (valWrap && valWrap._msInstance) ? valWrap._msInstance.getSelected() : [];
      if(colIdx !== null && vals.length){
        filtrosAtivos.push(headers[colIdx]+': '+vals.join(', '));
      }
    });

    var totalLinhas = todasLinhas.length;
    var nomeExibicao = displayListName(listName);
    var equipeLabel = currentEquipes.map(function(e){ return e.label; }).join(' + ');

    var doc = new jspdfNs.jsPDF({orientation: headers.length > 6 ? 'landscape' : 'portrait', unit:'pt', format:'a4'});
    var pageWidth = doc.internal.pageSize.getWidth();
    var pageHeight = doc.internal.pageSize.getHeight();
    var margin = 28;

    // ---- Faixa de cabeçalho ----
    doc.setFillColor(21,63,53);
    doc.rect(0,0,pageWidth,64,'F');
    doc.setTextColor(238,243,234);
    doc.setFont('helvetica','bold');
    doc.setFontSize(15);
    doc.text('Painel eMulti — Indicadores M1 e M2', margin, 26);
    doc.setFont('helvetica','normal');
    doc.setFontSize(10);
    doc.setTextColor(159,192,174);
    doc.text(equipeLabel, margin, 42);
    doc.setFontSize(8.5);
    doc.text('Gerado em '+new Date().toLocaleString('pt-BR'), pageWidth-margin, 26, {align:'right'});

    // ---- Título da lista + resumo dos filtros ----
    var y = 84;
    doc.setTextColor(21,63,53);
    doc.setFont('helvetica','bold');
    doc.setFontSize(13);
    doc.text(nomeExibicao, margin, y);
    y += 16;
    doc.setFont('helvetica','normal');
    doc.setFontSize(9);
    doc.setTextColor(81,96,90);
    if(filtrosAtivos.length){
      filtrosAtivos.forEach(function(linha){
        var quebradas = doc.splitTextToSize('• '+linha, pageWidth-margin*2);
        doc.text(quebradas, margin, y);
        y += 12*quebradas.length;
      });
    } else {
      doc.text('Sem filtros aplicados — exibindo todos os registros.', margin, y);
      y += 12;
    }
    doc.text(fmtInt(linhasVisiveis.length)+' de '+fmtInt(totalLinhas)+(totalLinhas===1?' linha no total.':' linhas no total.'), margin, y);
    y += 10;

    doc.autoTable({
      startY: y+6,
      head: [headers],
      body: linhasVisiveis,
      theme: 'grid',
      margin: {left:margin, right:margin, bottom:34},
      styles: {font:'helvetica', fontSize: headers.length > 9 ? 7 : (headers.length > 6 ? 7.8 : 8.6), cellPadding:4, overflow:'linebreak', textColor:[19,36,31], lineColor:[220,228,214], lineWidth:0.5},
      headStyles: {fillColor:[21,63,53], textColor:255, fontStyle:'bold'},
      alternateRowStyles: {fillColor:[241,244,238]},
      didDrawPage: function(){
        doc.setFontSize(8);
        doc.setTextColor(150,158,152);
        doc.text('Página '+doc.internal.getCurrentPageInfo().pageNumber, pageWidth-margin, pageHeight-14, {align:'right'});
      }
    });

    var arquivo = slugifyFileName(nomeExibicao)+'__'+slugifyFileName(equipeLabel)+'__'+slugifyFileName(new Date().toLocaleDateString('pt-BR'))+'.pdf';
    doc.save(arquivo);
  }

  // ---------- Gauge ----------
  function polar(cx,cy,r,angleDeg){
    var a = angleDeg * Math.PI/180;
    return {x: cx + r*Math.cos(a), y: cy - r*Math.sin(a)};
  }
  function arcPath(cx,cy,r,startAngle,endAngle){
    var p1 = polar(cx,cy,r,startAngle);
    var p2 = polar(cx,cy,r,endAngle);
    var large = Math.abs(startAngle-endAngle) > 180 ? 1 : 0;
    return "M "+p1.x+" "+p1.y+" A "+r+" "+r+" 0 "+large+" 1 "+p2.x+" "+p2.y;
  }
  function buildGauge(value, domainMax, bands, gaugeId){
    var cx=115,cy=122,r=92,thick=16;
    var bandsSvg = bands.map(function(b){
      var a1 = 180 - (b.from/domainMax)*180;
      var a2 = 180 - (b.to/domainMax)*180;
      return '<path d="'+arcPath(cx,cy,r,a1,a2)+'" stroke="'+b.color+'" stroke-width="'+thick+'" fill="none" stroke-linecap="round"/>';
    }).join('');
    var frac = (value===null || value===undefined || isNaN(value)) ? 0 : Math.max(0, Math.min(1, value/domainMax));
    var targetAngle = 180 - frac*180;
    var needleRotation = 180 - targetAngle; // graus a girar o ponteiro (que nasce apontando p/ 0)
    var needleLen = r - thick/2 - 6;
    var tipBase = polar(cx,cy,needleLen,180);
    // Ponteiro é desenhado sempre apontando para "0" (esquerda) e a rotação até
    // o valor real é feita via CSS puro (animation + custom property), em vez
    // de depender de JS aplicar o transform depois — isso evita que o ponteiro
    // fique "zerado" caso a atualização via JS não rode a tempo/corretamente.
    var needleSvg = '<g id="'+gaugeId+'" class="gauge-needle" style="transform-origin:'+cx+'px '+cy+'px;--target-angle:'+needleRotation+'deg;">'
      + '<line x1="'+cx+'" y1="'+cy+'" x2="'+tipBase.x+'" y2="'+tipBase.y+'" stroke="#13241F" stroke-width="3" stroke-linecap="round"/>'
      + '<circle cx="'+cx+'" cy="'+cy+'" r="5.5" fill="#13241F"/></g>';
    return '<svg class="gauge-svg" viewBox="0 0 230 148">'+bandsSvg+needleSvg+'</svg>';
  }
  function animateGauges(){
    // Mantida como no-op por compatibilidade com as chamadas existentes em
    // renderDashboard(); a animação agora é 100% CSS (ver .gauge-needle).
  }

  var CLASS_BANDS_M1 = [
    {from:0,to:1,color:arcHex("Regular")},
    {from:1,to:2,color:arcHex("Suficiente")},
    {from:2,to:3,color:arcHex("Bom")},
    {from:3,to:4,color:arcHex("Ótimo")}
  ];
  var CLASS_BANDS_M2 = [
    {from:0,to:1,color:arcHex("Regular")},
    {from:1,to:2.5,color:arcHex("Suficiente")},
    {from:2.5,to:5,color:arcHex("Bom")},
    {from:5,to:8,color:arcHex("Ótimo")}
  ];
  var CLASS_BANDS_NOTA = [
    {from:0,to:2.5,color:arcHex("Regular")},
    {from:2.5,to:5,color:arcHex("Suficiente")},
    {from:5,to:7.5,color:arcHex("Bom")},
    {from:7.5,to:10,color:arcHex("Ótimo")}
  ];

  function gaugeLegendHTML(items){
    return '<div class="gauge-legend">' + items.map(function(it){
      return '<span><i style="background:'+it.color+'"></i>'+it.label+' '+it.cond+'</span>';
    }).join('') + '</div>';
  }
  var LEGEND_M1 = [
    {label:'Ótimo',      cond:'&gt; 3',           color:arcHex('Ótimo')},
    {label:'Bom',        cond:'&gt; 2 e ≤ 3',     color:arcHex('Bom')},
    {label:'Suficiente', cond:'&gt; 1 e ≤ 2',     color:arcHex('Suficiente')},
    {label:'Regular',    cond:'≤ 1',              color:arcHex('Regular')}
  ];
  var LEGEND_M2 = [
    {label:'Ótimo',      cond:'&gt; 5%',              color:arcHex('Ótimo')},
    {label:'Bom',        cond:'&gt; 2,5% e ≤ 5%',     color:arcHex('Bom')},
    {label:'Suficiente', cond:'&gt; 1% e ≤ 2,5%',     color:arcHex('Suficiente')},
    {label:'Regular',    cond:'≤ 1%',                 color:arcHex('Regular')}
  ];
  var LEGEND_NOTA = [
    {label:'Ótimo',      cond:'&gt; 7,5',            color:arcHex('Ótimo')},
    {label:'Bom',        cond:'≥ 5 e ≤ 7,5',         color:arcHex('Bom')},
    {label:'Suficiente', cond:'&gt; 2,5 e &lt; 5',   color:arcHex('Suficiente')},
    {label:'Regular',    cond:'≤ 2,5',               color:arcHex('Regular')}
  ];

  function gaugeCardHTML(title, formula, value, domainMax, bands, gaugeId, valueHtml, classLabel, note, legend){
    return '<div class="card gauge-card">'
      + '<div class="gauge-header"><h3>'+title+'</h3><p class="formula">'+formula+'</p></div>'
      + buildGauge(value, domainMax, bands, gaugeId)
      + '<div class="gauge-value">'+valueHtml+'</div>'
      + '<span class="pill" style="background:'+pillHex(classLabel)+'">'+(classLabel||'—')+'</span>'
      + (note ? '<p class="gauge-note">'+note+'</p>' : '')
      + (legend ? gaugeLegendHTML(legend) : '')
      + '</div>';
  }

  // ---------- Meta do quadrimestre ----------
  // Ícone simples de alvo/meta usado no cabeçalho de cada bloco.
  var METAS_ICON_SVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8">'
    + '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.3"/><circle cx="12" cy="12" r="1"/></svg>';

  function fmtDecMeta(n){
    return (Math.round(n*10)/10).toLocaleString('pt-BR', {minimumFractionDigits:1, maximumFractionDigits:1});
  }

  // Um card de meta (ex.: "Bom" ou "Ótimo"): alvo a bater, ritmo médio
  // necessário até o fim do quadrimestre e quanto ainda falta.
  function metaCardHTML(cfg){
    var subLines =
        '<p class="meta-card-sub">Média/mês: '+fmtDecMeta(cfg.mediaMes)+'</p>'
      + '<p class="meta-card-sub">Média/semana: '+fmtDecMeta(cfg.mediaSemana)+'</p>';
    var faltamHtml;
    if(cfg.faltam<=0){
      faltamHtml = '<p class="meta-card-done">Meta já atingida ✓</p>';
    } else if(cfg.semanasRestantes<=0){
      faltamHtml = '<p class="meta-card-faltam">Faltam: '+fmtInt(cfg.faltam)+' '+cfg.unidadeFaltam+' (quadrimestre encerrado)</p>';
    } else {
      faltamHtml = '<p class="meta-card-faltam">Faltam: '+fmtInt(cfg.faltam)+' '+cfg.unidadeFaltam+'</p>'
        + '<p class="meta-card-sub">Média/semana: '+fmtDecMeta(cfg.mediaSemanaFaltam)+'</p>';
    }
    return '<div class="meta-card">'
      + '<h4 style="color:'+cfg.color+'">'+cfg.label+'</h4>'
      + '<div class="meta-card-value" style="color:'+cfg.color+'">'+fmtInt(cfg.alvo)+' <span>'+cfg.unidade+'</span></div>'
      + subLines + faltamHtml
      + '</div>';
  }

  // Bloco completo de meta do quadrimestre para um indicador (M1 ou M2):
  // cabeçalho com a base de cálculo + selo "Preliminar" (enquanto o
  // quadrimestre ainda não terminou) e os cards de cada faixa-alvo.
  function metaQuadrimestreHTML(titulo, base, baseLabel, cardsCfg, preliminar){
    var cardsHtml = cardsCfg.map(metaCardHTML).join('');
    return '<div class="meta-quad-wrap">'
      + '<div class="meta-quad-head">'
      +   '<span class="meta-icon">'+METAS_ICON_SVG+'</span>'
      +   '<div class="meta-quad-titles"><h3>'+titulo+'</h3><p>de <b>'+fmtInt(base)+'</b> '+baseLabel+'</p></div>'
      +   (preliminar ? '<span class="pill-preliminar"><i></i>Preliminar</span>' : '')
      + '</div>'
      + '<div class="meta-cards">'+cardsHtml+'</div>'
      + '</div>';
  }

  // Calcula os cards de meta (Bom/Ótimo) de um indicador para o
  // quadrimestre selecionado: alvo = limiar × base (denominador), ritmo
  // médio necessário (mês/semana) pra bater o alvo ao longo do
  // quadrimestre inteiro, e quanto falta + ritmo pro tempo que resta.
  function calcularMetasQuadrimestre(numerador, denominador, thresholds, unidade, unidadeFaltam){
    var meses = mesesDoQuadrimestre(quadSelecionado.ano, quadSelecionado.qIndex);
    var inicioQuad = new Date(meses[0].getFullYear(), meses[0].getMonth(), 1, 0,0,0,0);
    var fimQuad = new Date(meses[3].getFullYear(), meses[3].getMonth()+1, 0, 23,59,59,999);
    var hoje = new Date();
    var diasQuad = Math.round((fimQuad-inicioQuad)/86400000)+1;
    var semanasQuad = diasQuad/7;
    var diasRestantes = Math.max(0, Math.round((fimQuad-hoje)/86400000));
    var semanasRestantes = diasRestantes/7;
    var cards = thresholds.map(function(t){
      var alvo = Math.ceil(t.value*denominador);
      var faltam = Math.max(0, alvo-(numerador||0));
      return {
        label: t.label, color: t.color, unidade: unidade, unidadeFaltam: unidadeFaltam || unidade,
        alvo: alvo,
        mediaMes: alvo/4,
        mediaSemana: alvo/semanasQuad,
        faltam: faltam,
        semanasRestantes: semanasRestantes,
        mediaSemanaFaltam: semanasRestantes>0 ? faltam/semanasRestantes : 0
      };
    });
    return {cards: cards, preliminar: hoje < fimQuad};
  }

  var M1_META_THRESHOLDS = [
    {value:2,   label:'Bom (M1 ≥ 2,00)',   color:'var(--arc-bom)'},
    {value:3,   label:'Ótimo (M1 ≥ 3,00)', color:'var(--arc-otimo)'}
  ];
  var M2_META_THRESHOLDS = [
    {value:0.025, label:'Bom (M2 ≥ 2,50%)',  color:'var(--arc-bom)'},
    {value:0.05,  label:'Ótimo (M2 ≥ 5,00%)', color:'var(--arc-otimo)'}
  ];

  // ---------- Composition bars ----------
  function stackbar(segments, total){
    var t = total || segments.reduce(function(s,x){return s+(x.value||0);},0);
    var bars = segments.map(function(s){
      var pct = t>0 ? (s.value/t*100) : 0;
      return '<div class="seg" style="width:'+pct+'%;background:'+s.color+'"></div>';
    }).join('');
    var legend = segments.map(function(s){
      return '<span class="legend-item"><i style="background:'+s.color+'"></i>'+s.label+' ('+fmtInt(s.value)+')</span>';
    }).join('');
    return '<div class="stackbar">'+bars+'</div><div class="legend">'+legend+'</div>';
  }

  // ---------- Sparkline ----------
  // points: [{y, label, value}] — "value" (opcional) é o texto já formatado
  // (ex.: "2,45" ou "5,20%") mostrado acima de cada ponto da linha.
  function sparkline(points, color, opts){
    opts = opts || {};
    if(points.length < 2) return '<p class="footnote">Ainda não há leituras suficientes para mostrar a tendência.</p>';
    var hasAvg = !!opts.quadAvg;
    var W=320, padX=14, padTop=20;
    // Com linha de média, reserva uma faixa a mais (avgLabelGap) entre o
    // fundo da área de plotagem e a linha de rótulos dos meses, só pro
    // valor da média caber embaixo da linha tracejada sem encostar em nada.
    var plotH = hasAvg ? 54 : 66;
    var plotBottom = padTop + plotH;
    var avgLabelGap = hasAvg ? 22 : 0;
    var axisY = plotBottom + avgLabelGap + 10;
    var H = axisY + 4;

    var vals = points.map(function(p){return p.y;});
    var min = Math.min.apply(null, vals), max = Math.max.apply(null, vals);
    if(min===max){ min = min - 1; max = max + 1; }
    var stepX = (W-2*padX)/(points.length-1);
    var coords = points.map(function(p,i){
      var x = padX + i*stepX;
      var y = plotBottom - ((p.y-min)/(max-min))*(plotBottom-padTop);
      return {x:x,y:y};
    });
    var path = coords.map(function(c,i){ return (i===0?"M ":"L ")+c.x+" "+c.y; }).join(" ");
    // Pontos e rótulo numérico de cada mês: coloridos pela classificação
    // (Regular/Suficiente/Bom/Ótimo) DAQUELE valor específico — a linha
    // que os conecta continua na cor original do gráfico (só os valores
    // ganham a cor da faixa).
    // Cada mês vira um <g class="tp-point" data-month="..."> com um
    // círculo maior e invisível (área de toque/hover mais fácil de
    // acertar), o ponto, o valor e o rótulo do mês embaixo — tudo junto
    // pra dar/tirar destaque em bloco ao passar o mouse ou clicar (ver
    // setupTrendInteractivity), inclusive no card do outro indicador.
    var pointsSvg = points.map(function(p,i){
      var tone = opts.classify ? (CLASS_COLOR[opts.classify(p.y)] || color) : color;
      var ly = Math.max(9, coords[i].y - 8);
      var valTxt = (p.value!==undefined && p.value!==null && p.value!=='') ? '<text class="tp-val" x="'+coords[i].x+'" y="'+ly+'" font-size="9.5" font-weight="600" fill="'+tone+'" text-anchor="middle">'+escapeHtml(p.value)+'</text>' : '';
      return '<g class="tp-point" data-month="'+escapeHtml(p.label)+'">'
        + '<circle class="tp-hit" cx="'+coords[i].x+'" cy="'+coords[i].y+'" r="9" fill="transparent"/>'
        + '<circle class="tp-dot" cx="'+coords[i].x+'" cy="'+coords[i].y+'" r="3.2" fill="'+tone+'"/>'
        + valTxt
        + '<text class="tp-axis" x="'+coords[i].x+'" y="'+axisY+'" font-size="9" fill="var(--ink-soft)" text-anchor="middle">'+p.label+'</text>'
        + '</g>';
    }).join("");

    // "Montanha" de média de cada quadrimestre (Jan–Abr / Mai–Ago /
    // Set–Dez): agrupa os pontos exibidos que pertencem ao mesmo
    // quadrimestre (cada ponto já traz p.quadKey/p.quadLabel) e desenha,
    // atrás da linha de dados, um bloco reto (sem cantos arredondados) do
    // fundo do gráfico até a altura da média DESSES pontos — 30% opaco
    // (70% transparente) — com uma linha tracejada marcando o topo e o
    // valor logo ABAIXO dela (dentro da faixa reservada em avgLabelGap,
    // longe da linha de dados, dos rótulos de valor e da linha de meses).
    // A cor (área + linha + texto) segue a classificação da média
    // (Regular/Suficiente/Bom/Ótimo), igual às outras faixas do painel.
    var avgAreaSvg = '', avgLineSvg = '';
    if(hasAvg){
      var groups = [];
      points.forEach(function(p,i){
        var last = groups[groups.length-1];
        if(last && last.key === p.quadKey){ last.idx.push(i); }
        else { groups.push({key:p.quadKey, label:p.quadLabel, idx:[i]}); }
      });
      groups.forEach(function(g){
        var ys = g.idx.map(function(i){ return points[i].y; });
        var avg = ys.reduce(function(a,b){ return a+b; }, 0)/ys.length;
        var avgY = plotBottom - ((avg-min)/(max-min))*(plotBottom-padTop);
        var x1 = coords[g.idx[0]].x, x2 = coords[g.idx[g.idx.length-1]].x;
        if(g.idx.length===1){ x1 -= 12; x2 += 12; }
        x1 = Math.max(padX-4, x1); x2 = Math.min(W-padX+4, x2);
        var midX = (x1+x2)/2;
        var textY = Math.min(avgY + 16, plotBottom + avgLabelGap - 4);
        var faixa = opts.classify ? opts.classify(avg) : null;
        var tone = CLASS_COLOR[faixa] || '#7A5A2E';
        var valTxt = g.label+': '+fmtDec(avg,2)+(opts.suffix||'');
        var chipW = Math.max(34, valTxt.length*4.6+8);
        var chip = '<rect x="'+(midX-chipW/2)+'" y="'+(textY-9)+'" width="'+chipW+'" height="12" rx="3" fill="var(--surface)" opacity="0.92"/>';
        avgAreaSvg += '<polygon points="'+x1+','+plotBottom+' '+x1+','+avgY+' '+x2+','+avgY+' '+x2+','+plotBottom+'" fill="'+tone+'" opacity="0.3"/>';
        avgLineSvg += '<line x1="'+x1+'" y1="'+avgY+'" x2="'+x2+'" y2="'+avgY+'" stroke="'+tone+'" stroke-width="1.3" stroke-dasharray="3 3" opacity="0.9"/>'
          + chip
          + '<text x="'+midX+'" y="'+textY+'" font-size="8" font-weight="700" fill="'+tone+'" text-anchor="middle">'+escapeHtml(valTxt)+'</text>';
      });
    }

    return '<svg class="spark-svg trend-interactive" viewBox="0 0 '+W+' '+(H+14)+'">'
      + avgAreaSvg
      + '<path d="'+path+'" fill="none" stroke="'+color+'" stroke-width="2"/>' + pointsSvg
      + avgLineSvg + '</svg>';
  }

  // ---------- Interatividade da aba Tendência ----------
  // Ao passar o mouse (ou clicar/tocar, no celular) em cima de um ponto de
  // qualquer um dos dois gráficos (M1 ou M2), destaca em AMBOS os
  // containers só o ponto daquele mês: apaga (opacidade 0) o número dos
  // demais pontos e deixa o ponto/rótulo dos outros meses esmaecido, nos
  // dois gráficos ao mesmo tempo — permitindo comparar M1 e M2 do mesmo
  // mês lado a lado. Clique/toque "fixa" o destaque (pra quem não tem
  // hover); clicar de novo no mesmo ponto, ou fora dos gráficos, desfaz.
  function setupTrendInteractivity(){
    var svgs = document.querySelectorAll('.trend-interactive');
    if(!svgs.length) return;
    var pinnedMonth = null;

    function applyHighlight(month){
      svgs.forEach(function(svg){
        svg.classList.toggle('tp-hover-active', !!month);
        svg.querySelectorAll('.tp-point').forEach(function(pt){
          pt.classList.toggle('tp-active', !!month && pt.getAttribute('data-month') === month);
        });
      });
    }

    svgs.forEach(function(svg){
      svg.querySelectorAll('.tp-point').forEach(function(pt){
        var month = pt.getAttribute('data-month');
        pt.addEventListener('mouseenter', function(){
          if(!pinnedMonth) applyHighlight(month);
        });
        pt.addEventListener('mouseleave', function(){
          if(!pinnedMonth) applyHighlight(null);
        });
        pt.addEventListener('click', function(e){
          e.stopPropagation();
          pinnedMonth = (pinnedMonth === month) ? null : month;
          applyHighlight(pinnedMonth);
        });
      });
    });

    document.addEventListener('click', function(){
      if(pinnedMonth){
        pinnedMonth = null;
        applyHighlight(null);
      }
    });
  }

  // ---------- Tabs ----------
  var FILTER_BAR_TABS = {geral:true, m1:true, m2:true, tendencia:true};
  document.querySelectorAll('.tab').forEach(function(btn){
    btn.addEventListener('click', function(){
      document.querySelectorAll('.tab').forEach(function(b){ b.classList.toggle('active', b===btn); });
      var target = btn.getAttribute('data-tab');
      document.querySelectorAll('.tab-panel').forEach(function(p){
        p.classList.toggle('active', p.id === 'tab'+target.charAt(0).toUpperCase()+target.slice(1));
      });
      document.getElementById('filterBar').classList.toggle('hidden', !FILTER_BAR_TABS[target]);
    });
  });

  // ---------- Render ----------
  function renderDashboard(record, serieTendencia){
    serieTendencia = serieTendencia || [];
    document.getElementById('statusState').style.display = 'none';
    populateQuadSelect();
    document.getElementById('topEquipe').textContent = record.equipe || '—';
    document.getElementById('topPeriodo').textContent = record.periodo
      ? 'Período: ' + record.periodo.inicio + ' a ' + record.periodo.fim
      : '';
    document.getElementById('topUpdated').textContent = record.error
      ? 'Falha na última leitura'
      : 'Atualizado em ' + fmtDate(record.timestamp);

    if(record.error){
      document.getElementById('gaugeRow').innerHTML =
        '<div class="card" style="flex:1;"><p style="color:var(--pill-regular);margin:0;">Não encontramos os indicadores M1 e M2 nesta planilha. Confira se o link publicado é o correto.</p></div>';
      document.getElementById('compRow').innerHTML = '';
      return;
    }

    var d = record.data;

    // ---- Composição (4 cartões: Numerador/Denominador de M1 e M2) ----
    var numM1Bar = stackbar([
        {label:'Atendimentos individuais', value:d.atendimentosIndividuais, color:'#153F35'},
        {label:'Participações coletivas', value:d.participacoesColetivas, color:'#C68A3D'}
      ], d.numeradorM1);
    var denM1Bar = stackbar([
        {label:'Pessoas atendidas', value:d.denominadorM1, color:'#153F35'}
      ], d.denominadorM1);
    var numM2Bar = stackbar([
        {label:'Atividades coletivas compartilhadas', value:d.atividadesCompartilhadas, color:'#153F35'},
        {label:'Reuniões compartilhadas', value:d.reunioesCompartilhadas, color:'#C68A3D'}
      ], d.numeradorM2);
    var denM2Bar = stackbar([
        {label:'Atendimentos individuais (base)', value:(d.denominadorM2!=null && d.numeradorM2!=null) ? d.denominadorM2-d.numeradorM2 : d.atendimentosIndividuais, color:'#CBD3C4'},
        {label:'Atividades coletivas compartilhadas', value:d.atividadesCompartilhadas, color:'#153F35'},
        {label:'Reuniões compartilhadas', value:d.reunioesCompartilhadas, color:'#C68A3D'}
      ], d.denominadorM2);

    document.getElementById('compRow').innerHTML =
        '<div class="card comp-card"><h4>Numerador do M1</h4>'+numM1Bar+'</div>'
      + '<div class="card comp-card"><h4>Denominador do M1</h4>'+denM1Bar+'</div>'
      + '<div class="card comp-card"><h4>Numerador do M2</h4>'+numM2Bar+'</div>'
      + '<div class="card comp-card"><h4>Denominador do M2</h4>'+denM2Bar+'</div>';

    // ---- Visão geral: 3 gauges (M1, M2, Desempenho) ----
    document.getElementById('gaugeRow').innerHTML =
        gaugeCardHTML('M1 — Média de atendimentos por pessoa',
          'Atendimentos individuais + coletivos ÷ pessoas atendidas',
          d.m1, 4, CLASS_BANDS_M1, 'needle-geral-m1', fmtDec(d.m1,2), d.classificacaoM1,
          fmtInt(d.numeradorM1)+' atendimentos ÷ '+fmtInt(d.denominadorM1)+' pessoas', LEGEND_M1)
      + gaugeCardHTML('M2 — Ações compartilhadas',
          'Ações compartilhadas ÷ ações realizadas × 100',
          d.m2, 8, CLASS_BANDS_M2, 'needle-geral-m2', fmtDec(d.m2,2)+'<span class="unit">%</span>', d.classificacaoM2,
          fmtInt(d.numeradorM2)+' compartilhadas ÷ '+fmtInt(d.denominadorM2)+' ações', LEGEND_M2)
      + gaugeCardHTML('Desempenho quadrimestral',
          'Nota final = (Pontos M1 × 6 + Pontos M2 × 4) ÷ 10',
          d.notaFinal, 10, CLASS_BANDS_NOTA, 'needle-geral-nota', fmtDec(d.notaFinal,2), d.desempenho,
          'Pontos M1: '+fmtInt(d.pontosM1)+' · Pontos M2: '+fmtInt(d.pontosM2), LEGEND_NOTA);

    // ---- Meta do quadrimestre: alvo de atendimentos/ações compartilhadas
    // pra bater "Bom" e "Ótimo" em M1 e M2, com ritmo médio necessário. ----
    var metaM1 = calcularMetasQuadrimestre(d.numeradorM1, d.denominadorM1, M1_META_THRESHOLDS, 'atend.',
      'atendimentos (retornos) de pessoas que foram atendidas nos últimos 4 meses');
    var metaM2 = calcularMetasQuadrimestre(d.numeradorM2, d.denominadorM2, M2_META_THRESHOLDS, 'ações');
    document.getElementById('metaQuadRow').innerHTML =
        metaQuadrimestreHTML('Meta do quadrimestre — M1', d.denominadorM1, 'pessoas atendidas', metaM1.cards, metaM1.preliminar)
      + metaQuadrimestreHTML('Meta do quadrimestre — M2', d.denominadorM2, 'ações realizadas', metaM2.cards, metaM2.preliminar);

    // ---- Aba M1: gauge + composição do M1 + listas ----
    document.getElementById('gaugeRowM1').innerHTML =
      gaugeCardHTML('M1 — Média de atendimentos por pessoa',
        'Atendimentos individuais + coletivos ÷ pessoas atendidas',
        d.m1, 4, CLASS_BANDS_M1, 'needle-m1tab-m1', fmtDec(d.m1,2), d.classificacaoM1,
        fmtInt(d.numeradorM1)+' atendimentos ÷ '+fmtInt(d.denominadorM1)+' pessoas', LEGEND_M1);
    document.getElementById('compRowM1').innerHTML =
        '<div class="card comp-card"><h4>Numerador do M1</h4>'+numM1Bar+'</div>'
      + '<div class="card comp-card"><h4>Denominador do M1</h4>'+denM1Bar+'</div>';
    renderListsSection('listsM1', m1ListNames());

    // ---- Aba M2: gauge + composição do M2 + listas ----
    document.getElementById('gaugeRowM2').innerHTML =
      gaugeCardHTML('M2 — Ações compartilhadas',
        'Ações compartilhadas ÷ ações realizadas × 100',
        d.m2, 8, CLASS_BANDS_M2, 'needle-m2tab-m2', fmtDec(d.m2,2)+'<span class="unit">%</span>', d.classificacaoM2,
        fmtInt(d.numeradorM2)+' compartilhadas ÷ '+fmtInt(d.denominadorM2)+' ações', LEGEND_M2);
    document.getElementById('compRowM2').innerHTML =
        '<div class="card comp-card"><h4>Numerador do M2</h4>'+numM2Bar+'</div>'
      + '<div class="card comp-card"><h4>Denominador do M2</h4>'+denM2Bar+'</div>';
    renderListsSection('listsM2', m2ListNames());

    // Tendência mês a mês: cada ponto é o M1/M2 calculado com sua própria
    // janela móvel de JANELA_MESES meses terminando naquele mês (ver
    // calcularSerieTendencia) — não é mais o histórico de vezes que a
    // página foi atualizada.
    var trend = '';
    trend += '<div class="card trend-card"><h4>M1 mês a mês</h4>'
      + '<p class="cur">Mês de referência ('+refMonthLabel()+'): '+fmtDec(d.m1,2)+'</p>'
      + sparkline(serieTendencia.map(function(p){ return {y:p.m1, label:monthShortLabel(p.mes), value:fmtDec(p.m1,2), quadKey:quadKeyOfDate(p.mes), quadLabel:quadCode(p.mes)}; }).filter(function(p){return p.y!=null;}), '#153F35', {quadAvg:true, classify:classificarM1})
      + '<p class="footnote">Cada ponto já é a janela de '+JANELA_MESES+' meses terminando naquele mês. Linha tracejada = média do quadrimestre no período exibido.</p>'
      + '</div>';
    trend += '<div class="card trend-card"><h4>M2 (%) mês a mês</h4>'
      + '<p class="cur">Mês de referência ('+refMonthLabel()+'): '+fmtDec(d.m2,2)+'%</p>'
      + sparkline(serieTendencia.map(function(p){ return {y:p.m2, label:monthShortLabel(p.mes), value:fmtDec(p.m2,2)+'%', quadKey:quadKeyOfDate(p.mes), quadLabel:quadCode(p.mes)}; }).filter(function(p){return p.y!=null;}), '#C68A3D', {quadAvg:true, suffix:'%', classify:classificarM2})
      + '<p class="footnote">Cada ponto já é a janela de '+JANELA_MESES+' meses terminando naquele mês. Linha tracejada = média do quadrimestre no período exibido.</p>'
      + '</div>';
    document.getElementById('trendRow').innerHTML = trend;
    setupTrendInteractivity();

    // Série histórica (tabela): um mês por linha, cada um já calculado com
    // sua própria janela móvel de JANELA_MESES meses (mesmos pontos do
    // sparkline acima) — mostra numerador/denominador/classificação de
    // M1 e M2 e o "Desempenho quadrimestral" (síntese M1×6 + M2×4) mês a
    // mês, pra dar visibilidade à composição por trás de cada ponto do
    // gráfico.
    var trendHistoryRows = serieTendencia.map(function(p){
      var classeM1 = classificarM1(p.m1);
      var classeM2 = classificarM2(p.m2);
      var desemp = classificarDesempenho(p.notaFinal);
      function pill(txt){ return '<span class="pill" style="background:'+(CLASS_PILL_HEX[txt]||'#7A8A82')+'">'+escapeHtml(txt)+'</span>'; }
      return '<tr>'
        + '<td>'+escapeHtml(monthShortLabel(p.mes))+'</td>'
        + '<td>'+fmtInt(p.numeradorM1)+'</td>'
        + '<td>'+fmtInt(p.denominadorM1)+'</td>'
        + '<td>'+(p.m1!=null ? fmtDec(p.m1,2) : '—')+'</td>'
        + '<td>'+pill(classeM1)+'</td>'
        + '<td>'+fmtInt(p.numeradorM2)+'</td>'
        + '<td>'+fmtInt(p.denominadorM2)+'</td>'
        + '<td>'+(p.m2!=null ? fmtDec(p.m2,2)+'%' : '—')+'</td>'
        + '<td>'+pill(classeM2)+'</td>'
        + '<td>'+(p.notaFinal!=null ? fmtDec(p.notaFinal,2) : '—')+'</td>'
        + '<td>'+pill(desemp)+'</td>'
        + '</tr>';
    }).join('');
    document.getElementById('trendHistoryWrap').innerHTML =
        '<div class="card"><h4 style="margin:0 0 4px;font-size:14.5px;font-weight:500;">Série histórica — numerador, denominador e desempenho quadrimestral</h4>'
      + '<p class="footnote" style="margin:0 0 12px;">Um mês por linha, cada um com sua própria janela móvel de '+JANELA_MESES+' meses terminando naquele mês (mesmos pontos dos gráficos acima). "Desempenho quadrimestral" é a síntese própria M1×6 + M2×4 — ver Notas Metodológicas.</p>'
      + '<div class="table-wrap"><table class="data-table"><thead><tr>'
      +   '<th>Mês</th><th>Numerador M1</th><th>Denominador M1</th><th>M1</th><th>Classe M1</th>'
      +   '<th>Numerador M2</th><th>Denominador M2</th><th>M2 (%)</th><th>Classe M2</th><th>Nota do desempenho</th><th>Desempenho quadrimestral</th>'
      + '</tr></thead><tbody>'+trendHistoryRows+'</tbody></table></div></div>';

    var notesList = document.getElementById('notesList');
    if(record.notes && record.notes.length){
      notesList.innerHTML = record.notes.map(function(n){ return '<li>'+escapeHtml(n)+'</li>'; }).join('');
    } else {
      notesList.innerHTML = '<li style="list-style:none;margin-left:-20px;">Nenhuma nota disponível para esta leitura.</li>';
    }

    animateGauges();
    renderHistoryList();
  }

  // ---------- Storage ----------
  function loadHistoryArray(){
    return window.__historyCache || [];
  }

  function refreshHistoryFromStorage(cb){
    if(!STORAGE_AVAILABLE){
      window.__historyCache = memoryHistory;
      if(cb) cb(memoryHistory);
      return;
    }
    window.storage.get(STORAGE_KEY, false).then(function(res){
      var arr = [];
      if(res && res.value){
        try{ arr = JSON.parse(res.value); }catch(e){ arr = []; }
      }
      window.__historyCache = arr;
      if(cb) cb(arr);
    }).catch(function(){
      window.__historyCache = [];
      if(cb) cb([]);
    });
  }

  function saveHistoryArray(arr){
    window.__historyCache = arr;
    memoryHistory = arr;
    if(!STORAGE_AVAILABLE) return Promise.resolve();
    try{
      return window.storage.set(STORAGE_KEY, JSON.stringify(arr), false).catch(function(){});
    }catch(e){
      return Promise.resolve();
    }
  }

  function renderHistoryList(){
    var arr = loadHistoryArray().slice().sort(function(a,b){ return b.timestamp-a.timestamp; });
    var el = document.getElementById('historyList');
    var clearBtn = document.getElementById('clearHistory');
    if(!arr.length){
      el.innerHTML = '<p class="history-empty">Nenhuma leitura ainda.</p>';
      clearBtn.style.display = 'none';
      return;
    }
    clearBtn.style.display = 'block';
    el.innerHTML = arr.map(function(h){
      var dotColor = h.error ? '#9AA69E' : pillHex(h.data && h.data.desempenho);
      return '<div class="history-item'+(h.id===currentRecordId?' active':'')+'" data-id="'+h.id+'">'
        + '<span class="history-dot" style="background:'+dotColor+'"></span>'
        + '<span class="history-text"><span class="eq">'+escapeHtml(h.equipe)+'</span><span class="dt">'+fmtDate(h.timestamp)+'</span></span>'
        + '<button class="history-del" data-del="'+h.id+'" title="Remover">×</button>'
        + '</div>';
    }).join('');

    el.querySelectorAll('.history-item').forEach(function(item){
      item.addEventListener('click', function(e){
        if(e.target.classList.contains('history-del')) return;
        var id = item.getAttribute('data-id');
        var rec = loadHistoryArray().find(function(h){ return h.id === id; });
        if(rec){
          currentRecordId = id;
          renderDashboard(rec, calcularSerieTendencia(latestWb, anchorMonthDate(), TREND_MESES));
        }
      });
    });
    el.querySelectorAll('.history-del').forEach(function(btn){
      btn.addEventListener('click', function(e){
        e.stopPropagation();
        var id = btn.getAttribute('data-del');
        var arr2 = loadHistoryArray().filter(function(h){ return h.id !== id; });
        saveHistoryArray(arr2).then(function(){
          if(id === currentRecordId && arr2.length){
            var latest = arr2.slice().sort(function(a,b){return b.timestamp-a.timestamp;})[0];
            currentRecordId = latest.id;
            renderDashboard(latest);
          } else {
            renderHistoryList();
          }
        });
      });
    });
  }

  document.getElementById('clearHistory').addEventListener('click', function(){
    if(!confirm('Remover todo o histórico de leituras deste navegador?')) return;
    saveHistoryArray([]).then(function(){
      currentRecordId = null;
      renderHistoryList();
    });
  });

  // ---------- Fetch ----------
  var fetchStatusEl = document.getElementById('fetchStatus');
  var refreshBtn = document.getElementById('refreshBtn');
  var refreshLabel = document.getElementById('refreshLabel');

  function sameData(a,b){
    if(!a || !b) return false;
    return JSON.stringify(a) === JSON.stringify(b);
  }

  // Dados brutos (já filtrados pela equipe atual, mas SEM filtro de
  // período — o período é aplicado depois, em calcularIndicadoresDoPeriodo)
  // guardados aqui após o último fetch bem-sucedido. Trocar o "Mês de
  // referência" no seletor reusa este cache e recalcula tudo na hora, sem
  // precisar buscar a planilha de novo na rede.
  var latestWb = null;

  // Recalcula M1/M2/pontos/nota + a série de tendência pro mês de
  // referência atual (refMonthDates), a partir do cache latestWb.
  // saveHistory=true (usado logo após um fetch): grava uma nova "leitura"
  // no histórico se os dados mudaram desde a última do mesmo período/equipe.
  // saveHistory=false (usado ao trocar o seletor de mês): só recalcula e
  // renderiza na hora, sem criar entrada nova no histórico de leituras.
  function aplicarMesReferencia(saveHistory){
    if(!latestWb) return;

    var extracted, periodo;
    if(refMonthDates.length === 1){
      // Um único mês escolhido: NÃO é só aquele mês isolado — é a janela
      // móvel de JANELA_MESES meses TERMINANDO nesse mês (ex.: maio →
      // fev, mar, abr e maio, incluindo os dois extremos), a mesma janela
      // usada pela série de tendência (ver calcularJanelaPeriodo).
      var janelaMes = calcularJanelaPeriodo(refMonthDates[0]);
      extracted = calcularIndicadoresDoPeriodo(latestWb, janelaMes);
      periodo = {inicio: fmtBRDate(janelaMes.inicio), fim: fmtBRDate(janelaMes.fim)};
    } else if(refMonthDates.length > 1){
      // Vários meses escolhidos: o M1/M2 de CADA mês marcado já é o valor
      // com a janela móvel de JANELA_MESES meses terminando naquele mês
      // (mesma regra do mês único, acima) — os resultados dos meses
      // marcados entram na MÉDIA (mesmo princípio da média do
      // quadrimestre, ver mediaDeMeses), e os totais de contexto/"Pessoas
      // atendidas" somam o mês isolado (sem janela) de cada um, pra não
      // sobrepor dados de meses vizinhos quando as janelas se cruzam.
      var resultadosMensaisSel = refMonthDates.map(function(m){
        return calcularIndicadoresDoPeriodo(latestWb, periodoMesUnico(m));
      });
      var resultadosJanelaSel = refMonthDates.map(function(m){
        return calcularIndicadoresDoPeriodo(latestWb, calcularJanelaPeriodo(m));
      });
      extracted = mediaDeMeses(resultadosMensaisSel, resultadosJanelaSel);
      periodo = {
        inicio: fmtBRDate(periodoMesUnico(refMonthDates[0]).inicio),
        fim: fmtBRDate(periodoMesUnico(refMonthDates[refMonthDates.length-1]).fim)
      };
    } else {
      // Nenhum mês escolhido: média dos 4 meses do quadrimestre selecionado.
      // O M1/M2 de CADA mês usado na média já é o valor com a janela móvel
      // de JANELA_MESES meses terminando naquele mês (mesma regra oficial
      // usada na seleção de mês individual e no gráfico de tendência) —
      // por isso calculamos cada mês duas vezes: uma com a janela (pra
      // entrar na média de M1/M2) e outra isolada, só o mês em si (pra
      // somar contagens de contexto e montar "Pessoas atendidas" sem
      // sobrepor dados de meses vizinhos).
      var meses = mesesDoQuadrimestre(quadSelecionado.ano, quadSelecionado.qIndex);
      var resultadosMensais = meses.map(function(m){
        return calcularIndicadoresDoPeriodo(latestWb, periodoMesUnico(m));
      });
      var resultadosJanela = meses.map(function(m){
        return calcularIndicadoresDoPeriodo(latestWb, calcularJanelaPeriodo(m));
      });
      extracted = mediaDeMeses(resultadosMensais, resultadosJanela);
      periodo = {
        inicio: fmtBRDate(periodoMesUnico(meses[0]).inicio),
        fim: fmtBRDate(periodoMesUnico(meses[3]).fim)
      };
    }

    populateSheetsCache(latestWb);
    // "Pessoas atendidas" agora NÃO usa mais extracted.pessoasAtendidas
    // (ligado ao filtro de Mês do topo) — a lista, na aba Listas, é
    // recalculada direto por renderListCard/pessoasAtendidasParaMeses,
    // com o próprio filtro de mês (ver renderListsSection).

    var serie = calcularSerieTendencia(latestWb, anchorMonthDate(), TREND_MESES);

    if(quadMs) quadMs.setSelected([quadSelecionado.ano+'-'+quadSelecionado.qIndex]);
    if(mesMs) mesMs.setSelected(refMonthDates.map(monthOptionValue));
    var winEl = document.getElementById('refWindowLabel');
    if(winEl){
      winEl.innerHTML = refMonthDates.length
        ? 'Resultado de <b>'+refMonthLabel()+'</b> — janela de '+JANELA_MESES+' meses cada ('+periodo.inicio+' a '+periodo.fim+')'
        : 'Média de <b>'+QUAD_LABELS[quadSelecionado.qIndex]+'/'+quadSelecionado.ano+'</b> ('+periodo.inicio+' a '+periodo.fim+')';
    }

    if(!saveHistory){
      var base = currentRecordId ? loadHistoryArray().find(function(h){ return h.id === currentRecordId; }) : null;
      var record = {
        id: base ? base.id : 'tmp',
        timestamp: base ? base.timestamp : Date.now(),
        equipe: extracted.equipe,
        data: extracted.data,
        notes: extracted.notes,
        periodo: periodo
      };
      renderDashboard(record, serie);
      return;
    }

    var now = Date.now();
    var history = loadHistoryArray();
    var lastForEquipe = history.filter(function(h){
      return !h.error && h.equipe === extracted.equipe
        && h.periodo && h.periodo.inicio === periodo.inicio && h.periodo.fim === periodo.fim;
    }).sort(function(a,b){ return b.timestamp-a.timestamp; })[0];

    if(lastForEquipe && sameData(lastForEquipe.data, extracted.data)){
      currentRecordId = lastForEquipe.id;
      fetchStatusEl.textContent = 'Dados sem alterações desde a última leitura.';
      renderDashboard(lastForEquipe, serie);
      return;
    }

    var record = {
      id: 'u'+now+Math.random().toString(36).slice(2,7),
      timestamp: now,
      equipe: extracted.equipe,
      data: extracted.data,
      notes: extracted.notes,
      periodo: periodo
    };
    var arr = loadHistoryArray();
    arr.push(record);
    saveHistoryArray(arr).then(function(){
      currentRecordId = record.id;
      fetchStatusEl.textContent = 'Planilha lida e calculada com sucesso.';
      renderDashboard(record, serie);
    });
  }

  function fetchAndLoad(){
    refreshBtn.classList.add('loading');
    refreshBtn.disabled = true;
    refreshLabel.textContent = 'Atualizando…';
    fetchStatusEl.textContent = 'Buscando dados…';
    fetchStatusEl.className = 'fetch-status';

    fetchAllSheets()
      .then(function(results){
        var faltando = results.filter(function(r){ return !r.ok; });
        if(faltando.length){
          throw new Error('Não foi possível ler a(s) aba(s) "' + faltando.map(function(r){return r.name;}).join('", "')
            + '" (verifique se elas ainda existem com esse nome e se a planilha está com acesso "qualquer pessoa com o link pode visualizar").');
        }

        var wb = {SheetNames:[], Sheets:{}};
        results.forEach(function(r){
          var parsedRows = parseCsv(r.csvText);
          if(!parsedRows.length) return;
          // r.name é o nome REAL da aba (sem sufixo). Filtra as linhas pela
          // equipe selecionada e guarda no workbook sob a chave "sufixada"
          // — o resto do painel (cálculo, listas) continua lendo por essa
          // chave, sem precisar saber que a aba é compartilhada entre
          // equipes. Note: SEM filtro de período aqui — cada mês de
          // referência filtra por data na hora, em aplicarMesReferencia().
          var filtradas = filtrarLinhasPorEquipe(parsedRows, currentEquipes);
          var key = suffixedName(r.name);
          wb.Sheets[key] = filtradas;
          wb.SheetNames.push(key);
        });

        latestWb = wb;
        aplicarMesReferencia(true);
      })
      .catch(function(err){
        var msg = (err && err.message) ? err.message : 'verifique sua conexão e o link publicado.';
        fetchStatusEl.textContent = 'Não foi possível ler a planilha: ' + msg;
        fetchStatusEl.className = 'fetch-status err';
        var history = loadHistoryArray();
        if(!history.length){
          document.getElementById('statusState').innerHTML =
            '<h2>Não foi possível carregar</h2><p>' + escapeHtml(msg) + '</p>'
            + '<div class="ficha"><b>Verifique:</b> se a planilha continua publicada em "Arquivo → Compartilhar → Publicar na web" (incluindo todas as abas) e se o link ainda é válido.</div>'
            + '<button class="retry-btn" id="retryBtn">Tentar de novo</button>';
          var retry = document.getElementById('retryBtn');
          if(retry) retry.addEventListener('click', fetchAndLoad);
        }
      })
      .finally(function(){
        refreshBtn.classList.remove('loading');
        refreshBtn.disabled = false;
        refreshLabel.textContent = 'Atualizar agora';
      });
  }

  refreshBtn.addEventListener('click', fetchAndLoad);

  function renderEquipeSwitcher(){
    var equipeMs = createMultiSelect(document.getElementById('equipeMs'), {
      placeholder: 'Selecione',
      multi: true,
      search: false,
      showTags: true,
      onChange: function(keys){
        if(keys.length===0){
          // sempre precisa ficar pelo menos 1 equipe marcada
          equipeMs.setSelected(currentEquipes.map(function(e){ return e.key; }));
          return;
        }
        currentEquipes = EQUIPES.filter(function(eq){ return keys.indexOf(eq.key)>=0; });
        document.getElementById('statusState').style.display = '';
        fetchAndLoad();
      }
    });
    equipeMs.setOptions(EQUIPES.map(function(eq){ return {value: eq.key, label: eq.label}; }));
    equipeMs.setSelected(currentEquipes.map(function(e){ return e.key; }));
  }
  renderEquipeSwitcher();

  // ---------- Init ----------
  refreshHistoryFromStorage(function(arr){
    if(arr.length){
      var latest = arr.slice().sort(function(a,b){ return b.timestamp-a.timestamp; })[0];
      currentRecordId = latest.id;
      renderDashboard(latest);
    }
    fetchAndLoad();
  });
})();
