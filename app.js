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

    // "Sub" M1/M2: pro texto embaixo do gauge ("X atendimentos ÷ Y
    // pessoas"), usamos a MÉDIA das janelas móveis mensais (o mesmo
    // conjunto de números usado pra calcular m1/m2 acima) — não a
    // soma/média dos meses isolados, que não bate com o M1/M2 exibido.
    var subNumeradorM1 = Math.round(media('numeradorM1'));
    var subDenominadorM1 = Math.round(media('denominadorM1'));
    var subNumeradorM2 = Math.round(media('numeradorM2'));
    var subDenominadorM2 = Math.round(media('denominadorM2'));

    return {
      equipe: currentEquipes.map(function(e){ return e.label; }).join(' + '),
      data: {
        atendimentosIndividuais: soma('atendimentosIndividuais'),
        participacoesColetivas: soma('participacoesColetivas'),
        numeradorM1: soma('numeradorM1'),
        denominadorM1: soma('denominadorM1'),
        subNumeradorM1: subNumeradorM1,
        subDenominadorM1: subDenominadorM1,
        m1: m1,
        classificacaoM1: classificacaoM1,
        atividadesTotais: soma('atividadesTotais'),
        atividadesCompartilhadas: soma('atividadesCompartilhadas'),
        reunioesTotais: soma('reunioesTotais'),
        reunioesCompartilhadas: soma('reunioesCompartilhadas'),
        denominadorM2: soma('denominadorM2'),
        numeradorM2: soma('numeradorM2'),
        subNumeradorM2: subNumeradorM2,
        subDenominadorM2: subDenominadorM2,
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
        subNumeradorM1: numeradorM1,
        subDenominadorM1: denominadorM1,
        m1: m1,
        classificacaoM1: classificacaoM1,
        atividadesTotais: atividadesTotais,
        atividadesCompartilhadas: atividadesCompartilhadas,
        reunioesTotais: reunioesTotais,
        reunioesCompartilhadas: reunioesCompartilhadas,
        denominadorM2: denominadorM2,
        numeradorM2: numeradorM2,
        subNumeradorM2: numeradorM2,
        subDenominadorM2: denominadorM2,
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
