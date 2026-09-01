(function () {
  'use strict';

  const PHASES = {
    "1": ["A fazenda", "contexto"],
    "2": ["Pré-parto", "antes de nascer"],
    "3": ["Maternidade", "o parto"],
    "4": ["Colostragem", "primeiras horas"],
    "5": ["Nutrição", "até o desaleitamento"],
    "6": ["Sanidade", "doença e morte"],
    "7": ["Ambiente", "onde vive"],
    "8": ["Recria", "até a novilha"],
    "9": ["Custos", "o que custa"]
  };

  const NO_CHART_NUMERIC_IDS = new Set([
    "25", "26", "27", "31", "32", "33", "34", "49", "50", "51", "71",
    "92", "127", "153", "154", "155", "162", "163", "164", "176", "182", "184"
  ]);
  const FORCE_BAR_IDS = new Set(["84", "120", "172"]);
  const FULL_WIDTH_IDS = new Set(["44", "83", "103", "109"]);
  const HORIZONTAL_BAR_IDS = new Set(["41"]);
  const HIDDEN_QUESTION_IDS = new Set(["179", "180"]);
  const VACCINE_QUESTION_ID = "44";

  let data = null;
  let chapters = [];
  let currentChapter = 0;
  let currentView = 'intro';
  let query = '';

  const el = id => document.getElementById(id);
  const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, char => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[char]));
  const formatNumber = value => Number(value).toLocaleString('pt-BR', { maximumFractionDigits: 1 });
  const formatInteger = value => Math.round(Number(value) || 0).toLocaleString('pt-BR', { maximumFractionDigits: 0 });
  const formatPercent = value => `${Number(value || 0).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
  const plural = (n, singular, pluralForm) => n === 1 ? singular : pluralForm;
  const isNoChartNumeric = question => NO_CHART_NUMERIC_IDS.has(String(question.num));
  const metricValue = (question, value) => isNoChartNumeric(question) ? formatInteger(value) : formatNumber(value);

  if (window.ALTA_CRIA_DATA) {
    data = window.ALTA_CRIA_DATA;
    start();
  } else {
    fetch('dados.json')
      .then(response => {
        if (!response.ok) throw new Error('Arquivo de dados indisponível');
        return response.json();
      })
      .then(result => {
        data = result;
        start();
      })
      .catch(() => {
        el('vazio').hidden = false;
        el('vazio').innerHTML = '<b>Os dados ainda não foram publicados.</b><br>Assim que as fazendas responderem, os resultados aparecem aqui.';
      });
  }

  function start() {
    chapters = (Array.isArray(data.capitulos) ? data.capitulos : [])
      .filter(chapter => String(chapter.idx) !== '10')
      .map(chapter => ({
        ...chapter,
        perguntas: (chapter.perguntas || []).filter(question => !HIDDEN_QUESTION_IDS.has(String(question.num)))
      }));

    const brand = document.querySelector('.marca img');
    if (brand) {
      brand.src = 'assets/logo-alta-cria.png';
      brand.className = 'program-header-logo';
      brand.alt = 'Programa Alta CRIA';
      const altaLogo = document.createElement('img');
      altaLogo.src = 'assets/logo-alta-branca.png';
      altaLogo.alt = 'Alta Genetics';
      altaLogo.className = 'alta-header-logo';
      brand.insertAdjacentElement('afterend', altaLogo);
    }

    el('demo').hidden = true;

    el('nFaz').textContent = formatInteger(data.n_fazendas);
    const generatedAt = new Date(data.gerado_em);
    el('quando').textContent = Number.isNaN(generatedAt.getTime())
      ? ''
      : `atualizado em ${generatedAt.toLocaleDateString('pt-BR')}`;

    if (!chapters.length) {
      el('vazio').hidden = false;
      el('vazio').innerHTML = '<b>Ainda não há respostas suficientes.</b>';
      return;
    }

    buildNavigation();
    el('busca').addEventListener('input', event => {
      query = event.target.value.trim().toLocaleLowerCase('pt-BR');
      currentView = 'data';
      render();
    });
    el('voltarInicio').addEventListener('click', () => {
      currentView = 'intro';
      query = '';
      el('busca').value = '';
      render();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
    render();
  }

  function buildNavigation() {
    const chapterButtons = chapters.map((chapter, index) => {
      const phase = PHASES[chapter.idx] || [chapter.titulo, ''];
      return `<button class="fase" role="tab" data-chapter="${index}" aria-selected="false">
        <span class="et">${escapeHtml(phase[0])}</span>
      </button>`;
    }).join('');

    el('trilha').innerHTML = `${chapterButtons}
      <button class="fase autores-tab" role="tab" data-authors="true" aria-selected="false">
        <span class="et">Autores</span>
      </button>`;

    el('trilha').querySelectorAll('[data-chapter]').forEach(button => {
      button.addEventListener('click', () => openData(Number(button.dataset.chapter)));
    });

    el('trilha').querySelector('[data-authors]').addEventListener('click', () => {
      currentView = 'authors';
      query = '';
      el('busca').value = '';
      render();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }

  function openData(chapterIndex = 0) {
    currentView = 'data';
    currentChapter = chapterIndex;
    query = '';
    el('busca').value = '';
    render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function visibleQuestions() {
    if (!query) return chapters[currentChapter].perguntas;
    const matches = [];
    chapters.forEach(chapter => chapter.perguntas.forEach(question => {
      const searchable = `${question.titulo} ${question.num}`.toLocaleLowerCase('pt-BR');
      if (searchable.includes(query)) matches.push(question);
    }));
    return matches;
  }

  function render() {
    el('trilha').querySelectorAll('.fase').forEach(button => {
      const selected = currentView === 'authors'
          ? button.hasAttribute('data-authors')
          : !query && Number(button.dataset.chapter) === currentChapter;
      button.setAttribute('aria-selected', String(selected));
    });

    document.body.classList.toggle('intro-mode', currentView === 'intro');
    if (currentView === 'intro') {
      renderIntroduction();
      return;
    }
    if (currentView === 'authors') {
      renderAuthors();
      return;
    }

    el('busca').closest('.busca').hidden = false;
    el('grade').className = 'grade';

    const questions = visibleQuestions();
    if (query) {
      el('tituloCap').textContent = 'Resultados da busca';
    } else {
      const chapter = chapters[currentChapter];
      el('tituloCap').textContent = chapter.titulo;
    }
    el('contCap').textContent = '';

    el('vazio').hidden = questions.length > 0;
    if (!questions.length) {
      el('vazio').innerHTML = '<b>Nada encontrado.</b><br>Tente outra palavra.';
      el('grade').innerHTML = '';
      return;
    }

    el('grade').innerHTML = questions.map(questionCard).join('');
    bindSortButtons();
  }

  function renderIntroduction() {
    el('contCap').textContent = '';
    el('busca').closest('.busca').hidden = true;
    el('vazio').hidden = true;
    el('grade').className = 'intro-view';
    el('grade').innerHTML = `<section class="welcome-hero">
      <div class="welcome-brand-row">
        <img src="assets/logo-alta-cria.png" alt="Programa Alta CRIA" class="welcome-program-logo">
        <span class="welcome-divider" aria-hidden="true"></span>
        <img src="assets/logo-alta-branca.png" alt="Alta Genetics" class="welcome-alta-logo">
      </div>
      <div class="welcome-kicker">Caderno de Dados Anuais 2025/2026</div>
      <h2>Bem-vindo ao universo de dados do Programa Alta CRIA</h2>
      <p class="welcome-deck">Informação técnica, visão nacional e indicadores para apoiar decisões mais eficientes na criação de bezerras e novilhas leiteiras.</p>
      <div class="welcome-actions">
        <button class="access-data" id="accessData" type="button">Acessar os dados do questionário <span aria-hidden="true">→</span></button>
        <button class="access-data access-data-secondary" id="accessAnimals" type="button" disabled aria-disabled="true">Acessar os dados dos animais <span class="access-data-tag" aria-hidden="true">Em breve</span></button>
      </div>
    </section>

    <section class="welcome-message" aria-labelledby="welcome-title">
      <div class="message-marker" aria-hidden="true">“</div>
      <div class="message-copy">
        <p>A criação de bezerras e novilhas é uma das etapas mais estratégicas da pecuária leiteira, pois garante a reposição genética do rebanho e contribui para o desenvolvimento de animais mais produtivos e saudáveis.</p>
        <p>O gerenciamento de indicadores e a compreensão dos principais índices zootécnicos são fundamentais para definir metas, planejar estratégias e alcançar os objetivos da propriedade. Esse acompanhamento é essencial para facilitar a tomada de decisão, tornar os sistemas mais eficientes e lucrativos e identificar pontos que necessitam maior atenção.</p>
        <p>Com esse propósito, em 2017 foi criado o Programa Alta CRIA, que tem como objetivo coletar, organizar e analisar os principais índices zootécnicos e sanitários das fases de cria e recria. O programa gera informações e dados em âmbito nacional para auxiliar produtores e técnicos na gestão dos sistemas de criação.</p>
        <p>A interpretação dos resultados conta com um grupo de conselheiros especializados, que analisam tendências e inovações no setor. As devolutivas são disponibilizadas por meio de materiais técnicos como o Padrão Ouro de Criação de Bezerras e Novilhas Leiteiras, o Caderno de Dados Anuais e o Livro Perguntas e Respostas Alta CRIA. Além disso, o programa reúne uma ampla rede de técnicos e proprietários de fazendas comerciais em diversas regiões do Brasil.</p>
        <p>O Caderno de Dados Anuais Alta CRIA 2025/2026 foi desenvolvido a partir de questionários online respondidos por 187 fazendas e dos dados enviados por 184 fazendas participantes do programa. Os cálculos e resultados apresentados foram gerados com base nessas informações, analisadas de forma descritiva, e se referem a bezerras nascidas entre 1º de julho de 2025 e 30 de junho de 2026.</p>
        <span class="message-marker-close" aria-hidden="true">”</span>
      </div>
      <div class="message-signature">
        <img src="assets/autores/autor-rafael-azevedo-original.jpg" alt="Rafael Azevedo">
        <div>
          <strong>Rafael Azevedo</strong>
          <span>Gerente de Programas e Serviços Leite Alta</span>
          <span>Coordenador e Conselheiro do programa Alta CRIA</span>
          <span>Coordenador do Padrão Ouro de Criação de Bezerras e Novilhas Leiteiras</span>
        </div>
      </div>
    </section>

    <section class="thanks-panel">
      <div class="thanks-number"><strong>184</strong><span>fazendas enviaram seus dados</span></div>
      <div class="thanks-copy">
        <h3>Nosso agradecimento</h3>
        <p>Agradecemos às 184 fazendas que enviaram seus dados e às 187 fazendas que preencheram o questionário online no ciclo Alta CRIA 2025/2026.</p>
        <p>Este programa só é possível graças à confiança depositada em nossa equipe, que nos permite realizar análises consistentes e trabalhar pela padronização dos índices de cria e recria na pecuária leiteira brasileira.</p>
      </div>
    </section>

    <section class="sponsors" aria-labelledby="sponsor-title">
      <div class="sponsor-heading"><span id="sponsor-title">Parceiros do Programa</span></div>
      <div class="sponsor-tier master-tier">
        <h4>Patrocínio Master</h4>
        <div class="logo-slots"><div class="logo-slot master"><span>Logo Master</span></div><div class="logo-slot master"><span>Logo Master</span></div></div>
      </div>
      <div class="sponsor-tier">
        <h4>Patrocínio</h4>
        <div class="logo-slots four"><div class="logo-slot"><span>Logo</span></div><div class="logo-slot"><span>Logo</span></div><div class="logo-slot"><span>Logo</span></div><div class="logo-slot"><span>Logo</span></div></div>
      </div>
      <div class="sponsor-tier support-tier">
        <h4>Apoio</h4>
        <div class="logo-slots five"><div class="logo-slot compact"><span>Logo</span></div><div class="logo-slot compact"><span>Logo</span></div><div class="logo-slot compact"><span>Logo</span></div><div class="logo-slot compact"><span>Logo</span></div><div class="logo-slot compact"><span>Logo</span></div></div>
      </div>
    </section>`;

    el('accessData').addEventListener('click', () => openData(0));
  }

  function questionCard(question) {
    const optionCount = question.opcoes?.length || 0;
    const longLabel = question.opcoes?.some(option => String(option.rotulo).length > 42);
    const isKpiOnly = isNoChartNumeric(question);
    const isWide = isKpiOnly || FULL_WIDTH_IDS.has(String(question.num)) || question.tipo === 'multipla' || optionCount > 8 || longLabel;
    const header = `<h3>${escapeHtml(question.titulo)}</h3>
      <div class="analysis-meta">
        <span>${formatInteger(question.n)} ${plural(question.n, 'fazenda respondeu', 'fazendas responderam')}</span>
      </div>`;
    const content = question.tipo === 'numerico' || question.tipo === 'escala'
      ? numericContent(question)
      : categoricalContent(question);
    const reading = isKpiOnly ? '' : quickRead(question);
    return `<article class="cartao${isWide ? ' wide' : ''}${isKpiOnly ? ' kpi-only' : ''}" data-question-id="${escapeHtml(question.num)}">${header}${reading}${content}</article>`;
  }

  function quickRead(question) {
    if ((question.tipo === 'numerico' || question.tipo === 'escala') && question.n) {
      return `<div class="quick-read"><strong>Leitura rápida:</strong> mediana de ${metricValue(question, question.mediana)}, com 50% das respostas entre ${metricValue(question, question.p25)} e ${metricValue(question, question.p75)}.</div>`;
    }
    if (Array.isArray(question.opcoes) && question.opcoes.length) {
      const leader = question.opcoes.reduce((best, option) => Number(option.pct) > Number(best.pct) ? option : best, question.opcoes[0]);
      const prefix = question.tipo === 'multipla' ? 'Resposta mais citada' : 'Maior participação';
      return `<div class="quick-read"><strong>${prefix}:</strong> ${escapeHtml(leader.rotulo)} (${formatPercent(leader.pct)}).</div>`;
    }
    return '';
  }

  function categoricalContent(question) {
    if (!Array.isArray(question.opcoes) || !question.opcoes.length) return '<p class="analysis-meta">Sem respostas.</p>';

    const longLabel = question.opcoes.some(option => String(option.rotulo).length > 34);
    const useRanking = HORIZONTAL_BAR_IDS.has(String(question.num)) || (!FORCE_BAR_IDS.has(String(question.num)) && (question.tipo === 'multipla' || question.opcoes.length > 7 || longLabel));
    const note = question.tipo === 'multipla'
      ? '<p class="multiple-note">Pergunta de múltipla resposta: a soma dos percentuais pode ultrapassar 100%.</p>'
      : '';

    return useRanking
      ? `${rankingChart(question)}${note}`
      : `${columnChart(question.opcoes)}${note}`;
  }

  function columnChart(options) {
    const maximum = Math.max(...options.map(option => Number(option.pct) || 0));
    const items = options.map((option, index) => {
      const percentage = Math.max(0, Math.min(100, Number(option.pct) || 0));
      const isMaximum = percentage === maximum;
      return `<div class="column-item${isMaximum ? ' is-max' : ''}" style="--bar-order:${index}" title="${escapeHtml(option.rotulo)}: ${formatPercent(percentage)} (${formatInteger(option.n)})">
        <div class="column-zone">
          <div class="column-bar" style="height:${Math.max(1.5, percentage)}%">
            <span class="column-value">${formatPercent(percentage)}</span>
          </div>
        </div>
        <div class="column-label">${escapeHtml(option.rotulo)}</div>
      </div>`;
    }).join('');
    return `<div class="column-chart" style="--column-count:${options.length}" role="img" aria-label="Gráfico de colunas com a distribuição percentual das respostas">${items}</div>`;
  }

  function rankingChart(question) {
    const rows = question.opcoes.map(option => {
      const percentage = Math.max(0, Math.min(100, Number(option.pct) || 0));
      return `<div class="rank-row" data-label="${escapeHtml(String(option.rotulo).toLocaleLowerCase('pt-BR'))}" data-value="${percentage}">
        <span class="rank-label">${escapeHtml(option.rotulo)}</span>
        <span class="rank-track"><span class="rank-fill" style="width:${percentage}%"></span></span>
        <span class="rank-value">${formatPercent(percentage)}</span>
      </div>`;
    }).join('');
    const sortControl = String(question.num) === VACCINE_QUESTION_ID
      ? `<button class="sort-heading" type="button" data-sort-cycle data-mode="desc" title="Clique para reorganizar as vacinas">
          <span>Vacinas</span><small class="sort-mode">Maior → menor</small><span aria-hidden="true">↕</span>
        </button>`
      : '';
    return `${sortControl}<div class="ranking" role="img" aria-label="Ranking percentual das respostas">${rows}</div>`;
  }

  function numericContent(question) {
    if (!question.n) return '<p class="analysis-meta">Sem respostas.</p>';

    const stats = `<div class="stats">
      <div class="stat"><span class="stat-label">25% das respostas</span><span class="stat-value">${metricValue(question, question.p25)}</span></div>
      <div class="stat featured"><span class="stat-label">Mediana</span><span class="stat-value">${metricValue(question, question.mediana)}</span></div>
      <div class="stat"><span class="stat-label">75% das respostas</span><span class="stat-value">${metricValue(question, question.p75)}</span></div>
      <div class="stat"><span class="stat-label">Média</span><span class="stat-value">${metricValue(question, question.media)}</span></div>
    </div>`;

    if (isNoChartNumeric(question) || !Array.isArray(question.hist) || !question.hist.length) return stats;
    const maxCount = Math.max(...question.hist.map(item => Number(item.n) || 0)) || 1;
    const items = question.hist.map((item, index) => {
      const height = Math.max(2, (Number(item.n) || 0) / maxCount * 100);
      return `<div class="column-item" style="--bar-order:${index}" title="${escapeHtml(item.faixa)}: ${formatInteger(item.n)} fazendas">
        <div class="column-zone">
          <div class="column-bar" style="height:${height}%"><span class="column-value">${formatInteger(item.n)}</span></div>
        </div>
        <div class="column-label">${escapeHtml(item.faixa)}</div>
      </div>`;
    }).join('');
    return `${stats}<div class="column-chart distribution" style="--column-count:${question.hist.length}" role="img" aria-label="Distribuição das respostas por faixa">${items}</div>`;
  }

  function bindSortButtons() {
    el('grade').querySelectorAll('[data-sort-cycle]').forEach(button => {
      button.addEventListener('click', () => {
        const modes = ['desc', 'asc', 'alpha'];
        const currentIndex = modes.indexOf(button.dataset.mode);
        const mode = modes[(currentIndex + 1) % modes.length];
        button.dataset.mode = mode;
        const labels = { desc: 'Maior → menor', asc: 'Menor → maior', alpha: 'A → Z' };
        button.querySelector('.sort-mode').textContent = labels[mode];
        const ranking = button.nextElementSibling;
        const rows = [...ranking.querySelectorAll('.rank-row')];
        rows.sort((a, b) => {
          if (mode === 'alpha') return a.dataset.label.localeCompare(b.dataset.label, 'pt-BR');
          const delta = Number(a.dataset.value) - Number(b.dataset.value);
          return mode === 'asc' ? delta : -delta;
        });
        rows.forEach(row => ranking.appendChild(row));
      });
    });
  }

  function renderAuthors() {
    const authors = Array.isArray(window.ALTA_CRIA_AUTHORS) ? window.ALTA_CRIA_AUTHORS : [];
    el('tituloCap').textContent = 'Autores e conselheiros';
    el('contCap').textContent = '';
    el('busca').closest('.busca').hidden = true;
    el('vazio').hidden = authors.length > 0;
    el('grade').className = 'grade authors-view';

    const lead = `<section class="authors-lead">
      <div><span class="section-kicker">Corpo técnico</span><h3>Conhecimento construído em conjunto</h3><p>Um grupo multidisciplinar reúne pesquisa, assistência técnica e experiência prática na criação de bezerras e novilhas leiteiras.</p></div>
      <div class="authors-count"><strong>${authors.length}</strong><span>autores e conselheiros</span></div>
    </section>`;

    const cards = authors.map(author => `<article class="author-card">
      <div class="author-photo-frame"><img class="author-photo" src="${escapeHtml(author.photo)}" alt="Retrato de ${escapeHtml(author.name)}" loading="lazy"></div>
      <div class="author-copy">
        <h3>${escapeHtml(author.name)}</h3>
        <p>${escapeHtml(author.bio)}</p>
        <a href="mailto:${escapeHtml(author.email)}">${escapeHtml(author.email)}</a>
      </div>
    </article>`).join('');

    el('grade').innerHTML = `${lead}<div class="authors-grid">${cards}</div>`;
  }
})();
