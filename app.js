'use strict';

/* ============================================================
   DEFINIÇÃO DAS COLUNAS
   Cada objeto descreve uma coluna da planilha.

   Campos:
     id       - letra da coluna (A, B, C...)
     key      - chave no objeto de dados
     label    - nome exibido no cabeçalho
     type     - 'text' | 'integer' | 'currency' | 'percent'
     editable - true = input editável | false = resultado calculado
     group    - 'primary' | 'param' | 'result'
     width    - largura em px
   ============================================================ */
const COLUMNS = [
  // ── Inputs primários (A–G) ─────────────────────────────────
  { id: 'A', key: 'CODIGO',       label: 'CÓDIGO',     type: 'integer', editable: true,  group: 'primary', width: 80  },
  { id: 'B', key: 'COD_FC',       label: 'COD FC',     type: 'integer', editable: true,  group: 'primary', width: 90  },
  { id: 'C', key: 'FANTAS',       label: 'FANTASIA',   type: 'text',    editable: true,  group: 'primary', width: 130 },
  { id: 'D', key: 'DESCRICAO',    label: 'DESCRIÇÃO',  type: 'text',    editable: true,  group: 'primary', width: 200 },
  { id: 'E', key: 'REFERENCIA',   label: 'REFERÊNCIA', type: 'text',    editable: true,  group: 'primary', width: 100 },
  { id: 'F', key: 'COD_EMPRESA',  label: 'EMPRESA',    type: 'integer', editable: true,  group: 'primary', width: 70  },
  { id: 'G', key: 'CUE',          label: 'CUE (R$)',   type: 'currency',editable: true,  group: 'primary', width: 95  },
  // ── Parâmetros (H–K) ──────────────────────────────────────
  { id: 'H', key: 'OPER_PERC',    label: 'OPER. %',    type: 'percent', editable: true,  group: 'param',   width: 75  },
  { id: 'I', key: 'ICMS_PERC',    label: 'ICMS %',     type: 'percent', editable: true,  group: 'param',   width: 75  },
  { id: 'J', key: 'PC_PERC',      label: 'PC %',       type: 'percent', editable: true,  group: 'param',   width: 75  },
  { id: 'K', key: 'MAR_PERC',     label: 'MAR. %',     type: 'percent', editable: true,  group: 'param',   width: 75  },
  // ── Resultados calculados (L–O) ───────────────────────────
  { id: 'L', key: 'VPC',          label: 'VPC (R$)',   type: 'currency',editable: true,  group: 'param',   width: 90  },
  { id: 'M', key: 'PRECO',        label: 'PREÇO (R$)', type: 'currency',editable: false, group: 'result',  width: 100 },
  { id: 'N', key: 'PRVD1',        label: 'PV D1 (R$)', type: 'currency',editable: true,  group: 'param',   width: 100 },
  { id: 'O', key: 'MARGEM_ATUAL', label: 'MARGEM %',        type: 'percent', editable: false, group: 'result',  width: 110 },
  { id: 'P', key: 'VERBA_LOJA',  label: 'VLR VERBA LOJA',  type: 'currency',editable: false, group: 'result',  width: 120 },
  { id: 'Q', key: 'VERBA_SITE',  label: 'VLR VERBA SITE',  type: 'currency',editable: false, group: 'result',  width: 120 },
];

/* Mapeamento de índice Excel → chave de dados para importação.
   null = coluna calculada (ignorada no import) */
const EXCEL_IMPORT_MAP = {
  0:  'CODIGO',
  1:  'COD_FC',
  2:  'FANTAS',
  3:  'DESCRICAO',
  4:  'REFERENCIA',
  5:  'COD_EMPRESA',
  6:  'CUE',
  7:  'OPER_PERC',
  8:  'ICMS_PERC',
  9:  'PC_PERC',
  10: 'MAR_PERC',
  11: 'VPC',  // L = VPC → editável, importado do Excel
  12: null,  // M = PRECO     → calculado
  13: 'PRVD1',  // N = PRVD1 → editável, importado do Excel
  14: null,  // O = MARGEM    → calculado
};

// Valores padrão dos parâmetros (editáveis por linha)
const DEFAULT_PARAMS = {
  OPER_PERC: 0.215,   // 21,5%
  ICMS_PERC: 0.205,   // 20,5%
  PC_PERC:   0.0925,  // 9,25%
  MAR_PERC:  0.05,    // 5,00%
  VPC:       0,       // R$ 0,00
};

// Estado principal da tabela
let tableData = [];

// Estado dos filtros
const filterState = { search: '', empresa: '', margem: '' };

function getFilteredIndices() {
  return tableData.reduce((acc, row, idx) => {
    if (filterState.search) {
      const q = filterState.search.toLowerCase();
      const hit = ['FANTAS', 'DESCRICAO', 'REFERENCIA', 'CODIGO', 'COD_FC'].some(k =>
        String(row[k] || '').toLowerCase().includes(q)
      );
      if (!hit) return acc;
    }
    if (filterState.empresa !== '' && String(row.COD_EMPRESA) !== filterState.empresa) return acc;
    if (filterState.margem) {
      const cls = classeMargem(row.MARGEM_ATUAL, num(row.MAR_PERC));
      const map = { positiva: 'margem-positiva', atencao: 'margem-atencao', negativa: 'margem-negativa' };
      if (cls !== map[filterState.margem]) return acc;
    }
    acc.push(idx);
    return acc;
  }, []);
}

/* ============================================================
   FUNÇÕES DE CÁLCULO
   ──────────────────────────────────────────────────────────
   COMO ADICIONAR/CORRIGIR UMA FÓRMULA:
   1. Encontre a função correspondente abaixo
   2. Substitua o bloco "TODO" pelo cálculo correto
   3. Salve o arquivo — os resultados atualizam automaticamente
   ============================================================ */

// L – VPC é editável pelo usuário; não há cálculo automático para este campo.

/**
 * M – PREÇO mínimo de venda
 *
 * Fórmula extraída da planilha:
 *   M = (G - L) / (1 - (K + H + J + I))
 *
 * Onde:
 *   G = CUE        (custo unitário estimado)
 *   L = VPC        (verba de preço de custo)
 *   H = OPER. %    (percentual operacional)
 *   I = ICMS %     (percentual de ICMS)
 *   J = PC %       (PIS + COFINS)
 *   K = MAR. %     (margem desejada)
 *
 * @param {Object} row
 * @returns {number|null}
 */
function calcularPreco(row) {
  const G = num(row.CUE);
  const L = num(row.VPC);
  const H = num(row.OPER_PERC);
  const I = num(row.ICMS_PERC);
  const J = num(row.PC_PERC);
  const K = num(row.MAR_PERC);

  const denominador = 1 - (K + H + J + I);

  if (Math.abs(denominador) < 0.0001) return null; // Evita divisão por zero

  return (G - L) / denominador;
}

// N – PRVD1 é editável pelo usuário; não há cálculo automático para este campo.

/**
 * O – MARGEM ATUAL
 *
 * Fórmula extraída da planilha:
 *   O = 1 - (H + I + J) - (G - L) / N
 *
 * Retorna null se N for zero ou vazio (impossível calcular).
 *
 * Onde:
 *   G = CUE        (custo)
 *   L = VPC        (verba de custo)
 *   H = OPER. %
 *   I = ICMS %
 *   J = PC %
 *   N = PRVD1      (preço de venda D1)
 *
 * @param {Object} row
 * @returns {number|null}
 */
function calcularMargem(row) {
  const G = num(row.CUE);
  const L = num(row.VPC);
  const H = num(row.OPER_PERC);
  const I = num(row.ICMS_PERC);
  const J = num(row.PC_PERC);
  const N = parseFloat(row.PRVD1);

  if (!N || N === 0) return null;

  return 1 - (H + I + J) - (G - L) / N;
}

/**
 * P – VLR VERBA LOJA
 *
 * Fórmula extraída da planilha:
 *   P = G - N * (1 - H - I - J - K)
 *
 * Representa a verba necessária para viabilizar o preço N na loja,
 * mantendo a margem K. Retorna null se o resultado for <= 0
 * (não há necessidade de verba nesse cenário).
 */
function calcularVerbaLoja(row) {
  const G = num(row.CUE);
  const H = num(row.OPER_PERC);
  const I = num(row.ICMS_PERC);
  const J = num(row.PC_PERC);
  const K = num(row.MAR_PERC);
  const N = num(row.PRVD1);

  const resultado = G - N * (1 - H - I - J - K);
  return resultado > 0 ? resultado : null;
}

/**
 * Q – VLR VERBA SITE
 *
 * Fórmula extraída da planilha:
 *   Q = G - N * (1 - 0.167 - I - J - K)
 *
 * Igual à verba loja, mas usa taxa operacional fixa de 16,7% para o site
 * (em vez do OPER. % configurável da loja).
 */
function calcularVerbaSite(row) {
  const G          = num(row.CUE);
  const I          = num(row.ICMS_PERC);
  const J          = num(row.PC_PERC);
  const K          = num(row.MAR_PERC);
  const N          = num(row.PRVD1);
  const OPER_SITE  = 0.167; // Taxa operacional fixa do canal site

  const resultado = G - N * (1 - OPER_SITE - I - J - K);
  return resultado > 0 ? resultado : null;
}

/**
 * Recalcula todos os campos derivados de uma linha.
 * VPC (L) e PRVD1 (N) são editáveis — não recalculados aqui.
 */
function calcularLinha(row) {
  row.PRECO        = calcularPreco(row);
  row.MARGEM_ATUAL = calcularMargem(row);
  row.VERBA_LOJA   = calcularVerbaLoja(row);
  row.VERBA_SITE   = calcularVerbaSite(row);
  return row;
}

/* ============================================================
   UTILITÁRIOS
   ============================================================ */

/** Converte para número, retornando 0 se inválido */
function num(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

/** Formata valor como moeda BRL */
function fmtMoeda(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = parseFloat(v);
  if (isNaN(n)) return '';
  return 'R$ ' + n.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

/** Formata valor como percentual (ex: 0.215 → "21,50%") */
function fmtPercent(v) {
  if (v === null || v === undefined || v === '') return '-';
  const n = parseFloat(v);
  if (isNaN(n)) return '-';
  return (n * 100).toFixed(2).replace('.', ',') + '%';
}

/** Formata valor para exibição de acordo com o tipo da coluna */
function fmtValue(v, type) {
  if (v === null || v === undefined || v === '') return '';
  switch (type) {
    case 'currency': return fmtMoeda(v);
    case 'percent':  return fmtPercent(v);
    case 'integer':  return String(Math.round(Number(v)));
    default:         return String(v);
  }
}

/**
 * Converte o valor digitado pelo usuário para armazenamento interno.
 * Percentuais: usuário digita "21,5" → armazenado como 0.215
 */
function parseInput(v, type) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim().replace(',', '.');
  if (s === '') return '';
  switch (type) {
    case 'percent': {
      const n = parseFloat(s);
      return isNaN(n) ? '' : n / 100; // Converte % para decimal
    }
    case 'currency': {
      const n = parseFloat(s);
      return isNaN(n) ? '' : n;
    }
    case 'integer': {
      const n = parseInt(s, 10);
      return isNaN(n) ? '' : n;
    }
    default: return s;
  }
}

/**
 * Converte valor interno para exibição no input.
 * Percentuais: 0.215 → "21.5" (o usuário edita como percentual)
 */
function inputDisplayValue(v, type) {
  if (v === null || v === undefined || v === '') return '';
  if (type === 'percent') {
    const n = parseFloat(v);
    if (isNaN(n)) return '';
    // Remove zeros desnecessários à direita
    return String(parseFloat((n * 100).toFixed(6)));
  }
  if (type === 'currency') {
    const n = parseFloat(v);
    return isNaN(n) ? '' : String(parseFloat(n.toFixed(6)));
  }
  return String(v);
}

/** Escapa HTML para uso em atributos */
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ============================================================
   CLASSIFICAÇÃO DE COR
   ============================================================ */

/**
 * Retorna a classe CSS para a célula de MARGEM ATUAL.
 *
 * Regras:
 *   margem >= meta (K)      → verde  (margem-positiva)
 *   0 <= margem < meta      → amarelo (margem-atencao)
 *   margem < 0              → vermelho (margem-negativa)
 *   null                    → cinza (margem-neutra)
 */
function classeMargem(margem, meta) {
  if (margem === null || margem === undefined) return 'margem-neutra';
  if (margem < 0)               return 'margem-negativa';
  if (margem >= (meta || 0))    return 'margem-positiva';
  return 'margem-atencao';
}

/* ============================================================
   RENDERIZAÇÃO
   ============================================================ */

/** Monta o cabeçalho da tabela (executado uma única vez) */
function renderHeader() {
  const thead = document.getElementById('table-head');

  // ── Linha 1: grupos ────────────────────────────────────────
  const primaryCols = COLUMNS.filter(c => c.group === 'primary').length;
  const paramCols   = COLUMNS.filter(c => c.group === 'param').length;
  const resultCols  = COLUMNS.filter(c => c.group === 'result').length;

  let groupRow = '<tr class="group-row">';
  groupRow += '<th class="group-action" rowspan="2"></th>'; // coluna #
  groupRow += `<th class="group-primary" colspan="${primaryCols}">Inputs Primários (A–G)</th>`;
  groupRow += `<th class="group-param"   colspan="${paramCols}">Parâmetros (H–K)</th>`;
  groupRow += `<th class="group-result"  colspan="${resultCols}">Resultados (L–O)</th>`;
  groupRow += '<th class="group-action" rowspan="2"></th>'; // botão excluir
  groupRow += '</tr>';

  // ── Linha 2: nomes das colunas ─────────────────────────────
  let colRow = '<tr class="col-row">';
  COLUMNS.forEach(col => {
    const w = col.width;
    colRow += `<th class="col-${col.group}" style="width:${w}px;min-width:${w}px" title="${col.label}">
      <span class="col-letter">${col.id}</span>
      <span class="col-name">${col.label}</span>
    </th>`;
  });
  colRow += '</tr>';

  thead.innerHTML = groupRow + colRow;
}

/** Renderiza todo o corpo da tabela */
function renderTable() {
  const tbody      = document.getElementById('table-body');
  const emptyState = document.getElementById('empty-state');

  if (tableData.length === 0) {
    tbody.innerHTML = '';
    emptyState.style.display = 'block';
    document.getElementById('stats-panel').style.display = 'none';
    document.getElementById('filter-count').textContent = '';
    popularFiltroEmpresa();
    return;
  }

  emptyState.style.display = 'none';
  popularFiltroEmpresa();

  const indices = getFilteredIndices();
  let html = '';
  indices.forEach((originalIdx, pos) => {
    html += buildRow(tableData[originalIdx], originalIdx, pos + 1);
  });

  tbody.innerHTML = html;

  const countEl = document.getElementById('filter-count');
  if (indices.length < tableData.length) {
    countEl.textContent = `${indices.length} de ${tableData.length} linha(s)`;
  } else {
    countEl.textContent = `${tableData.length} linha(s)`;
  }

  atualizarEstatisticas(indices);
  vincularEventos();
}

/** Constrói o HTML de uma única linha */
function buildRow(row, originalIdx, displayNum) {
  let html = `<tr data-row="${originalIdx}">`;
  html += `<td class="row-num">${displayNum}</td>`;

  COLUMNS.forEach(col => {
    if (col.editable) {
      const displayVal  = inputDisplayValue(row[col.key], col.type);
      const isText      = col.type === 'text';
      const placeholder = col.type === 'percent' ? '0' : '-';
      const title       = col.type === 'percent' ? `${col.label} (ex: 21.5 para 21,5%)` : col.label;
      const hasReplicate = col.key === 'PRVD1' || col.key === 'MAR_PERC';

      html += `<td class="col-${col.group}${hasReplicate ? ' td-prvd1' : ''}">
        <input
          class="cell-input${isText ? ' text-cell' : ''}"
          type="text"
          inputmode="${isText ? 'text' : 'decimal'}"
          data-row="${originalIdx}"
          data-col="${col.key}"
          data-type="${col.type}"
          value="${esc(displayVal)}"
          placeholder="${placeholder}"
          title="${title}"
        >${hasReplicate ? `<button class="btn-replicate" data-row="${originalIdx}" data-col="${col.key}" title="Replicar para todas as linhas">&#8595; todas</button>` : ''}
      </td>`;
    } else {
      html += buildResultCell(col, row);
    }
  });

  html += `<td style="padding:0;background:#fafafa;border:1px solid var(--border)">
    <button class="btn-del-row" data-row="${originalIdx}" title="Excluir linha">&times;</button>
  </td>`;

  html += '</tr>';
  return html;
}

/** Constrói o HTML de uma célula de resultado calculado */
function buildResultCell(col, row) {
  const val = row[col.key];

  if (col.key === 'MARGEM_ATUAL') {
    const classe = classeMargem(val, num(row.MAR_PERC));
    const texto  = val !== null ? fmtPercent(val) : '-';
    return `<td class="result-cell col-result ${classe}" data-col="${col.key}">${texto}</td>`;
  }

  // VERBA LOJA e VERBA SITE:
  //   tem valor  → amarelo (verba necessária = atenção)
  //   nulo/vazio → verde   (sem necessidade de verba)
  if (col.key === 'VERBA_LOJA' || col.key === 'VERBA_SITE') {
    const classe = val !== null ? 'margem-atencao' : 'margem-positiva';
    const texto  = val !== null ? fmtMoeda(val) : '';
    return `<td class="result-cell col-result ${classe}" data-col="${col.key}">${texto}</td>`;
  }

  const texto = val !== null && val !== undefined ? fmtValue(val, col.type) : '';
  return `<td class="result-cell col-result" data-col="${col.key}">${texto}</td>`;
}

/** Atualiza somente as células calculadas de uma linha (sem re-renderizar o <tbody>) */
function atualizarCelulasResultado(tr, row) {
  COLUMNS.filter(c => !c.editable).forEach(col => {
    const td = tr.querySelector(`td[data-col="${col.key}"]`);
    if (!td) return;

    if (col.key === 'MARGEM_ATUAL') {
      const classe = classeMargem(row[col.key], num(row.MAR_PERC));
      td.className   = `result-cell col-result ${classe}`;
      td.textContent = row[col.key] !== null ? fmtPercent(row[col.key]) : '-';
    } else if (col.key === 'VERBA_LOJA' || col.key === 'VERBA_SITE') {
      const val    = row[col.key];
      td.className   = `result-cell col-result ${val !== null ? 'margem-atencao' : 'margem-positiva'}`;
      td.textContent = val !== null ? fmtMoeda(val) : '';
    } else {
      td.className   = 'result-cell col-result';
      td.textContent = row[col.key] !== null && row[col.key] !== undefined
        ? fmtValue(row[col.key], col.type)
        : '';
    }
  });
}

/* ============================================================
   ESTATÍSTICAS
   ============================================================ */
function atualizarEstatisticas(indices) {
  const panel = document.getElementById('stats-panel');
  if (tableData.length === 0) { panel.style.display = 'none'; return; }
  panel.style.display = 'flex';

  const visivel  = (indices || tableData.map((_, i) => i)).map(i => tableData[i]);
  const total    = visivel.length;
  const positiva = visivel.filter(r => r.MARGEM_ATUAL !== null && r.MARGEM_ATUAL >= num(r.MAR_PERC)).length;
  const negativa = visivel.filter(r => r.MARGEM_ATUAL !== null && r.MARGEM_ATUAL < 0).length;
  const atencao  = visivel.filter(r => {
    if (r.MARGEM_ATUAL === null) return false;
    return r.MARGEM_ATUAL >= 0 && r.MARGEM_ATUAL < num(r.MAR_PERC);
  }).length;

  document.getElementById('stat-total').textContent    = total;
  document.getElementById('stat-positive').textContent = positiva;
  document.getElementById('stat-negative').textContent = negativa;
  document.getElementById('stat-warning').textContent  = atencao;
}

/* ============================================================
   EVENTOS
   ============================================================ */

/** Vincula os eventos de edição e exclusão após renderização */
function vincularEventos() {
  document.querySelectorAll('#table-body .cell-input').forEach(input => {
    input.addEventListener('change',  onCellChange);
    input.addEventListener('keydown', onCellKeydown);
  });

  document.querySelectorAll('#table-body .btn-del-row').forEach(btn => {
    btn.addEventListener('click', onDeleteRow);
  });

  document.querySelectorAll('#table-body .btn-replicate').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const rowIdx = parseInt(btn.dataset.row, 10);
      const colKey = btn.dataset.col;
      const valor  = tableData[rowIdx][colKey];
      tableData.forEach(row => { row[colKey] = valor; calcularLinha(row); });
      renderTable();
    });
  });
}

/** Chamado quando o usuário altera o valor de uma célula */
function onCellChange(e) {
  const el     = e.target;
  const rowIdx = parseInt(el.dataset.row, 10);
  const colKey = el.dataset.col;
  const type   = el.dataset.type;

  // Persiste o valor convertido para decimal/número
  tableData[rowIdx][colKey] = parseInput(el.value, type);

  // Recalcula a linha
  calcularLinha(tableData[rowIdx]);

  // Atualiza somente as células de resultado (sem re-renderizar tudo)
  const tr = document.querySelector(`tr[data-row="${rowIdx}"]`);
  if (tr) atualizarCelulasResultado(tr, tableData[rowIdx]);

  atualizarEstatisticas();
}

/** Navegação via teclado (Enter/Tab = avança, Seta = move) */
function onCellKeydown(e) {
  const el     = e.target;
  const rowIdx = parseInt(el.dataset.row, 10);
  const colKey = el.dataset.col;
  const colIdx = COLUMNS.findIndex(c => c.key === colKey);

  if (e.key === 'Enter' || (e.key === 'Tab' && !e.shiftKey)) {
    e.preventDefault();
    moverFoco(rowIdx, colIdx, 1, 0);
  } else if (e.key === 'Tab' && e.shiftKey) {
    e.preventDefault();
    moverFoco(rowIdx, colIdx, -1, 0);
  } else if (e.key === 'ArrowDown') {
    e.preventDefault();
    moverFoco(rowIdx, colIdx, 0, 1);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    moverFoco(rowIdx, colIdx, 0, -1);
  }
}

/** Move o foco para a próxima célula editável */
function moverFoco(row, col, dCol, dRow) {
  let novaCol = col + dCol;
  // Pula colunas não-editáveis
  while (novaCol >= 0 && novaCol < COLUMNS.length && !COLUMNS[novaCol].editable) {
    novaCol += dCol !== 0 ? Math.sign(dCol) : 1;
  }
  const novaLinha = Math.max(0, Math.min(tableData.length - 1, row + dRow));
  novaCol = Math.max(0, Math.min(COLUMNS.length - 1, novaCol));

  if (!COLUMNS[novaCol]?.editable) return;

  const target = document.querySelector(
    `input[data-row="${novaLinha}"][data-col="${COLUMNS[novaCol].key}"]`
  );
  if (target) { target.focus(); target.select(); }
}

/** Exclui uma linha da tabela */
function onDeleteRow(e) {
  const rowIdx = parseInt(e.target.closest('[data-row]').dataset.row, 10);
  tableData.splice(rowIdx, 1);
  renderTable();
}

/* ============================================================
   IMPORTAÇÃO DE EXCEL
   ============================================================ */

/** Lê e processa um arquivo Excel (.xlsx / .xls) */
function importarExcel(file) {
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const workbook  = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Converte para array de arrays, mantendo valores brutos
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: true,
      });

      if (rawData.length < 2) {
        alert('Arquivo sem dados. Verifique o arquivo Excel selecionado.');
        return;
      }

      const novasLinhas = [];

      // Processa a partir da linha 1 (linha 0 = cabeçalho)
      for (let r = 1; r < rawData.length; r++) {
        const excelRow = rawData[r];

        // Ignora linhas completamente vazias
        if (!excelRow || excelRow.every(v => v === '' || v === null || v === undefined)) continue;

        const row = criarLinhaVazia();

        Object.entries(EXCEL_IMPORT_MAP).forEach(([colIdx, chave]) => {
          if (chave === null) return; // Coluna calculada → ignorar

          const valor = excelRow[parseInt(colIdx, 10)];
          if (valor !== undefined && valor !== '') {
            // Strings numéricas são convertidas automaticamente pelo JS
            row[chave] = typeof valor === 'number' ? valor : String(valor).trim();
          }
        });

        calcularLinha(row);
        novasLinhas.push(row);
      }

      if (novasLinhas.length === 0) {
        alert('Nenhum dado válido encontrado. Verifique se a planilha tem dados a partir da linha 2.');
        return;
      }

      tableData = novasLinhas;
      renderTable();

    } catch (err) {
      console.error('Erro na importação:', err);
      alert('Erro ao ler o arquivo: ' + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

/* ============================================================
   GERENCIAMENTO DE LINHAS
   ============================================================ */

/** Cria um objeto de linha com campos em branco e parâmetros com valores padrão */
function criarLinhaVazia() {
  const row = {};
  COLUMNS.forEach(col => { row[col.key] = ''; });
  Object.assign(row, DEFAULT_PARAMS);
  return row;
}

/** Adiciona uma nova linha vazia e foca o primeiro campo */
function adicionarLinha() {
  const row = criarLinhaVazia();
  calcularLinha(row);
  tableData.push(row);
  renderTable();

  // Foca o campo CODIGO da nova linha
  setTimeout(() => {
    const novoIdx = tableData.length - 1;
    const input   = document.querySelector(`input[data-row="${novoIdx}"][data-col="CODIGO"]`);
    if (input) {
      input.focus();
      input.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }, 50);
}

/** Limpa todos os dados após confirmação */
function limparDados() {
  if (tableData.length > 0 && !confirm('Deseja limpar todos os dados da tabela?')) return;
  tableData = [];
  renderTable();
}

/* ============================================================
   EXPORTAR RESULTADOS
   ============================================================ */
function formatarData(d) {
  if (!d) return '';
  const [y, m, dia] = d.split('-');
  return `${dia}/${m}/${y}`;
}

function exportarResultados(dtInicio, dtFim) {
  const cabecalho = [
    'COD_EMPRESA', 'CANAL_VENDA', 'CODIGO', 'TIPO', 'QTDE',
    'VLR_VERBA', 'DT_INICIO', 'DT_FIM', 'ATIVO', 'DT_ATUALIZACAO', 'DESCRICAO'
  ];

  const linhas = [cabecalho];

  // Canal 1 = Loja (coluna P = VERBA_LOJA)
  tableData.forEach(row => {
    if (row.VERBA_LOJA === null || row.VERBA_LOJA === undefined) return;
    linhas.push([
      row.COD_EMPRESA,
      1,
      row.CODIGO,
      'P',
      10,
      parseFloat(row.VERBA_LOJA.toFixed(2)),
      dtInicio,
      dtFim,
      1,
      'SYSDATE',
      row.DESCRICAO
    ]);
  });

  // Canal 2 = Site (coluna Q = VERBA_SITE)
  tableData.forEach(row => {
    if (row.VERBA_SITE === null || row.VERBA_SITE === undefined) return;
    linhas.push([
      row.COD_EMPRESA,
      2,
      row.CODIGO,
      'P',
      10,
      parseFloat(row.VERBA_SITE.toFixed(2)),
      dtInicio,
      dtFim,
      1,
      'SYSDATE',
      row.DESCRICAO
    ]);
  });

  if (linhas.length === 1) {
    alert('Nenhuma linha com verba calculada para exportar.');
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(linhas);
  ws['!cols'] = [
    { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 6 }, { wch: 6 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 6 }, { wch: 14 }, { wch: 40 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Resultados');
  XLSX.writeFile(wb, 'resultados_verba.xlsx');
}

function exportarResultadosFCM(dtInicio, dtFim) {
  const cabecalho = [
    'FORNECEDOR_FANTAS', 'TIPO', 'DT_INICIO', 'DT_FIM', 'SELL_OUT_S_N',
    'COD_PRODUTO', 'COD_FILIAL', 'DESCONTO_CUE', 'QTDE',
    'LOJA', 'ECOMMERCE', 'VENDA_CORPORATIVA', 'CANAL_VENDA'
  ];

  const linhas = [cabecalho];

  // Canal 1 = Loja (coluna P)
  tableData.forEach(row => {
    if (row.VERBA_LOJA === null || row.VERBA_LOJA === undefined) return;
    linhas.push([
      row.FANTAS, 'P', dtInicio, dtFim, 'N',
      row.CODIGO, row.COD_EMPRESA,
      parseFloat(row.VERBA_LOJA.toFixed(2)),
      10, 'S', 'N', 'N', 1
    ]);
  });

  // Canal 2 = Site (coluna Q)
  tableData.forEach(row => {
    if (row.VERBA_SITE === null || row.VERBA_SITE === undefined) return;
    linhas.push([
      row.FANTAS, 'P', dtInicio, dtFim, 'N',
      row.CODIGO, row.COD_EMPRESA,
      parseFloat(row.VERBA_SITE.toFixed(2)),
      10, 'N', 'S', 'N', 2
    ]);
  });

  if (linhas.length === 1) { alert('Nenhuma linha com verba calculada para exportar.'); return; }

  const ws = XLSX.utils.aoa_to_sheet(linhas);
  ws['!cols'] = [
    { wch: 20 }, { wch: 6 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 6 },
    { wch: 6 }, { wch: 10 }, { wch: 18 }, { wch: 12 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Resultados FCM');
  XLSX.writeFile(wb, 'resultados_verba_FCM.xlsx');
}

/* ============================================================
   FILTROS — EMPRESA DROPDOWN
   ============================================================ */
function popularFiltroEmpresa() {
  const sel = document.getElementById('filter-empresa');
  if (!sel) return;
  const currentVal = sel.value;
  const empresas = [...new Set(
    tableData.map(r => r.COD_EMPRESA).filter(v => v !== '' && v !== null && v !== undefined)
  )].sort((a, b) => Number(a) - Number(b));

  sel.innerHTML = '<option value="">Todas</option>';
  empresas.forEach(e => {
    const opt = document.createElement('option');
    opt.value = String(e);
    opt.textContent = e;
    sel.appendChild(opt);
  });
  sel.value = currentVal;
}

/* ============================================================
   EXPORTAR MODELO
   ============================================================ */
function exportarModelo() {
  const cabecalho = [
    'CÓDIGO', 'COD FC', 'FANTASIA', 'DESCRIÇÃO', 'REFERÊNCIA',
    'EMPRESA', 'CUE (R$)', 'OPER. %', 'ICMS %', 'PC %', 'MAR. %', 'VPC (R$)', 'PV D1 (R$)'
  ];
  const empresas = [1, 2, 3, 4, 5, 6, 7, 8, 9, 92, 93];

  const linhas = [cabecalho];
  empresas.forEach(emp => {
    const linha = new Array(cabecalho.length).fill('');
    linha[5] = emp; // Coluna F = EMPRESA
    linhas.push(linha);
  });

  const ws = XLSX.utils.aoa_to_sheet(linhas);

  // Larguras de coluna
  ws['!cols'] = [
    { wch: 10 }, { wch: 10 }, { wch: 20 }, { wch: 30 }, { wch: 14 },
    { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 12 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Verba');
  XLSX.writeFile(wb, 'modelo_verba.xlsx');
}

/* ============================================================
   INICIALIZAÇÃO
   ============================================================ */
document.addEventListener('DOMContentLoaded', function () {
  renderHeader();
  renderTable();

  // Botão "Importar Excel"
  document.getElementById('btn-import').addEventListener('click', () => {
    document.getElementById('file-input').click();
  });

  // Seleção de arquivo
  document.getElementById('file-input').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (file) {
      importarExcel(file);
      this.value = ''; // Permite re-importar o mesmo arquivo
    }
  });

  // Botão "+ Linha"
  document.getElementById('btn-add-row').addEventListener('click', adicionarLinha);

  // Botão "Limpar"
  document.getElementById('btn-clear').addEventListener('click', limparDados);

  // Botão "Exportar Modelo"
  document.getElementById('btn-export-modelo').addEventListener('click', exportarModelo);

  // Modal exportar resultados
  const modalEl = document.getElementById('modal-export');

  document.getElementById('btn-export-resultados').addEventListener('click', () => {
    if (tableData.length === 0) { alert('Nenhum dado carregado para exportar.'); return; }
    modalEl.style.display = 'flex';
  });

  function fecharModal() { modalEl.style.display = 'none'; }

  document.getElementById('modal-close').addEventListener('click', fecharModal);
  document.getElementById('modal-cancel').addEventListener('click', fecharModal);
  modalEl.addEventListener('click', (e) => { if (e.target === modalEl) fecharModal(); });

  document.getElementById('modal-confirm').addEventListener('click', () => {
    const dtInicio = formatarData(document.getElementById('export-dt-inicio').value);
    const dtFim    = formatarData(document.getElementById('export-dt-fim').value);
    if (!dtInicio || !dtFim) { alert('Preencha as datas DT_INICIO e DT_FIM antes de exportar.'); return; }

    const tipo = document.querySelector('input[name="export-type"]:checked').value;
    if (tipo === 'fcm') {
      exportarResultadosFCM(dtInicio, dtFim);
    } else {
      exportarResultados(dtInicio, dtFim);
    }
    fecharModal();
  });

  // Filtros
  document.getElementById('filter-search').addEventListener('input', function () {
    filterState.search = this.value.trim();
    renderTable();
  });

  document.getElementById('filter-empresa').addEventListener('change', function () {
    filterState.empresa = this.value;
    renderTable();
  });

  document.getElementById('filter-margem').addEventListener('change', function () {
    filterState.margem = this.value;
    renderTable();
  });

  document.getElementById('btn-clear-filters').addEventListener('click', function () {
    filterState.search  = '';
    filterState.empresa = '';
    filterState.margem  = '';
    document.getElementById('filter-search').value  = '';
    document.getElementById('filter-empresa').value = '';
    document.getElementById('filter-margem').value  = '';
    renderTable();
  });
});
