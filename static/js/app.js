/* ============================================
   基站工参管理器 — 前端逻辑
   ============================================ */

let currentPage = 1, lastTotalResults = 0, PER_PAGE = 50;
// 预览数据: [{filename, carrier, tech, sheets:[{name,rows,recommended}]}]
let previewData = [];

const searchInput = document.getElementById('searchInput');

searchInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') { currentPage = 1; doSearch(); }
  // Escape：清空输入框并重新聚焦
  if (e.key === 'Escape') {
    searchInput.value = '';
    searchInput.select();
    e.preventDefault();
  }
});

// 全局快捷键
document.addEventListener('keydown', e => {
  const tag = document.activeElement.tagName;
  const inInput = (tag === 'INPUT' || tag === 'TEXTAREA');

  // Ctrl+A：无论焦点在哪，都聚焦到搜索框并全选内容
  if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
    if (!inInput || document.activeElement !== searchInput) {
      e.preventDefault();
      searchInput.focus();
      searchInput.select();
      return;
    }
  }

  // 任意可打印字符（不带 Ctrl/Meta/Alt）：跳转到搜索框接收输入
  if (!inInput && !e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
    searchInput.focus();
  }
});

// ======== 统计 & 文件卡片 ========
async function loadStats() {
  try {
    const d = await (await fetch('/api/stats')).json();
    document.getElementById('totalCount').textContent = d.total.toLocaleString();
    document.getElementById('stationCount').textContent = (d.stations || 0).toLocaleString();
    document.getElementById('fileCount').textContent = d.files;
    renderFileCards(d.sources);
  } catch(e) {}
}

function renderFileCards(sources) {
  const fc = document.getElementById('fileCards');
  if (!sources || !sources.length) { fc.innerHTML = ''; return; }
  fc.innerHTML = sources.map(s => {
    const sheets = s.sheets || [];
    const sheetCount = sheets.length;
    const cardId = 'card_' + s.filename.replace(/[^a-zA-Z0-9]/g,'_');
    return `
    <div class="file-card">
      <div class="file-card-header">
        <div class="file-card-name" title="${esc(s.filename)}">${esc(trunc(s.filename, 28))}</div>
        <button class="btn btn-danger btn-xs" onclick="removeFile('${esc(s.filename)}')" title="关闭整个文件">✕</button>
      </div>
      <div class="file-card-meta">
        <span class="carrier-tag carrier-${s.carrier === '中国电信' ? 'dx' : s.carrier === '中国联通' ? 'lt' : 'yd'}">${esc(s.carrier||'')}</span>
        <span class="tech-tag tech-${(s.tech||'').includes('5G')?'5g':'4g'}">${esc(s.tech||'')}</span>
        <span>${s.count.toLocaleString()} 条</span>
      </div>
      ${sheetCount > 0 ? `
      <div class="file-card-sheets" id="${cardId}_toggle" onclick="toggleSheets('${cardId}')">
        <span class="arrow">&#x25B6;</span>
        <span>${sheetCount} 个工作表</span>
      </div>
      <div class="sheet-list" id="${cardId}_sheets">
        ${sheets.map(sn => `
          <div class="sheet-item">
            <span class="sheet-name">${esc(sn)}</span>
            <button class="btn btn-outline btn-xs" style="border-color:#fca5a5;color:#dc2626"
              onclick="removeSheet('${esc(s.filename)}','${esc(sn)}')">移除</button>
          </div>`).join('')}
      </div>` : ''}
    </div>`}).join('');
}

function toggleSheets(cardId) {
  const toggle = document.getElementById(cardId + '_toggle');
  const list = document.getElementById(cardId + '_sheets');
  if (!toggle || !list) return;
  list.classList.toggle('show');
  toggle.classList.toggle('expanded');
}

// ======== 上传 → 预览弹窗 ========
async function handleFiles(files) {
  if (!files.length) return;

  showProgress('正在读取文件...', '解析 sheet 信息，请稍候');

  const fd = new FormData();
  for (let i = 0; i < files.length; i++) fd.append('files', files[i]);
  document.getElementById('fileInput').value = '';

  try {
    const d = await (await fetch('/api/preview', { method: 'POST', body: fd })).json();
    hideProgress();
    if (d.error) { showToast(d.error, 'error'); return; }
    previewData = d.files || [];
    if (!previewData.length) { showToast('未识别到可导入的文件', 'error'); return; }
    openPreview();
  } catch(e) {
    hideProgress();
    showToast('读取文件失败: ' + e.message, 'error');
  }
}

function openPreview() {
  const body = document.getElementById('previewBody');
  if (!previewData.length) return;

  body.innerHTML = previewData.map((f, fi) => {
    if (f.error) return `<div class="file-section">
      <div class="file-section-head" style="color:var(--danger)">${esc(f.filename)} — ${esc(f.error)}</div>
    </div>`;

    const sheetItems = f.sheets.map((s, si) => `
      <label class="sheet-check-item">
        <input type="checkbox" id="chk_${fi}_${si}" data-fi="${fi}" data-si="${si}"
          ${s.recommended ? 'checked' : ''}>
        <span class="sc-name">${esc(s.name)}${s.recommended ? '<span class="sc-recommended">推荐</span>' : ''}</span>
        <span class="sc-rows">${s.rows > 0 ? s.rows.toLocaleString() + ' 行' : ''}</span>
      </label>`).join('');

    return `<div class="file-section">
      <div class="file-section-head">
        ${esc(trunc(f.filename, 40))}
        <span class="file-section-meta">${esc(f.carrier||'')} ${esc(f.tech||'')}</span>
      </div>
      <div class="sheet-check-list">
        <div class="select-all-row">
          <a onclick="selectAll(${fi},true)">全选</a>
          <a onclick="selectAll(${fi},false)">全不选</a>
          <a onclick="selectAll(${fi},'recommended')">仅推荐</a>
        </div>
        ${sheetItems}
      </div>
    </div>`;
  }).join('');

  document.getElementById('previewOverlay').style.display = 'flex';
}

function selectAll(fi, mode) {
  const checks = document.querySelectorAll(`input[data-fi="${fi}"]`);
  checks.forEach(cb => {
    const si = parseInt(cb.getAttribute('data-si'));
    if (mode === 'recommended') cb.checked = previewData[fi].sheets[si].recommended;
    else cb.checked = !!mode;
  });
}

function closePreview() {
  document.getElementById('previewOverlay').style.display = 'none';
  previewData = [];
}

// ======== 确认导入 ========
async function confirmImport() {
  // 收集每个文件选中的 sheet
  const selections = [];
  previewData.forEach((f, fi) => {
    if (f.error) return;
    const selected = [];
    f.sheets.forEach((s, si) => {
      const cb = document.getElementById(`chk_${fi}_${si}`);
      if (cb && cb.checked) selected.push(s.name);
    });
    if (selected.length) selections.push({ filename: f.filename, sheets: selected });
  });

  if (!selections.length) { showToast('请至少勾选一个工作表', 'error'); return; }

  closePreview();
  showProgress('正在导入工参数据...', '准备中');

  // 启动后台导入，轮询进度
  const r = await fetch('/api/import-sheets-async', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ files: selections })
  });
  const d = await r.json();
  if (d.pending) {
    pollImportProgress();
  } else {
    hideProgress();
    showToast('导入完成', 'success');
    await loadStats(); currentPage = 1; doSearch();
  }
}

// ======== 进度轮询 ========
async function pollImportProgress() {
  try {
    const d = await (await fetch('/api/import-status')).json();
    if (d.total_files > 0) {
      const pct = Math.round(d.current_index / d.total_files * 100);
      document.getElementById('progressTitle').textContent = `正在导入工参数据... (${pct}%)`;
    }
    let msg = `(${d.current_index}/${d.total_files}) ${d.current_file}`;
    if (d.current_sheet) msg += ` → ${d.current_sheet}`;
    msg += ` | 已导入 ${d.imported_records.toLocaleString()} 条`;
    document.getElementById('progressDetail').textContent = msg;
    document.getElementById('totalCount').textContent = d.imported_records.toLocaleString();

    if (d.done || !d.running) {
      hideProgress();
      showToast(`导入完成，共 ${d.imported_records.toLocaleString()} 条记录`, 'success');
      await loadStats(); currentPage = 1; doSearch();
      return;
    }
    setTimeout(pollImportProgress, 500);
  } catch(e) { setTimeout(pollImportProgress, 1000); }
}

// ======== 移除文件 / 移除 sheet ========
async function removeFile(filename) {
  if (!confirm(`确定要移除文件「${filename}」的所有数据吗？`)) return;
  try {
    const r = await fetch('/api/remove-sheet', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename })
    });
    const d = await r.json();
    if (d.ok) { showToast('已移除', 'info'); await loadStats(); currentPage = 1; doSearch(); }
  } catch(e) { showToast('操作失败: ' + e.message, 'error'); }
}

async function removeSheet(filename, sheet) {
  if (!confirm(`确定要移除「${filename}」中的工作表「${sheet}」吗？`)) return;
  try {
    const r = await fetch('/api/remove-sheet', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename, sheets: [sheet] })
    });
    const d = await r.json();
    if (d.ok) { showToast('已移除工作表', 'info'); await loadStats(); currentPage = 1; doSearch(); }
  } catch(e) { showToast('操作失败: ' + e.message, 'error'); }
}

// ======== 搜索 & 渲染 ========
async function doSearch(page) {
  currentPage = typeof page === 'number' ? page : 1;
  const q = document.getElementById('searchInput').value.trim();
  try {
    const resp = await fetch(`/api/search?q=${encodeURIComponent(q)}&page=${currentPage}&per_page=${PER_PAGE}`);
    if (!resp.ok) throw new Error(`服务器错误 (${resp.status})`);
    const d = await resp.json();
    lastTotalResults = d.total;
    document.getElementById('searchResultCount').textContent = d.total > 0 ? `共 ${d.total.toLocaleString()} 条` : '';
    document.getElementById('exportBtn').style.display = d.total > 0 ? 'inline-flex' : 'none';
    renderTable(d.results, q);
    renderPagination(d.total, d.page, d.pages);
  } catch(e) {
    console.error('搜索失败:', e);
    showToast('搜索失败: ' + e.message, 'error');
  }
}

function renderTable(data, q) {
  const tb = document.getElementById('tableBody');
  if (!data || !data.length) {
    const empty = q ? '未找到匹配的工参记录，请尝试其他关键词' : '暂无数据，请导入工参文件';
    tb.innerHTML = `<tr><td colspan="16"><div class="empty-state"><div class="icon">&#x1F50D;</div><h3>${empty}</h3><p style="margin-top:8px">支持按小区名、PCI、基站ID、基站名搜索</p></div></td></tr>`;
    document.getElementById('tableTitle').innerHTML = '&#x1F4CB; 工参数据';
    return;
  }
  let html = '';
  for (const r of data) {
    const tech = r['\u6280\u672f\u5236\u5f0f'] || '';
    const tCls = tech.includes('5G') ? 'tech-5g' : 'tech-4g';
    const techLabel = tech.includes('5G') ? '5G' : '4G';
    const cc = r['\u8fd0\u8425\u5546'] || '';
    const shortCarrier = cc.replace('\u4e2d\u56fd', '');
    const cCls = cc.includes('\u7535\u4fe1') ? 'carrier-dx' : cc.includes('\u8054\u901a') ? 'carrier-lt' : 'carrier-yd';
    html += `<tr>
      <td><span class="tech-tag ${tCls}">${techLabel}</span></td>
      <td><span class="carrier-tag ${cCls}">${shortCarrier}</span></td>
      <td>${(() => { const v = r['\u8bbe\u5907\u5546']||''; let vc='vendor-other'; if(v.includes('\u534e\u4e3a'))vc='vendor-hw'; else if(v.includes('\u5927\u5510'))vc='vendor-dt'; else if(v.includes('\u8bfa\u57fa\u4e9a')||v.includes('Nokia'))vc='vendor-ns'; else if(v.includes('\u7231\u7acb\u4fe1')||v.includes('Ericsson'))vc='vendor-er'; else if(v.includes('\u4e2d\u5174'))vc='vendor-zt'; return `<span class="vendor-tag ${vc}">${esc(v)}</span>`; })()}</td>
      <td title="${esc(r['\u57fa\u7ad9\u540d']||'')}">${esc(trunc(r['\u57fa\u7ad9\u540d'],30))}</td>
      <td>${esc(r['\u57fa\u7ad9ID']||'')}</td>
      <td title="${esc(r['\u5c0f\u533a\u540d']||'')}">${esc(trunc(r['\u5c0f\u533a\u540d'],35))}</td>
      <td style="font-family:monospace">${esc(r['PCI']||'')}</td>
      <td style="font-family:monospace">${esc(r['\u5c0f\u533aID']||'')}</td>
      <td>${esc(r['\u4e0b\u884c\u9891\u70b9']||'')}</td>
      <td>${esc(r['\u4e0b\u503e\u89d2']||'')}</td>
      <td>${esc(r['\u6302\u9ad8']||'')}</td>
      <td>${esc(r['\u65b9\u4f4d\u89d2']||'')}</td>
      <td>${esc(r['\u7ecf\u5ea6']||'')}</td>
      <td>${esc(r['\u7eac\u5ea6']||'')}</td>
      <td>${(() => { const s = r['\u5171\u4eab']||''; if(!s) return ''; const cls = s.includes('\u975e\u5171\u4eab') ? 'share-no' : 'share-yes'; return `<span class="share-tag ${cls}">${esc(s)}</span>`; })()}</td>
      <td style="font-size:11px;color:var(--text-sec)">${esc(trunc(r['_\u6587\u4ef6\u540d']||'',25))}</td>
    </tr>`;
  }
  tb.innerHTML = html;
  document.getElementById('tableTitle').innerHTML = '&#x1F4CB; 工参数据' + (q ? ` &mdash; 搜索: <b>${esc(q)}</b> (共 ${lastTotalResults.toLocaleString()} 条)` : '');
}

function renderPagination(total, page, pages) {
  const pg = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML = ''; return; }
  let html = '';
  html += `<button class="page-btn" onclick="doSearch(1)"${page<=1?' disabled':''}>&laquo;</button>`;
  html += `<button class="page-btn" onclick="doSearch(${page-1})"${page<=1?' disabled':''}>&lsaquo;</button>`;
  const start = Math.max(1, page-2), end = Math.min(pages, page+2);
  if (start > 1) html += '<span class="page-info">...</span>';
  for (let i = start; i <= end; i++)
    html += `<button class="page-btn${i===page?' active':''}" onclick="doSearch(${i})">${i}</button>`;
  if (end < pages) html += '<span class="page-info">...</span>';
  html += `<button class="page-btn" onclick="doSearch(${page+1})"${page>=pages?' disabled':''}>&rsaquo;</button>`;
  html += `<button class="page-btn" onclick="doSearch(${pages})"${page>=pages?' disabled':''}>&raquo;</button>`;
  html += `<span class="page-info">${page} / ${pages} 页</span>`;
  html += `<span class="page-info">跳至 <input type="number" class="page-jump" id="pageJump" min="1" max="${pages}" placeholder="${page}" onkeydown="if(event.key==='Enter')jumpPage(${pages})"> 页</span>`;
  pg.innerHTML = html;
}

function jumpPage(totalPages) {
  const inp = document.getElementById('pageJump');
  const page = parseInt(inp.value, 10);
  if (!isNaN(page) && page >= 1 && page <= totalPages) {
    doSearch(page);
  } else {
    inp.value = '';
    inp.placeholder = '1~' + totalPages;
  }
}

// ======== 清空 ========
async function clearAll() {
  if (!confirm('确定要清空所有已导入的数据（同时删除本地缓存）？')) return;
  try {
    const d = await (await fetch('/api/clear', { method: 'POST' })).json();
    if (d.ok) {
      showToast('数据已清空', 'info');
      currentPage = 1;
      document.getElementById('searchInput').value = '';
      await loadStats(); doSearch();
    }
  } catch(e) { showToast('清空失败: ' + e.message, 'error'); }
}

// ======== 导出 ========
function exportCSV() {
  const q = document.getElementById('searchInput').value.trim();
  const a = document.createElement('a');
  a.href = '/api/export?q=' + encodeURIComponent(q);
  a.download = '工参查询结果_' + new Date().toISOString().slice(0,10) + '.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

// ======== 工具函数 ========
function showProgress(title, detail) {
  document.getElementById('progressTitle').textContent = title;
  document.getElementById('progressDetail').textContent = detail;
  document.getElementById('progressOverlay').style.display = 'flex';
}
function hideProgress() { document.getElementById('progressOverlay').style.display = 'none'; }
function trunc(s, n) { return s && s.length > n ? s.slice(0,n) + '...' : s || ''; }
function esc(s) {
  if (!s) return '';
  const d = document.createElement('div'); d.textContent = s; return d.innerHTML;
}
function showToast(msg, type) {
  const t = document.createElement('div');
  t.className = 'toast toast-' + type; t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3000);
}

// ======== 列宽拖拽调整 ========
(function(){
  const table = document.getElementById('dataTable');
  if (!table) return;
  const ths = table.querySelectorAll('thead th');
  ths.forEach(th => {
    const grip = document.createElement('div');
    grip.style.cssText = 'position:absolute;right:0;top:0;bottom:0;width:6px;cursor:col-resize;z-index:3';
    th.style.position = 'relative';
    th.appendChild(grip);
    let startX, startW;
    grip.addEventListener('mousedown', e => {
      e.preventDefault();
      startX = e.clientX;
      startW = th.offsetWidth;
      const onMove = ev => {
        const w = Math.max(60, startW + ev.clientX - startX);
        th.style.width = w + 'px';
        th.style.minWidth = w + 'px';
        th.style.maxWidth = w + 'px';
      };
      const onUp = () => {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  });
})();

// ======== 单元格点击复制 ========
document.getElementById('tableBody').addEventListener('click', function(e) {
  const td = e.target.closest('td');
  if (!td) return;
  const text = td.innerText.trim();
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    showToast('已复制: ' + text, 'info');
  }).catch(() => {
    showToast('复制失败', 'error');
  });
});

// ======== 报告工具 ========
let reportTab = 'query';  // 当前报告标签: 'query' | 'nearby'

function openReportTool() {
  document.getElementById('reportOverlay').style.display = 'flex';
  switchReportTab('query');
}

function closeReportTool() {
  document.getElementById('reportOverlay').style.display = 'none';
}

function switchReportTab(tab) {
  reportTab = tab;
  document.getElementById('tabBtnQuery').classList.toggle('active', tab === 'query');
  document.getElementById('tabBtnNearby').classList.toggle('active', tab === 'nearby');
  document.getElementById('panelQuery').style.display = tab === 'query' ? 'block' : 'none';
  document.getElementById('panelNearby').style.display = tab === 'nearby' ? 'block' : 'none';
}

// ---- 小区查询 ----
async function queryCell() {
  const q = document.getElementById('reportCellInput').value.trim();
  if (!q) { showToast('请输入小区名', 'error'); return; }

  const container = document.getElementById('reportCellResults');
  container.innerHTML = '<div class="empty-state"><div class="spinner"></div><p>查询中...</p></div>';

  try {
    const resp = await fetch(`/api/query-cell?q=${encodeURIComponent(q)}&limit=20`);
    const d = await resp.json();
    if (!d.results || !d.results.length) {
      container.innerHTML = '<div class="empty-state"><div class="icon">&#x1F50D;</div><h3>未找到匹配的小区</h3><p>请尝试更精确的小区名</p></div>';
      return;
    }
    container.innerHTML = d.results.map((r, i) => renderReportCard(r, i, 'cell')).join('');
  } catch(e) {
    container.innerHTML = '<div class="empty-state"><div class="icon">&#x26A0;</div><h3>查询失败</h3><p>' + esc(e.message) + '</p></div>';
  }
}

// ---- 附近基站查询 ----
async function queryNearby() {
  const lng = document.getElementById('reportLngInput').value.trim();
  const lat = document.getElementById('reportLatInput').value.trim();
  const radius = document.getElementById('reportRadiusInput').value.trim() || '5';

  if (!lng || !lat) { showToast('请输入经纬度', 'error'); return; }

  const container = document.getElementById('reportNearbyResults');
  container.innerHTML = '<div class="empty-state"><div class="spinner"></div><p>查询中...</p></div>';

  try {
    const resp = await fetch(`/api/cell-by-location?lng=${encodeURIComponent(lng)}&lat=${encodeURIComponent(lat)}&radius=${encodeURIComponent(radius)}`);
    const d = await resp.json();
    if (d.error) {
      container.innerHTML = '<div class="empty-state"><div class="icon">&#x26A0;</div><h3>' + esc(d.error) + '</h3></div>';
      return;
    }
    if (!d.results || !d.results.length) {
      container.innerHTML = '<div class="empty-state"><div class="icon">&#x1F4CD;</div><h3>未找到附近基站</h3><p>尝试扩大搜索半径</p></div>';
      return;
    }
    container.innerHTML = d.results.map((r, i) => renderReportCard(r, i, 'nearby')).join('');
  } catch(e) {
    container.innerHTML = '<div class="empty-state"><div class="icon">&#x26A0;</div><h3>查询失败</h3><p>' + esc(e.message) + '</p></div>';
  }
}

// ---- 渲染单张结果卡片 ----
function renderReportCard(r, idx, mode) {
  const techTag = (r['技术制式'] || '').includes('5G') ? '<span class="tech-tag tech-5g">5G</span>' : '<span class="tech-tag tech-4g">4G</span>';
  const cc = r['运营商'] || '';
  const shortCarrier = cc.replace('中国', '');
  const cCls = cc.includes('电信') ? 'carrier-dx' : cc.includes('联通') ? 'carrier-lt' : 'carrier-yd';

  const kmInfo = mode === 'nearby' ? `<span class="sc-recommended">${r.distance_km} km</span>` : '';

  return `
  <div class="report-result-card">
    <div class="report-result-head">
      <span class="carrier-tag ${cCls}">${esc(shortCarrier)}</span>
      ${techTag}
      <span class="cell-name">${esc(r['小区名'] || '—')}</span>
      ${kmInfo}
    </div>
    <div class="report-result-body">
      <div class="kv"><span class="k">基站名:</span><span class="v">${esc(trunc(r['基站名'], 28))}</span></div>
      <div class="kv"><span class="k">基站ID:</span><span class="v">${esc(r['基站ID'])}</span></div>
      <div class="kv"><span class="k">小区ID:</span><span class="v">${esc(r['小区ID'])}</span></div>
      <div class="kv"><span class="k">PCI:</span><span class="v">${esc(r['PCI'])}</span></div>
      <div class="kv"><span class="k">下行频点:</span><span class="v">${esc(r['下行频点'])}</span></div>
      <div class="kv"><span class="k">频段:</span><span class="v">${esc(r['频段'])}</span></div>
      <div class="kv"><span class="k">下倾角:</span><span class="v">${esc(r['下倾角'])}</span></div>
      <div class="kv"><span class="k">挂高:</span><span class="v">${esc(r['挂高'])}</span></div>
      <div class="kv"><span class="k">方位角:</span><span class="v">${esc(r['方位角'])}</span></div>
      <div class="kv"><span class="k">设备商:</span><span class="v">${esc(r['设备商'])}</span></div>
      <div class="kv"><span class="k">共享:</span><span class="v">${esc(r['共享'])}</span></div>
    </div>
    <div class="report-result-actions">
      <button class="btn btn-outline btn-xs" onclick="copyCellInfo(${idx})" title="复制工参信息">&#x1F4CB; 复制</button>
      <button class="btn btn-outline btn-xs" onclick="jumpToCell('${esc(r['小区名']||'')}')" title="在主搜索框中搜索此小区">&#x1F50D; 搜索</button>
    </div>
  </div>`;
}

// ---- 复制小区信息到剪贴板 ----
function copyCellInfo(idx) {
  const container = document.getElementById(reportTab === 'nearby' ? 'reportNearbyResults' : 'reportCellResults');
  const cards = container.querySelectorAll('.report-result-card');
  if (!cards[idx]) return;
  const card = cards[idx];
  const kvPairs = card.querySelectorAll('.kv');
  let text = '';
  kvPairs.forEach(kv => {
    const k = kv.querySelector('.k').textContent.replace(':', '');
    const v = kv.querySelector('.v').textContent;
    text += `${k}: ${v}\n`;
  });
  navigator.clipboard.writeText(text.trim()).then(() => {
    showToast('工参信息已复制', 'success');
  }).catch(() => showToast('复制失败', 'error'));
}

// ---- 跳转到主搜索框搜索此小区 ----
function jumpToCell(cellName) {
  closeReportTool();
  document.getElementById('searchInput').value = cellName;
  currentPage = 1;
  doSearch();
  document.getElementById('searchInput').focus();
  document.getElementById('searchInput').select();
}

// ======== 启动 ========
loadStats(); doSearch();
