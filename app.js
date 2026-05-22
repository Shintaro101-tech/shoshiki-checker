// UI 制御

const dropArea  = document.getElementById("dropArea");
const fileInput = document.getElementById("fileInput");
const resultEl  = document.getElementById("result");
const errorEl   = document.getElementById("error");
const fileNameEl = document.getElementById("fileName");
const summaryEl  = document.getElementById("summary");
const checksEl   = document.getElementById("checks");
const resetBtn   = document.getElementById("resetBtn");

let filterMode = "all"; // 'all' | 'issues'

["dragenter", "dragover"].forEach(ev => {
  dropArea.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation();
    dropArea.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach(ev => {
  dropArea.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation();
    dropArea.classList.remove("dragover");
  });
});

dropArea.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

dropArea.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

resetBtn.addEventListener("click", () => {
  resultEl.classList.add("hidden");
  errorEl.classList.add("hidden");
  dropArea.classList.remove("hidden");
  fileInput.value = "";
});

async function handleFile(file) {
  errorEl.classList.add("hidden");

  const name = file.name.toLowerCase();
  const isDocx = name.endsWith(".docx");
  const isPdf  = name.endsWith(".pdf");

  if (!isDocx && !isPdf) {
    showError(".docx または .pdf ファイルを選択してください。");
    return;
  }

  try {
    const result = isDocx
      ? await Checker.checkDocx(file)
      : await PdfChecker.checkPdf(file);
    window.__lastResult = result;
    renderResult(result);
  } catch (err) {
    console.error(err);
    showError(`ファイルの読み込みに失敗しました: ${err.message}`);
  }
}

function showError(msg) {
  errorEl.textContent = msg;
  errorEl.classList.remove("hidden");
}

function renderResult(result) {
  const modeBadge = result.mode === "pdf"
    ? ' <span class="mode-badge pdf">PDFモード（簡易チェック）</span>'
    : ' <span class="mode-badge docx">Wordモード（詳細チェック）</span>';
  fileNameEl.innerHTML = escapeHtml(result.fileName) + modeBadge;

  const counts = { ok: 0, error: 0, warn: 0 };
  for (const sec of result.sections) {
    for (const it of sec.items) counts[it.status]++;
  }

  summaryEl.innerHTML = `
    <div class="summary-item error">
      <span class="summary-count">${counts.error}</span>
      <span class="summary-label">🔴 要修正</span>
    </div>
    <div class="summary-item warn">
      <span class="summary-count">${counts.warn}</span>
      <span class="summary-label">🟡 要確認</span>
    </div>
    <div class="summary-item ok">
      <span class="summary-count">${counts.ok}</span>
      <span class="summary-label">✅ OK</span>
    </div>
    <div class="filter-buttons">
      <button class="filter-btn ${filterMode === "all" ? "active" : ""}" data-mode="all">すべて表示</button>
      <button class="filter-btn ${filterMode === "issues" ? "active" : ""}" data-mode="issues">要修正・要確認のみ</button>
    </div>
  `;

  summaryEl.querySelectorAll(".filter-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      filterMode = btn.dataset.mode;
      renderResult(window.__lastResult);
    });
  });

  checksEl.innerHTML = result.sections.map(renderSection).filter(Boolean).join("");

  dropArea.classList.add("hidden");
  resultEl.classList.remove("hidden");
}

function renderSection(sec) {
  const items = sec.items.filter(it => {
    if (filterMode === "issues") return it.status !== "ok";
    return true;
  });
  if (items.length === 0) return "";

  const refHtml = sec.reference
    ? `<div class="section-ref">📖 ${escapeHtml(sec.reference)}</div>`
    : "";

  return `
    <div class="check-section">
      <h3>${escapeHtml(sec.title)}</h3>
      ${refHtml}
      ${items.map(renderItem).join("")}
    </div>
  `;
}

function renderItem(item) {
  const icon = { ok: "✅", error: "🔴", warn: "🟡" }[item.status];

  // 該当箇所（単一）
  let locationHtml = "";
  if (item.location && !item.locations) {
    locationHtml = renderLocationLine(item.location);
  }

  // 該当箇所（複数）
  let locationsHtml = "";
  if (item.locations && item.locations.length > 0) {
    const list = item.locations.slice(0, 20).map(loc => renderLocationLine({ ...loc.location, snippet: loc.snippet })).join("");
    const more = item.locations.length > 20 ? `<div class="loc-more">…他 ${item.locations.length - 20} 件</div>` : "";
    locationsHtml = `<div class="loc-list">${list}${more}</div>`;
  }

  // 現状 / 正しい
  let bodyHtml = "";
  if (item.status === "ok") {
    if (item.current) {
      bodyHtml = `<div class="state ok">✓ ${escapeHtml(item.current)}</div>`;
    }
  } else {
    if (item.current != null) {
      bodyHtml += `<div class="state-row"><span class="state-label state-cur">現状</span><span class="state-val">${escapeHtml(item.current)}</span></div>`;
    }
    if (item.expected != null) {
      bodyHtml += `<div class="state-row"><span class="state-label state-exp">正しい</span><span class="state-val">${escapeHtml(item.expected)}</span></div>`;
    }
    if (item.hint) {
      bodyHtml += `<div class="hint">💡 ${escapeHtml(item.hint)}</div>`;
    }
  }

  // マニュアル参照（item 単位で異なる場合のみ表示。同じならセクション参照に従う）
  const refHtml = item.reference
    ? `<div class="item-ref">📖 ${escapeHtml(item.reference)}</div>`
    : "";

  return `
    <div class="check-item ${item.status}">
      <div class="check-status">${icon}</div>
      <div class="check-body">
        <div class="check-label">${escapeHtml(item.label)}</div>
        ${locationHtml}
        ${bodyHtml}
        ${locationsHtml}
        ${refHtml}
      </div>
    </div>
  `;
}

function renderLocationLine(loc) {
  if (!loc) return "";
  const parts = [];
  if (loc.area) {
    parts.push(escapeHtml(loc.area));
  } else if (loc.line) {
    const lineStr = loc.lineEnd && loc.lineEnd !== loc.line
      ? `${loc.line}〜${loc.lineEnd}行目`
      : `${loc.line}行目`;
    parts.push(`本文 p.${loc.page} / ${lineStr}（推定）`);
  } else if (loc.paragraphIndex != null) {
    parts.push(`${loc.paragraphIndex + 1}段落目`);
  }
  if (loc.snippet) parts.push(`「${escapeHtml(loc.snippet)}」`);
  if (parts.length === 0) return "";
  return `<div class="loc">📍 ${parts.join(" ／ ")}</div>`;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
