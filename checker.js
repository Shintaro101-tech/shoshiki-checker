// 書式チェッカー（長期研修員用）— チェックロジック
// 仕様は「研究報告書書式設定（所内共通）長研員用」R8 改訂版に基づく

const TWIPS_PER_MM = 56.6929;  // 1mm = 56.6929 twips（1twip = 1/1440インチ）
const TOLERANCE_MM = 0.3;      // 余白の許容誤差（mm）
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// フォント名の正規化（半角/全角・大小・スペース無視）
function normalizeFontName(name) {
  if (!name) return "";
  return name.replace(/\s+/g, "").toLowerCase()
    .replace(/ゴシック/g, "gothic")
    .replace(/明朝/g, "mincho");
}

const FONT_GOTHIC  = ["bizudgothic", "bizudゴシック"].map(normalizeFontName);
const FONT_MINCHO  = ["bizudminchomedium", "bizud明朝medium"].map(normalizeFontName);
const FONT_TIMES   = ["timesnewroman"].map(normalizeFontName);

function isGothic(name) { return FONT_GOTHIC.includes(normalizeFontName(name)); }
function isMincho(name) { return FONT_MINCHO.includes(normalizeFontName(name)); }
function isTimes(name)  { return FONT_TIMES.includes(normalizeFontName(name)); }

// XML パース用ユーティリティ
function parseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function twipsToMm(twips) { return twips / TWIPS_PER_MM; }

// w: prefixed attribute getter（DOMParser は属性に namespaceURI を持たないので getAttribute で取る）
function wAttr(el, name) {
  if (!el) return null;
  return el.getAttribute("w:" + name) || el.getAttribute(name);
}

// マニュアルの該当箇所（書式設定マニュアル R8 改訂版）
const REF = {
  PAGE_SETUP:   "マニュアル P.1「1. 基本設定」",
  HEADER:       "マニュアル P.1「1. 基本設定」ヘッダーの挿入",
  SUBJECT:      "マニュアル P.1「2. 原稿の書き方」表「教科・領域名」",
  THEME:        "マニュアル P.1「2. 原稿の書き方」表「テーマ」",
  SUBTHEME:    "マニュアル P.1「2. 原稿の書き方」表「サブテーマ」",
  AUTHOR:       "マニュアル P.1「2. 原稿の書き方」表「所属・役職・氏名」",
  ABSTRACT:     "マニュアル P.1「2. 原稿の書き方」表「要旨（Abstract）」",
  KEYWORD:      "マニュアル P.1「2. 原稿の書き方」表「キーワード」",
  BIG_HEAD:     "マニュアル P.2「2. 原稿の書き方」表「（大見出し）」",
  SMALL_HEAD:   "マニュアル P.2「2. 原稿の書き方」表「（小見出し）」",
  FIG_TITLE:    "マニュアル P.2「2. 原稿の書き方」表「（図表タイトル）」",
  BODY:         "マニュアル P.2「2. 原稿の書き方」表「本文」",
  CLOSING:      "マニュアル P.2「2. 原稿の書き方」表「（注記・引用文献・生成AI利用）」",
  AI_USE:       "マニュアル P.4「5. 生成AI利用」",
  REFERENCES:   "マニュアル P.3-4「4. 引用文献」",
  PROHIBITED:   "マニュアル P.5「6. その他注意事項」①（特殊な文字、文字飾りは原則使用しない）",
  LEADING_SP:   "マニュアル P.5「6. その他注意事項」④（行の先頭にスペースが入る場合は、スペースを削除する）",
  HALFWIDTH_NUM:"マニュアル P.2「2. 原稿の書き方」表「本文（数字）」（数字は半角を用いる）",
  HALFWIDTH_EN: "マニュアル P.2「2. 原稿の書き方」表「本文（英単語・略語）」（アルファベット表記はすべて半角を用いる）",
  ANGLE_BRACKET:"マニュアル P.1「2. 原稿の書き方」表「教科・領域名」（不等号記号を使用しない）",
};

// 結果オブジェクト生成
// status: 'ok' | 'error' | 'warn'
// current: 現状の値（短く）
// expected: 規定値（短く）
// hint: 補足の説明（修正方法など）
// reference: マニュアルの該当箇所
// location: { paragraph, line, page, snippet, area } または locations[] の配列
function makeCheck(label, status, options = {}) {
  return {
    label,
    status,
    current: options.current ?? null,
    expected: options.expected ?? null,
    hint: options.hint ?? null,
    reference: options.reference ?? null,
    location: options.location ?? null,
    locations: options.locations ?? null,
  };
}

// .docx を解凍してチェック実行
async function checkDocx(file) {
  const zip = await JSZip.loadAsync(file);
  const ctx = await extractContext(zip);
  const sections = [];

  sections.push({ ...checkPageSetup(ctx), reference: REF.PAGE_SETUP });
  sections.push({ ...checkHeader(ctx), reference: REF.HEADER });
  sections.push({ ...checkOpeningStructure(ctx) });
  sections.push({ ...checkAbstractAndKeyword(ctx) });
  sections.push({ ...checkBodyAndHeadings(ctx) });
  sections.push({ ...checkClosingStructure(ctx) });
  sections.push({ ...checkReferences(ctx), reference: REF.REFERENCES });
  sections.push({ ...checkProhibited(ctx) });

  return { fileName: file.name, sections };
}

// === コンテキスト抽出 ===
async function extractContext(zip) {
  const documentXml = await zip.file("word/document.xml").async("string");
  const stylesXml   = await zip.file("word/styles.xml").async("string");
  const settingsXml = zip.file("word/settings.xml")
    ? await zip.file("word/settings.xml").async("string") : null;
  const relsXml = await zip.file("word/_rels/document.xml.rels").async("string");

  const docDoc = parseXml(documentXml);

  const sectPrs = Array.from(docDoc.getElementsByTagNameNS(W_NS, "sectPr"));
  const firstSectPr = sectPrs[0] || null;

  // ヘッダー参照
  const headerRefs = {};
  if (firstSectPr) {
    Array.from(firstSectPr.getElementsByTagNameNS(W_NS, "headerReference"))
      .forEach((ref) => {
        const type = wAttr(ref, "type") || "default";
        const rid  = ref.getAttribute("r:id");
        headerRefs[type] = rid;
      });
  }

  // rels で headerN.xml の場所を解決
  const relsDoc = parseXml(relsXml);
  const ridToTarget = {};
  Array.from(relsDoc.getElementsByTagName("Relationship"))
    .forEach((rel) => { ridToTarget[rel.getAttribute("Id")] = rel.getAttribute("Target"); });

  const headers = {};
  for (const [type, rid] of Object.entries(headerRefs)) {
    const target = ridToTarget[rid];
    if (!target) continue;
    const f = zip.file(`word/${target}`);
    if (f) headers[type] = await f.async("string");
  }

  // スタイル解析
  const styles = parseStyles(stylesXml);

  // 段落リスト（本文のみ）
  const paragraphs = parseParagraphs(docDoc, styles);

  // テキストボックス内の段落（図表タイトル等）
  const textboxParagraphs = parseTextboxParagraphs(docDoc, styles);

  // 行・ページ位置の付与
  attachLineLocation(paragraphs);

  return {
    documentXml, stylesXml, settingsXml,
    docDoc, sectPrs, firstSectPr,
    headers, styles, paragraphs, textboxParagraphs,
  };
}

// テキストボックス内の段落を解析（図表タイトルなど）
// AlternateContent の Choice 側のみを採用し、Fallback の重複を避ける
function parseTextboxParagraphs(docDoc, stylesObj) {
  const result = [];
  const tbxContents = Array.from(docDoc.getElementsByTagNameNS(W_NS, "txbxContent"));
  for (const tbx of tbxContents) {
    // Fallback 配下のものはスキップ
    let inFallback = false;
    let cur = tbx.parentNode;
    while (cur) {
      if (cur.localName === "Fallback") { inFallback = true; break; }
      cur = cur.parentNode;
    }
    if (inFallback) continue;

    const pElems = Array.from(tbx.getElementsByTagNameNS(W_NS, "p"));
    for (const p of pElems) {
      const parsed = parseSingleParagraph(p, stylesObj, /*allowExcludedAncestor=*/true);
      if (parsed.text.trim().length === 0) continue;
      parsed.location = { area: "テキストボックス内" };
      result.push(parsed);
    }
  }
  return result;
}

// 段落要素を1つ解析（本文用とテキストボックス用で共通）
function parseSingleParagraph(p, stylesObj, allowExcludedAncestor) {
  const pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
  const jcEl = pPr ? pPr.getElementsByTagNameNS(W_NS, "jc")[0] : null;
  const alignment = jcEl ? wAttr(jcEl, "val") : "left";
  const pStyleEl = pPr ? pPr.getElementsByTagNameNS(W_NS, "pStyle")[0] : null;
  const pStyleId = pStyleEl ? wAttr(pStyleEl, "val") : null;

  const runs = [];
  for (const r of Array.from(p.getElementsByTagNameNS(W_NS, "r"))) {
    if (!allowExcludedAncestor && isInExcludedAncestor(r)) continue;
    // テキストボックス内の段落を解析する場合、自分の直接の親より下のテキストボックスは見ない
    if (allowExcludedAncestor && isNestedExclusion(r, p)) continue;

    const rPrEl = r.getElementsByTagNameNS(W_NS, "rPr")[0];
    const runRPr = rPrEl ? readRPr(rPrEl) : null;
    const texts = Array.from(r.getElementsByTagNameNS(W_NS, "t"))
      .map(t => t.textContent || "").join("");
    const resolved = resolveRPr(runRPr, pStyleId, stylesObj);
    runs.push({ text: texts, resolved, runRPrRaw: runRPr });
  }

  const text = runs.map(r => r.text).join("");
  return { text, alignment, pStyleId, runs };
}

// p の内側でさらに別のテキストボックスがある場合の除外チェック
function isNestedExclusion(el, ancestorP) {
  let cur = el.parentNode;
  while (cur && cur !== ancestorP) {
    if (cur.localName === "txbxContent") return true;
    if (cur.localName === "Fallback") return true;
    cur = cur.parentNode;
  }
  return false;
}

// 段落リストに行番号・ページ番号の推定値を付与
// 長期研修サンプルの構造: 1段組(冒頭)→ Section Break(Continuous) → 2段組(本文)
// 行番号は2段組セクション開始時点から1から振り直し、各カラム50行で連続カウント
function attachLineLocation(paragraphs) {
  // 「キーワード」段落の次に 2段組が始まる前提（または明示的なsectPr）
  // キーワード段落の index を見つける
  const kwIdx = paragraphs.findIndex(p => /^キーワード/.test(p.text.trim()));
  const bodyStartIdx = kwIdx >= 0 ? kwIdx + 1 : -1;

  let cumLine = 0;
  const CHARS_PER_LINE = 21;       // 2段組での1段あたり字数
  const LINES_PER_COLUMN = 50;     // 1段あたり行数
  const COLUMNS_PER_PAGE = 2;      // 2段組

  paragraphs.forEach((p, idx) => {
    if (idx < bodyStartIdx || bodyStartIdx < 0) {
      // 本文セクション前は「原稿冒頭」エリア
      p.location = { area: "原稿冒頭部", paragraphIndex: idx, page: 1 };
      return;
    }
    const lines = estimateLines(p.text, CHARS_PER_LINE);
    const startLine = cumLine + 1;
    cumLine += lines;
    // 2段組: 100行で1ページ完了、カラム1=1-50, カラム2=51-100
    const page = Math.floor((startLine - 1) / (LINES_PER_COLUMN * COLUMNS_PER_PAGE)) + 1;
    p.location = {
      paragraphIndex: idx,
      line: startLine,
      lineEnd: cumLine,
      page,
    };
  });
}

// テキストが何行を占めるか（セルカウントベース）
function estimateLines(text, charsPerLine) {
  if (!text || !text.trim()) return 1;
  let cells = 0;
  for (const ch of text) {
    // 半角範囲（ASCII / 半角カナ）は 0.5 セル、それ以外は 1 セル
    if (/[ -~｡-ﾟ]/.test(ch)) cells += 0.5;
    else cells += 1;
  }
  return Math.max(1, Math.ceil(cells / charsPerLine));
}

// location オブジェクトから表示用文字列を作る
function formatLocation(loc, snippet) {
  if (!loc) return null;
  if (loc.area === "header") {
    return { area: "原稿冒頭部", snippet };
  }
  if (loc.area === "body" && loc.line) {
    const lineStr = loc.lineEnd && loc.lineEnd !== loc.line
      ? `${loc.line}〜${loc.lineEnd}行目`
      : `${loc.line}行目`;
    return { area: `本文 p.${loc.page} ${lineStr}（推定）`, snippet };
  }
  return { area: `${loc.paragraphIndex + 1}段落目`, snippet };
}

// === スタイル定義の解析 ===
function parseStyles(stylesXml) {
  const doc = parseXml(stylesXml);
  const styles = {};
  let docDefaults = { sz: null, font: null };

  // docDefaults
  const def = doc.getElementsByTagNameNS(W_NS, "rPrDefault")[0];
  if (def) {
    const rPr = def.getElementsByTagNameNS(W_NS, "rPr")[0];
    if (rPr) {
      docDefaults = readRPr(rPr);
    }
  }

  // 各スタイル
  for (const st of Array.from(doc.getElementsByTagNameNS(W_NS, "style"))) {
    const id = wAttr(st, "styleId");
    if (!id) continue;
    const isDefault = wAttr(st, "default") === "1";
    const type = wAttr(st, "type");
    const basedOn = wAttr(st.getElementsByTagNameNS(W_NS, "basedOn")[0], "val");
    const rPrEl = st.getElementsByTagNameNS(W_NS, "rPr")[0];
    const rPr = rPrEl ? readRPr(rPrEl) : {};
    styles[id] = { id, type, isDefault, basedOn, rPr };
  }

  // 既定 paragraph スタイル（default="1" type="paragraph"）を特定
  let defaultParaStyle = null;
  for (const s of Object.values(styles)) {
    if (s.type === "paragraph" && s.isDefault) defaultParaStyle = s.id;
  }

  return { docDefaults, styles, defaultParaStyle };
}

// rPr 要素を読み取り、{fontEa, fontAscii, sz, bold, underline} を返す
function readRPr(rPrEl) {
  const r = {};
  const rFonts = rPrEl.getElementsByTagNameNS(W_NS, "rFonts")[0];
  if (rFonts) {
    r.fontAscii = wAttr(rFonts, "ascii");
    r.fontEa    = wAttr(rFonts, "eastAsia");
    r.fontHAnsi = wAttr(rFonts, "hAnsi");
  }
  const sz = rPrEl.getElementsByTagNameNS(W_NS, "sz")[0];
  if (sz) r.sz = parseInt(wAttr(sz, "val"));

  // 太字: <w:b/> （val なしまたは val="true"）。val="false" / "0" はOFF
  const b = rPrEl.getElementsByTagNameNS(W_NS, "b")[0];
  if (b) {
    const val = wAttr(b, "val");
    r.bold = (val == null || val === "true" || val === "1");
  }

  // 下線: <w:u w:val="..."/> "none" 以外は下線あり
  const u = rPrEl.getElementsByTagNameNS(W_NS, "u")[0];
  if (u) {
    const val = wAttr(u, "val");
    r.underline = (val && val !== "none");
  }
  return r;
}

// スタイル継承を辿って最終的な rPr を解決
// 重要: 段落のpPr→rPr（paragraph mark formatting）は段落マーク（¶）の書式であり、
// 段落内の run には継承されない。run のフォントは以下の順で解決される:
//   run rPr > paragraph style chain > default paragraph style > docDefaults
function resolveRPr(runRPr, paraStyleId, stylesObj) {
  const merged = {};
  const layers = [];

  layers.push(stylesObj.docDefaults);
  if (stylesObj.defaultParaStyle && stylesObj.styles[stylesObj.defaultParaStyle]) {
    layers.push(stylesObj.styles[stylesObj.defaultParaStyle].rPr);
  }
  if (paraStyleId) {
    const chain = [];
    let cur = stylesObj.styles[paraStyleId];
    while (cur) {
      chain.unshift(cur.rPr);
      cur = cur.basedOn ? stylesObj.styles[cur.basedOn] : null;
    }
    for (const c of chain) layers.push(c);
  }
  if (runRPr) layers.push(runRPr);

  for (const l of layers) {
    if (!l) continue;
    for (const k of ["fontAscii", "fontEa", "fontHAnsi", "sz", "bold", "underline"]) {
      if (l[k] != null) merged[k] = l[k];
    }
  }
  return merged;
}

// === 段落リスト構築（本文 = body 直下の <w:p> 群）===
function parseParagraphs(docDoc, stylesObj) {
  const result = [];
  const body = docDoc.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) return result;

  const pElems = Array.from(body.childNodes).filter(
    n => n.nodeType === 1 && n.localName === "p"
  );

  pElems.forEach((p, idx) => {
    const parsed = parseSingleParagraph(p, stylesObj, /*allowExcludedAncestor=*/false);
    result.push({ index: idx, ...parsed });
  });

  return result;
}

// ラン・テキストが「本文の流れ」に含まれない場所にあるかを判定。
// 除外対象:
//  - テキストボックス内 (<w:txbxContent>)
//  - 図やオブジェクト内 (<w:drawing>, <w:pict>, <w:object>)
//  - AlternateContent の Fallback 内（Choice と重複するため）
//  - フィールド・コメントなど（<w:fldChar>系は今は対象外）
function isInExcludedAncestor(el) {
  let cur = el.parentNode;
  while (cur) {
    const ln = cur.localName;
    if (ln === "txbxContent" || ln === "drawing" || ln === "pict" ||
        ln === "object" || ln === "Fallback") return true;
    cur = cur.parentNode;
  }
  return false;
}

// 解決済みサイズの取得（pt）
function getSizePt(resolved) {
  return resolved.sz != null ? resolved.sz / 2 : null;
}
function getEaFont(resolved) {
  return resolved.fontEa || resolved.fontAscii || null;
}
function getAsciiFont(resolved) {
  return resolved.fontAscii || resolved.fontEa || null;
}

// 段落の最初の意味のあるテキストの run を返す
function firstSignificantRun(p) {
  return p.runs.find(r => r.text.trim().length > 0) || null;
}

// === 1. 基本ページ設定 ===
function checkPageSetup(ctx) {
  const items = [];
  const sect = ctx.firstSectPr;
  if (!sect) {
    items.push(makeCheck("ページ設定が見つかりません", "error"));
    return { title: "1. ページ基本設定", items };
  }

  const pgMar = sect.getElementsByTagNameNS(W_NS, "pgMar")[0];
  if (pgMar) {
    const top    = parseInt(wAttr(pgMar, "top"));
    const bottom = parseInt(wAttr(pgMar, "bottom"));
    const left   = parseInt(wAttr(pgMar, "left"));
    const right  = parseInt(wAttr(pgMar, "right"));
    const header = parseInt(wAttr(pgMar, "header"));
    const footer = parseInt(wAttr(pgMar, "footer"));
    const gutter = parseInt(wAttr(pgMar, "gutter") || "0");

    items.push(marginCheck("上余白", top, 20));
    items.push(marginCheck("下余白", bottom, 18));
    items.push(marginCheck("左余白", left, 18));
    items.push(marginCheck("右余白", right, 18));
    items.push(marginCheck("ヘッダー位置", header, 10));
    items.push(marginCheck("フッター位置", footer, 10));
    items.push(marginCheck("とじしろ", gutter, 0));
  }

  const docGrid = sect.getElementsByTagNameNS(W_NS, "docGrid")[0];
  if (docGrid) {
    const linePitch = parseInt(wAttr(docGrid, "linePitch") || "0");
    const charSpace = parseInt(wAttr(docGrid, "charSpace") || "0");
    items.push(numericCheck("行送り", linePitch / 20, 14.65, "pt", 0.05));
    items.push(charSpaceCheck("字送り（charSpace値）", charSpace, 2021));
  }

  const pgSz = sect.getElementsByTagNameNS(W_NS, "pgSz")[0];
  if (pgSz && docGrid && pgMar) {
    const pageH = parseInt(wAttr(pgSz, "h"));
    const top    = parseInt(wAttr(pgMar, "top"));
    const bottom = parseInt(wAttr(pgMar, "bottom"));
    const linePitch = parseInt(wAttr(docGrid, "linePitch") || "0");
    const usableH = pageH - top - bottom;
    const lines = Math.round(usableH / linePitch);
    items.push(numericCheck("行数（計算値）", lines, 50, "行", 0));
  }

  return { title: "1. ページ基本設定", items };
}

function marginCheck(label, twips, expectedMm) {
  const mm = twipsToMm(twips);
  const ok = Math.abs(mm - expectedMm) <= TOLERANCE_MM;
  return makeCheck(label, ok ? "ok" : "error", {
    current: `${mm.toFixed(1)} mm`,
    expected: `${expectedMm} mm`,
    hint: ok ? null : `Wordの「レイアウト → 余白 → ユーザー設定の余白」で ${expectedMm}mm に設定してください。`,
    reference: REF.PAGE_SETUP,
  });
}

function numericCheck(label, actual, expected, unit, tolerance, hint) {
  const ok = Math.abs(actual - expected) <= tolerance;
  return makeCheck(label, ok ? "ok" : "error", {
    current: `${actual.toFixed(2)} ${unit}`,
    expected: `${expected} ${unit}`,
    hint: ok ? null : (hint || `${expected}${unit} に設定してください。`),
    reference: REF.PAGE_SETUP,
  });
}

function charSpaceCheck(label, actual, expected) {
  const ok = Math.abs(actual - expected) <= 5;
  return makeCheck(label, ok ? "ok" : "warn", {
    current: `内部値 ${actual}`,
    expected: `内部値 ${expected}（= 10.5pt 相当）`,
    hint: ok ? null : `Wordの「レイアウト → ページ設定 → 文字数と行数」で字送りを 10.5pt に設定してください。`,
    reference: REF.PAGE_SETUP,
  });
}

// === 2. ヘッダー内容 ===
function checkHeader(ctx) {
  const items = [];
  const oddXml  = ctx.headers["default"];
  const evenXml = ctx.headers["even"];

  if (!oddXml && !evenXml) {
    items.push(makeCheck("ヘッダーが設定されていません", "error"));
    return { title: "2. ヘッダー内容", items };
  }

  if (oddXml) {
    const texts = extractTexts(oddXml);
    const joined = texts.join("");

    const hasCenter = /沖縄県立総合教育センター/.test(joined);
    const hasBan    = /[^\s]+班/.test(joined);
    const banMatch  = extractMatch(joined, /沖縄県立総合教育センター[^\s]*?班/);
    items.push(makeCheck(
      "奇数ページ左上：班名",
      hasCenter && hasBan ? "ok" : "error",
      {
        current: banMatch || (joined.slice(0, 40) || "（未記載）"),
        expected: "沖縄県立総合教育センター○班",
        hint: hasCenter && hasBan ? null : "奇数ページのヘッダー左上に「沖縄県立総合教育センター」+ 班名を記載してください。",
        location: { area: "ヘッダー（奇数ページ・左上）" },
      }
    ));

    const shuroku = /長期研修員研究集録/.test(joined);
    const dai     = /第\s*\S+\s*集/.test(joined);
    const monthMatch = joined.match(/(\d{1,4}|令和\s*\S+)\s*年\s*([０-９0-9])\s*月/);
    const monthOK = monthMatch && /^[３3９9]$/.test(monthMatch[2]);
    const shurokuMatch = extractMatch(joined, /長期研修員研究集録[\s\S]*?月/) || "（未記載）";

    let s = "ok", hint = null;
    if (!shuroku || !dai || !monthMatch) {
      s = "error";
      hint = "ヘッダー右上に「長期研修員研究集録 第○集 ○○年○月」の形式で記載してください。";
    } else if (!monthOK) {
      s = "warn";
      hint = "月は 9月（前期）または 3月（後期・1年・インターバル）にしてください。";
    }
    items.push(makeCheck("奇数ページ右上：長期研修員研究集録", s, {
      current: shurokuMatch,
      expected: "長期研修員研究集録 第○集 ○○年○月（月は3月 or 9月）",
      hint,
      location: { area: "ヘッダー（奇数ページ・右上）" },
    }));

    items.push(...checkHeaderFonts(oddXml, "奇数ページヘッダー"));
  } else {
    items.push(makeCheck("奇数ページヘッダーが設定されていません", "error"));
  }

  if (evenXml) {
    const texts = extractTexts(evenXml);
    const joined = texts.join("").trim();
    if (joined.length === 0) {
      items.push(makeCheck("偶数ページヘッダー：執筆者名と研究テーマ", "error", {
        current: "（空）",
        expected: "左上：執筆者名 / 右上：研究テーマ",
        hint: "偶数ページのヘッダーに、左上に執筆者名、右上に研究テーマ（メインのみ）を記載してください。",
        location: { area: "ヘッダー（偶数ページ）" },
      }));
    } else {
      items.push(makeCheck("偶数ページヘッダー：執筆者名と研究テーマ", "warn", {
        current: joined.slice(0, 60) + (joined.length > 60 ? "…" : ""),
        expected: "左上：執筆者名 / 右上：研究テーマ",
        hint: "左上に執筆者名・右上に研究テーマが含まれているか目視で確認してください。",
        location: { area: "ヘッダー（偶数ページ）" },
      }));
    }
    items.push(...checkHeaderFonts(evenXml, "偶数ページヘッダー"));
  } else {
    items.push(makeCheck("偶数ページヘッダーが設定されていません", "error", {
      hint: "偶数ページにもヘッダーを設定してください（「ヘッダーは奇数及び偶数ページ別左右に設定」）。",
    }));
  }

  return { title: "2. ヘッダー内容", items };
}

function extractTexts(xmlStr) {
  const doc = parseXml(xmlStr);
  return Array.from(doc.getElementsByTagNameNS(W_NS, "t"))
    .map(t => t.textContent || "");
}

function checkHeaderFonts(xmlStr, label) {
  const doc = parseXml(xmlStr);
  const runs = Array.from(doc.getElementsByTagNameNS(W_NS, "r"));
  const fontIssues = [];
  const sizeIssues = [];

  for (const r of runs) {
    const text = Array.from(r.getElementsByTagNameNS(W_NS, "t"))
      .map(t => t.textContent).join("");
    if (!text.trim()) continue;

    const rFonts = r.getElementsByTagNameNS(W_NS, "rFonts")[0];
    const ea = rFonts ? (wAttr(rFonts, "eastAsia") || wAttr(rFonts, "ascii")) : null;
    if (ea && !isGothic(ea)) {
      fontIssues.push({ text: text.slice(0,20), font: ea });
    }

    const sz = r.getElementsByTagNameNS(W_NS, "sz")[0];
    if (sz) {
      const pt = parseInt(wAttr(sz, "val")) / 2;
      if (pt !== 9) sizeIssues.push({ text: text.slice(0,20), size: pt });
    }
  }

  return [
    makeCheck(`${label}：フォント`,
      fontIssues.length === 0 ? "ok" : "error",
      {
        current: fontIssues.length === 0
          ? "BIZ UDゴシック"
          : fontIssues.map(f => `「${f.text}」→ ${f.font}`).join(" / "),
        expected: "BIZ UDゴシック",
        hint: fontIssues.length === 0 ? null : "ヘッダー内のフォントを「BIZ UDゴシック」に変更してください。",
        location: { area: `ヘッダー（${label.replace("：フォント", "")}）` },
      }),
    makeCheck(`${label}：サイズ`,
      sizeIssues.length === 0 ? "ok" : "error",
      {
        current: sizeIssues.length === 0 ? "9pt" : sizeIssues.map(f => `「${f.text}」→ ${f.size}pt`).join(" / "),
        expected: "9pt",
        hint: sizeIssues.length === 0 ? null : "ヘッダー内のサイズを 9pt に変更してください。",
      }),
  ];
}

function extractMatch(text, regex) {
  const m = text.match(regex);
  return m ? m[0] : null;
}

// === 3. 原稿冒頭部（教科・領域名／テーマ／サブテーマ／所属氏名） ===
function checkOpeningStructure(ctx) {
  const items = [];
  const paras = ctx.paragraphs.filter(p => p.runs.length > 0 || p.text.length > 0);

  // 非空段落だけ取り出す
  const nonEmpty = ctx.paragraphs.filter(p => p.text.trim().length > 0);

  if (nonEmpty.length < 4) {
    items.push(makeCheck("原稿の冒頭部が確認できません（教科領域名・テーマ・サブテーマ・所属氏名）", "error"));
    return { title: "3. 原稿冒頭部", items };
  }

  // [0] 教科・領域名
  const subject = nonEmpty[0];
  items.push(...checkSubject(subject));

  // [1] テーマ
  const theme = nonEmpty[1];
  items.push(...checkTheme(theme));

  // [2..] サブテーマ（中央揃え・12pt・両端に —）— 0個以上、テーマ直後に連続
  let i = 2;
  const subThemes = [];
  while (i < nonEmpty.length) {
    const p = nonEmpty[i];
    const t = p.text.trim();
    if (p.alignment === "center" && (/^[—－―]/.test(t) || /[—－―]$/.test(t))) {
      subThemes.push(p);
      i++;
    } else break;
  }
  if (subThemes.length > 0) {
    items.push(...checkSubTheme(subThemes));
  } else {
    items.push(makeCheck("サブテーマ", "warn", {
      detail: "サブテーマが検出されませんでした。任意項目ですが、ある場合は両端に「—」を付けてください。"
    }));
  }

  // 所属・役職・氏名: 残りの中から「右揃え」段落を探す
  const nameP = nonEmpty.slice(i).find(p => p.alignment === "right");
  if (nameP) {
    items.push(...checkAuthor(nameP));
  } else {
    items.push(makeCheck("所属・役職・氏名（長期研修：右揃え）", "error", {
      detail: "右揃えの所属・役職・氏名行が見つかりません。長期研修は右揃えで記載してください。"
    }));
  }

  return { title: "3. 原稿冒頭部", items };
}

function checkSubject(p) {
  const items = [];
  const text = p.text.trim();
  const loc = { ...p.location, snippet: text.slice(0, 40) };
  const REF_ = REF.SUBJECT;

  const ok_bracket = /^〈.+〉$/.test(text);
  items.push(makeCheck("教科・領域名：〈 〉でくくる", ok_bracket ? "ok" : "error", {
    current: text || "（未記載）",
    expected: "〈○○〉（例: 〈音楽〉）",
    hint: ok_bracket ? null : "1行目の教科名を全角の山括弧「〈 」「 〉」でくくってください。半角の「< >」は使えません。",
    location: loc, reference: REF_,
  }));

  const alignOk = (p.alignment === "left" || p.alignment == null);
  items.push(makeCheck("教科・領域名：左揃え", alignOk ? "ok" : "error", {
    current: alignmentLabel(p.alignment),
    expected: "左揃え",
    hint: alignOk ? null : "Wordで該当行を選択し、ホームタブの「左揃え」をクリックしてください。",
    location: loc, reference: REF_,
  }));

  const run = firstSignificantRun(p);
  if (run) {
    items.push({ ...fontCheck("教科・領域名：フォント", run, isGothic, "BIZ UDゴシック", loc), reference: REF_ });
    items.push({ ...sizeCheck("教科・領域名：サイズ", run, 12, loc), reference: REF_ });
  }

  if (/[<>]/.test(p.text)) {
    items.push(makeCheck("教科・領域名：半角不等号 < >", "error", {
      current: "< または > を使用中",
      expected: "〈 〉（全角山括弧）",
      hint: "半角の「< >」を全角の「〈 〉」に変更してください。",
      location: loc, reference: REF.ANGLE_BRACKET,
    }));
  }
  return items;
}

function checkTheme(p) {
  const items = [];
  const loc = { ...p.location, snippet: p.text.slice(0, 40) };
  items.push(makeCheck("テーマ：中央揃え", p.alignment === "center" ? "ok" : "error", {
    current: alignmentLabel(p.alignment), expected: "中央揃え",
    hint: p.alignment === "center" ? null : "テーマ行を選択し「中央揃え」にしてください。",
    location: loc, reference: REF.THEME,
  }));
  const run = firstSignificantRun(p);
  if (run) {
    items.push({ ...fontCheck("テーマ：フォント", run, isGothic, "BIZ UDゴシック", loc), reference: REF.THEME });
    items.push({ ...sizeCheck("テーマ：サイズ", run, 16, loc), reference: REF.THEME });
  }
  return items;
}

function checkSubTheme(paras) {
  const items = [];
  const joined = paras.map(p => p.text).join("");
  const loc = { ...paras[0].location, snippet: joined.slice(0, 40) };

  const hasOpen  = /^[—－―]/.test(paras[0].text.trim());
  const hasClose = /[—－―]$/.test(paras[paras.length-1].text.trim());

  items.push(makeCheck("サブテーマ：両端に「—」（全角ハイフン200%）", hasOpen && hasClose ? "ok" : "error", {
    current: `先頭: ${hasOpen ? "—あり" : "なし"} / 末尾: ${hasClose ? "—あり" : "なし"}`,
    expected: "— サブテーマ —（両端に全角ハイフン）",
    hint: hasOpen && hasClose ? null : "サブテーマの先頭と末尾に全角ハイフン「—」を入れ、文字幅200%に設定してください。",
    location: loc, reference: REF.SUBTHEME,
  }));

  for (const p of paras) {
    items.push(makeCheck("サブテーマ：中央揃え", p.alignment === "center" ? "ok" : "error", {
      current: alignmentLabel(p.alignment), expected: "中央揃え",
      location: { ...p.location, snippet: p.text.slice(0, 40) }, reference: REF.SUBTHEME,
    }));
  }

  const run = firstSignificantRun(paras[0]);
  if (run) {
    items.push({ ...fontCheck("サブテーマ：フォント", run, isGothic, "BIZ UDゴシック", loc), reference: REF.SUBTHEME });
    items.push({ ...sizeCheck("サブテーマ：サイズ", run, 12, loc), reference: REF.SUBTHEME });
  }
  return items;
}

function checkAuthor(p) {
  const items = [];
  const loc = { ...p.location, snippet: p.text.slice(0, 40) };
  items.push(makeCheck("所属・役職・氏名：右揃え", p.alignment === "right" ? "ok" : "error", {
    current: alignmentLabel(p.alignment), expected: "右揃え（長期研修）",
    hint: p.alignment === "right" ? null : "長期研修の所属・役職・氏名は右揃えにしてください。",
    location: loc, reference: REF.AUTHOR,
  }));
  const run = firstSignificantRun(p);
  if (run) {
    items.push({ ...fontCheck("所属・役職・氏名：フォント", run, isGothic, "BIZ UDゴシック", loc), reference: REF.AUTHOR });
    items.push({ ...sizeCheck("所属・役職・氏名：サイズ", run, 12, loc), reference: REF.AUTHOR });
  }
  if (!/　$/.test(p.text)) {
    items.push(makeCheck("所属・役職・氏名：末尾に全角スペース", "warn", {
      current: "末尾に全角スペースなし", expected: "末尾に全角スペース「　」",
      hint: "氏名の最後に全角スペースを1つ追加してください。",
      location: loc, reference: REF.AUTHOR,
    }));
  }
  return items;
}

function alignmentLabel(a) {
  return ({ left: "左揃え", right: "右揃え", center: "中央揃え", both: "両端揃え" }[a]) || "左揃え";
}

// === 4. 要旨・キーワード ===
function checkAbstractAndKeyword(ctx) {
  const items = [];
  const nonEmpty = ctx.paragraphs.filter(p => p.text.trim().length > 0);

  // キーワード段落を探す（「キーワード」で始まる）
  const keywordIdx = nonEmpty.findIndex(p => /^キーワード/.test(p.text.trim()));
  if (keywordIdx === -1) {
    items.push(makeCheck("キーワード行が見つかりません", "error"));
    return { title: "4. 要旨・キーワード", items };
  }

  // 要旨はキーワードの直前にある（300字程度の本文）
  // 「所属氏名」の段落の後、キーワードの前で、最も文字数が多い段落を要旨と判定
  const authorIdx = nonEmpty.findIndex(p => p.alignment === "right");
  const startIdx = authorIdx >= 0 ? authorIdx + 1 : 0;
  const between = nonEmpty.slice(startIdx, keywordIdx);
  const abstract = between.slice().sort((a,b) => b.text.length - a.text.length)[0];

  if (!abstract) {
    items.push(makeCheck("要旨が見つかりません", "error"));
  } else {
    items.push(...checkAbstract(abstract));
  }

  items.push(...checkKeyword(nonEmpty[keywordIdx]));
  return { title: "4. 要旨・キーワード", items };
}

function checkAbstract(p) {
  const items = [];
  const text = p.text;
  const len = text.length;
  const loc = { ...p.location, snippet: text.slice(0, 50) + "…" };
  const R = REF.ABSTRACT;

  const inRange = len >= 200 && len <= 450;
  items.push(makeCheck("要旨：文字数（300字程度）", inRange ? "ok" : "warn", {
    current: `${len}字`, expected: "300字程度（200〜450字）",
    hint: inRange ? null : len < 200
      ? `${len}字は短すぎます。300字程度を目安に記述してください。`
      : `${len}字は長すぎます。300字程度を目安に推敲してください。`,
    location: loc, reference: R,
  }));

  const alignOk = (p.alignment === "left" || p.alignment == null);
  items.push(makeCheck("要旨：左揃え", alignOk ? "ok" : "error", {
    current: alignmentLabel(p.alignment), expected: "左揃え",
    location: loc, reference: R,
  }));

  const run = firstSignificantRun(p);
  if (run) {
    items.push({ ...fontCheck("要旨：フォント", run, isMincho, "BIZ UD明朝 Medium", loc), reference: R });
    items.push({ ...sizeCheck("要旨：サイズ", run, 10, loc), reference: R });
  }
  return items;
}

function checkKeyword(p) {
  const items = [];
  const text = p.text;
  const loc = { ...p.location, snippet: text.slice(0, 40) };
  const R = REF.KEYWORD;

  const ok_space = /^キーワード　　/.test(text);
  items.push(makeCheck("キーワード：見出し後の全角スペース2つ", ok_space ? "ok" : "error", {
    current: text.slice(0, 12) + "…",
    expected: "キーワード　　○○○ ○○○ ○○○",
    hint: ok_space ? null : "「キーワード」の直後に全角スペースを2つ入れてください。",
    location: loc, reference: R,
  }));

  const contentMatch = text.match(/^キーワード　*(.+)$/);
  if (contentMatch) {
    const content = contentMatch[1].trim();
    const tokens = content.split(/　+/).filter(t => t.length > 0);
    const ok_count = tokens.length >= 3 && tokens.length <= 5;
    items.push(makeCheck("キーワード：個数（3〜5個）", ok_count ? "ok" : "error", {
      current: `${tokens.length}個：${tokens.join(" / ")}`,
      expected: "3〜5個（全角スペース区切り）",
      hint: ok_count ? null : "キーワードは3〜5個に調整してください。",
      location: loc, reference: R,
    }));
  }

  // 「キーワード」を含む run までを見出し、それ以降を内容とみなす
  const headerNGs = [], contentNGs = [];
  let headerSizeNG = null, contentSizeNG = null;
  let passedKeyword = false;
  for (const r of p.runs) {
    if (!r.text.trim()) continue;
    const ea = getEaFont(r.resolved);
    const sz = getSizePt(r.resolved);
    if (!passedKeyword) {
      if (ea && !isGothic(ea)) headerNGs.push({ text: r.text, font: ea });
      if (sz !== null && sz !== 10) headerSizeNG = sz;
      if (/キーワード/.test(r.text)) passedKeyword = true;
    } else {
      if (ea && !isMincho(ea)) contentNGs.push({ text: r.text.slice(0,20), font: ea });
      if (sz !== null && sz !== 10) contentSizeNG = sz;
    }
  }
  items.push(makeCheck("キーワード見出し：フォント",
    headerNGs.length === 0 ? "ok" : "error",
    {
      current: headerNGs.length === 0 ? "BIZ UDゴシック" : headerNGs.map(n=>`「${n.text}」→ ${n.font}`).join(" / "),
      expected: "BIZ UDゴシック",
      location: loc, reference: R,
    }));
  items.push(makeCheck("キーワード見出し：サイズ",
    headerSizeNG === null ? "ok" : "error",
    { current: headerSizeNG === null ? "10pt" : `${headerSizeNG}pt`, expected: "10pt", location: loc, reference: R }));
  items.push(makeCheck("キーワード内容：フォント",
    contentNGs.length === 0 ? "ok" : "error",
    {
      current: contentNGs.length === 0 ? "BIZ UD明朝 Medium" : contentNGs.map(n=>`「${n.text}」→ ${n.font}`).join(" / "),
      expected: "BIZ UD明朝 Medium",
      location: loc, reference: R,
    }));
  items.push(makeCheck("キーワード内容：サイズ",
    contentSizeNG === null ? "ok" : "error",
    { current: contentSizeNG === null ? "10pt" : `${contentSizeNG}pt`, expected: "10pt", location: loc, reference: R }));
  return items;
}

// === 5. 本文と見出し ===
function checkBodyAndHeadings(ctx) {
  const items = [];
  const paras = ctx.paragraphs;

  // キーワード段落以降を本文領域とする
  const kwIdx = paras.findIndex(p => /^キーワード/.test(p.text.trim()));
  if (kwIdx === -1) {
    items.push(makeCheck("本文領域の特定ができませんでした", "warn"));
    return { title: "5. 本文・見出し・図表タイトル", items };
  }

  // 末尾セクション（〈注記〉〈引用文献〉〈生成AI利用〉）の手前まで
  const endIdx = paras.findIndex((p, i) => i > kwIdx && /^〈(注記|引用文献|生成\s*AI\s*利用)〉/.test(p.text.trim()));
  const bodyParas = paras.slice(kwIdx + 1, endIdx === -1 ? undefined : endIdx);

  const bigHeads = [];   // 大見出し: "N. xxx" または "N　xxx"
  const smallHeads = []; // 小見出し: "N.M. xxx" または "N.M xxx"（末尾ピリオドの有無を許容）
  const bodyTexts = [];  // 本文

  // 小見出しを先に判定（大見出しのパターンとも前方一致するため）
  for (const p of bodyParas) {
    const t = p.text.trim();
    if (!t) continue;
    if (/^\d+\.\d+\.?[　\s]/.test(t)) {
      smallHeads.push(p);
    } else if (/^\d+\.[　\s]/.test(t)) {
      bigHeads.push(p);
    } else {
      bodyTexts.push(p);
    }
  }

  // 図表タイトル: テキストボックス内の段落から「図N」「表N」で始まるものを抽出
  const figTitles = (ctx.textboxParagraphs || []).filter(p => /^(図|表)\s*\d+/.test(p.text.trim()));

  // 大見出し
  if (bigHeads.length === 0) {
    items.push(makeCheck("大見出し（1. はじめに 等）", "warn", {
      current: "未検出",
      expected: "1. ○○○、2. ○○○、…",
      hint: "「1. はじめに」「2. 方法」のような大見出しが見つかりません。",
      reference: REF.BIG_HEAD,
    }));
  } else {
    items.push(...checkBatchFontSize(bigHeads, "大見出し", isGothic, "BIZ UDゴシック", 12, REF.BIG_HEAD));

    const nums = bigHeads.map(h => parseInt(h.text.match(/^(\d+)\./)[1]));
    const sequential = nums.every((n, i) => n === i + 1);
    items.push(makeCheck("大見出し：連番", sequential ? "ok" : "warn", {
      current: nums.join(", "), expected: "1, 2, 3, … と昇順",
      hint: sequential ? null : "大見出しの番号が連番になっていません。順序を確認してください。",
      reference: REF.BIG_HEAD,
    }));

    const punctNGs = bigHeads.filter(h => !/^\d+\.　/.test(h.text));
    items.push(makeCheck("大見出し：数字の後（半角ピリオド＋全角スペース）",
      punctNGs.length === 0 ? "ok" : "error", {
      current: punctNGs.length === 0 ? "全て正しい" : punctNGs.slice(0,3).map(h => `「${h.text.slice(0,30)}」`).join(" / "),
      expected: "1.（半角ピリオド）＋ 全角スペース ＋ 見出し文字",
      hint: punctNGs.length === 0 ? null : "数字の後は半角ピリオドと全角スペース1つを入れてください。",
      location: punctNGs[0] ? { ...punctNGs[0].location, snippet: punctNGs[0].text.slice(0,30) } : null,
      reference: REF.BIG_HEAD,
    }));
  }

  // 小見出し
  if (smallHeads.length > 0) {
    items.push(...checkBatchFontSize(smallHeads, "小見出し", isGothic, "BIZ UDゴシック", 10, REF.SMALL_HEAD));

    // 数字の後（N.M. + 全角スペース）形式チェック
    const punctNGs = smallHeads.filter(h => !/^\d+\.\d+\.　/.test(h.text));
    if (punctNGs.length > 0) {
      items.push(makeCheck("小見出し：数字の後（N.M. + 全角スペース）", "warn", {
        current: punctNGs.slice(0,3).map(h => `「${h.text.slice(0,30)}」`).join(" / "),
        expected: "1.1. （末尾ピリオド）＋ 全角スペース ＋ 見出し文字",
        hint: "小見出しは『1.1.』のように末尾にピリオドを付け、その後に全角スペースを入れてください。",
        location: punctNGs[0] ? { ...punctNGs[0].location, snippet: punctNGs[0].text.slice(0,30) } : null,
        reference: REF.SMALL_HEAD,
      }));
    }
  }

  // 図表タイトル
  if (figTitles.length > 0) {
    items.push(...checkBatchFontSize(figTitles, "図表タイトル", isGothic, "BIZ UDゴシック", 9, REF.FIG_TITLE));
    const alignNGs = figTitles.filter(h => h.alignment !== "center");
    items.push(makeCheck("図表タイトル：中央揃え",
      alignNGs.length === 0 ? "ok" : "warn", {
      current: alignNGs.length === 0 ? `全${figTitles.length}件中央揃え` : `${alignNGs.length}件が中央揃えでない`,
      expected: "中央揃え",
      hint: alignNGs.length === 0 ? null : "テキストボックスを選択し、テキストを中央揃えにしてください。",
      location: alignNGs[0] ? { ...alignNGs[0].location, snippet: alignNGs[0].text.slice(0,30) } : null,
      reference: REF.FIG_TITLE,
    }));
  }

  // 本文
  if (bodyTexts.length > 0) {
    const fontNGs = [], szNGs = [];
    for (const b of bodyTexts) {
      for (const r of b.runs) {
        if (!r.text.trim()) continue;
        const sz = getSizePt(r.resolved);
        if (sz !== null && sz !== 10) {
          szNGs.push({ text: r.text.slice(0,30), size: sz, location: b.location });
        }
        const hasJa = /[぀-ゟ゠-ヿ一-鿿]/.test(r.text);
        if (hasJa) {
          const ea = getEaFont(r.resolved);
          if (ea && !isMincho(ea)) {
            fontNGs.push({ text: r.text.slice(0,30), font: ea, location: b.location });
          }
        }
      }
    }
    items.push(makeCheck("本文：フォント（日本語部分）",
      fontNGs.length === 0 ? "ok" : "error", {
      current: fontNGs.length === 0
        ? "BIZ UD明朝 Medium"
        : fontNGs.slice(0,3).map(n=>`「${n.text}」→ ${n.font}`).join(" / ") + (fontNGs.length>3?` 他${fontNGs.length-3}件`:""),
      expected: "BIZ UD明朝 Medium",
      hint: fontNGs.length === 0 ? null : "本文の日本語部分のフォントを BIZ UD明朝 Medium に統一してください。",
      location: fontNGs[0] ? { ...fontNGs[0].location, snippet: fontNGs[0].text } : null,
      reference: REF.BODY,
    }));
    items.push(makeCheck("本文：サイズ",
      szNGs.length === 0 ? "ok" : "error", {
      current: szNGs.length === 0
        ? "10pt"
        : szNGs.slice(0,3).map(n=>`「${n.text}」→ ${n.size}pt`).join(" / ") + (szNGs.length>3?` 他${szNGs.length-3}件`:""),
      expected: "10pt",
      hint: szNGs.length === 0 ? null : "本文のサイズを 10pt に統一してください。",
      location: szNGs[0] ? { ...szNGs[0].location, snippet: szNGs[0].text } : null,
      reference: REF.BODY,
    }));
  }

  return { title: "5. 本文・見出し・図表タイトル", items };
}

// 見出し類の一括フォント・サイズチェック（first NG の位置を返す）
function checkBatchFontSize(paras, label, fontPredicate, expectedFontName, expectedSize, reference) {
  const fontNGs = [];
  const szNGs = [];
  for (const p of paras) {
    const run = firstSignificantRun(p);
    if (!run) continue;
    const ea = getEaFont(run.resolved);
    const sz = getSizePt(run.resolved);
    if (ea && !fontPredicate(ea)) {
      fontNGs.push({ text: p.text.slice(0,30), font: ea, location: p.location });
    }
    if (sz !== null && sz !== expectedSize) {
      szNGs.push({ text: p.text.slice(0,30), size: sz, location: p.location });
    }
  }
  return [
    makeCheck(`${label}：フォント`,
      fontNGs.length === 0 ? "ok" : "error", {
      current: fontNGs.length === 0
        ? expectedFontName
        : fontNGs.slice(0,3).map(n=>`「${n.text}」→ ${n.font}`).join(" / ") + (fontNGs.length>3?` 他${fontNGs.length-3}件`:""),
      expected: expectedFontName,
      hint: fontNGs.length === 0 ? null : `${label}のフォントを ${expectedFontName} に変更してください。`,
      location: fontNGs[0] ? { ...fontNGs[0].location, snippet: fontNGs[0].text } : null,
      reference,
    }),
    makeCheck(`${label}：サイズ`,
      szNGs.length === 0 ? "ok" : "error", {
      current: szNGs.length === 0
        ? `${expectedSize}pt`
        : szNGs.slice(0,3).map(n=>`「${n.text}」→ ${n.size}pt`).join(" / ") + (szNGs.length>3?` 他${szNGs.length-3}件`:""),
      expected: `${expectedSize}pt`,
      hint: szNGs.length === 0 ? null : `${label}のサイズを ${expectedSize}pt に変更してください。`,
      location: szNGs[0] ? { ...szNGs[0].location, snippet: szNGs[0].text } : null,
      reference,
    }),
  ];
}

// === 6. 末尾構造（〈注記〉〈引用文献〉〈生成AI利用〉） ===
function checkClosingStructure(ctx) {
  const items = [];
  const paras = ctx.paragraphs;
  const text = paras.map(p => p.text).join("\n");

  const hasNote   = /〈注記〉/.test(text);
  const hasRef    = /〈引用文献〉/.test(text);
  const hasAI     = /〈生成\s*AI\s*利用〉/.test(text);
  const hasAIFull = /〈生成\s*[ＡAＩI]+\s*利用〉/.test(text);

  items.push(makeCheck("〈引用文献〉セクション", hasRef ? "ok" : "error", {
    current: hasRef ? "あり" : "なし",
    expected: "〈引用文献〉として末尾に記載",
    hint: hasRef ? null : "末尾に〈引用文献〉セクションを追加してください。",
    reference: REF.REFERENCES,
  }));

  items.push(makeCheck("〈注記〉セクション（任意）", hasNote ? "ok" : "warn", {
    current: hasNote ? "あり" : "なし",
    expected: "注記がある場合は〈注記〉として末尾に記載",
    hint: hasNote ? null : "本文中に注記（注1）等）がある場合は〈注記〉セクションを末尾にまとめてください。",
    reference: REF.CLOSING,
  }));

  if (hasAIFull && !hasAI) {
    items.push(makeCheck("〈生成AI利用〉表記の AI 文字", "error", {
      current: "〈生成ＡＩ利用〉（全角ＡＩ）",
      expected: "〈生成AI利用〉（半角AI）",
      hint: "「ＡＩ」を半角の「AI」に修正してください。",
      reference: REF.AI_USE,
    }));
  } else if (hasAI) {
    items.push(makeCheck("〈生成AI利用〉セクション（任意）", "ok", { current: "あり", reference: REF.AI_USE }));
  } else {
    items.push(makeCheck("〈生成AI利用〉セクション（任意）", "warn", {
      current: "なし",
      expected: "生成AIを利用した場合は〈生成AI利用〉として記載",
      hint: "生成AIを利用した場合は規定の文言（『本論文の執筆にあたり…』）を〈生成AI利用〉として末尾に記載してください。",
      reference: REF.AI_USE,
    }));
  }

  const order = [];
  for (const p of paras) {
    if (/^〈注記〉/.test(p.text.trim())) order.push("注記");
    else if (/^〈引用文献〉/.test(p.text.trim())) order.push("引用文献");
    else if (/^〈生成\s*[AＡ][IＩ]\s*利用〉/.test(p.text.trim())) order.push("生成AI利用");
  }
  const expectedOrder = ["注記","引用文献","生成AI利用"].filter(x => order.includes(x));
  const orderOK = JSON.stringify(order) === JSON.stringify(expectedOrder);
  if (order.length >= 2) {
    items.push(makeCheck("末尾見出しの順序", orderOK ? "ok" : "warn", {
      current: order.join(" → "),
      expected: expectedOrder.join(" → "),
      hint: orderOK ? null : "末尾の見出しは「注記 → 引用文献 → 生成AI利用」の順に並べてください。",
      reference: REF.CLOSING,
    }));
  }

  // 末尾見出しのフォント・サイズチェック（12pt BIZ UDゴシック）
  const closingHeads = paras.filter(p =>
    /^〈(注記|引用文献|生成\s*[AＡ][IＩ]\s*利用)〉/.test(p.text.trim())
  );
  if (closingHeads.length > 0) {
    items.push(...checkBatchFontSize(closingHeads, "末尾見出し", isGothic, "BIZ UDゴシック", 12, REF.CLOSING));

    // 〈 〉でくくられているか
    const bracketNGs = closingHeads.filter(p => !/^〈.+〉$/.test(p.text.trim()));
    if (bracketNGs.length > 0) {
      items.push(makeCheck("末尾見出し：〈 〉でくくる", "error", {
        current: bracketNGs.map(p => `「${p.text.slice(0,20)}」`).join(" / "),
        expected: "〈注記〉〈引用文献〉〈生成AI利用〉",
        hint: "末尾見出しは全角の山括弧〈 〉でくくってください。",
        reference: REF.CLOSING,
      }));
    }
  }

  return { title: "6. 末尾構造", items };
}

// === 引用文献リストのチェック ===
function checkReferences(ctx) {
  const items = [];
  const paras = ctx.paragraphs;

  // 〈引用文献〉見出しを探す
  const refHeadIdx = paras.findIndex(p => /^〈引用文献〉/.test(p.text.trim()));
  if (refHeadIdx === -1) {
    items.push(makeCheck("〈引用文献〉セクションが見つかりません", "error", {
      hint: "末尾に〈引用文献〉セクションを追加してください。",
    }));
    return { title: "6.5. 引用文献リスト", items };
  }

  // 次の末尾セクション見出し（〈生成AI利用〉等）まで
  const endIdx = paras.findIndex((p, i) => i > refHeadIdx && /^〈[^〉]*〉/.test(p.text.trim()));
  const entries = paras.slice(refHeadIdx + 1, endIdx === -1 ? undefined : endIdx)
    .filter(p => p.text.trim().length > 0);

  if (entries.length === 0) {
    items.push(makeCheck("引用文献リスト：項目数", "warn", {
      current: "0件",
      expected: "1件以上",
      hint: "〈引用文献〉セクション内に文献を記載してください。",
    }));
    return { title: "6.5. 引用文献リスト", items };
  }

  items.push(makeCheck("引用文献リスト：項目数", "ok", {
    current: `${entries.length}件`,
  }));

  // --- A. フォント・サイズチェック ---
  const fontNGs = [];   // 日本語部分が BIZ UD明朝 Medium 以外
  const sizeNGs = [];   // 9pt 以外
  const enFontNGs = []; // 英文部分が Times New Roman 以外

  for (const p of entries) {
    for (const r of p.runs) {
      if (!r.text.trim()) continue;
      const sz = getSizePt(r.resolved);
      if (sz !== null && sz !== 9) {
        sizeNGs.push({ snippet: r.text.slice(0,30), location: p.location });
      }
      const hasJa = /[぀-ゟ゠-ヿ一-鿿]/.test(r.text);
      const hasAlphaOnly = /^[\x20-\x7E]+$/.test(r.text) && /[A-Za-z]/.test(r.text);
      if (hasJa) {
        const ea = getEaFont(r.resolved);
        if (ea && !isMincho(ea)) {
          fontNGs.push({ snippet: `「${r.text.slice(0,30)}」→ ${ea}`, location: p.location });
        }
      } else if (hasAlphaOnly) {
        // 英文ラン: Times New Roman か
        const asc = getAsciiFont(r.resolved);
        if (asc && !isTimes(asc)) {
          enFontNGs.push({ snippet: `「${r.text.slice(0,30)}」→ ${asc}`, location: p.location });
        }
      }
    }
  }

  items.push(makeCheck("文献リスト：サイズ（9pt）",
    sizeNGs.length === 0 ? "ok" : "error", {
    current: sizeNGs.length === 0 ? "9pt" : `${sizeNGs.length}件が9pt以外`,
    expected: "9pt",
    hint: sizeNGs.length === 0 ? null : "文献リスト全体を選択し、サイズを 9pt に変更してください。",
    locations: sizeNGs.length > 0 ? sizeNGs.slice(0,20) : null,
  }));
  items.push(makeCheck("文献リスト：日本語フォント（BIZ UD明朝 Medium）",
    fontNGs.length === 0 ? "ok" : "error", {
    current: fontNGs.length === 0 ? "BIZ UD明朝 Medium" : `${fontNGs.length}件`,
    expected: "BIZ UD明朝 Medium",
    hint: fontNGs.length === 0 ? null : "文献リストの日本語部分のフォントを BIZ UD明朝 Medium に変更してください。",
    locations: fontNGs.length > 0 ? fontNGs.slice(0,20) : null,
  }));
  items.push(makeCheck("文献リスト：英文フォント（Times New Roman）",
    enFontNGs.length === 0 ? "ok" : "error", {
    current: enFontNGs.length === 0 ? "Times New Roman" : `${enFontNGs.length}件`,
    expected: "Times New Roman",
    hint: enFontNGs.length === 0 ? null : "英文の文献は Times New Roman に変更してください。",
    locations: enFontNGs.length > 0 ? enFontNGs.slice(0,20) : null,
  }));

  // --- B. 各項目の表記チェック ---
  // 発行年は全角括弧 （YYYY）
  // 書籍名は 『』、論文名は「」、Web は「」+ URL + （YYYY年M月参照）
  const yearParenNGs = [];     // 半角括弧で発行年を書いている
  const yearMissing = [];      // 発行年が見つからない
  const webMissingDate = [];   // URL があるが参照年月がない
  const webHasUnderline = [];  // Web の URL に下線あり
  const punctMidNGs = [];      // 著者間の区切りが中黒/全角コンマ以外

  for (const p of entries) {
    const text = p.text;
    const loc = { ...p.location, snippet: text.slice(0,40) };

    // 全角括弧での発行年
    const hasFwYear = /（\s*\d{4}\s*）|（\s*\d{4}[a-z]?\s*）/.test(text);
    // 半角括弧での発行年
    const hasHwYear = /\(\s*\d{4}\s*\)|\(\s*\d{4}[a-z]?\s*\)/.test(text);
    if (!hasFwYear && !hasHwYear) {
      // 発行年自体がない（短い項目など）
      if (text.length > 10) yearMissing.push({ snippet: text.slice(0,40), location: p.location });
    } else if (!hasFwYear && hasHwYear) {
      yearParenNGs.push({ snippet: text.slice(0,40), location: p.location });
    }

    // URL を含むか
    const hasUrl = /https?:\/\//.test(text);
    if (hasUrl) {
      // 末尾に「（YYYY年M月参照）」があるか
      const refDate = /（\s*\d{4}\s*年\s*\d{1,2}\s*月\s*参照\s*）/.test(text);
      if (!refDate) {
        webMissingDate.push({ snippet: text.slice(0,40), location: p.location });
      }
      // 下線が引かれていないか
      for (const r of p.runs) {
        if (r.resolved.underline && /https?:\/\//.test(r.text)) {
          webHasUnderline.push({ snippet: r.text.slice(0,40), location: p.location });
          break;
        }
      }
    }

    // 著者間の区切り（カンマ or 中黒）
    // 半角カンマで「著者間」を区切っているのが違反。ただし以下は OK:
    //  - 単著で「Krashen, S. D.」のように "姓, 名のイニシャル" 形式の半角カンマ
    //  - 全角コンマ「，」または中黒「・」で区切られているもの
    // 違反パターン例:
    //  - 欧文複数著者: "Smith, J., Brown, B." → 半角カンマで著者間を分けている
    //  - 和文複数著者: "大熊信彦,酒井美恵子" → 半角カンマで著者間を分けている
    const beforeYear = text.split(/[（(]\s*\d{4}/)[0] || "";
    // 欧文: 「, 」の後に「[A-Z][a-z]+」（大文字始まりの単語＝姓）が続く場合は複数著者
    const enMultiAuthor = /,\s*[A-Z][a-z]+/.test(beforeYear);
    // 和文: 「,」の後に日本語文字が続く場合は複数著者
    const jaMultiAuthor = /,\s*[぀-ヿ一-鿿]/.test(beforeYear);
    if ((enMultiAuthor || jaMultiAuthor) && !/[，・]/.test(beforeYear)) {
      punctMidNGs.push({ snippet: text.slice(0,40), location: p.location });
    }
  }

  items.push(makeCheck("文献リスト：発行年の括弧",
    yearParenNGs.length === 0 ? "ok" : "error", {
    current: yearParenNGs.length === 0 ? "全角括弧（YYYY）" : `${yearParenNGs.length}件が半角括弧`,
    expected: "全角括弧 （YYYY） を使用",
    hint: yearParenNGs.length === 0 ? null : "発行年を全角括弧（）でくくってください。",
    locations: yearParenNGs.length > 0 ? yearParenNGs : null,
  }));

  if (yearMissing.length > 0) {
    items.push(makeCheck("文献リスト：発行年の記載", "warn", {
      current: `${yearMissing.length}件で発行年が確認できません`,
      expected: "著者名（発行年）の形式",
      hint: "各文献に発行年を記載してください。",
      locations: yearMissing,
    }));
  }

  if (punctMidNGs.length > 0) {
    items.push(makeCheck("文献リスト：著者間の区切り", "error", {
      current: `${punctMidNGs.length}件で半角カンマ使用`,
      expected: "中黒（・）／欧文は全角コンマ（，）",
      hint: "著者間の区切りは「・」（和文）または「，」（欧文）を使ってください。",
      locations: punctMidNGs,
    }));
  }

  // URL を含む項目があるかどうかで Web チェックを表示
  const hasAnyUrl = entries.some(p => /https?:\/\//.test(p.text));
  if (hasAnyUrl) {
    items.push(makeCheck("Web資料：参照年月の記載",
      webMissingDate.length === 0 ? "ok" : "error", {
      current: webMissingDate.length === 0
        ? "全Web資料に参照年月あり"
        : `${webMissingDate.length}件で「（YYYY年M月参照）」が見つかりません`,
      expected: "URL の後に「（YYYY年M月参照）」",
      hint: webMissingDate.length === 0 ? null : "Web資料の末尾に参照した年月（西暦）を全角括弧でくくって記載してください。",
      locations: webMissingDate.length > 0 ? webMissingDate : null,
    }));
    items.push(makeCheck("Web資料：URLに下線がない",
      webHasUnderline.length === 0 ? "ok" : "error", {
      current: webHasUnderline.length === 0 ? "下線なし" : `${webHasUnderline.length}件に下線あり`,
      expected: "下線なし・黒文字",
      hint: webHasUnderline.length === 0 ? null : "URL の下線を解除し、文字色を黒に変更してください。",
      locations: webHasUnderline.length > 0 ? webHasUnderline : null,
    }));
  }

  // --- C. 並び順（目視確認を促す）---
  // 自動判定は欧文混在で誤判定が起きやすいため、各項目の先頭著者名を一覧表示して目視で確認してもらう
  const heads = entries.map((p, i) => {
    const t = p.text.replace(/^[　 \s]+/, "");
    const head = t.match(/^([^（(]+?)(?=（|\()/);
    return { snippet: `${i+1}. ${head ? head[1].trim() : t.slice(0,20)}`, location: p.location };
  });
  items.push(makeCheck("文献リスト：著者名の並び順", "warn", {
    current: `${heads.length}件`,
    expected: "著者名の五十音順（Web も含めて一括、同一著者は発行年の新しい順）",
    hint: "並び順は手動で確認してください。下記に各文献の先頭著者を一覧表示します。",
    locations: heads,
  }));

  return { title: "6.5. 引用文献リスト", items };
}

// === 7. 禁止事項のチェック ===
function checkProhibited(ctx) {
  const items = [];

  const boldHits = [], underlineHits = [];
  for (const p of ctx.paragraphs) {
    for (const r of p.runs) {
      if (!r.text.trim()) continue;
      if (r.resolved.bold)      boldHits.push({ snippet: r.text.slice(0,30), location: p.location });
      if (r.resolved.underline) underlineHits.push({ snippet: r.text.slice(0,30), location: p.location });
    }
  }
  items.push(makeCheck("太字の使用", boldHits.length === 0 ? "ok" : "error", {
    current: boldHits.length === 0 ? "なし" : `${boldHits.length}件`,
    expected: "太字は使用しない",
    hint: boldHits.length === 0 ? null : "該当箇所の太字（Ctrl+B）を解除してください。",
    locations: boldHits.length > 0 ? boldHits : null,
    reference: REF.PROHIBITED,
  }));
  items.push(makeCheck("下線の使用", underlineHits.length === 0 ? "ok" : "error", {
    current: underlineHits.length === 0 ? "なし" : `${underlineHits.length}件`,
    expected: "下線は使用しない（Web資料のリンクも下線なし）",
    hint: underlineHits.length === 0 ? null : "該当箇所の下線（Ctrl+U）を解除してください。",
    locations: underlineHits.length > 0 ? underlineHits : null,
    reference: REF.PROHIBITED,
  }));

  // 特殊文字
  const specialHits = findCharOccurrences(ctx, /[℡㌕㌻㌃㌍㌔㌖㌘㌢㌣㌤㌥㌦㌧㌨㌩㌪㌫㌬㌭㌮㌯㌰㌱㌲㌳㌴㌵㌶㌷㌸㌹㌺㌻㌼㌽㌾㌿]/g);
  items.push(makeCheck("特殊文字（℡ ㌕ ㌻ 等）", specialHits.length === 0 ? "ok" : "error", {
    current: specialHits.length === 0 ? "なし" : `${specialHits.length}件`,
    expected: "使用しない",
    hint: specialHits.length === 0 ? null : "該当文字を通常の表記（例: ℡ → TEL）に書き換えてください。",
    locations: specialHits.length > 0 ? specialHits : null,
    reference: REF.PROHIBITED,
  }));

  // 全角アルファベット
  const fwAlphaHits = findCharOccurrences(ctx, /[ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａ-ｚ]/g);
  items.push(makeCheck("全角アルファベットの混入", fwAlphaHits.length === 0 ? "ok" : "error", {
    current: fwAlphaHits.length === 0 ? "なし" : `${fwAlphaHits.length}件`,
    expected: "英字は半角を使用",
    hint: fwAlphaHits.length === 0 ? null : "全角アルファベットを半角に変換してください。",
    locations: fwAlphaHits.length > 0 ? fwAlphaHits : null,
    reference: REF.HALFWIDTH_EN,
  }));

  // 全角数字
  const fwDigitHits = findCharOccurrences(ctx, /[０-９]/g);
  items.push(makeCheck("全角数字の混入", fwDigitHits.length === 0 ? "ok" : "warn", {
    current: fwDigitHits.length === 0 ? "なし" : `${fwDigitHits.length}件`,
    expected: "数字は半角を使用",
    hint: fwDigitHits.length === 0 ? null : "原則として数字は半角ですが、「第２学年」「図１」など慣例的に全角が使われる箇所はそのままで問題ありません。各検出箇所を確認してください。",
    locations: fwDigitHits.length > 0 ? fwDigitHits : null,
    reference: REF.HALFWIDTH_NUM,
  }));

  // 半角不等号
  const angleHits = findCharOccurrences(ctx, /[<>]/g);
  items.push(makeCheck("半角不等号 < > の使用", angleHits.length === 0 ? "ok" : "error", {
    current: angleHits.length === 0 ? "なし" : `${angleHits.length}件`,
    expected: "〈 〉（全角山括弧）を使用",
    hint: angleHits.length === 0 ? null : "半角の「< >」を全角の「〈 〉」に置換してください。",
    locations: angleHits.length > 0 ? angleHits : null,
    reference: REF.ANGLE_BRACKET,
  }));

  // 行頭スペース
  const leadingHits = [];
  for (const p of ctx.paragraphs) {
    if (/^[　 ]/.test(p.text) && p.text.trim().length > 0) {
      leadingHits.push({ snippet: p.text.slice(0,40), location: p.location });
    }
  }
  items.push(makeCheck("行頭スペース", leadingHits.length === 0 ? "ok" : "warn", {
    current: leadingHits.length === 0 ? "なし" : `${leadingHits.length}件`,
    expected: "行の先頭にスペースを入れない",
    hint: leadingHits.length === 0 ? null : "行頭のスペースを削除してください。",
    locations: leadingHits.length > 0 ? leadingHits : null,
    reference: REF.LEADING_SP,
  }));

  return { title: "7. 禁止事項チェック", items };
}

// 段落を走査して指定パターンの出現箇所をリストアップ
function findCharOccurrences(ctx, pattern) {
  const hits = [];
  for (const p of ctx.paragraphs) {
    const text = p.text;
    const regex = new RegExp(pattern.source, "g");
    let m;
    while ((m = regex.exec(text)) !== null) {
      const idx = m.index;
      const start = Math.max(0, idx - 12);
      const end = Math.min(text.length, idx + 12);
      const before = text.slice(start, idx);
      const target = m[0];
      const after = text.slice(idx + target.length, end);
      hits.push({
        snippet: `${start > 0 ? "…" : ""}${before}【${target}】${after}${end < text.length ? "…" : ""}`,
        location: p.location,
      });
      if (hits.length > 50) return hits;
    }
  }
  return hits;
}

// === フォント・サイズ共通チェック関数 ===
function fontCheck(label, run, predicate, expectedName, location) {
  const ea = getEaFont(run.resolved);
  if (!ea) {
    return makeCheck(label, "warn", {
      current: "（未指定／スタイル継承）",
      expected: expectedName,
      hint: "フォントが明示的に設定されていません。該当箇所を選択して明示的に指定すると確実です。",
      location,
    });
  }
  const ok = predicate(ea);
  return makeCheck(label, ok ? "ok" : "error", {
    current: ea,
    expected: expectedName,
    hint: ok ? null : `該当箇所を選択し、フォントを「${expectedName}」に変更してください。`,
    location,
  });
}

function sizeCheck(label, run, expectedPt, location) {
  const sz = getSizePt(run.resolved);
  if (sz == null) {
    return makeCheck(label, "warn", {
      current: "（未指定／スタイル継承）",
      expected: `${expectedPt}pt`,
      hint: "サイズが明示的に設定されていません。該当箇所を選択して明示的に指定すると確実です。",
      location,
    });
  }
  const ok = sz === expectedPt;
  return makeCheck(label, ok ? "ok" : "error", {
    current: `${sz}pt`,
    expected: `${expectedPt}pt`,
    hint: ok ? null : `該当箇所のサイズを ${expectedPt}pt に変更してください。`,
    location,
  });
}

// グローバル公開
window.Checker = { checkDocx };
