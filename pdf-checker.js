// PDF 書式チェッカー
// PDF からは Word の「設定値」（余白・字送り・スタイル継承等）は取得できないため、
// 取得可能な情報（テキスト・フォント名・サイズ・座標）に基づくチェックに限定する

const PDF_REF = {
  PAGE_SETUP: "マニュアル P.1「1. 基本設定」",
  HEADER:     "マニュアル P.1「1. 基本設定」ヘッダーの挿入",
  SUBJECT:    "マニュアル P.1「2. 原稿の書き方」表「教科・領域名」",
  THEME:      "マニュアル P.1「2. 原稿の書き方」表「テーマ」",
  ABSTRACT:   "マニュアル P.1「2. 原稿の書き方」表「要旨」",
  KEYWORD:    "マニュアル P.1「2. 原稿の書き方」表「キーワード」",
  BODY:       "マニュアル P.2「2. 原稿の書き方」表「本文」",
  CLOSING:    "マニュアル P.2「2. 原稿の書き方」表「（注記・引用文献・生成AI利用）」",
  AI_USE:     "マニュアル P.4「5. 生成AI利用」",
  REFERENCES: "マニュアル P.3-4「4. 引用文献」",
  PROHIBITED: "マニュアル P.5「6. その他注意事項」①",
  LEADING_SP: "マニュアル P.5「6. その他注意事項」④",
  HALFWIDTH_NUM: "マニュアル P.2「2. 原稿の書き方」表「本文（数字）」",
  HALFWIDTH_EN:  "マニュアル P.2「2. 原稿の書き方」表「本文（英単語・略語）」",
  ANGLE_BRACKET: "マニュアル P.1「2. 原稿の書き方」表「教科・領域名」",
};

function pdfCheck(label, status, opts = {}) {
  return {
    label, status,
    current: opts.current ?? null,
    expected: opts.expected ?? null,
    hint: opts.hint ?? null,
    reference: opts.reference ?? null,
    location: opts.location ?? null,
    locations: opts.locations ?? null,
  };
}

// フォント名の正規化
function pdfNormalizeFont(name) {
  if (!name) return "";
  return name.replace(/^[A-Z]{6}\+/, "")  // PDFサブセットプレフィックス除去
    .replace(/[\s,-]+/g, "")
    .toLowerCase();
}
function pdfIsGothic(name) {
  const n = pdfNormalizeFont(name);
  return /bizudgothic|bizudゴシック/.test(n);
}
function pdfIsMincho(name) {
  const n = pdfNormalizeFont(name);
  return /bizudminchomedium|bizud明朝medium/.test(n);
}
function pdfIsTimes(name) {
  return /timesnewroman/i.test(pdfNormalizeFont(name));
}
function pdfIsBold(name) {
  return /bold/i.test(name || "");
}

async function checkPdf(file) {
  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const sections = [];

  // 全ページのテキストと書式情報を抽出
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 1.0 });
    const content = await page.getTextContent();
    const items = content.items.map(it => ({
      text: it.str,
      x: it.transform[4],
      y: it.transform[5],
      width: it.width,
      height: it.height,
      fontName: it.fontName,
      hasEOL: it.hasEOL,
    }));
    // フォント名は内部ID。実際の名前は styles から取得
    const styles = content.styles;
    for (const it of items) {
      const st = styles[it.fontName];
      if (st && st.fontFamily) it.fontFamily = st.fontFamily;
    }
    pages.push({ pageNum: i, viewport, items });
  }

  // 全テキスト連結
  const allText = pages.map(p => p.items.map(it => it.text).join("")).join("\n");

  sections.push(buildPdfNotice(pdf, pages));
  sections.push(buildPdfFontUsage(pages));
  sections.push(buildPdfStructure(pages, allText));
  sections.push(buildPdfReferences(pages, allText));
  sections.push(buildPdfProhibited(pages));

  return { fileName: file.name, sections, mode: "pdf" };
}

// === セクション: PDF モード説明 ===
function buildPdfNotice(pdf, pages) {
  const items = [];
  items.push(pdfCheck("PDF モードで解析中", "warn", {
    current: `${pdf.numPages}ページ / ${pages.reduce((n,p)=>n+p.items.length,0)} テキストアイテム`,
    expected: "—",
    hint:
      "PDF からは余白・字送り・行送り・スタイル継承の「設定値」を取得できないため、" +
      "テキスト内容・フォント・サイズ・文字種に基づくチェックのみ実行します。" +
      "より詳細なチェックが必要な場合は Word ファイル（.docx）をご利用ください。",
  }));
  return { title: "0. PDF モード", items };
}

// === セクション: フォント使用状況 ===
function buildPdfFontUsage(pages) {
  const items = [];
  const fontSet = new Map(); // family -> {count, examples}
  for (const p of pages) {
    for (const it of p.items) {
      if (!it.text || !it.text.trim()) continue;
      const key = it.fontFamily || it.fontName || "(unknown)";
      if (!fontSet.has(key)) fontSet.set(key, { count: 0, examples: [] });
      const e = fontSet.get(key);
      e.count++;
      if (e.examples.length < 2) e.examples.push(it.text.slice(0, 20));
    }
  }

  // 期待されるフォント: BIZ UDゴシック、BIZ UD明朝 Medium、Times New Roman
  const fonts = Array.from(fontSet.entries());
  const hasGothic = fonts.some(([f]) => pdfIsGothic(f));
  const hasMincho = fonts.some(([f]) => pdfIsMincho(f));
  const hasTimes  = fonts.some(([f]) => pdfIsTimes(f));

  items.push(pdfCheck("BIZ UDゴシックの使用",
    hasGothic ? "ok" : "warn", {
    current: hasGothic ? "使用あり" : "未検出",
    expected: "見出し・ヘッダーで使用",
    hint: hasGothic ? null : "見出しやヘッダーは BIZ UDゴシックで作成する必要があります。",
  }));
  items.push(pdfCheck("BIZ UD明朝 Medium の使用",
    hasMincho ? "ok" : "warn", {
    current: hasMincho ? "使用あり" : "未検出",
    expected: "本文・要旨・キーワード内容で使用",
    hint: hasMincho ? null : "本文は BIZ UD明朝 Medium で作成する必要があります。",
  }));
  items.push(pdfCheck("Times New Roman の使用",
    hasTimes ? "ok" : "warn", {
    current: hasTimes ? "使用あり" : "未検出",
    expected: "英文本文・英文文献で使用",
    hint: hasTimes ? null : "英文がある場合は Times New Roman を使用してください。英文が存在しない場合は問題ありません。",
  }));

  // 想定外フォントの一覧
  const unexpectedFonts = fonts.filter(([f]) =>
    !pdfIsGothic(f) && !pdfIsMincho(f) && !pdfIsTimes(f)
  );
  if (unexpectedFonts.length > 0) {
    items.push(pdfCheck("規定外のフォント検出", "warn", {
      current: unexpectedFonts.slice(0, 8).map(([f, e]) => `${f}（${e.count}件、例: 「${e.examples[0]}」）`).join(" / "),
      expected: "BIZ UDゴシック / BIZ UD明朝 Medium / Times New Roman のいずれか",
      hint: "規定外のフォントが含まれています。各箇所を確認してください。",
    }));
  } else {
    items.push(pdfCheck("規定外のフォント", "ok", { current: "なし" }));
  }

  // 太字（フォント名にBold等を含む）
  const boldFonts = fonts.filter(([f]) => pdfIsBold(f));
  if (boldFonts.length > 0) {
    items.push(pdfCheck("太字フォントの使用", "error", {
      current: boldFonts.map(([f, e]) => `${f}（${e.count}件）`).join(" / "),
      expected: "太字は使用しない",
      hint: "太字フォントが検出されました。通常フォントに変更してください。",
      reference: PDF_REF.PROHIBITED,
    }));
  } else {
    items.push(pdfCheck("太字フォントの使用", "ok", { current: "なし", reference: PDF_REF.PROHIBITED }));
  }

  return { title: "1. フォント使用状況", items, reference: PDF_REF.BODY };
}

// === セクション: 構造（教科名・テーマ・要旨・キーワード・末尾） ===
function buildPdfStructure(pages, allText) {
  const items = [];

  // 教科・領域名: 〈○○〉のパターン
  const subjectMatch = allText.match(/〈([^〉]+)〉/);
  items.push(pdfCheck("教科・領域名：〈 〉でくくる",
    subjectMatch ? "ok" : "warn", {
    current: subjectMatch ? `〈${subjectMatch[1]}〉` : "未検出",
    expected: "〈○○〉（例: 〈音楽〉）",
    hint: subjectMatch ? null : "教科・領域名が〈 〉で囲まれていません。冒頭に〈音楽〉などの形式で記載してください。",
    reference: PDF_REF.SUBJECT,
  }));

  // 半角不等号の使用
  const angleBracket = (allText.match(/[<>]/g) || []).length;
  items.push(pdfCheck("半角不等号 < > の使用",
    angleBracket === 0 ? "ok" : "error", {
    current: angleBracket === 0 ? "なし" : `${angleBracket}件`,
    expected: "〈 〉（全角山括弧）を使用",
    hint: angleBracket === 0 ? null : "半角の「< >」を全角の「〈 〉」に置換してください。",
    reference: PDF_REF.ANGLE_BRACKET,
  }));

  // キーワード
  const hasKeyword = /キーワード\s/.test(allText);
  items.push(pdfCheck("キーワード行",
    hasKeyword ? "ok" : "warn", {
    current: hasKeyword ? "あり" : "未検出",
    expected: "キーワード　　○○○ ○○○ ○○○",
    hint: hasKeyword ? null : "「キーワード」見出しが見つかりません。",
    reference: PDF_REF.KEYWORD,
  }));

  // 末尾構造
  const hasRefList = /〈引用文献〉/.test(allText);
  const hasAI = /〈生成\s*AI\s*利用〉/.test(allText);
  const hasAIFull = /〈生成\s*[ＡAＩI]+\s*利用〉/.test(allText);
  const hasNote = /〈注記〉/.test(allText);

  items.push(pdfCheck("〈引用文献〉セクション",
    hasRefList ? "ok" : "error", {
    current: hasRefList ? "あり" : "なし",
    expected: "末尾に〈引用文献〉として記載",
    hint: hasRefList ? null : "末尾に〈引用文献〉セクションを追加してください。",
    reference: PDF_REF.REFERENCES,
  }));
  items.push(pdfCheck("〈注記〉セクション（任意）",
    hasNote ? "ok" : "warn", {
    current: hasNote ? "あり" : "なし",
    expected: "注記がある場合は〈注記〉として記載",
    reference: PDF_REF.CLOSING,
  }));
  if (hasAIFull && !hasAI) {
    items.push(pdfCheck("〈生成AI利用〉表記の AI 文字", "error", {
      current: "〈生成ＡＩ利用〉（全角ＡＩ）",
      expected: "〈生成AI利用〉（半角AI）",
      hint: "「ＡＩ」を半角の「AI」に修正してください。",
      reference: PDF_REF.AI_USE,
    }));
  } else if (hasAI) {
    items.push(pdfCheck("〈生成AI利用〉セクション（任意）", "ok", { current: "あり", reference: PDF_REF.AI_USE }));
  } else {
    items.push(pdfCheck("〈生成AI利用〉セクション（任意）", "warn", {
      current: "なし", expected: "生成AIを利用した場合は〈生成AI利用〉として記載",
      reference: PDF_REF.AI_USE,
    }));
  }

  // 大見出し（1. xxx 形式の検出）
  const bigHeads = [];
  const bigHeadRe = /(?:^|\n)\s*(\d+)\.\s*([^\n]{2,40})/g;
  let m;
  while ((m = bigHeadRe.exec(allText)) !== null) {
    bigHeads.push({ num: parseInt(m[1]), text: m[2].trim() });
    if (bigHeads.length > 30) break;
  }
  if (bigHeads.length > 0) {
    items.push(pdfCheck("大見出し（連番）",
      "warn", {
      current: `${bigHeads.length}件検出: ${bigHeads.slice(0,5).map(h => `${h.num}.${h.text.slice(0,15)}`).join(" / ")}${bigHeads.length>5?"…":""}`,
      expected: "1. ○○○、2. ○○○、…",
      hint: "PDF からは正確な見出し抽出が難しいため、内容を目視確認してください。",
      reference: "マニュアル P.2「2. 原稿の書き方」表「（大見出し）」",
    }));
  }

  return { title: "2. 文書構造", items };
}

// === セクション: 引用文献 ===
function buildPdfReferences(pages, allText) {
  const items = [];
  // 〈引用文献〉以降〈生成AI利用〉まで
  const refStart = allText.search(/〈引用文献〉/);
  if (refStart < 0) {
    return { title: "3. 引用文献", items, reference: PDF_REF.REFERENCES };
  }
  let refEnd = allText.indexOf("〈", refStart + 5);
  if (refEnd < 0) refEnd = allText.length;
  const refText = allText.slice(refStart, refEnd);

  // 発行年の括弧が全角か
  const fwYears = (refText.match(/（\s*\d{4}/g) || []).length;
  const hwYears = (refText.match(/\(\s*\d{4}/g) || []).length;
  if (hwYears > 0) {
    items.push(pdfCheck("文献リスト：発行年の括弧", "error", {
      current: `半角括弧 ${hwYears}件 / 全角括弧 ${fwYears}件`,
      expected: "全角括弧 （YYYY）",
      hint: "発行年は全角括弧（）でくくってください。",
      reference: PDF_REF.REFERENCES,
    }));
  } else if (fwYears > 0) {
    items.push(pdfCheck("文献リスト：発行年の括弧", "ok", {
      current: `全角括弧 ${fwYears}件`,
      reference: PDF_REF.REFERENCES,
    }));
  }

  // URLがあるなら参照年月
  const urls = (refText.match(/https?:\/\/\S+/g) || []).length;
  if (urls > 0) {
    const refDates = (refText.match(/（\s*\d{4}\s*年\s*\d{1,2}\s*月\s*参照\s*）/g) || []).length;
    items.push(pdfCheck("Web資料：参照年月の記載",
      refDates >= urls ? "ok" : "warn", {
      current: `URL ${urls}件 / 参照年月 ${refDates}件`,
      expected: "各URLに「（YYYY年M月参照）」",
      hint: refDates >= urls ? null : "Web資料の末尾に参照年月を記載してください。",
      reference: PDF_REF.REFERENCES,
    }));
  }

  return { title: "3. 引用文献", items, reference: PDF_REF.REFERENCES };
}

// === セクション: 禁止事項 ===
function buildPdfProhibited(pages) {
  const items = [];
  const allText = pages.map(p => p.items.map(it => it.text).join("")).join("\n");

  // 特殊文字
  const specialMatches = allText.match(/[℡㌕㌻㌃㌍㌔㌖㌘㌢㌣㌤㌥㌦㌧㌨㌩㌪㌫㌬㌭㌮㌯㌰㌱㌲㌳㌴㌵㌶㌷㌸㌹㌺㌻㌼㌽㌾㌿]/g);
  items.push(pdfCheck("特殊文字（℡ ㌕ ㌻ 等）",
    !specialMatches ? "ok" : "error", {
    current: specialMatches ? `${specialMatches.length}件：${Array.from(new Set(specialMatches)).join(" ")}` : "なし",
    expected: "使用しない",
    hint: specialMatches ? "通常の表記（例: ℡ → TEL）に書き換えてください。" : null,
    reference: PDF_REF.PROHIBITED,
  }));

  // 全角アルファベット
  const fwAlpha = (allText.match(/[ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａ-ｚ]/g) || []);
  items.push(pdfCheck("全角アルファベットの混入",
    fwAlpha.length === 0 ? "ok" : "error", {
    current: fwAlpha.length === 0 ? "なし" : `${fwAlpha.length}文字`,
    expected: "英字は半角を使用",
    hint: fwAlpha.length === 0 ? null : "全角アルファベットを半角に変換してください。",
    reference: PDF_REF.HALFWIDTH_EN,
  }));

  // 全角数字
  const fwDigit = (allText.match(/[０-９]/g) || []);
  items.push(pdfCheck("全角数字の混入",
    fwDigit.length === 0 ? "ok" : "warn", {
    current: fwDigit.length === 0 ? "なし" : `${fwDigit.length}文字：${Array.from(new Set(fwDigit)).join(" ")}`,
    expected: "数字は半角を使用",
    hint: fwDigit.length === 0 ? null : "原則として数字は半角ですが、「第２学年」「図１」など慣例的に全角の箇所はそのままで問題ありません。",
    reference: PDF_REF.HALFWIDTH_NUM,
  }));

  return { title: "4. 禁止事項チェック", items };
}

window.PdfChecker = { checkPdf };
