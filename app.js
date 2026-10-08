"use strict";
// 社内パトロール点検簿をつくる：様式・現場・点検の情報・区分を選ぶ → 入力済みの点検簿（xlsx）をダウンロード。
// 原本と データ\ は OneDrive「社内共有フォルダ\05_現場パトロール」から読むだけ（書き込まない）。
const CFG = window.PATROL_CONFIG;
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Files.ReadWrite.All"];   // 全体工程表で管理者の同意を済ませた権限（このページは読むだけ）
const EMBED = window.self !== window.top;          // ポータル（Google サイト）に埋め込まれているか
const $ = id => document.getElementById(id);
const S = { form: "shita", site: null, cfg: null, sites: [], files: {}, list: [], hidden: {} };
const HIDDEN_FILE = "データ/現場の非表示.json";   // 終わった現場（一覧に出さない）。保存はフォルダを編集できる人だけ

/* ---------- サインインと Graph ---------- */
let msalApp = null;
async function token() {
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      if (EMBED) return (await msalApp.acquireTokenPopup({ scopes: SCOPES })).accessToken;
      await msalApp.acquireTokenRedirect({ scopes: SCOPES });
    }
    throw e;
  }
}
async function g(path) {
  const res = await fetch(GRAPH + path, { headers: { Authorization: "Bearer " + await token() } });
  if (!res.ok) { const e = new Error(`${path} … ${res.status}`); e.status = res.status; throw e; }
  return res.json();
}
let base = "";
async function findFolder() {
  if (CFG.folder.itemId) { base = `/drives/${CFG.folder.driveId}/items/${CFG.folder.itemId}`; return; }
  // ID がまだ無いとき：フォルダの持ち主の OneDrive から path で探して、ID を知らせる
  const enc = CFG.folder.path.split("/").map(encodeURIComponent).join("/");
  const it = await g(`/me/drive/root:/${enc}`);
  base = `/drives/${it.parentReference.driveId}/items/${it.id}`;
  $("idNote").hidden = false;
  $("idNote").textContent = `（管理用）config.js の folder に入れる値：driveId "${it.parentReference.driveId}"、itemId "${it.id}"`;
}
async function upload(name, text) {
  const enc = name.split("/").map(encodeURIComponent).join("/");
  const res = await fetch(`${GRAPH}${base}:/${enc}:/content`, {
    method: "PUT", body: text,
    headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json" } });
  if (!res.ok) { const e = new Error(`${name} を保存できません（${res.status}）`); e.status = res.status; throw e; }
  return res.json();
}
async function download(name, as) {
  const enc = name.split("/").map(encodeURIComponent).join("/");
  const meta = await g(`${base}:/${enc}`);
  const res = await fetch(meta["@microsoft.graph.downloadUrl"]);
  if (!res.ok) throw new Error(`${name} を読めません（${res.status}）`);
  return as === "json" ? res.json() : res.arrayBuffer();
}

/* ---------- 画面 ---------- */
const FORM_LABEL = { moto: "元請工事用", shita: "下請工事用" };
const FIELDS = [   // 画面の欄 → 基本マスタの鍵（点検表のセルの名前と同じ）
  ["koujimei", "工事名（正式）", "formal"], ["shanai", "社内工事名", "name"],
  ["start", "工期（自）", "start", "date"], ["end", "工期（至）", "end", "date"],
  ["staff", "工事担当責任者", "staff"], ["amount", "請負金（千円）", "amount", "money"],
  ["client", "元請名", "client"], ["client_person", "元請担当者", "client_person"],
  ["place", "工事場所", "place"],
];
const clientLabel = () => S.form === "moto" ? ["発注者", "発注者担当者"] : ["元請名", "元請担当者"];

function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v; else if (k.startsWith("on")) e[k] = v; else e.setAttribute(k, v);
  }
  kids.flat().forEach(k => e.append(k));
  return e;
}

function renderForm() {
  document.querySelectorAll("#formSel button").forEach(b => b.classList.toggle("on", b.dataset.form === S.form));
  const [cl, cp] = clientLabel();
  $("lab_client").textContent = cl; $("lab_client_person").textContent = cp;
  renderSlots();
}

function renderSiteList() {
  const q = $("siteQ").value.trim().normalize("NFKC").toLowerCase();
  const list = $("siteList");
  list.innerHTML = "";
  const hits = S.sites.filter(r => !S.hidden[r.code])
    .filter(r => !q || `${r.code} ${r.name} ${r.client} ${r.staff} ${r.place}`.normalize("NFKC").toLowerCase().includes(q));
  let lastP = null, og = null;
  for (const r of hits) {
    if (r.period !== lastP) { og = el("optgroup", { label: `${r.period}期` }); list.append(og); lastP = r.period; }
    og.append(el("option", { value: r.code }, `${r.code}　${r.name}　（${r.client || "―"}／${r.staff || "―"}）`));
  }
  if (S.site) list.value = S.site.code;
  $("siteCount").textContent = `${hits.length} 件`;
}

function pickSite(code) {
  S.site = S.sites.find(r => r.code === code) || null;
  for (const [key, , src, kind] of FIELDS) {
    let v = S.site ? S.site[src] : "";
    if (kind === "money") v = v ? Math.round(v / 1000).toLocaleString("ja-JP") : "";
    $("f_" + key).value = v ?? "";
  }
  $("siteSummary").textContent = S.site ? (S.site.summary || "") : "";
  $("siteDetail").hidden = !S.site;
  renderSlots();
}

/* ---------- 一覧を整理：終わった現場にチェック → 保存すると一覧に出さない ---------- */
function renderSiteAdmin() {
  const box = $("siteAdminList");
  box.innerHTML = "";
  for (const r of S.sites) {
    const cb = el("input", { type: "checkbox", "data-code": r.code });
    cb.checked = !!S.hidden[r.code];
    box.append(el("label", { class: "adminRow" }, cb, ` ${r.code}　${r.name}　（${r.client || "―"}／${r.staff || "―"}）`));
  }
}
async function saveSiteAdmin() {
  const hidden = {};
  document.querySelectorAll("#siteAdminList input:checked").forEach(cb => {
    const c = cb.dataset.code;
    hidden[c] = S.hidden[c] || { name: (S.sites.find(r => r.code === c) || {}).name || "", when: new Date().toISOString().slice(0, 10) };
  });
  $("siteAdminMsg").textContent = "保存しています…";
  try {
    await upload(HIDDEN_FILE, JSON.stringify({ note: "社内パトロール点検簿：一覧に出さない（終わった）現場", hidden }, null, 1));
    S.hidden = hidden;
    $("siteAdminMsg").textContent = `保存しました（隠す現場 ${Object.keys(hidden).length} 件）。`;
    renderSiteList();
  } catch (e) {
    $("siteAdminMsg").textContent = (e.status === 403 || e.status === 401)
      ? "保存できません。一覧の整理は、フォルダを編集できる人（管理者）だけができます。"
      : "保存できませんでした：" + e.message;
  }
}

/* ---------- 区分：選んだ区分の一覧（同じ区分は 1 回だけ）を、1枚目の左→右（→2枚目の左→右）に詰める ---------- */
const MARU = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮";
const uniq = a => [...new Set(a)];
const choices = () => S.cfg.forms[S.form].choices;
const findChoice = name => choices().find(c => c.name === name);
const sizeOf = name => (findChoice(name) || { n: 0 }).n;
// 区分が使う行の数：項目の数と、区分の名前（縦書き）が入る行の数の大きいほう。足りない分は下に空欄の行
const CHAR_MM = 3.3;    // 縦書き 8pt の 1 文字の高さ（行間こみ）
const rowsFor = (name, h) => Math.max(sizeOf(name), Math.ceil(name.length * CHAR_MM / h));

function fillSelect(sel, opts, value, blank) {
  sel.innerHTML = "";
  if (blank) sel.append(el("option", { value: "" }, blank));
  opts.forEach(([v, t, dis]) => sel.append(el("option", dis ? { value: v, disabled: "" } : { value: v }, t)));
  sel.value = value;
  if (sel.value !== value) sel.value = "";
}

// 入れる場所：[{page, side, rows(行の数)}]
function columns() {
  const f = S.cfg.forms[S.form];
  const cols = [{ page: 1, side: "L", lay: f.layout.L }, { page: 1, side: "R", lay: f.layout.R }];
  if ($("allow2").checked) cols.push({ page: 2, side: "L", lay: f.layout2.L }, { page: 2, side: "R", lay: f.layout2.R });
  return cols.map(c => ({ ...c, cap: c.lay.rows.length }));
}
// 順に詰める：入らなければ次の列へ（戻らない）。どこにも入らなければ over
function place(list) {
  const cols = columns().map(c => ({ ...c, used: 0, groups: [] }));
  let ci = 0;
  const where = {};
  for (const name of list) {
    while (ci < cols.length && cols[ci].used + rowsFor(name, cols[ci].lay.h) > cols[ci].cap) ci++;
    if (ci >= cols.length) { where[name] = null; continue; }
    const n = rowsFor(name, cols[ci].lay.h);
    cols[ci].groups.push({ name, start: cols[ci].used, n });
    cols[ci].used += n;
    where[name] = cols[ci];
  }
  return { cols, where };
}
const colLabel = c => c ? `${c.page}枚目 ${c.side === "L" ? "左" : "右"}` : "入りきらない";

function renderChosen() {
  if (!S.cfg) return;
  const { cols, where } = place(S.list);
  const box = $("chosen");
  box.innerHTML = "";
  S.list.forEach((name, i) => {
    const c = findChoice(name), w = where[name];
    const up = el("button", { class: "mini", title: "上へ", onclick: () => { if (i > 0) { [S.list[i - 1], S.list[i]] = [S.list[i], S.list[i - 1]]; renderChosen(); } } }, "↑");
    const dn = el("button", { class: "mini", title: "下へ", onclick: () => { if (i < S.list.length - 1) { [S.list[i + 1], S.list[i]] = [S.list[i], S.list[i + 1]]; renderChosen(); } } }, "↓");
    const rm = el("button", { class: "mini", title: "外す", onclick: () => { S.list.splice(i, 1); renderChosen(); } }, "×");
    const items = el("div", { class: "items", hidden: "" });
    (c ? c.items : []).forEach((t, k) => items.append(el("div", {}, `${MARU[k]} ${t}`)));
    const nm = el("button", { class: "link name", onclick: () => { items.hidden = !items.hidden; } },
      `${name}（${c ? c.n : "?"}項目）`);
    box.append(el("div", { class: "chosenRow" + (w ? "" : " over") },
      el("span", { class: "place" }, colLabel(w)), nm, el("span", { class: "btns" }, up, dn, rm), items));
  });
  // 行数のようす
  const meter = cols.map(c => `${c.page}枚目${c.side === "L" ? "左" : "右"} ${c.used}/${c.cap}行`).join("　");
  const over = S.list.filter(n => !where[n]);
  $("meter").textContent = meter;
  $("overMsg").textContent = over.length
    ? `入りきらない区分があります：${over.join("、")}。外すか、${$("allow2").checked ? "ほかを外してください" : "「2枚目を使ってよい」にしてください"}。`
    : "";
  $("bMake").disabled = over.length > 0;
  renderPicker();
  renderSuggest();
}

// 追加：大項目 → 中項目 → 小項目（選び済みは選べない）
const COMMON = "__common__";
const pick = { big: "", mid: null };
function renderPicker() {
  const sBig = $("pBig"), sMid = $("pMid"), sSmall = $("pSmall");
  fillSelect(sBig, uniq(choices().map(c => c.path[0])).map(b => [b, b]), pick.big, "大項目");
  const mids = uniq(choices().filter(c => c.path[0] === pick.big).map(c => c.path[1]));
  sMid.hidden = !pick.big || (mids.length === 1 && mids[0] === "");
  if (sMid.hidden) pick.mid = mids[0] ?? null;
  fillSelect(sMid, mids.map(m => [m || COMMON, m || "（共通）"]), pick.mid === null ? "" : (pick.mid || COMMON), "中項目");
  const smalls = choices().filter(c => c.path[0] === pick.big && c.path[1] === pick.mid);
  sSmall.hidden = !pick.big || pick.mid === null;
  fillSelect(sSmall, smalls.map(c => [c.name, `${c.name}（${c.n}項目）${S.list.includes(c.name) ? " ✓選び済み" : ""}`,
    S.list.includes(c.name)]), "", "小項目を選ぶと追加");
}
function addKubun(name) {
  if (!name || S.list.includes(name)) return;
  S.list.push(name);
  renderChosen();
}

/* ---------- おすすめ：作業内容（重く）と工事概要（軽く）に出てくる目印の言葉で点数をつける ---------- */
const nz = s => (s || "").normalize("NFKC").toLowerCase();
function suggestions() {
  const work = nz($("f_sagyou").value), summ = nz(S.site ? `${S.site.summary} ${S.site.name}` : "");
  return choices().map(c => {
    let score = 0;
    const hit = [];
    for (const k of c.kw || []) {
      const kk = nz(k);
      if (work.includes(kk)) { score += 3; hit.push(k); }
      else if (summ.includes(kk)) { score += 1; hit.push(k); }
    }
    return { c, score, hit: uniq(hit), chosen: S.list.includes(c.name) };
  }).filter(x => x.score > 0).sort((a, b) => b.score - a.score || a.c.n - b.c.n);
}
function renderSuggest() {
  const box = $("suggest");
  if (!box || !S.cfg) return;
  box.innerHTML = "";
  const list = suggestions().slice(0, 14);
  if (!list.length) {
    box.append(el("span", { class: "muted small" }, "作業内容を書くと、ここに必要そうな区分が出ます。"));
    return;
  }
  for (const x of list) {
    box.append(el("button", { class: "chip" + (x.chosen ? " done" : ""), title: "目印：" + x.hit.join("・"),
      onclick: () => addKubun(x.c.name) },
      `${x.chosen ? "✓ " : "＋ "}${x.c.name}`, el("span", { class: "stars" }, "★".repeat(Math.min(3, Math.ceil(x.score / 3))))));
  }
}

function resetKubun() {
  S.list = S.cfg.forms[S.form].defaults.filter(n => findChoice(n) && sizeOf(n) > 0);
  renderChosen();
}

function renderSlots() {   // 様式・現場が変わったとき
  if (!S.cfg) return;
  if (S.listForm !== S.form) { S.listForm = S.form; resetKubun(); return; }
  $("summaryHint").textContent = S.site && S.site.summary ? `工事概要（工事工程表）：${S.site.summary}` : "";
  renderChosen();
}

/* ---------- 点検簿をつくる ---------- */
const serial = iso => {           // "2026-02-01" → Excel の日付の数
  const [y, m, d] = iso.split("-").map(Number);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
};

// 1 つの列（lay）に、groups を上から書く。空いた行は空欄。区分欄の結合も付け直す
function writeColumn(xml, lay, groups) {
  const [c1, c2] = lay.cat;
  // 区分欄の書式（原本の最初の区分の欄。太線の付いた書式なら元の書式に戻して使う）
  const catStyle = (() => { const s0 = XlsxEdit.cellStyle(xml, lay.cat_style_cell), tops = S.cfg.forms[S.form].top_style;
    for (const [b0, v] of Object.entries(tops)) if (Object.values(v).includes(s0)) return b0; return s0; })();
  const rows = lay.rows, ends = lay.ends;
  const inCol = new Set(rows.map(String));
  // この列の区分欄の結合を外す
  const re = new RegExp(`^${c1}(\\d+):${c2}(\\d+)$`);
  const merges = [];
  let i = 0;
  for (const g of groups) {
    const c = findChoice(g.name);
    // 空いた行（前の区分との間）は無い：詰めて入れる
    for (let k = 0; k < g.n; k++, i++) {
      const has = k < c.items.length;    // 項目の行（番号つき）／名前を入れるための空欄の行
      xml = XlsxEdit.setCell(xml, `${lay.no}${rows[i]}`, has ? MARU[k] : null);
      xml = XlsxEdit.setCell(xml, `${lay.item}${rows[i]}`, has ? c.items[k] : null);
      xml = XlsxEdit.setCell(xml, `${c1}${rows[i]}`, k === 0 ? g.name : null, catStyle);
    }
    merges.push(`${c1}${rows[i - g.n]}:${c2}${ends[i - 1]}`);
  }
  for (; i < rows.length; i++) {      // 残りは空欄（手書きで足せる）
    xml = XlsxEdit.setCell(xml, `${lay.no}${rows[i]}`, null);
    xml = XlsxEdit.setCell(xml, `${lay.item}${rows[i]}`, null);
    xml = XlsxEdit.setCell(xml, `${c1}${rows[i]}`, null, catStyle);
    merges.push(`${c1}${rows[i]}:${c2}${ends[i]}`);
  }
  // 区分の境目の太線：区分の始まりの行の上端を「上が太線」、その前の行の下端を「下が太線」に（両方なら上下）。
  // 境目でない行は元の書式に戻す。区分欄の書式も、太線の付かない元の書式を使う
  const top = S.cfg.forms[S.form].top_style, back = {};
  for (const [b0, v] of Object.entries(top)) for (const id of Object.values(v)) back[id] = b0;
  const starts = new Set();
  let acc = 0;
  for (const g of groups) { starts.add(acc); acc += g.n; }
  if (acc < rows.length) starts.add(acc);
  const flags = {};            // 行の番号（シート）→ "t" / "b" / "tb"
  const addFlag = (r, f) => { flags[r] = flags[r] && flags[r] !== f ? "tb" : f; };
  rows.forEach((r, k) => {
    if (starts.has(k)) { addFlag(r, "t"); if (k > 0) addFlag(ends[k - 1], "b"); }
  });
  const edge = new Set([...rows, ...ends]);
  for (const r of edge) {
    for (let c = lay.span[0]; c <= lay.span[1]; c++) {
      const ref = colName(c) + r, cur = XlsxEdit.cellStyle(xml, ref);
      if (cur === undefined) continue;
      const base = back[cur] || cur, f = flags[r];
      const want = f ? ((top[base] || {})[f] || base) : base;
      if (want !== cur) xml = XlsxEdit.setStyle(xml, ref, want);
    }
  }
  return XlsxEdit.remerge(xml, r => { const m = r.match(re); return m && inCol.has(m[1]); }, merges);
}
const colName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

async function build() {
  const btn = $("bMake");
  btn.disabled = true; $("makeMsg").textContent = "点検簿をつくっています…";
  try {
    const f = S.cfg.forms[S.form];
    const buf = S.files[f.file] || (S.files[f.file] = await download(f.file));
    const zip = XlsxEdit.readZip(buf.slice(0));
    const entry = n => zip.find(e => e.name === n);
    const t1 = entry(f.sheets.tenken), t2 = entry(f.sheets.page2), wbx = entry("xl/workbook.xml");
    let x1 = await XlsxEdit.text(t1), x2 = await XlsxEdit.text(t2), xw = await XlsxEdit.text(wbx);
    const C = f.cells;
    const put = (key, v) => { x1 = XlsxEdit.setCell(x1, C[key], v); };

    // 点検の情報
    const day = $("f_date").value;
    if (day) put("date", serial(day));
    const t0 = $("f_time1").value, t9 = $("f_time2").value;
    if (t0 || t9) put("jikoku", `${t0 || "　："}　～　${t9 || "　："}`);
    if ($("f_tenki").value) put("tenki", $("f_tenki").value);
    if ($("f_tenkensha").value.trim()) put("tenkensha", $("f_tenkensha").value.trim());
    if ($("f_tachiai").value.trim()) put("tachiai", $("f_tachiai").value.trim());
    if ($("f_sagyou").value.trim()) put("sagyou", $("f_sagyou").value.trim().replace(/\s*\n\s*/g, "、"));

    // 現場：選択欄にも入れ、各欄には画面の値をそのまま書く（原本の現場マスタが古くても困らない）
    if (S.site) {
      put("genba", `${S.site.code} ${S.site.name}`);
      put("code", S.site.code);
      for (const [key, , , kind] of FIELDS) {
        const v = $("f_" + key).value.trim();
        if (kind === "date") put(key, v ? serial(v) : S.cfg.date_blank);
        else if (kind === "money") put(key, v === "" ? null : Number(v.normalize("NFKC").replace(/[,\s]/g, "")));
        else put(key, v || null);
      }
    }

    // 区分：詰めて書く。2枚目を使わないときは 2枚目のシートを隠す
    const { cols, where } = place(S.list);
    if (S.list.some(n => !where[n])) throw new Error("入りきらない区分があります");
    const byKey = Object.fromEntries(cols.map(c => [`${c.page}${c.side}`, c.groups]));
    x1 = writeColumn(x1, f.layout.L, byKey["1L"] || []);
    x1 = writeColumn(x1, f.layout.R, byKey["1R"] || []);
    const use2 = (byKey["2L"] || []).length + (byKey["2R"] || []).length > 0;
    if (use2) {
      x2 = writeColumn(x2, f.layout2.L, byKey["2L"] || []);
      x2 = writeColumn(x2, f.layout2.R, byKey["2R"] || []);
    }
    put("shiteki_head", use2 ? "指摘・指導事項　／　是正確認　　（2枚目の分もここに記入）" : "指摘・指導事項　／　是正確認");
    xw = xw.replace(new RegExp(`(<sheet name="${f.page2_name}"[^>]*?state=")[a-zA-Z]+(")`), `$1${use2 ? "visible" : "hidden"}$2`);

    XlsxEdit.setText(t1, x1); XlsxEdit.setText(t2, x2); XlsxEdit.setText(wbx, xw);
    const blob = XlsxEdit.writeZip(zip);
    const ymd = (day || new Date().toISOString().slice(0, 10)).replace(/-/g, "");
    const who = S.site ? `${S.site.code}${S.site.name}` : "現場未定";
    const name = `パトロール点検簿_${FORM_LABEL[S.form].replace("工事用", "")}_${who}_${ymd}.xlsx`.replace(/[\\/:*?"<>|]/g, "_");
    const a = el("a", { href: URL.createObjectURL(blob), download: name });
    document.body.append(a); a.click(); a.remove();
    $("makeMsg").innerHTML = "";
    $("makeMsg").append(el("b", {}, `「${name}」をダウンロードしました${use2 ? "（2枚）" : "（1枚）"}。`), el("br"),
      "開いたら、上に出る黄色い帯の「編集を有効にする」を押してから印刷してください。");
  } catch (e) {
    $("makeMsg").textContent = "つくれませんでした：" + e.message;
  } finally {
    btn.disabled = false;
    renderChosen();
  }
}

/* ---------- はじめ ---------- */
// サインインした人の名前を、点検者の欄に最初から入れる（書き換えてよい）
function signedIn(acc) {
  const name = acc.name || acc.username;
  $("who").textContent = name;
  if (!$("f_tenkensha").value) $("f_tenkensha").value = name;
}

async function load() {
  $("loading").hidden = false;
  try {
    await findFolder();
    const [cfg, master, hid] = await Promise.all([download("データ/点検簿の設定.json", "json"),
                                                  download("データ/現場マスタ.json", "json"),
                                                  download(HIDDEN_FILE, "json").catch(() => ({ hidden: {} }))]);
    S.hidden = hid.hidden || {};
    S.cfg = cfg;
    S.sites = master.sites.slice().sort((a, b) => (b.period - a.period) || a.code.localeCompare(b.code));
    $("masterInfo").textContent = `現場 ${S.sites.length - Object.keys(S.hidden).filter(c => S.sites.some(r => r.code === c)).length} 件（${(master.meta.made || "").slice(0, 10)} 作成）`;
    const ts = $("f_tenki");
    ts.innerHTML = ""; ts.append(el("option", { value: "" }, "（手書き：晴・曇・雨・雪）"));
    cfg.tenki.forEach(t => ts.append(el("option", { value: t }, t)));
    $("main").hidden = false;
    renderSiteList(); renderForm();
  } catch (e) {
    $("err").hidden = false;
    $("err").textContent = (e.status === 404 || e.status === 403)
      ? "点検簿のフォルダが見えません。管理者に「05_現場パトロール」フォルダの共有を頼んでください。（" + e.message + "）"
      : "読み込めませんでした：" + e.message;
  } finally {
    $("loading").hidden = true;
  }
}

async function start() {
  // 点検日は空欄から選んでもらう（空のままなら手書き用の「令和　年　月　日」）
  document.querySelectorAll("#formSel button").forEach(b => b.onclick = () => { S.form = b.dataset.form; renderForm(); });
  $("siteQ").oninput = renderSiteList;
  $("siteList").onchange = () => pickSite($("siteList").value);
  $("bNoSite").onclick = () => { $("siteList").value = ""; pickSite(null); };
  $("bMake").onclick = build;
  $("f_sagyou").oninput = renderSuggest;
  $("allow2").onchange = renderChosen;
  $("bSiteAdmin").onclick = () => { const p = $("siteAdmin"); p.hidden = !p.hidden; if (!p.hidden) renderSiteAdmin(); };
  $("bSiteAdminSave").onclick = saveSiteAdmin;
  $("bReset").onclick = resetKubun;
  $("pBig").onchange = () => { pick.big = $("pBig").value; pick.mid = null; renderPicker(); };
  $("pMid").onchange = () => { const v = $("pMid").value; pick.mid = v === COMMON ? "" : (v || null); renderPicker(); };
  $("pSmall").onchange = () => addKubun($("pSmall").value);
  // 請負金（千円）はカンマ区切りで見せる（全角の数字・カンマも受ける）
  $("f_amount").onblur = () => {
    const n = Number($("f_amount").value.normalize("NFKC").replace(/[,\s]/g, ""));
    if ($("f_amount").value.trim() && !isNaN(n)) $("f_amount").value = n.toLocaleString("ja-JP");
  };

  msalApp = new msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId,
            redirectUri: new URL("./", location.href).href },
    cache: { cacheLocation: "localStorage" }
  });
  await msalApp.initialize();
  try { const r = await msalApp.handleRedirectPromise(); if (r) msalApp.setActiveAccount(r.account); }
  catch (e) { $("signinMsg").textContent = "サインインできませんでした：" + e.message; }
  const acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!acc) {
    $("signin").hidden = false;
    $("bSignIn").onclick = async () => {
      try {
        if (EMBED) {
          const r = await msalApp.loginPopup({ scopes: SCOPES });
          msalApp.setActiveAccount(r.account);
          $("signin").hidden = true; signedIn(r.account);
          load();
        } else await msalApp.loginRedirect({ scopes: SCOPES });
      } catch (e) { $("signinMsg").textContent = "サインインできませんでした：" + e.message; }
    };
    return;
  }
  msalApp.setActiveAccount(acc);
  signedIn(acc);
  load();
}
start();
