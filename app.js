"use strict";
// 社内パトロール点検簿をつくる：様式・現場・点検の情報・区分を選ぶ → 入力済みの点検簿（xlsx）をダウンロード。
// 原本と データ\ は OneDrive「社内共有フォルダ\05_現場パトロール」から読むだけ（書き込まない）。
const CFG = window.PATROL_CONFIG;
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Files.ReadWrite.All"];   // 全体工程表で管理者の同意を済ませた権限（このページは読むだけ）
const EMBED = window.self !== window.top;          // ポータル（Google サイト）に埋め込まれているか
const $ = id => document.getElementById(id);
const S = { form: "shita", site: null, cfg: null, sites: [], files: {} };

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
  const hits = S.sites.filter(r => !q || `${r.code} ${r.name} ${r.client} ${r.staff} ${r.place}`.normalize("NFKC").toLowerCase().includes(q));
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

/* ---------- 区分：大項目 → 中項目 → 小項目 ---------- */
const MARU = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫";
const uniq = a => [...new Set(a)];
const choices = () => S.cfg.forms[S.form].choices;
const findChoice = name => choices().find(c => c.name === name);

function fillSelect(sel, opts, value, blank) {
  sel.innerHTML = "";
  if (blank) sel.append(el("option", { value: "" }, blank));
  opts.forEach(([v, t, cls]) => sel.append(el("option", cls ? { value: v, class: cls } : { value: v }, t)));
  sel.value = value;
  if (sel.value !== value) sel.value = "";
}

function setChoice(page, idx, name) {
  (S.choice[page] = S.choice[page] || [])[idx] = name;
}

function slotRow(slot, idx, page) {
  const cur = (S.choice[page] || [])[idx];
  const name = cur !== undefined ? cur : slot.default;
  setChoice(page, idx, name);
  const c0 = findChoice(name);
  const sBig = el("select"), sMid = el("select"), sSmall = el("select");
  const items = el("div", { class: "items" });
  // 中項目：null＝まだ選んでいない、""＝（共通）。select の値では "（共通）" を COMMON で表す
  const COMMON = "__common__";
  let big = c0 ? c0.path[0] : "", mid = c0 ? c0.path[1] : null;

  const draw = () => {
    fillSelect(sBig, uniq(choices().map(c => c.path[0])).map(b => [b, b]), big, "大項目（使わない＝手書き）");
    const mids = uniq(choices().filter(c => c.path[0] === big).map(c => c.path[1]));
    sMid.hidden = !big || (mids.length === 1 && mids[0] === "");
    fillSelect(sMid, mids.map(m => [m || COMMON, m || "（共通）"]), mid === null ? "" : (mid || COMMON), "中項目");
    const smalls = choices().filter(c => c.path[0] === big && (sMid.hidden || c.path[1] === mid));
    sSmall.hidden = !big;
    fillSelect(sSmall, smalls.map(c => [c.name, `${c.name}（${c.n}項目）${c.n > slot.rows ? " ※入りきらない" : ""}`,
      c.n > slot.rows ? "over" : ""]), (S.choice[page] || [])[idx] || "", "小項目");
    const c = findChoice(sSmall.value);
    items.innerHTML = "";
    if (!c) { items.append(el("span", { class: "muted" }, big ? "小項目を選んでください" : "空欄（手書き用）")); return; }
    c.items.slice(0, slot.rows).forEach((t, i) => items.append(el("div", {}, `${MARU[i]} ${t}`)));
    if (c.n > slot.rows) items.append(el("div", { class: "warn" }, `※ この枠は ${slot.rows} 行。残り ${c.n - slot.rows} 項目は入りません`));
  };
  sBig.onchange = () => { big = sBig.value; mid = null; setChoice(page, idx, "");
    const mids = uniq(choices().filter(c => c.path[0] === big).map(c => c.path[1]));
    if (mids.length === 1) mid = mids[0];
    const smalls = choices().filter(c => c.path[0] === big && c.path[1] === mid);
    if (smalls.length === 1) setChoice(page, idx, smalls[0].name);
    draw(); renderSuggest(); };
  sMid.onchange = () => { mid = sMid.value === COMMON ? "" : (sMid.value || null); setChoice(page, idx, "");
    const smalls = choices().filter(c => c.path[0] === big && c.path[1] === mid);
    if (smalls.length === 1) setChoice(page, idx, smalls[0].name);
    draw(); renderSuggest(); };
  sSmall.onchange = () => { setChoice(page, idx, sSmall.value); draw(); renderSuggest(); };
  draw();
  return el("div", { class: "slot" },
    el("div", { class: "slotHead" }, el("span", { class: "rows" }, `${slot.rows}行`), sBig, sMid, sSmall), items);
}

/* ---------- おすすめ：作業内容（重く）と工事概要（軽く）に出てくる目印の言葉で点数をつける ---------- */
const nz = s => (s || "").normalize("NFKC").toLowerCase();
function suggestions() {
  const work = nz($("f_sagyou").value), summ = nz(S.site ? `${S.site.summary} ${S.site.name}` : "");
  const chosen = new Set([...(S.choice.tenken || []), ...(S.choice.koushu || [])].filter(Boolean));
  return choices().map(c => {
    let score = 0;
    const hit = [];
    for (const k of c.kw || []) {
      const kk = nz(k);
      if (work.includes(kk)) { score += 3; hit.push(k); }
      else if (summ.includes(kk)) { score += 1; hit.push(k); }
    }
    return { c, score, hit: uniq(hit), chosen: chosen.has(c.name) };
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
    const b = el("button", { class: "chip" + (x.chosen ? " done" : ""), title: "目印：" + x.hit.join("・") },
      `${x.chosen ? "✓ " : ""}${x.c.name}`, el("span", { class: "stars" }, "★".repeat(Math.min(3, Math.ceil(x.score / 3)))));
    b.onclick = () => {
      if (x.chosen) return;
      const slots = S.cfg.forms[S.form].koushu_slots;
      const i = slots.findIndex((s, k) => !(S.choice.koushu || [])[k]);
      if (i < 0) { alert("2枚目の枠が全部うまっています。どれかを「使わない」にしてから押してください。"); return; }
      setChoice("koushu", i, x.c.name);
      renderSlots();
    };
    box.append(b);
  }
}

function renderSlots() {
  if (!S.cfg) return;
  const f = S.cfg.forms[S.form];
  if (S.choiceForm !== S.form) { S.choice = {}; S.choiceForm = S.form; }
  const p1 = $("slots1"), p2 = $("slots2");
  p1.innerHTML = ""; p2.innerHTML = "";
  f.slots.forEach((s, i) => p1.append(slotRow(s, i, "tenken")));
  f.koushu_slots.forEach((s, i) => p2.append(slotRow(s, i, "koushu")));
  $("summaryHint").textContent = S.site && S.site.summary ? `工事概要（工事工程表）：${S.site.summary}` : "";
  renderSuggest();
}

/* ---------- 点検簿をつくる ---------- */
const serial = iso => {           // "2026-02-01" → Excel の日付の数
  const [y, m, d] = iso.split("-").map(Number);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
};

async function build() {
  const btn = $("bMake");
  btn.disabled = true; $("makeMsg").textContent = "点検簿をつくっています…";
  try {
    const f = S.cfg.forms[S.form];
    const buf = S.files[f.file] || (S.files[f.file] = await download(f.file));
    const zip = XlsxEdit.readZip(buf.slice(0));
    const entry = n => zip.find(e => e.name === n);
    const t1 = entry(f.sheets.tenken), t2 = entry(f.sheets.koushu);
    let x1 = await XlsxEdit.text(t1), x2 = await XlsxEdit.text(t2);
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
    // 区分
    f.slots.forEach((s, i) => { x1 = XlsxEdit.setCell(x1, s.cell, (S.choice.tenken || [])[i] || null); });
    f.koushu_slots.forEach((s, i) => { x2 = XlsxEdit.setCell(x2, s.cell, (S.choice.koushu || [])[i] || null); });

    XlsxEdit.setText(t1, x1); XlsxEdit.setText(t2, x2);
    const blob = XlsxEdit.writeZip(zip);
    const ymd = (day || new Date().toISOString().slice(0, 10)).replace(/-/g, "");
    const who = S.site ? `${S.site.code}${S.site.name}` : "現場未定";
    const name = `パトロール点検簿_${FORM_LABEL[S.form].replace("工事用", "")}_${who}_${ymd}.xlsx`.replace(/[\\/:*?"<>|]/g, "_");
    const a = el("a", { href: URL.createObjectURL(blob), download: name });
    document.body.append(a); a.click(); a.remove();
    $("makeMsg").innerHTML = "";
    $("makeMsg").append(el("b", {}, `「${name}」をダウンロードしました。`), el("br"),
      "開いたら、上に出る黄色い帯の「編集を有効にする」を押してから印刷してください。");
  } catch (e) {
    $("makeMsg").textContent = "つくれませんでした：" + e.message;
  } finally {
    btn.disabled = false;
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
    const [cfg, master] = await Promise.all([download("データ/点検簿の設定.json", "json"),
                                             download("データ/現場マスタ.json", "json")]);
    S.cfg = cfg;
    S.sites = master.sites.slice().sort((a, b) => (b.period - a.period) || a.code.localeCompare(b.code));
    $("masterInfo").textContent = `現場 ${S.sites.length} 件（${(master.meta.made || "").slice(0, 10)} 作成）`;
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
