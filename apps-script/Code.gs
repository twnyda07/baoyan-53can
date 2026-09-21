/* =========================================================================
   善財童子五十三參 ── 抽籤／面相結果寄送信箱（Google Apps Script 後端）
   ------------------------------------------------------------------------
   ‧ 前端（mailer.js）只送「編號」：第幾參、面相等第、菩薩與功課的索引。
     信件內容一律由本檔依 Data.gs（＝網站同一份 signs.js / face-data.js）重建，
     所以這個公開端點不可能被拿來夾帶任意內容寄信。
   ‧ 每封寄出同時記錄到 Google 試算表，做為名單。
   ‧ 首次使用請先執行一次 setup()，會自動建立試算表並印出網址。
   ========================================================================= */

const CFG = {
  SITE:        "https://twnyda07.github.io/baoyan-53can/",
  SENDER_NAME: "寶嚴禪寺",
  REPLY_TO:    "",                    // 留空＝用部署者的 Gmail 位址
  SS_NAME:     "善財五十三參．寄送名單",
  SHEET_NAME:  "寄送紀錄",
  MAX_PER_EMAIL_PER_DAY: 5,           // 同一信箱每日最多寄幾封（防濫用）
  MAX_TOTAL_PER_DAY:     400          // 全站每日上限（Workspace 每日配額 1500）
};

const HEAD = ["時間","姓名","Email","類型","第幾參","善知識","有緣菩薩","三分數(智/業/福)","來源頁","狀態"];

/* ============ 入口 ============ */

function doGet() {
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8"><div style="font-family:sans-serif;padding:24px">' +
    '善財五十三參 · 寄送服務運作中。</div>');
}

function doPost(e) {
  let out;
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const req = validate_(p);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      checkQuota_(req.email);
      if (MailApp.getRemainingDailyQuota() < 5) {
        throw new Error("今日寄信量已達上限，請明日再試，或直接截圖保存結果。");
      }
      const mail = buildMail_(req);
      MailApp.sendEmail({
        to: req.email,
        subject: mail.subject,
        htmlBody: mail.html,
        body: mail.text,
        name: CFG.SENDER_NAME,
        replyTo: CFG.REPLY_TO || undefined
      });
      bumpQuota_(req.email);
      log_(req, "已寄出");
    } finally {
      lock.releaseLock();
    }
    out = { ok: true };
  } catch (err) {
    out = { ok: false, msg: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
                       .setMimeType(ContentService.MimeType.JSON);
}

/* ============ 驗證（只接受編號，超出範圍一律擋掉）============ */

function validate_(p) {
  const name = String(p.name || "").replace(/[\x00-\x1f<>]/g, "").trim();
  const email = String(p.email || "").trim();
  if (!name)                  throw new Error("請留下您的姓名或稱呼。");
  if (name.length > 30)       throw new Error("姓名過長。");
  if (email.length > 80 || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email))
                              throw new Error("Email 格式不正確。");

  const mode = (p.mode === "face") ? "face" : "draw";
  const n = Number(p.n);
  const sign = SIGNS.filter(function (s) { return s.n === n; })[0];
  if (!sign)                  throw new Error("找不到這一參的籤文。");

  const req = { name: name, email: email, mode: mode, sign: sign,
                page: String(p.page || "").slice(0, 200) };

  if (mode === "face") {
    const f = p.face || {};
    const lv = {};
    TRAIT_KEYS.forEach(function (k) {
      const v = f.lv && f.lv[k];
      lv[k] = (v === "high" || v === "low") ? v : "mid";
    });
    const sc = (f.scores || []).map(function (v) {
      v = Math.round(Number(v));
      return (isFinite(v) && v >= 0 && v <= 100) ? v : 0;
    });
    if (sc.length !== 3)      throw new Error("面相分數格式有誤。");
    const bi = idx_(f.bodhi, BODHISATTVAS.length);
    const hw = [idx_((f.hw || [])[0], JIAO_POOL.length),
                idx_((f.hw || [])[1], FUDE_POOL.length),
                idx_((f.hw || [])[2], CHAN_POOL.length)];
    const ms = ["上停","中停","下停"].indexOf(f.maxStop) >= 0 ? f.maxStop : "中停";
    req.face = { lv: lv, scores: sc, bodhi: BODHISATTVAS[bi], maxStop: ms,
                 homework: [JIAO_POOL[hw[0]], FUDE_POOL[hw[1]], CHAN_POOL[hw[2]]] };
  }
  return req;
}

function idx_(v, len) {
  v = Math.floor(Number(v));
  return (isFinite(v) && v >= 0 && v < len) ? v : 0;
}

/* ============ 每日配額（防止端點被濫用）============ */

function today_() {
  return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd");
}
function checkQuota_(email) {
  const props = PropertiesService.getScriptProperties();
  const d = today_();
  const total = Number(props.getProperty("q:" + d + ":_total") || 0);
  if (total >= CFG.MAX_TOTAL_PER_DAY)
    throw new Error("今日寄送量已滿，請明日再試，或先截圖保存結果。");
  const mine = Number(props.getProperty("q:" + d + ":" + email.toLowerCase()) || 0);
  if (mine >= CFG.MAX_PER_EMAIL_PER_DAY)
    throw new Error("這個信箱今天已寄送 " + CFG.MAX_PER_EMAIL_PER_DAY + " 封，請明日再試。");
}
function bumpQuota_(email) {
  const props = PropertiesService.getScriptProperties();
  const d = today_(), tk = "q:" + d + ":_total", ek = "q:" + d + ":" + email.toLowerCase();
  const first = !props.getProperty(tk);
  props.setProperty(tk, String(Number(props.getProperty(tk) || 0) + 1));
  props.setProperty(ek, String(Number(props.getProperty(ek) || 0) + 1));
  if (first) {   // 每天第一封時順手清掉舊日期的計數
    const all = props.getProperties();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf("q:") === 0 && k.indexOf("q:" + d + ":") !== 0) props.deleteProperty(k);
    });
  }
}

/* ============ 名單試算表 ============ */

function setup() {
  const sh = sheet_();
  const url = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SS_ID")).getUrl();
  Logger.log("名單試算表已就緒：" + url);
  return url;
}

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SS_ID"), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(CFG.SS_NAME);
    props.setProperty("SS_ID", ss.getId());
  }
  let sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0].getIndex() === 1 && ss.getSheets().length === 1
       ? ss.getSheets()[0].setName(CFG.SHEET_NAME)
       : ss.insertSheet(CFG.SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight("bold").setBackground("#f3e0a3");
    sh.setFrozenRows(1);
  }
  return sh;
}

function log_(req, status) {
  try {
    sheet_().appendRow([
      new Date(), req.name, req.email,
      req.mode === "face" ? "面相觀照" : "抽籤",
      req.sign.n, req.sign.name,
      req.face ? req.face.bodhi.name : "",
      req.face ? req.face.scores.join(" / ") : "",
      req.page, status
    ]);
  } catch (e) { /* 記錄失敗不影響寄信 */ }
}

/* ============ 信件內容（依 Data.gs 重建，與網站同一份文案）============ */

function esc_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function img_(n) {
  return CFG.SITE + "images/" + (n < 10 ? "0" + n : String(n)) + ".jpg";
}

const C = { ink:"#2b2417", soft:"#6d6350", gold:"#9c7a1e", goldL:"#d4af37",
            navy:"#131a40", paper:"#fffdf7", band:"#f6efdc", line:"#e3d5a8" };

function h3_(t) {
  return '<div style="font:700 13px/1.6 serif;letter-spacing:.24em;color:' + C.gold +
         ';border-bottom:1px solid ' + C.line + ';padding-bottom:7px;margin:22px 0 12px;text-align:center">' +
         t + '</div>';
}
function row_(k, v) {
  return '<tr><td width="58" valign="top" style="font:700 13.5px/1.75 serif;color:' + C.gold +
         ';padding:4px 8px 4px 0;white-space:nowrap">' + k + '</td>' +
         '<td valign="top" style="font:400 13.5px/1.75 serif;color:' + C.ink + ';padding:4px 0">' + v + '</td></tr>';
}
function box_(title, body) {
  return '<div style="background:' + C.band + ';border:1px solid ' + C.line +
         ';border-radius:10px;padding:13px 15px;margin-top:12px">' +
         (title ? '<div style="font:700 12.5px/1.6 serif;letter-spacing:.2em;color:' + C.gold + ';margin-bottom:5px">' + title + '</div>' : '') +
         '<div style="font:400 14.5px/1.95 serif;color:' + C.ink + '">' + body + '</div></div>';
}

function buildMail_(req) {
  const s = req.sign, f = req.face;
  const src = sourceLink(s);
  const isFace = req.mode === "face";
  const subject = "【寶嚴禪寺】" + (isFace ? "面相觀照" : "善知識籤") +
                  "．第 " + s.n + " 參 " + s.name + " — " + esc_(req.name) + " 的牌卡";

  const aspects = [["💼 事業工作", s.career], ["💗 感情姻緣", s.love],
                   ["💰 財運求財", s.wealth], ["🌿 健康身體", s.health],
                   ["📖 考試升學", s.exam],   ["🏠 遷移搬遷", s.move],
                   ["🧭 出行外出", s.travel], ["🎯 謀望求事", s.plan]];
  const aspectHTML = '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px">' +
    aspects.map(function (a) {
      return '<tr><td style="padding:5px 0;border-top:1px dashed ' + C.line + '">' +
             '<div style="font:700 13.5px/1.7 serif;color:' + C.gold + '">' + a[0] + '</div>' +
             '<div style="font:400 13.5px/1.75 serif;color:' + C.ink + '">' + esc_(a[1]) + '</div></td></tr>';
    }).join("") + '</table>';

  let faceHTML = "";
  if (isFace) {
    faceHTML =
      h3_("面　相　觀　照") +
      '<table width="100%" cellpadding="0" cellspacing="0">' +
      TRAIT_KEYS.map(function (k) { return row_(k, esc_(READINGS[k][f.lv[k]])); }).join("") +
      '</table>' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px"><tr>' +
      [["智慧", f.scores[0]], ["事業", f.scores[1]], ["福德", f.scores[2]]].map(function (x) {
        return '<td align="center" width="33%" style="padding:10px 4px;background:' + C.band +
               ';border:1px solid ' + C.line + ';border-radius:10px">' +
               '<div style="font:700 24px/1.2 serif;color:' + C.gold + '">' + x[1] + '</div>' +
               '<div style="font:400 12px/1.6 serif;color:' + C.soft + ';letter-spacing:.12em">' + x[0] + '</div></td>';
      }).join('<td width="8"></td>') + '</tr></table>' +
      '<div style="font:400 13px/1.8 serif;color:' + C.soft + ';margin-top:10px">' + esc_(STOP_FOCUS[f.maxStop]) + '</div>' +

      '<div style="margin-top:20px;padding:16px;border:1.5px solid ' + C.goldL +
      ';border-radius:12px;background:' + C.band + '">' +
      '<div style="font:400 12px/1.6 serif;letter-spacing:.28em;color:' + C.gold + ';text-align:center">與 您 有 緣 的 菩 薩</div>' +
      '<div style="font:700 21px/1.5 serif;color:' + C.gold + ';text-align:center;margin-top:6px">' + f.bodhi.icon + '　' + f.bodhi.name + '</div>' +
      '<div style="font:400 13px/1.7 serif;color:' + C.soft + ';text-align:center">' + esc_(f.bodhi.symbol) + '</div>' +
      '<div style="font:400 13.5px/1.9 serif;color:' + C.ink + ';margin-top:12px">' + faceMessage(f.bodhi.name) + '</div>' +
      '<div style="font:400 12px/1.6 serif;letter-spacing:.28em;color:' + C.gold + ';text-align:center;margin:14px 0 6px">每 日 功 課</div>' +
      '<table width="100%" cellpadding="0" cellspacing="0">' +
        row_("教　理", esc_(f.homework[0])) + row_("福　德", esc_(f.homework[1])) + row_("禪　定", esc_(f.homework[2])) +
      '</table>' +
      '<div style="text-align:center;margin-top:12px"><span style="display:inline-block;background:' + C.goldL +
      ';color:#2a1d02;font:700 13px/1.6 serif;letter-spacing:.12em;border-radius:20px;padding:5px 16px">✦　' +
      esc_(f.bodhi.reward) + '　✦</span></div></div>';
  }

  const html =
  '<div style="margin:0;padding:24px 12px;background:#efe8d6">' +
  '<table align="center" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:' + C.paper +
  ';border:1px solid ' + C.line + ';border-radius:16px;overflow:hidden">' +
  '<tr><td style="background:' + C.navy + ';padding:20px 24px;text-align:center">' +
    '<div style="font:400 12px/1.6 serif;letter-spacing:.5em;color:#f3e0a3">寶 嚴 禪 寺</div>' +
    '<div style="font:700 20px/1.5 serif;color:#f3e0a3;margin-top:6px">善財童子五十三參．' + (isFace ? "面相觀照" : "善知識籤") + '</div>' +
  '</td></tr>' +
  '<tr><td style="padding:24px">' +
    '<div style="font:400 14.5px/1.9 serif;color:' + C.ink + '">' + esc_(req.name) + ' 菩薩 您好：<br>' +
    '這是您在寶嚴禪寺「善財童子五十三參」' + (isFace ? "面相觀照" : "虔心抽卡") + '所得的牌卡，謹寄呈如下。</div>' +

    '<div style="text-align:center;margin-top:18px">' +
      '<img src="' + img_(s.n) + '" width="240" alt="' + esc_(s.name) + '" style="width:240px;max-width:70%;border:1px solid ' + C.line + ';border-radius:12px">' +
      '<div style="font:400 13px/1.7 serif;color:' + C.soft + ';margin-top:12px">第 <b style="color:' + C.gold + ';font-size:16px">' + s.n + '</b> 參　／　共五十三參' +
      (isFace ? '' : '　·　<b style="color:' + C.gold + '">' + esc_(s.luck) + '</b>') + '</div>' +
      '<div style="font:700 25px/1.5 serif;color:' + C.gold + ';margin-top:4px">' + esc_(s.name) + '</div>' +
      '<div style="font:400 13px/1.7 serif;color:' + C.soft + '">參訪：' + esc_(s.place || "") + '</div>' +
      '<div style="font:400 13px/1.9 serif;color:' + C.ink + ';margin-top:8px">獨特法門　' + (s.keys || []).map(esc_).join("　·　") + '</div>' +
    '</div>' +

    '<div style="margin:18px auto 0;padding:14px 8px;border-top:1px solid ' + C.line + ';border-bottom:1px solid ' + C.line + ';text-align:center">' +
      '<div style="font:700 20px/2 serif;letter-spacing:.3em;color:' + C.gold + '">' + esc_(s.poem[0]) + '</div>' +
      '<div style="font:700 20px/2 serif;letter-spacing:.3em;color:' + C.gold + '">' + esc_(s.poem[1]) + '</div>' +
    '</div>' +

    (isFace ? '<div style="font:400 13.5px/1.8 serif;color:' + C.ink + ';margin-top:12px"><b style="color:' + C.gold + '">所應法門　</b>' + esc_(s.gate) + '</div>' : '') +
    box_(isFace ? "善知識開示" : "解　曰", esc_(s.jie)) +

    (s.challenge ?
      box_("此　籤　勉　勵", esc_(s.challenge)) +
      '<div style="font:700 12.5px/1.6 serif;letter-spacing:.14em;color:' + C.gold + ';margin:14px 0 4px">精進之道 ・ 教理／福德／禪定</div>' +
      '<table width="100%" cellpadding="0" cellspacing="0">' +
        row_("教　理", esc_(s.overcome.jiao)) + row_("福　德", esc_(s.overcome.fude)) + row_("禪　定", esc_(s.overcome.chan)) +
      '</table>' : '') +

    faceHTML +

    h3_("八　面　向　籤　解") + aspectHTML +

    '<div style="text-align:center;margin-top:22px">' +
      '<a href="' + src.url + '" style="display:inline-block;background:' + C.goldL + ';color:#2a1d02;text-decoration:none;font:700 13.5px/1.6 serif;border-radius:30px;padding:10px 20px">▶　聆聽出處 · ' + esc_(src.label) + '</a>' +
    '</div>' +
    '<div style="text-align:center;margin-top:12px">' +
      '<a href="' + CFG.SITE + '" style="color:' + C.gold + ';font:400 13px/1.7 serif">再抽一張 · 善財童子五十三參</a>' +
    '</div>' +
  '</td></tr>' +
  '<tr><td style="background:' + C.band + ';padding:16px 24px;font:400 11.5px/1.85 serif;color:' + C.soft + ';text-align:center">' +
    '籤文依《大方廣佛華嚴經‧入法界品》五十三參發展，為勸善省思之方便語，非命理斷定。<br>' +
    (isFace ? '您的照片全程留在您的裝置上運算，未曾上傳，本信亦不含您的影像。<br>' : '') +
    '本信因您在網站上主動留下 Email 而寄出；若不願再收到法訊，直接回信告知即可。<br>' +
    '© 寶嚴禪寺　Baoyan Chan Monastery' +
  '</td></tr></table></div>';

  /* 純文字版（不支援 HTML 的信箱） */
  let text = req.name + " 菩薩 您好：\n\n" +
    "善財童子五十三參．" + (isFace ? "面相觀照" : "善知識籤") + "\n" +
    "第 " + s.n + " 參　" + s.name + (isFace ? "" : "（" + s.luck + "）") + "\n" +
    "參訪：" + (s.place || "") + "\n" +
    "獨特法門：" + (s.keys || []).join("、") + "\n\n" +
    s.poem[0] + "\n" + s.poem[1] + "\n\n" +
    (isFace ? "所應法門：" + s.gate + "\n\n" : "") +
    (isFace ? "善知識開示：" : "解曰：") + s.jie + "\n\n" +
    (s.challenge ? "此籤勉勵：" + s.challenge + "\n" +
      "‧教理：" + s.overcome.jiao + "\n‧福德：" + s.overcome.fude + "\n‧禪定：" + s.overcome.chan + "\n\n" : "");
  if (isFace) {
    text += "【面相觀照】\n" +
      TRAIT_KEYS.map(function (k) { return k + "：" + READINGS[k][f.lv[k]]; }).join("\n") + "\n" +
      "智慧 " + f.scores[0] + "／事業 " + f.scores[1] + "／福德 " + f.scores[2] + "\n" +
      STOP_FOCUS[f.maxStop] + "\n\n" +
      "【與您有緣的菩薩】" + f.bodhi.name + "（" + f.bodhi.symbol + "）\n" +
      "每日功課：\n‧教理：" + f.homework[0] + "\n‧福德：" + f.homework[1] + "\n‧禪定：" + f.homework[2] + "\n\n";
  }
  text += "【八面向籤解】\n" +
    aspects.map(function (a) { return a[0] + "：" + a[1]; }).join("\n") + "\n\n" +
    "聆聽出處：" + src.label + "\n" + src.url + "\n" +
    "再抽一張：" + CFG.SITE + "\n\n" +
    "籤文依《華嚴經‧入法界品》五十三參發展，為勸善省思之方便語，非命理斷定。\n" +
    "© 寶嚴禪寺";

  return { subject: subject, html: html, text: text };
}

/* ============ 自我測試（在編輯器直接執行，寄一封給自己）============ */

function testDraw() {
  const me = Session.getActiveUser().getEmail();
  const mail = buildMail_(validate_({ name: "測試", email: me, mode: "draw", n: 7 }));
  MailApp.sendEmail({ to: me, subject: mail.subject, htmlBody: mail.html, body: mail.text, name: CFG.SENDER_NAME });
  Logger.log("已寄出抽籤測試信到 " + me);
}

function testFace() {
  const me = Session.getActiveUser().getEmail();
  const mail = buildMail_(validate_({
    name: "測試", email: me, mode: "face", n: 21,
    face: { lv: { 上停:"high", 中停:"mid", 下停:"low", 氣色:"mid", 神采:"high", 端正:"mid" },
            scores: [93, 85, 74], maxStop: "上停", bodhi: 2, hw: [1, 3, 0] }
  }));
  MailApp.sendEmail({ to: me, subject: mail.subject, htmlBody: mail.html, body: mail.text, name: CFG.SENDER_NAME });
  Logger.log("已寄出面相測試信到 " + me);
}
