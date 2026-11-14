/* =========================================================================
   寄送牌卡結果到 Email（抽籤頁／面相頁共用）
   後端：Google Apps Script 網頁應用程式（見 apps-script/ 資料夾）
   ------------------------------------------------------------------------
   部署完 Apps Script 後，把「網頁應用程式網址」貼到下面 MAIL_API 即可啟用；
   留空時整個寄送區塊不會出現（網站仍可正常抽籤／觀照）。
   ========================================================================= */

const MAIL_API = "https://script.google.com/macros/s/AKfycbxY3gPydDZYhdyaxEn4Ta5tr60bnP0RiggJSiLI6-x9GPDvzXl7VjF-8jsJAYnpZ6vQAg/exec";

const BaoyanMail = (function(){
  const LS_KEY = "baoyan53_contact";
  let styled = false;

  const CSS = `
  .mailbox{margin:20px auto 0; max-width:440px; padding:18px 18px 16px; text-align:center;
    background:linear-gradient(180deg,rgba(212,175,55,.13),rgba(8,12,30,.55));
    border:1.5px solid var(--gold,#d4af37); border-radius:14px;
    box-shadow:inset 0 0 22px rgba(212,175,55,.12), 0 0 16px rgba(212,175,55,.18);}
  .mailbox-h{color:var(--gold,#d4af37); font-size:14px; letter-spacing:.3em; text-indent:.3em; opacity:.95;}
  .mailbox-p{font-size:14px; line-height:1.85; color:var(--ink-soft,#b9b08f); margin-top:9px;}
  .mailbox-f{margin-top:13px; display:flex; flex-direction:column; gap:9px;}
  .mailbox-f input[type=text], .mailbox-f input[type=email]{
    font-family:inherit; font-size:16px; color:var(--ink,#ede6cf);
    background:rgba(8,12,30,.62); border:1px solid rgba(212,175,55,.45); border-radius:10px;
    padding:11px 14px; outline:none; transition:border-color .15s, box-shadow .15s; width:100%;}
  .mailbox-f input::placeholder{color:rgba(185,176,143,.6);}
  .mailbox-f input:focus{border-color:var(--gold,#d4af37); box-shadow:0 0 0 2px rgba(212,175,55,.18);}
  .mailbox-agree{display:flex; align-items:flex-start; gap:8px; text-align:left;
    font-size:14px; line-height:1.7; color:var(--ink-soft,#b9b08f); padding:2px 2px 0; min-height:44px; cursor:pointer;}
  .mailbox-agree input{margin-top:3px; accent-color:var(--gold,#d4af37); width:20px; height:20px; flex:0 0 auto;}
  .mailbox-btn{appearance:none; border:1.5px solid #fff2cf; cursor:pointer; font-family:inherit;
    background:linear-gradient(180deg,#f3e0a3,#d4af37 55%,#9c7a1e); color:#3a2a05;
    font-weight:700; font-size:15.5px; letter-spacing:.24em; text-indent:.24em; padding:12px 26px;
    border-radius:40px; box-shadow:0 6px 18px rgba(0,0,0,.4), inset 0 1px 0 #fff6d8; transition:transform .15s;}
  .mailbox-btn:hover{transform:translateY(-2px);}
  .mailbox-btn:active{transform:translateY(1px);}
  .mailbox-btn[disabled]{opacity:.55; cursor:default; transform:none;}
  .mailbox-msg{min-height:20px; margin-top:11px; font-size:15px; line-height:1.75;}
  .mailbox-msg.ok{color:#bff0c8;}
  .mailbox-msg.bad{color:#ffd9a8;}
  .mailbox-note{font-size:13.5px; line-height:1.75; color:var(--ink-soft,#b9b08f); opacity:.85; margin-top:10px;}
  .mailbox.done .mailbox-f, .mailbox.done .mailbox-agree{display:none;}
  .mailbox [hidden]{display:none !important;}
  .mailbox-alt{margin-top:12px; padding-top:12px; border-top:1px dashed rgba(212,175,55,.4);}
  .mailbox-alt-t{font-size:14.5px; line-height:1.75; color:var(--ink,#ede6cf); margin-bottom:9px;}
  .mailbox-alt-b{display:block; width:100%; box-sizing:border-box; margin-top:8px; font-family:inherit; font-size:14.5px;
    letter-spacing:.08em; padding:11px 14px; border-radius:30px; cursor:pointer; text-decoration:none; text-align:center;
    color:var(--gold,#d4af37); background:rgba(8,12,30,.55); border:1.5px solid var(--gold,#d4af37);}
  .mailbox-alt-b.primary{color:#3a2a05; font-weight:700; border-color:#fff2cf;
    background:linear-gradient(180deg,#f3e0a3,#d4af37 55%,#9c7a1e);}
  `;

  function injectCSS(){
    if(styled) return; styled = true;
    const st = document.createElement("style"); st.textContent = CSS; document.head.appendChild(st);
  }

  /* 不記住任何人的姓名與 Email：道場的平板／電腦是共用的，
     若自動帶出上一位的位址，下一位沒注意按下去就會把牌卡寄到別人信箱。
     （順手清掉舊版本可能已經存在這台裝置上的資料） */
  function forgetContact(){
    try{ localStorage.removeItem(LS_KEY); }catch(_){ }
  }

  /* 每次掛載給兩個欄位一組隨機的 name：Chrome 對 autocomplete=off 常不理會，
     會依欄位名稱猜這是「姓名／Email」而跳出上一位存在瀏覽器裡的位址。
     名稱每次都不一樣，猜不中就不會跳。 */
  function rnd(){ return Math.random().toString(36).slice(2,10); }

  function validEmail(v){ return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v); }

  /* 送出：payload 只帶「編號」不帶文案，信件內容由伺服器端依同一份資料重建，
     避免此公開端點被拿來夾帶任意內容寄信。
     payload 另帶一個 rid（本次送出的識別碼）：Apps Script 的轉址偶爾會讓 fetch 失敗，
     我們會重試，而後端看到同一個 rid 就不會重複寄信。 */
  /* 每次連線的等待上限。沒有上限時，公共 Wi-Fi 擋住 script.google.com（不回應）
     按鈕會一直卡在「寄送中…」好幾分鐘。第一次給長一點：Apps Script 冷啟動加寄信本來就要幾秒。
     全部加起來不超過 TOTAL（約 20 秒）；第 SLOW_HINT 毫秒還沒結果，就先把「用我的郵件 App 寄給自己」打開，不讓人乾等。 */
  const WAIT = [12000, 7000], TOTAL = 20000, SLOW_HINT = 8000;
  const NET_FAIL = "連不上寄信服務（公共 Wi-Fi 常會擋住）。請直接按下方「用我的郵件 App 寄給自己」；" +
                   "或關掉 Wi-Fi、改用行動網路，再按一次「寄送結果」（不會重複寄）。";
  /* 讀不到後台回應時，無法知道信到底有沒有寄出（後台壞掉、要重新授權、被 Wi-Fi 登入頁攔下都長這樣），
     一律不可說「已寄出」。 */
  const UNKNOWN = "無法確認信是否送達。若兩分鐘內沒收到信（也請看看「促銷／垃圾郵件」匣），" +
                  "請按下方「用我的郵件 App 寄給自己」，內容都已幫您填好。";

  /* 送一次並讀完回應；超過 ms 毫秒就中止（讀回應本身也算在內） */
  async function once(body, ms, blind){
    const ctl = (typeof AbortController === "function") ? new AbortController() : null;
    const timer = ctl ? setTimeout(function(){ ctl.abort(); }, ms) : 0;
    try{
      const opt = {method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"},  // text/plain 避免 CORS preflight
                   body: body, signal: ctl ? ctl.signal : undefined};
      if(blind){ opt.mode = "no-cors"; await fetch(MAIL_API, opt); return ""; }
      opt.redirect = "follow";
      const res = await fetch(MAIL_API, opt);
      return await res.text();
    }finally{ clearTimeout(timer); }
  }

  /* 回傳：{ok:true} 確定寄出／{ok:false,msg} 確定沒寄出／{ok:null,msg} 無法確認（一律要給備援） */
  async function post(payload, onSlow){
    const body = JSON.stringify(payload);
    const t0 = Date.now(), left = function(){ return TOTAL - (Date.now() - t0); };
    const slow = setTimeout(function(){ if(onSlow) onSlow(); }, SLOW_HINT);
    try{
      let txt = null;
      for(let i=0; i<WAIT.length && txt===null && left() > 2500; i++){
        if(i) await new Promise(r=>setTimeout(r, 800));
        try{ txt = await once(body, Math.min(WAIT[i], left() - 2000), false); }catch(_){ /* 逾時或被擋，續試 */ }
      }
      if(txt !== null){
        let j = null;
        try{ j = JSON.parse(txt); }catch(_){ }
        if(!j || typeof j !== "object")   // 讀得到但不是後台的回應（例如 Wi-Fi 登入頁、Google 錯誤頁）
          return {ok:null, msg:UNKNOWN};
        if(j.ok === true) return {ok:true, dup:!!j.dup};
        if(j.pending) return {ok:null, msg:UNKNOWN};
        return {ok:false, msg:j.msg || ""};
      }
      /* 讀不到回應（App 內建瀏覽器擋跨網域時會這樣）→ 用 no-cors 再送一次（同 rid 後台不會重複寄）。
         no-cors 看不到回應內容，送得出去不代表寄得出去，所以只能說「無法確認」，並打開備援。 */
      try{ await once(body, Math.max(left(), 1500), true); return {ok:null, blind:true, msg:UNKNOWN}; }
      catch(_){ return {ok:false, msg:NET_FAIL}; }
    }finally{ clearTimeout(slow); }
  }

  /* ── 備援：寄信服務連不上（公共 Wi-Fi 擋住、或當日配額用完）時，
        改用訪客手機自己的郵件 App／分享，把畫面上的結果寄給自己。完全不經過 script.google.com。 */
  function resultText(root){
    const out = [];
    Array.prototype.forEach.call(root.children, function(el){
      if(el.matches(".mailbox, .nextstop, .row, button, script, style") || !el.getClientRects().length) return;
      const t = String(el.innerText || "").replace(/[ \t　]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
      if(t) out.push(t);
    });
    let s = out.join("\n\n");
    if(s.length > 6000) s = s.slice(0, 6000) + "\n……（全文請回網站觀看）";
    /* innerText 不含網址：把結果裡的外部連結（如見輝法師開示影片）另列在後 */
    const links = [];
    Array.prototype.forEach.call(root.querySelectorAll("a[href^='http']"), function(a){
      if(a.closest(".mailbox, .nextstop, .row")) return;
      const t = String(a.innerText || "").replace(/\s+/g, " ").trim();
      if(links.indexOf(a.href) < 0) links.push(t ? "▶ " + t + "\n" + a.href : a.href);
    });
    if(links.length) s += "\n\n" + links.join("\n\n");
    return s;
  }
  function altPack(root, title, name){
    return { subject: "【寶嚴禪寺】" + title + " — " + name,
             body: name + " 菩薩 您好：\n\n" + resultText(root) + "\n\n" +
                   location.origin + location.pathname + "\n© 寶嚴山寶嚴禪寺" };
  }
  function mailtoHref(email, p){
    return "mailto:" + encodeURIComponent(email).replace(/%40/g, "@") +
      "?subject=" + encodeURIComponent(p.subject) +
      "&body=" + encodeURIComponent(p.body.replace(/\r?\n/g, "\r\n"));
  }

  function newRid(){
    return "r" + Date.now().toString(36) + Math.random().toString(36).slice(2,10);
  }

  /* 在 container 內加上寄送區塊。
     payload：{mode:"draw"|"face", n:1..53, face:{...}}  由各頁自行組好 */
  function mount(container, payload){
    if(!MAIL_API){ console.warn("[BaoyanMail] 尚未設定 MAIL_API，寄送區塊不顯示。"); return; }
    injectCSS();
    forgetContact();
    const fid = rnd();
    const box = document.createElement("div");
    box.className = "mailbox";
    box.innerHTML = `
      <div class="mailbox-h">寄 一 份 到 我 的 信 箱</div>
      <p class="mailbox-p">留下姓名與 Email，這張${payload.mode==="face"?"面相觀照":"善知識"}牌卡的完整內容<br>即會寄到您的信箱，隨時回味、依之用功。</p>
      <div class="mailbox-f">
        <input type="text" class="m-name" name="n-${fid}" maxlength="30" placeholder="您的姓名（稱呼）" autocomplete="off" autocorrect="off" spellcheck="false">
        <input type="email" class="m-mail" name="e-${fid}" maxlength="80" placeholder="您的 Email" autocomplete="off" autocorrect="off" spellcheck="false" inputmode="email">
        <label class="mailbox-agree"><input type="checkbox" class="m-ok" checked>
          <span>我同意寶嚴禪寺以此 Email 寄送本次結果；日後若有法會、課程等法訊，亦歡迎通知我。</span></label>
        <button class="mailbox-btn" type="button">寄 送 結 果</button>
      </div>
      <div class="mailbox-msg"></div>
      <div class="mailbox-alt" hidden>
        <div class="mailbox-alt-t">直接用您自己的手機保存這份結果：</div>
        <a class="mailbox-alt-b primary m-alt-mail" href="#">用我的郵件 App 寄給自己</a>
        <button class="mailbox-alt-b m-alt-share" type="button" hidden>分享到 LINE／其他 App 保存</button>
      </div>
      <div class="mailbox-note">※ 我們只保存您的姓名與 Email 供寄送與法訊之用${payload.mode==="face"?"；您的照片全程留在本機，不會上傳、也不會出現在信中":""}。<br>※ 這台裝置不會記住您填的資料（共用裝置請每位自行填寫）。</div>`;
    container.appendChild(box);

    const nameEl = box.querySelector(".m-name"), mailEl = box.querySelector(".m-mail"),
          okEl = box.querySelector(".m-ok"), btn = box.querySelector(".mailbox-btn"),
          msg = box.querySelector(".mailbox-msg");

    function say(text, cls){ msg.textContent = text; msg.className = "mailbox-msg " + (cls||""); }
    /* 本次送出的識別碼只認「同一組姓名＋Email」：若連線逾時但其實已寄出，
       使用者改正打錯的 Email 再按，必須換新的 rid，否則後端會當成重複而不寄到新信箱。 */
    let rid = null, ridFor = null;

    const alt = box.querySelector(".mailbox-alt"), altMail = box.querySelector(".m-alt-mail"),
          altShare = box.querySelector(".m-alt-share");
    const altTitle = payload.mode === "face" ? "面相觀照牌卡" : "善知識籤牌卡";
    let altP = null;
    function hideAlt(){ alt.hidden = true; altMail.setAttribute("href", "#"); altP = null; }
    function showAlt(name, email){
      altP = altPack(container, altTitle, name);
      altMail.setAttribute("href", mailtoHref(email, altP));
      altShare.hidden = (typeof navigator.share !== "function");
      alt.hidden = false;
      try{ alt.scrollIntoView({behavior:"smooth", block:"nearest"}); }catch(_){}
    }
    altMail.addEventListener("click", function(e){ if(!altP) e.preventDefault(); });
    altShare.addEventListener("click", function(){
      if(altP && navigator.share) navigator.share({title: altP.subject, text: altP.body}).catch(function(){});
    });
    nameEl.addEventListener("input", hideAlt); mailEl.addEventListener("input", hideAlt);

    btn.addEventListener("click", async function(){
      const name = nameEl.value.trim(), email = mailEl.value.trim();
      if(name.length < 1){ say("請留下您的姓名或稱呼。","bad"); nameEl.focus(); return; }
      if(!validEmail(email)){ say("Email 格式似乎不正確，請再檢查一次。","bad"); mailEl.focus(); return; }
      if(!okEl.checked){ say("請先勾選同意，我們才能寄信給您。","bad"); return; }

      btn.disabled = true; const label = btn.textContent; btn.textContent = "寄 送 中 …";
      say("正在為您寄出…","");
      if(ridFor !== name + "\n" + email){ rid = newRid(); ridFor = name + "\n" + email; }   // 同一次送出重試共用，避免重複寄信
      const data = Object.assign({}, payload, {name:name, email:email, page:location.href, rid:rid});
      const res = await post(data, function(){
        say("網路較慢，仍在為您寄出…。不想等的話，可以直接按下方「用我的郵件 App 寄給自己」。","");
        showAlt(name, email);
      });
      if(!box.isConnected) return;   // 等待中已換下一位（結果區已清掉），不再顯示上一位的訊息
      if(res && res.ok === true){
        nameEl.value = ""; mailEl.value = "";   // 立刻清掉，後面的人看不到也帶不走
        hideAlt();
        box.classList.add("done");
        say("已寄出！請查收信箱（約一分鐘內；若沒看到請看看「促銷／垃圾郵件」匣）。","ok");
      }else if(res && res.ok === null){
        /* 無法確認：不說「已寄出」，一定給備援；按鈕恢復可再按（同一組姓名＋Email 共用 rid，後台不會重複寄） */
        btn.disabled = false; btn.textContent = label;
        say(res.msg || "無法確認信是否送達，請按下方「用我的郵件 App 寄給自己」。", "bad");
        msg.title = res.blind ? "no-cors 盲送：看不到後台回應" : "";
        showAlt(name, email);
      }else{
        btn.disabled = false; btn.textContent = label;
        /* 後端程式錯誤（英文技術訊息）不直接丟給信眾看：改用中文說明並導向下方備援；原文留在 title 與 console 供工作人員查 */
        const raw = (res && res.msg) || "";
        const tech = raw && !/[\u4e00-\u9fff]/.test(raw);
        if(tech) console.warn("[BaoyanMail] 後端錯誤：", raw);
        say(tech ? "寄信服務暫停中。請按下方「用我的郵件 App 寄給自己」——收件人與內容都已幫您填好，在郵件畫面按「傳送」即可。"
                 : (raw || "寄送失敗，請稍後再試。"), "bad");
        msg.title = tech ? raw : "";
        showAlt(name, email);
      }
    });
  }

  return {mount:mount};
})();
