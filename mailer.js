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
  .mailbox-h{color:var(--gold,#d4af37); font-size:12.5px; letter-spacing:.3em; text-indent:.3em; opacity:.95;}
  .mailbox-p{font-size:13px; line-height:1.85; color:var(--ink-soft,#b9b08f); margin-top:9px;}
  .mailbox-f{margin-top:13px; display:flex; flex-direction:column; gap:9px;}
  .mailbox-f input[type=text], .mailbox-f input[type=email]{
    font-family:inherit; font-size:15px; color:var(--ink,#ede6cf);
    background:rgba(8,12,30,.62); border:1px solid rgba(212,175,55,.45); border-radius:10px;
    padding:11px 14px; outline:none; transition:border-color .15s, box-shadow .15s; width:100%;}
  .mailbox-f input::placeholder{color:rgba(185,176,143,.6);}
  .mailbox-f input:focus{border-color:var(--gold,#d4af37); box-shadow:0 0 0 2px rgba(212,175,55,.18);}
  .mailbox-agree{display:flex; align-items:flex-start; gap:8px; text-align:left;
    font-size:12.5px; line-height:1.7; color:var(--ink-soft,#b9b08f); padding:2px 2px 0;}
  .mailbox-agree input{margin-top:3px; accent-color:var(--gold,#d4af37); width:15px; height:15px; flex:0 0 auto;}
  .mailbox-btn{appearance:none; border:1.5px solid #fff2cf; cursor:pointer; font-family:inherit;
    background:linear-gradient(180deg,#f3e0a3,#d4af37 55%,#9c7a1e); color:#3a2a05;
    font-weight:700; font-size:15.5px; letter-spacing:.24em; text-indent:.24em; padding:12px 26px;
    border-radius:40px; box-shadow:0 6px 18px rgba(0,0,0,.4), inset 0 1px 0 #fff6d8; transition:transform .15s;}
  .mailbox-btn:hover{transform:translateY(-2px);}
  .mailbox-btn:active{transform:translateY(1px);}
  .mailbox-btn[disabled]{opacity:.55; cursor:default; transform:none;}
  .mailbox-msg{min-height:20px; margin-top:11px; font-size:13.5px; line-height:1.75;}
  .mailbox-msg.ok{color:#bff0c8;}
  .mailbox-msg.bad{color:#ffd9a8;}
  .mailbox-note{font-size:11.5px; line-height:1.7; color:var(--ink-soft,#b9b08f); opacity:.72; margin-top:10px;}
  .mailbox.done .mailbox-f, .mailbox.done .mailbox-agree{display:none;}
  `;

  function injectCSS(){
    if(styled) return; styled = true;
    const st = document.createElement("style"); st.textContent = CSS; document.head.appendChild(st);
  }

  function loadContact(){
    try{ return JSON.parse(localStorage.getItem(LS_KEY)) || {}; }catch(_){ return {}; }
  }
  function saveContact(name,email){
    try{ localStorage.setItem(LS_KEY, JSON.stringify({name:name,email:email})); }catch(_){ }
  }

  function validEmail(v){ return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v); }

  /* 送出：payload 只帶「編號」不帶文案，信件內容由伺服器端依同一份資料重建，
     避免此公開端點被拿來夾帶任意內容寄信。 */
  async function post(payload){
    const body = JSON.stringify(payload);
    try{
      const res = await fetch(MAIL_API, {
        method:"POST",
        headers:{"Content-Type":"text/plain;charset=utf-8"},  // 用 text/plain 避免 CORS preflight
        body: body,
        redirect:"follow"
      });
      const txt = await res.text();
      try{ return JSON.parse(txt); }catch(_){ return {ok:false, msg:"回應格式有誤，請稍後再試。"}; }
    }catch(e){
      // 少數瀏覽器（App 內建瀏覽器）讀不到跨網域回應 → 改用 no-cors 補送一次
      try{
        await fetch(MAIL_API, {method:"POST", mode:"no-cors",
          headers:{"Content-Type":"text/plain;charset=utf-8"}, body: body});
        return {ok:true, blind:true};
      }catch(e2){
        return {ok:false, msg:"連線失敗，請確認網路後再試一次。"};
      }
    }
  }

  /* 在 container 內加上寄送區塊。
     payload：{mode:"draw"|"face", n:1..53, face:{...}}  由各頁自行組好 */
  function mount(container, payload){
    if(!MAIL_API){ console.warn("[BaoyanMail] 尚未設定 MAIL_API，寄送區塊不顯示。"); return; }
    injectCSS();
    const saved = loadContact();
    const box = document.createElement("div");
    box.className = "mailbox";
    box.innerHTML = `
      <div class="mailbox-h">寄 一 份 到 我 的 信 箱</div>
      <p class="mailbox-p">留下姓名與 Email，這張${payload.mode==="face"?"面相觀照":"善知識"}牌卡的完整內容<br>即會寄到您的信箱，隨時回味、依之用功。</p>
      <div class="mailbox-f">
        <input type="text" class="m-name" maxlength="30" placeholder="您的姓名（稱呼）" value="${(saved.name||"").replace(/"/g,"&quot;")}" autocomplete="name">
        <input type="email" class="m-mail" maxlength="80" placeholder="您的 Email" value="${(saved.email||"").replace(/"/g,"&quot;")}" autocomplete="email" inputmode="email">
        <label class="mailbox-agree"><input type="checkbox" class="m-ok" checked>
          <span>我同意寶嚴禪寺以此 Email 寄送本次結果；日後若有法會、課程等法訊，亦歡迎通知我。</span></label>
        <button class="mailbox-btn" type="button">寄 送 結 果</button>
      </div>
      <div class="mailbox-msg"></div>
      <div class="mailbox-note">※ 我們只保存您的姓名與 Email 供寄送與法訊之用${payload.mode==="face"?"；您的照片全程留在本機，不會上傳、也不會出現在信中":""}。</div>`;
    container.appendChild(box);

    const nameEl = box.querySelector(".m-name"), mailEl = box.querySelector(".m-mail"),
          okEl = box.querySelector(".m-ok"), btn = box.querySelector(".mailbox-btn"),
          msg = box.querySelector(".mailbox-msg");

    function say(text, cls){ msg.textContent = text; msg.className = "mailbox-msg " + (cls||""); }

    btn.addEventListener("click", async function(){
      const name = nameEl.value.trim(), email = mailEl.value.trim();
      if(name.length < 1){ say("請留下您的姓名或稱呼。","bad"); nameEl.focus(); return; }
      if(!validEmail(email)){ say("Email 格式似乎不正確，請再檢查一次。","bad"); mailEl.focus(); return; }
      if(!okEl.checked){ say("請先勾選同意，我們才能寄信給您。","bad"); return; }

      btn.disabled = true; const label = btn.textContent; btn.textContent = "寄 送 中 …";
      say("正在為您寄出…","");
      const data = Object.assign({}, payload, {name:name, email:email, page:location.href});
      const res = await post(data);
      if(res && res.ok){
        saveContact(name, email);
        box.classList.add("done");
        say("已寄出！請查收信箱（約一分鐘內；若沒看到請看看「促銷／垃圾郵件」匣）。","ok");
      }else{
        btn.disabled = false; btn.textContent = label;
        say((res && res.msg) || "寄送失敗，請稍後再試。","bad");
      }
    });
  }

  return {mount:mount};
})();
