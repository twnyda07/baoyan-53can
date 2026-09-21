#!/usr/bin/env node
/* 把網站的 signs.js ＋ face-data.js 打包成 Apps Script 用的 Data.gs
   —— 信件內容與網站永遠同一份文案，改網站資料後重跑一次即可。
   用法：node apps-script/build.js   （在專案根目錄執行）                        */

const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const parts = ["signs.js", "face-data.js"].map(function (f) {
  return "/* ===== 以下內容自動產生自 " + f + "，請勿直接修改 ===== */\n" +
         fs.readFileSync(path.join(root, f), "utf8");
});

const out =
"/* =========================================================================\n" +
"   Data.gs ── 由 apps-script/build.js 自動產生，請勿手動編輯。\n" +
"   來源：signs.js（53 參籤文）、face-data.js（面相文案／菩薩／功課池）\n" +
"   改完來源檔後，在專案根目錄執行： node apps-script/build.js\n" +
"   ========================================================================= */\n\n" +
parts.join("\n\n");

const dest = path.join(__dirname, "Data.gs");
fs.writeFileSync(dest, out, "utf8");
console.log("已產生 " + dest + "（" + Math.round(out.length / 1024) + " KB，" + "共 " +
            (out.match(/"n":\s*\d+/g) || []).length + " 參）");
