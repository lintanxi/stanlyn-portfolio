/* ──────────────────────────────────────────────────────────────
   gate.js — password gate for case study pages.
   Include as the FIRST script in <head> (no defer/async —
     it must run before the body paints): <script src="gate.js"></script>
   Note: client-side only. Deters casual access; not real security.
   ────────────────────────────────────────────────────────────── */
(function () {
  var KEY = 'sl_cs_unlock';
  var TOKEN = 'ok';
  /* SHA-256 of the passphrase */
  var HASH = '92f473809e5979a7da2b7f52771b3a9e3d7105dcb0f24ae333c5bf279b863d73';
  var ALT = 'c2w6c3Bhcms=';

  /* Already unlocked this browser? Do nothing. */
  try {
    if (localStorage.getItem(KEY) === TOKEN) return;
  } catch (e) {}

  /* Hide the page immediately so nothing flashes before the gate paints. */
  var hider = document.createElement('style');
  hider.id = 'gate-hider';
  hider.textContent = 'html{visibility:hidden}html.gate-on{visibility:visible}' +
    /* everything hides except the gate and the site header, so the brand stays clickable */
    'html.gate-on body>*:not(#gate):not(header){display:none!important}' +
    'html.gate-on,html.gate-on body{overflow:hidden;height:100%}';
  (document.head || document.documentElement).appendChild(hider);

  var css = [
    '#gate{position:fixed;inset:0;z-index:10;display:flex;align-items:center;justify-content:center;',
    'padding:24px;padding-top:calc(var(--gate-top,62px) + 24px);background:#f0efeb;color:#222;',
    'font-family:"Everyday Sans","Helvetica Neue",Arial,sans-serif;-webkit-font-smoothing:antialiased}',
    '#gate .box{width:100%;max-width:380px;text-align:center}',
    '#gate .mark{width:34px;height:34px;margin:0 auto 22px;color:#3a9d3a}',
    '#gate h1{font-size:22px;font-weight:700;letter-spacing:-.015em;line-height:1.3;margin:0 0 8px}',
    '#gate p{font-size:15px;line-height:1.6;color:#515357;margin:0 0 24px}',
    '#gate form{display:flex;gap:8px}',
    '#gate input{flex:1;min-width:0;font:inherit;font-size:16px;padding:12px 14px;border:1px solid #d8d7cd;',
    'border-radius:10px;background:#fff;color:#222}',
    '#gate input::placeholder{color:#8b8d91}',
    '#gate input:focus{outline:none;border-color:#0a5fd6;box-shadow:0 0 0 3px rgba(10,95,214,.16)}',
    '#gate button{font:inherit;font-size:15px;font-weight:600;padding:12px 20px;border:0;border-radius:10px;',
    'background:#0a5fd6;color:#fff;cursor:pointer}',
    '#gate button:hover{background:#002e99}',
    '#gate button:focus-visible{outline:3px solid #002e99;outline-offset:2px}',
    '#gate .err{min-height:20px;margin:12px 0 0;font-size:14px;color:#b3261e;opacity:0;transition:opacity .15s}',
    '#gate .err.on{opacity:1}',
    '#gate .shake{animation:gateShake .32s}',
    '@keyframes gateShake{0%,100%{transform:translateX(0)}25%{transform:translateX(-6px)}75%{transform:translateX(6px)}}',
    '@media(prefers-reduced-motion:reduce){#gate .shake{animation:none}}',
    '#gate .foot{margin:26px 0 0;font-size:14px}',
    '#gate .foot a{color:#515357;text-decoration:none;border-bottom:1px solid #d8d7cd}',
    '#gate .foot a:hover{color:#222;border-color:#222}'
  ].join('');

  var HTML = '<div class="box">' +
    '<svg class="mark" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" ' +
    'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<rect x="4" y="10.5" width="16" height="10.5" rx="2.5"></rect>' +
    '<path d="M8 10.5V7a4 4 0 0 1 8 0v3.5"></path></svg>' +
    '<h1>This case study is private</h1>' +
    '<p>Enter the password to continue. Don’t have it? Email me and I’ll share it.</p>' +
    '<form novalidate>' +
    '<label for="gate-pw" style="position:absolute;left:-9999px">Password</label>' +
    '<input id="gate-pw" type="password" placeholder="Password" autocomplete="current-password" ' +
    'autocapitalize="off" autocorrect="off" spellcheck="false">' +
    '<button type="submit">Enter</button>' +
    '</form>' +
    '<p class="err" role="alert">That password isn’t right. Try again.</p>' +
    '<p class="foot"><a href="mailto:stanlyn.lu@gmail.com">Request access</a></p>' +
    '</div>';

  /* Primary check: SHA-256 via WebCrypto (needs a secure context).
     Fallback keeps local file:// previews working, where subtle is absent. */
  function check(pw) {
    if (window.crypto && crypto.subtle && window.isSecureContext) {
      var enc = new TextEncoder().encode(pw);
      return crypto.subtle.digest('SHA-256', enc).then(function (buf) {
        var hex = Array.prototype.map.call(new Uint8Array(buf), function (b) {
          return b.toString(16).padStart(2, '0');
        }).join('');
        return hex === HASH;
      }).catch(function () {
        return btoa('sl:' + pw) === ALT;
      });
    }
    return Promise.resolve(btoa('sl:' + pw) === ALT);
  }

  function unlock() {
    try { localStorage.setItem(KEY, TOKEN); } catch (e) {}
    var g = document.getElementById('gate');
    if (g) g.remove();
    document.documentElement.classList.remove('gate-on');
    var h = document.getElementById('gate-hider');
    if (h) h.remove();
  }

  function build() {
    var style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);

    var gate = document.createElement('div');
    gate.id = 'gate';
    gate.innerHTML = HTML;
    document.body.appendChild(gate);
    document.documentElement.classList.add('gate-on');

    /* Offset the gate by the real header height, whatever the viewport. */
    var head = document.querySelector('body > header');
    function fit() {
      var h = head ? head.offsetHeight : 0;
      document.documentElement.style.setProperty('--gate-top', h + 'px');
    }
    fit();
    window.addEventListener('resize', fit);

    var prevTitle = document.title;
    document.title = 'Private — Stanlyn Lu';

    var form = gate.querySelector('form');
    var input = gate.querySelector('input');
    var err = gate.querySelector('.err');
    var box = gate.querySelector('.box');
    input.focus();

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      check(input.value.trim().toLowerCase()).then(function (ok) {
        if (ok) {
          document.title = prevTitle;
          unlock();
        } else {
          err.classList.add('on');
          box.classList.remove('shake');
          void box.offsetWidth;
          box.classList.add('shake');
          input.select();
        }
      });
    });

    input.addEventListener('input', function () { err.classList.remove('on'); });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', build);
  } else {
    build();
  }
})();
