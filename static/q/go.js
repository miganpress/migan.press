/*
 * 미간 QR 리디렉트
 * ------------------------------------------------------------
 * 목적지를 바꾸려면 아래 DEST 표의 주소만 고치면 됩니다.
 * 이미 인쇄된 QR도 다음 접속부터 새 목적지로 갑니다.
 *
 * 새 QR 자리를 추가하려면
 *   1) DEST에 한 줄 추가 (예: tote: "https://...")
 *   2) q/card.html을 복사해 q/tote.html로 저장
 * ------------------------------------------------------------
 */

// 구글 애널리틱스 측정 ID (build.py와 같은 값)
var GA_ID = "G-ZNCEWNR0QF";

// QR 자리별 목적지
var DEST = {
  card:  "https://www.instagram.com/migan.press/",  // 명함
  "mg001-bph": "https://www.instagram.com/migan.press/",  // 『책 공포증』 표3
  goods: "https://www.instagram.com/migan.press/"   // 굿즈
};

// 표에 없는 자리로 들어왔을 때 갈 곳
var FALLBACK = "https://miganpress.com/";

/* ---------- 아래는 고칠 필요 없음 ---------- */
(function () {
  var id = (location.pathname.split("/").pop() || "").replace(/\.html$/, "").toLowerCase();
  var target = DEST[id] || FALLBACK;

  // 자바스크립트가 안 될 때를 위한 링크도 같은 목적지로
  var a = document.getElementById("fallback");
  if (a) a.href = target;

  var done = false;
  function go() {
    if (done) return;
    done = true;
    location.replace(target);
  }

  // GA가 막히거나 늦어도 0.8초 뒤에는 무조건 이동
  setTimeout(go, 800);

  if (!GA_ID || GA_ID.indexOf("XXXX") > -1) { go(); return; }

  window.dataLayer = window.dataLayer || [];
  window.gtag = function () { dataLayer.push(arguments); };
  gtag("js", new Date());
  gtag("config", GA_ID, { send_page_view: false });
  gtag("event", "qr_scan", {
    qr_id: id,
    qr_target: target,
    transport_type: "beacon",
    event_callback: go
  });

  var s = document.createElement("script");
  s.async = true;
  s.src = "https://www.googletagmanager.com/gtag/js?id=" + GA_ID;
  s.onerror = go;
  document.head.appendChild(s);
})();
