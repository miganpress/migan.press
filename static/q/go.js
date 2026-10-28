/*
 * 미간 QR 리디렉트: 스캔을 구글 애널리틱스에 기록하고 목적지로 넘깁니다.
 * 목적지는 여기서 정하지 않습니다. 시트 'QR' 탭에서 고친 뒤 '사이트 업데이트'를 누르세요.
 * (build.py가 시트를 읽어 /q/주소.html 페이지마다 목적지를 넣어 둡니다.)
 */
var GA_ID = "G-ZNCEWNR0QF"; // build.py와 같은 값

(function () {
  var a = document.getElementById("fallback");
  if (!a) return;
  var id = a.getAttribute("data-qr") || "";
  var target = a.href;

  var done = false;
  function go() {
    if (done) return;
    done = true;
    location.replace(target);
  }

  // GA가 막히거나 늦어도 0.8초 뒤에는 무조건 이동
  setTimeout(go, 800);

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
