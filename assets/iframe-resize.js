/* Excel.Flo – meldet die Seitenhöhe per postMessage an eine einbettende Elternseite
 * (z. B. Ablefy-Lektion mit iframe), damit dort kein unnötiger Scrollbalken entsteht.
 * Läuft nur, wenn die Seite tatsächlich in einem iframe angezeigt wird; hat sonst keinen Effekt.
 */

(function () {
  "use strict";

  if (window.self === window.top) return; // nicht eingebettet

  function reportHeight() {
    const height = document.documentElement.scrollHeight;
    window.parent.postMessage({ type: "excelflo:resize", height }, "*");
  }

  if (window.ResizeObserver) {
    new ResizeObserver(reportHeight).observe(document.documentElement);
  } else {
    window.addEventListener("load", reportHeight);
    setInterval(reportHeight, 1000);
  }

  window.addEventListener("load", reportHeight);
  document.addEventListener("DOMContentLoaded", reportHeight);
})();
