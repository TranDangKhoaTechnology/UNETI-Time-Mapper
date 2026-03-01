(function () {
  "use strict";

  function applyAndSelfPolyfill(jq) {
    if (!jq || !jq.fn) return false;
    if (typeof jq.fn.andSelf === "function") return true;
    if (typeof jq.fn.addBack !== "function") return false;

    jq.fn.andSelf = jq.fn.addBack;
    return true;
  }

  function tryApply() {
    const okJq = applyAndSelfPolyfill(window.jQuery);
    const okDollar = window.$ === window.jQuery ? okJq : applyAndSelfPolyfill(window.$);
    return okJq || okDollar;
  }

  if (tryApply()) {
    return;
  }

  let attempts = 0;
  const timer = setInterval(() => {
    attempts += 1;
    const done = tryApply();
    if (done || attempts >= 200) {
      clearInterval(timer);
    }
  }, 25);
})();
