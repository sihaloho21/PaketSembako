(function (global) {
  if (!global || !global.location) {
    return;
  }

  // Set endpoint API default per host agar pengguna tidak perlu ?apiBase=.
  // Isi URL Web App Apps Script sekali di sini.
  var API_BASE_BY_HOST = {
    "kasir-paketsembako.netlify.app": "",
  };

  function sanitize(value) {
    var text = value === null || value === undefined ? "" : String(value).trim();
    if (!text) {
      return "";
    }
    return text
      .replace(/[?#].*$/, "")
      .replace(/\/+$/, "");
  }

  function getMetaApiBase() {
    if (!global.document || typeof global.document.querySelector !== "function") {
      return "";
    }
    var meta = global.document.querySelector('meta[name="app-api-base"]');
    if (!meta) {
      return "";
    }
    return sanitize(meta.getAttribute("content"));
  }

  var host = String(global.location.hostname || "").toLowerCase();
  var configured = sanitize(global.APP_API_BASE);
  var fromMeta = getMetaApiBase();
  var fromHost = sanitize(API_BASE_BY_HOST[host]);
  var resolved = configured || fromMeta || fromHost;

  if (resolved) {
    global.APP_API_BASE = resolved;
  }
})(window);
