/* =====================================================================
   PoliMusic :: app.js
   Utilidades de interfaz: toasts, buscador, confirmaciones, alertas.
   ===================================================================== */
(function () {
  "use strict";

  /* ---------- Toasts ---------- */
  var stack = document.getElementById("toastStack");
  var ICONS = {
    ok: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8.5 12.5 2.4 2.4 4.6-5"/></svg>',
    error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/></svg>',
    info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><circle cx="12" cy="7.8" r=".6" fill="currentColor"/></svg>'
  };

  function toast(message, type, ttl) {
    if (!stack) return;
    type = type || "info";
    ttl = ttl || 3400;
    var el = document.createElement("div");
    el.className = "toast toast-" + type;
    el.setAttribute("role", "status");
    el.innerHTML = (ICONS[type] || ICONS.info) + "<span></span>";
    el.querySelector("span").textContent = message;
    stack.appendChild(el);
    setTimeout(function () {
      el.classList.add("hide");
      setTimeout(function () { if (el.parentNode) el.parentNode.removeChild(el); }, 320);
    }, ttl);
  }
  window.PMToast = toast;

  /* ---------- Cerrar alertas flash ---------- */
  document.querySelectorAll(".alert-close").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var alert = btn.closest(".alert");
      if (alert) {
        alert.style.transition = "opacity .25s ease, transform .25s ease";
        alert.style.opacity = "0";
        alert.style.transform = "translateY(-6px)";
        setTimeout(function () { if (alert.parentNode) alert.parentNode.removeChild(alert); }, 260);
      }
    });
  });

  /* ---------- Buscador generico (filtra filas de tabla o tarjetas) ---------- */
  var search = document.getElementById("songSearch");
  if (search) {
    search.addEventListener("input", function () {
      var q = search.value.trim().toLowerCase();
      var rows = document.querySelectorAll("#songTableBody tr.song-row");
      var visible = 0;
      rows.forEach(function (row) {
        var hay = (row.getAttribute("data-search") || "").toLowerCase();
        var show = q === "" || hay.indexOf(q) !== -1;
        row.style.display = show ? "" : "none";
        if (show) visible++;
      });
      var empty = document.getElementById("searchEmpty");
      if (empty) empty.style.display = visible === 0 ? "" : "none";
    });
  }

  /* ---------- Confirmacion de acciones destructivas ---------- */
  document.querySelectorAll("[data-confirm]").forEach(function (el) {
    el.addEventListener("click", function (ev) {
      if (!window.confirm(el.getAttribute("data-confirm"))) {
        ev.preventDefault();
      }
    });
  });

  /* ---------- Marcar enlace de navegacion activo por URL (respaldo) ---------- */
  (function markNav() {
    var page = (window.location.pathname.split("/").pop() || "index.asp").toLowerCase();
    document.querySelectorAll(".nav-item").forEach(function (a) {
      var href = (a.getAttribute("href") || "").toLowerCase();
      if (href && href.indexOf("?") === -1 && href === page) a.classList.add("active");
    });
  })();

  /* ---------- Anti doble-envio en formularios ---------- */
  document.querySelectorAll("form[data-guard]").forEach(function (form) {
    form.addEventListener("submit", function () {
      var btn = form.querySelector("button[type=submit]");
      if (btn && !btn.disabled) {
        setTimeout(function () { btn.disabled = true; btn.style.opacity = ".65"; }, 0);
      }
    });
  });
})();
