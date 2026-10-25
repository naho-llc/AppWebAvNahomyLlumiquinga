/* =====================================================================
   PoliMusic :: player.js
   Reproductor HTML5 persistente: lista desde la tabla de canciones,
   controles play/pause, anterior/siguiente, barra de progreso,
   volumen y registro de reproducciones (play.asp).
   ===================================================================== */
(function () {
  "use strict";

  var audio   = document.getElementById("audio");
  var tbody   = document.getElementById("songTableBody");
  if (!audio || !tbody) return;

  var elCover  = document.getElementById("pCover");
  var elTitle  = document.getElementById("pTitle");
  var elArtist = document.getElementById("pArtist");
  var btnPlay  = document.getElementById("btnPlay");
  var btnPrev  = document.getElementById("btnPrev");
  var btnNext  = document.getElementById("btnNext");
  var btnMute  = document.getElementById("btnMute");
  var seek     = document.getElementById("seek");
  var vol      = document.getElementById("vol");
  var tCur     = document.getElementById("tCur");
  var tDur     = document.getElementById("tDur");

  var ICON_PLAY  = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5.2v13.6c0 .8.9 1.3 1.6.9l10.8-6.8c.6-.4.6-1.4 0-1.8L9.6 4.3C8.9 3.9 8 4.4 8 5.2Z"/></svg>';
  var ICON_PAUSE = '<svg viewBox="0 0 24 24" fill="currentColor"><rect x="6.5" y="4.5" width="4.4" height="15" rx="1.4"/><rect x="13.1" y="4.5" width="4.4" height="15" rx="1.4"/></svg>';
  var ICON_PREV  = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M6 5.5a1 1 0 0 1 2 0v13a1 1 0 0 1-2 0v-13Z"/><path d="M19 5.8v12.4c0 .8-.9 1.3-1.6.9L8.9 13.4a1.1 1.1 0 0 1 0-1.8l8.5-5.7c.7-.4 1.6.1 1.6.9Z"/></svg>';
  var ICON_NEXT  = '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M18 5.5a1 1 0 0 1 2 0v13a1 1 0 0 1-2 0v-13Z"/><path d="M5 5.8v12.4c0 .8.9 1.3 1.6.9l8.5-5.7a1.1 1.1 0 0 0 0-1.8L6.6 4.9c-.7-.4-1.6.1-1.6.9Z"/></svg>';
  var ICON_VOL   = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M11 5 6.5 9H3v6h3.5L11 19V5Z"/><path d="M15 9.5a3.5 3.5 0 0 1 0 5"/><path d="M17.5 7a7 7 0 0 1 0 10"/></svg>';
  var ICON_MUTE  = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M11 5 6.5 9H3v6h3.5L11 19V5Z"/><path d="m16 10 5 4"/><path d="m21 10-5 4"/></svg>';

  /* ---------- Lista de reproduccion desde el DOM ---------- */
  var playlist = [];
  var rows = tbody.querySelectorAll("tr.song-row");
  Array.prototype.forEach.call(rows, function (row, i) {
    var btn = row.querySelector(".row-play");
    playlist.push({
      index: i,
      row: row,
      src: row.getAttribute("data-src") || "",
      title: row.getAttribute("data-title") || "Pista",
      artist: row.getAttribute("data-artist") || "Artista desconocido",
      cover: row.getAttribute("data-cover") || "assets/img/default-cover.svg",
      id: row.getAttribute("data-id") || "0",
      button: btn
    });
  });
  playlist = playlist.filter(function (t) { return t.src !== ""; });

  var current = -1;
  var seeking = false;

  function fmt(sec) {
    if (!isFinite(sec) || sec < 0) sec = 0;
    var m = Math.floor(sec / 60);
    var s = Math.floor(sec % 60);
    return m + ":" + (s < 10 ? "0" : "") + s;
  }

  function toast(msg, type) {
    if (window.PMToast) window.PMToast(msg, type || "error");
  }

  function clearPlaying() {
    playlist.forEach(function (t) { t.row.classList.remove("playing"); });
  }

  function loadTrack(i, autoplay) {
    if (i < 0 || i >= playlist.length) return;
    current = i;
    var t = playlist[i];
    clearPlaying();
    t.row.classList.add("playing");
    elCover.src = t.cover;
    elCover.alt = t.title;
    elTitle.textContent = t.title;
    elArtist.textContent = t.artist;
    audio.src = t.src;
    audio.load();
    if (autoplay) {
      play();
    } else {
      btnPlay.innerHTML = ICON_PLAY;
    }
    seek.value = 0;
    tCur.textContent = "0:00";
    tDur.textContent = "0:00";
  }

  function play() {
    if (current < 0) {
      if (playlist.length > 0) loadTrack(0, true);
      return;
    }
    var p = audio.play();
    if (p && typeof p.catch === "function") {
      p.catch(function () {
        toast("No se pudo reproducir el archivo de audio.", "error");
      });
    }
  }

  function toggle() {
    if (current < 0) { play(); return; }
    if (audio.paused) play();
    else audio.pause();
  }

  function next() {
    if (playlist.length === 0) return;
    loadTrack((current + 1) % playlist.length, true);
  }

  function prev() {
    if (playlist.length === 0) return;
    if (audio.currentTime > 3) { audio.currentTime = 0; return; }
    loadTrack((current - 1 + playlist.length) % playlist.length, true);
  }

  function registerPlay(id) {
    if (!id || id === "0") return;
    try {
      var xhr = new XMLHttpRequest();
      xhr.open("GET", "play.asp?id=" + encodeURIComponent(id), true);
      xhr.send(null);
    } catch (e) { /* silencioso: el contador es best-effort */ }
  }

  /* ---------- Eventos de audio ---------- */
  audio.addEventListener("play", function () {
    btnPlay.innerHTML = ICON_PAUSE;
  });
  audio.addEventListener("pause", function () {
    btnPlay.innerHTML = ICON_PLAY;
  });
  audio.addEventListener("ended", next);
  audio.addEventListener("error", function () {
    toast("Archivo de audio no disponible: " + (playlist[current] ? playlist[current].title : ""), "error");
  });
  audio.addEventListener("loadedmetadata", function () {
    tDur.textContent = fmt(audio.duration);
    seek.max = isFinite(audio.duration) ? Math.floor(audio.duration) : 100;
  });
  audio.addEventListener("timeupdate", function () {
    if (!seeking && isFinite(audio.duration)) {
      seek.value = Math.floor(audio.currentTime);
      tCur.textContent = fmt(audio.currentTime);
    }
  });

  seek.addEventListener("input", function () { seeking = true; });
  seek.addEventListener("change", function () {
    if (isFinite(audio.duration)) audio.currentTime = Number(seek.value);
    seeking = false;
  });

  vol.addEventListener("input", function () {
    audio.volume = Number(vol.value) / 100;
    audio.muted = audio.volume === 0;
    btnMute.innerHTML = audio.muted ? ICON_MUTE : ICON_VOL;
  });

  btnMute.addEventListener("click", function () {
    audio.muted = !audio.muted;
    btnMute.innerHTML = audio.muted ? ICON_MUTE : ICON_VOL;
  });

  btnPlay.addEventListener("click", toggle);
  btnPrev.addEventListener("click", prev);
  btnNext.addEventListener("click", next);

  /* ---------- Clic en filas / botones de la tabla ---------- */
  Array.prototype.forEach.call(rows, function (row) {
    var idx = -1;
    playlist.forEach(function (t, i) { if (t.row === row) idx = i; });

    row.addEventListener("click", function (ev) {
      if (ev.target.closest("a, button, input, .no-play")) return;
      if (idx < 0) {
        toast("Esta cancion no tiene archivo de audio disponible.", "error");
        return;
      }
      if (idx === current && !audio.paused) {
        audio.pause();
      } else if (idx === current && audio.paused) {
        play();
      } else {
        loadTrack(idx, true);
        registerPlay(playlist[idx].id);
      }
    });

    var btn = row.querySelector(".row-play");
    if (btn) {
      btn.addEventListener("click", function (ev) {
        ev.stopPropagation();
        if (idx < 0) {
          toast("Archivo MP3 no encontrado en la carpeta de recursos.", "error");
          return;
        }
        if (idx === current && !audio.paused) {
          audio.pause();
        } else if (idx === current && audio.paused) {
          play();
        } else {
          loadTrack(idx, true);
          registerPlay(playlist[idx].id);
        }
      });
    }
  });

  /* ---------- Atajos de teclado ---------- */
  document.addEventListener("keydown", function (ev) {
    var tag = (ev.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea" || tag === "select") return;
    if (ev.code === "Space") { ev.preventDefault(); toggle(); }
    else if (ev.code === "ArrowRight" && ev.shiftKey) { next(); }
    else if (ev.code === "ArrowLeft" && ev.shiftKey) { prev(); }
  });

  /* ---------- Estado inicial ---------- */
  btnPlay.innerHTML = ICON_PLAY;
  btnPrev.innerHTML = ICON_PREV;
  btnNext.innerHTML = ICON_NEXT;
  btnMute.innerHTML = ICON_VOL;
  audio.volume = Number(vol.value) / 100;
  if (playlist.length === 0) {
    elTitle.textContent = "Sin canciones disponibles";
    elArtist.textContent = "Agrega canciones para comenzar";
  }

  window.PMPlayer = { loadTrack: loadTrack, next: next, prev: prev, toggle: toggle };
})();
