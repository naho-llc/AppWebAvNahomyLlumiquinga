<%
'=========================================================================
' PoliMusic :: index.asp
' Pantalla principal: mensaje de bienvenida personalizado, listado de
' canciones ("Song List") y reproductor HTML5 persistente.
' Los administradores ven ademas las acciones CRUD sobre canciones.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = "Song List"
ACTIVE_NAV = "songs"

Dim songsRS, artistMap, albumMap
Set songsRS = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG ORDER BY ID_SONG", Null)
Set artistMap = LoadArtistMap()
Set albumMap = LoadAlbumMap()
%>
<!--#include file="includes/layout.asp"-->

      <%'--- Hero de bienvenida personalizado -------------------------------%>
      <section class="hero-welcome">
        <p class="hero-kicker">PoliMusic &middot; Tu biblioteca</p>
        <h2 class="hero-title">Bienvenido, <em><%= H(Session("USER_NAME")) %></em></h2>
        <p class="hero-text">
          Has iniciado sesion como <b><%= H(RoleLabel()) %></b>. Explora el listado de canciones,
          reproduce tu musica favorita y <% If IsAdmin() Then %>administra canciones, artistas y albumes.<% Else %>descubre nuevos sonidos.<% End If %>
        </p>
      </section>

      <%'--- Encabezado de seccion ----------------------------------------%>
      <div class="section-head">
        <div>
          <h2 class="section-title">Song List</h2>
          <p class="section-sub" id="songCount"></p>
        </div>
        <div class="section-actions">
          <div class="search-box">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.8-3.8"/></svg>
            <input type="search" id="songSearch" placeholder="Buscar cancion, artista o album..." aria-label="Buscar canciones">
          </div>
          <% If IsAdmin() Then %>
          <a class="btn btn-primary" href="song_form.asp?action=new">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
            Nueva cancion
          </a>
          <% End If %>
        </div>
      </div>

      <%'--- Tabla de canciones -------------------------------------------%>
      <div class="table-wrap">
        <div class="table-scroll">
          <table class="song-table">
            <thead>
              <tr>
                <th class="col-index">#</th>
                <th class="col-cover">Portada</th>
                <th class="col-title">Cancion</th>
                <th class="col-artist">Artista(s)</th>
                <th class="col-album">Album</th>
                <th class="col-plays">Reproducciones</th>
                <th class="col-actions">Acciones</th>
              </tr>
            </thead>
            <tbody id="songTableBody">
            <%
            Dim i, sId, sName, sPath, sPlays, sUrl, coverUrl, artists, albumTitle, hasAudio
            i = 0
            Do Until songsRS.EOF
                i = i + 1
                sId    = songsRS("ID_SONG")
                sName  = Trim(songsRS("SONG_NAME") & "")
                sPath  = Trim(songsRS("SONG_PATH") & "")
                sPlays = songsRS("PLAYS")
                If IsNull(sPlays) Or Not IsNumeric(sPlays) Then sPlays = 0

                sUrl    = SongUrl(sPath)
                hasAudio = (sUrl <> "")

                If albumMap.Exists(CStr(sId)) Then
                    albumTitle = albumMap(CStr(sId))(0)
                    coverUrl   = albumMap(CStr(sId))(1)
                Else
                    albumTitle = ""
                    coverUrl   = DefaultCoverUrl()
                End If
                If albumTitle = "" Then albumTitle = "Sin album"

                If artistMap.Exists(CStr(sId)) Then
                    artists = artistMap(CStr(sId))
                Else
                    artists = "Artista desconocido"
                End If
            %>
              <tr class="song-row" data-id="<%= sId %>" data-src="<%= Attr(sUrl) %>"
                  data-title="<%= Attr(sName) %>" data-artist="<%= Attr(artists) %>"
                  data-cover="<%= Attr(coverUrl) %>"
                  data-search="<%= Attr(sName & " " & artists & " " & albumTitle) %>">
                <td class="col-index" data-label="#">
                  <span class="eq" aria-hidden="true"><span></span><span></span><span></span><span></span></span>
                  <span class="row-num"><%= i %></span>
                </td>
                <td class="col-cover" data-label="Portada">
                  <img class="song-cover" src="<%= Attr(coverUrl) %>" alt="Portada de <%= Attr(sName) %>" loading="lazy">
                </td>
                <td class="col-title" data-label="Cancion">
                  <span class="song-title"><%= H(sName) %></span>
                  <span class="song-path"><%= H(sPath) %></span>
                </td>
                <td class="col-artist" data-label="Artista"><%= H(artists) %></td>
                <td class="col-album" data-label="Album"><%= H(albumTitle) %></td>
                <td class="col-plays" data-label="Reproducciones"><%= sPlays %></td>
                <td class="col-actions" data-label="Acciones">
                  <% If hasAudio Then %>
                  <button type="button" class="row-play" title="Reproducir <%= Attr(sName) %>" aria-label="Reproducir">
                    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5.2v13.6c0 .8.9 1.3 1.6.9l10.8-6.8c.6-.4.6-1.4 0-1.8L9.6 4.3C8.9 3.9 8 4.4 8 5.2Z"/></svg>
                  </button>
                  <% Else %>
                  <span class="badge-na" title="El archivo MP3 no se encuentra en la carpeta de recursos">No disponible</span>
                  <% End If %>
                  <% If IsAdmin() Then %>
                  <a class="icon-btn accent" href="song_form.asp?action=edit&amp;id=<%= sId %>" title="Editar cancion" aria-label="Editar">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3.5a2.1 2.1 0 0 1 3 3L8 18.5l-4 1 1-4L17 3.5Z"/></svg>
                  </a>
                  <a class="icon-btn danger" href="song_delete.asp?id=<%= sId %>" title="Eliminar cancion" aria-label="Eliminar">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"/><path d="M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
                  </a>
                  <% End If %>
                </td>
              </tr>
            <%
                songsRS.MoveNext
            Loop
            songsRS.Close
            Set songsRS = Nothing
            %>
            </tbody>
          </table>
        </div>
        <% If i = 0 Then %>
        <div class="empty-state">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18V6l10-2v12"/><circle cx="6.5" cy="18" r="2.5"/><circle cx="16.5" cy="16" r="2.5"/></svg>
          <h3>No hay canciones registradas</h3>
          <p>Cuando el administrador agregue canciones apareceran aqui.</p>
        </div>
        <% End If %>
        <div class="empty-state" id="searchEmpty" style="display:none">
          <h3>Sin resultados</h3>
          <p>Ninguna cancion coincide con tu busqueda.</p>
        </div>
      </div>

      <%'--- Reproductor persistente ---------------------------------------%>
      <div class="player" id="player">
        <div class="player-track">
          <img class="player-cover" id="pCover" src="assets/img/default-cover.svg" alt="">
          <div class="player-meta">
            <strong class="player-title" id="pTitle">Selecciona una cancion</strong>
            <span class="player-artist" id="pArtist">PoliMusic</span>
          </div>
        </div>
        <div class="player-controls">
          <div class="ctrl-row">
            <button type="button" class="ctrl-btn" id="btnPrev" title="Anterior" aria-label="Cancion anterior"></button>
            <button type="button" class="ctrl-btn ctrl-play" id="btnPlay" title="Reproducir / Pausar" aria-label="Reproducir o pausar"></button>
            <button type="button" class="ctrl-btn" id="btnNext" title="Siguiente" aria-label="Cancion siguiente"></button>
          </div>
          <div class="seek-row">
            <span class="time-label" id="tCur">0:00</span>
            <input type="range" id="seek" min="0" max="100" value="0" aria-label="Progreso de la cancion">
            <span class="time-label" id="tDur">0:00</span>
          </div>
        </div>
        <div class="player-volume">
          <button type="button" class="ctrl-btn" id="btnMute" title="Silenciar" aria-label="Silenciar"></button>
          <input type="range" id="vol" min="0" max="100" value="80" aria-label="Volumen">
        </div>
        <audio id="audio" preload="none"></audio>
      </div>

      <script>
        document.getElementById("songCount").textContent =
          "<%= i %> cancion<%= IIf(i = 1, "", "es") %> en tu biblioteca";
      </script>
      <script src="js/player.js"></script>

<!--#include file="includes/layout_foot.asp"-->
