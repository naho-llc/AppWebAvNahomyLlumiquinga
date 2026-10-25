<%
'=========================================================================
' PoliMusic :: artists.asp
' Listado de artistas (TBL_ARTIST). Los administradores pueden crear,
' editar y eliminar.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = "Artistas"
ACTIVE_NAV = "artists"

Dim rs
Set rs = DbQuery("SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
                 "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
                 "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME", Null)
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title">Artistas</h2>
          <p class="section-sub">Quienes hacen sonar PoliMusic</p>
        </div>
        <div class="section-actions">
          <% If IsAdmin() Then %>
          <a class="btn btn-primary" href="artist_form.asp?action=new">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
            Nuevo artista
          </a>
          <% End If %>
        </div>
      </div>

      <div class="card-grid">
      <%
      Dim any
      any = False
      Do Until rs.EOF
          any = True
          Dim aId, aName, aLast, aImg, aCount, aFull
          aId    = rs("ID_ARTIST")
          aName  = Trim(rs("ARTIST_NAME") & "")
          aLast  = Trim(rs("ARTIST_LASTNAME") & "")
          aImg   = ResourceUrl(rs("ARTIST_IMAGE") & "", DefaultAvatarUrl())
          aCount = rs("SONG_COUNT")
          If IsNull(aCount) Then aCount = 0
          aFull  = Trim(aName & " " & aLast)
          If aFull = "" Then aFull = "Sin nombre"
      %>
        <article class="entity-card">
          <img class="entity-cover round" src="<%= Attr(aImg) %>" alt="Foto de <%= Attr(aFull) %>" loading="lazy">
          <h3 class="entity-name"><%= H(aFull) %></h3>
          <p class="entity-meta"><%= aCount %> cancion<%= IIf(aCount = 1, "", "es") %> en el catalogo</p>
          <% If IsAdmin() Then %>
          <div class="entity-actions">
            <a class="icon-btn accent" href="artist_form.asp?action=edit&amp;id=<%= aId %>" title="Editar artista" aria-label="Editar">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3.5a2.1 2.1 0 0 1 3 3L8 18.5l-4 1 1-4L17 3.5Z"/></svg>
            </a>
            <a class="icon-btn danger" href="artist_delete.asp?id=<%= aId %>" title="Eliminar artista" aria-label="Eliminar">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"/><path d="M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
            </a>
          </div>
          <% End If %>
        </article>
      <%
          rs.MoveNext
      Loop
      rs.Close
      Set rs = Nothing
      %>
      </div>

      <% If Not any Then %>
      <div class="empty-state">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4.5 21c1.4-3.6 4.2-5.4 7.5-5.4s6.1 1.8 7.5 5.4"/></svg>
        <h3>No hay artistas registrados</h3>
        <p>El administrador puede agregar artistas desde el boton "Nuevo artista".</p>
      </div>
      <% End If %>

<!--#include file="includes/layout_foot.asp"-->
