<%
'=========================================================================
' PoliMusic :: albums.asp
' Listado de albumes (TBL_ALBUM). Los administradores pueden crear,
' editar y eliminar.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = "Albumes"
ACTIVE_NAV = "albums"

Dim rs
Set rs = DbQuery("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
                 "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS TRACK_COUNT " & _
                 "FROM TBL_ALBUM al ORDER BY al.RELEASE_YEAR, al.TITLE_ALBUM", Null)
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title">Albumes</h2>
          <p class="section-sub">Colecciones completas para escuchar de principio a fin</p>
        </div>
        <div class="section-actions">
          <% If IsAdmin() Then %>
          <a class="btn btn-primary" href="album_form.asp?action=new">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
            Nuevo album
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
          Dim alId, alTitle, alYear, alImg, alCount
          alId    = rs("ID_ALBUM")
          alTitle = Trim(rs("TITLE_ALBUM") & "")
          alYear  = rs("RELEASE_YEAR")
          alImg   = ResourceUrl(rs("ALBUM_IMAGE_PATH") & "", DefaultCoverUrl())
          alCount = rs("TRACK_COUNT")
          If IsNull(alCount) Then alCount = 0
          If IsNull(alYear) Or Not IsNumeric(alYear) Then alYear = "S/f" Else alYear = CStr(alYear)
      %>
        <article class="entity-card">
          <img class="entity-cover" src="<%= Attr(alImg) %>" alt="Portada de <%= Attr(alTitle) %>" loading="lazy">
          <h3 class="entity-name"><%= H(alTitle) %></h3>
          <p class="entity-meta"><%= alYear %> &middot; <%= alCount %> cancion<%= IIf(alCount = 1, "", "es") %></p>
          <% If IsAdmin() Then %>
          <div class="entity-actions">
            <a class="icon-btn accent" href="album_form.asp?action=edit&amp;id=<%= alId %>" title="Editar album" aria-label="Editar">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3.5a2.1 2.1 0 0 1 3 3L8 18.5l-4 1 1-4L17 3.5Z"/></svg>
            </a>
            <a class="icon-btn danger" href="album_delete.asp?id=<%= alId %>" title="Eliminar album" aria-label="Eliminar">
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
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="2.6"/></svg>
        <h3>No hay albumes registrados</h3>
        <p>El administrador puede agregar albumes desde el boton "Nuevo album".</p>
      </div>
      <% End If %>

<!--#include file="includes/layout_foot.asp"-->
