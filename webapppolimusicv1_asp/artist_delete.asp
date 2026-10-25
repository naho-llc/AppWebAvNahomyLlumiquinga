<%
'=========================================================================
' PoliMusic :: artist_delete.asp  (solo administradores)
' Confirmacion y eliminacion de un artista (TBL_ARTIST).
' Elimina primero las referencias en TBL_PLAYER.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call RequireAdmin()

Dim artistId, artist, errMsg, songCount
artistId = Trim(Request.QueryString("id") & "")
errMsg = ""

If Not IsNumeric(artistId) Then
    SetFlash "Identificador de artista invalido.", "error"
    Response.Redirect "artists.asp"
    Response.End
End If

Set artist = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(CLng(artistId)))
If artist.EOF Then
    artist.Close
    SetFlash "El artista no existe.", "error"
    Response.Redirect "artists.asp"
    Response.End
End If

Dim aName, aLast, aImg
aName = Trim(artist("ARTIST_NAME") & "")
aLast = Trim(artist("ARTIST_LASTNAME") & "")
aImg  = ResourceUrl(artist("ARTIST_IMAGE") & "", DefaultAvatarUrl())
songCount = DbScalar("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(CLng(artistId)))
If IsNull(songCount) Then songCount = 0
artist.Close
Set artist = Nothing

Dim aFull
aFull = Trim(aName & " " & aLast)
If aFull = "" Then aFull = "Sin nombre"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    On Error Resume Next
    conn.BeginTrans
    If Err.Number <> 0 Then
        errMsg = "No se pudo iniciar la transaccion: " & Err.Description
        Err.Clear
    Else
        DbExec "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(CLng(artistId))
        DbExec "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(CLng(artistId))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            errMsg = "Error al eliminar el artista: " & Err.Description
            Err.Clear
        Else
            conn.CommitTrans
            On Error GoTo 0
            SetFlash "Artista eliminado: " & aFull & ".", "ok"
            Response.Redirect "artists.asp"
            Response.End
        End If
    End If
    On Error GoTo 0
End If

Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = "Eliminar artista"
ACTIVE_NAV = "artists"
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title">Eliminar artista</h2>
          <p class="section-sub">Esta accion no se puede deshacer</p>
        </div>
      </div>

      <% If errMsg <> "" Then %>
      <div class="alert alert-error" role="alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/></svg>
        <span><%= H(errMsg) %></span>
      </div>
      <% End If %>

      <div class="form-card confirm-box">
        <div class="warn-note">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3.5 2.5 20h19L12 3.5Z"/><path d="M12 10v4.5"/><circle cx="12" cy="17.4" r=".6" fill="currentColor"/></svg>
          <span>Se eliminara el artista de TBL_ARTIST y sus <%= songCount %> referencia(s) en TBL_PLAYER. Las canciones no se eliminan.</span>
        </div>

        <div class="confirm-object">
          <img class="round" style="border-radius:50%" src="<%= Attr(aImg) %>" alt="">
          <div>
            <div class="conf-name"><%= H(aFull) %></div>
            <div class="conf-meta">ID <%= artistId %></div>
          </div>
        </div>

        <form method="post" action="artist_delete.asp?id=<%= artistId %>" data-guard>
          <div class="form-actions">
            <button class="btn btn-danger" type="submit">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"/><path d="M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
              Si, eliminar artista
            </button>
            <a class="btn btn-ghost" href="artists.asp">Cancelar</a>
          </div>
        </form>
      </div>

<!--#include file="includes/layout_foot.asp"-->
