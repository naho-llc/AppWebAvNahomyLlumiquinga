<%
'=========================================================================
' PoliMusic :: album_delete.asp  (solo administradores)
' Confirmacion y eliminacion de un album (TBL_ALBUM).
' Elimina primero las referencias en TBL_TRACKLIST.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call RequireAdmin()

Dim albumId, album, errMsg, trackCount
albumId = Trim(Request.QueryString("id") & "")
errMsg = ""

If Not IsNumeric(albumId) Then
    SetFlash "Identificador de album invalido.", "error"
    Response.Redirect "albums.asp"
    Response.End
End If

Set album = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(CLng(albumId)))
If album.EOF Then
    album.Close
    SetFlash "El album no existe.", "error"
    Response.Redirect "albums.asp"
    Response.End
End If

Dim alTitle, alYear, alImg
alTitle = Trim(album("TITLE_ALBUM") & "")
alImg   = ResourceUrl(album("ALBUM_IMAGE_PATH") & "", DefaultCoverUrl())
trackCount = DbScalar("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(CLng(albumId)))
If IsNull(trackCount) Then trackCount = 0
If IsNull(album("RELEASE_YEAR")) Then alYear = "S/f" Else alYear = CStr(album("RELEASE_YEAR"))
album.Close
Set album = Nothing

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    On Error Resume Next
    conn.BeginTrans
    If Err.Number <> 0 Then
        errMsg = "No se pudo iniciar la transaccion: " & Err.Description
        Err.Clear
    Else
        DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(CLng(albumId))
        DbExec "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(CLng(albumId))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            errMsg = "Error al eliminar el album: " & Err.Description
            Err.Clear
        Else
            conn.CommitTrans
            On Error GoTo 0
            SetFlash "Album eliminado: " & alTitle & ".", "ok"
            Response.Redirect "albums.asp"
            Response.End
        End If
    End If
    On Error GoTo 0
End If

Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = "Eliminar album"
ACTIVE_NAV = "albums"
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title">Eliminar album</h2>
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
          <span>Se eliminara el album de TBL_ALBUM y sus <%= trackCount %> referencia(s) en TBL_TRACKLIST. Las canciones no se eliminan.</span>
        </div>

        <div class="confirm-object">
          <img src="<%= Attr(alImg) %>" alt="">
          <div>
            <div class="conf-name"><%= H(alTitle) %></div>
            <div class="conf-meta"><%= H(alYear) %></div>
          </div>
        </div>

        <form method="post" action="album_delete.asp?id=<%= albumId %>" data-guard>
          <div class="form-actions">
            <button class="btn btn-danger" type="submit">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"/><path d="M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
              Si, eliminar album
            </button>
            <a class="btn btn-ghost" href="albums.asp">Cancelar</a>
          </div>
        </form>
      </div>

<!--#include file="includes/layout_foot.asp"-->
