<%
'=========================================================================
' PoliMusic :: song_delete.asp  (solo administradores)
' Confirmacion y eliminacion de una cancion. Elimina en transaccion las
' dependencias (TBL_PLAYLIST_SONG, TBL_TRACKLIST, TBL_PLAYER).
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call RequireAdmin()

Dim songId, song, errMsg
songId = Trim(Request.QueryString("id") & "")
errMsg = ""

If Not IsNumeric(songId) Then
    SetFlash "Identificador de cancion invalido.", "error"
    Response.Redirect "index.asp"
    Response.End
End If

Set song = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH FROM TBL_SONG WHERE ID_SONG = ?", Array(CLng(songId)))
If song.EOF Then
    song.Close
    SetFlash "La cancion no existe.", "error"
    Response.Redirect "index.asp"
    Response.End
End If

Dim sName, sPath
sName = Trim(song("SONG_NAME") & "")
sPath = Trim(song("SONG_PATH") & "")
song.Close
Set song = Nothing

'--- Eliminar (POST) -------------------------------------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    On Error Resume Next
    conn.BeginTrans
    If Err.Number <> 0 Then
        errMsg = "No se pudo iniciar la transaccion: " & Err.Description
        Err.Clear
    Else
        DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array(CLng(songId))
        DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(CLng(songId))
        DbExec "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(CLng(songId))
        DbExec "DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(CLng(songId))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            errMsg = "Error al eliminar la cancion: " & Err.Description
            Err.Clear
        Else
            conn.CommitTrans
            On Error GoTo 0
            SetFlash "Cancion eliminada: " & sName & ".", "ok"
            Response.Redirect "index.asp"
            Response.End
        End If
    End If
    On Error GoTo 0
End If

Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = "Eliminar cancion"
ACTIVE_NAV = "songs"
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title">Eliminar cancion</h2>
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
          <span>Se eliminara la cancion de TBL_SONG y todas sus referencias en listas de reproduccion, tracklists y artistas.</span>
        </div>

        <div class="confirm-object">
          <img src="<%= Attr(ResourceUrl(sPath, DefaultCoverUrl())) %>" alt="">
          <div>
            <div class="conf-name"><%= H(sName) %></div>
            <div class="conf-meta"><%= H(sPath) %></div>
          </div>
        </div>

        <form method="post" action="song_delete.asp?id=<%= songId %>" data-guard>
          <div class="form-actions">
            <button class="btn btn-danger" type="submit">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"/><path d="M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
              Si, eliminar cancion
            </button>
            <a class="btn btn-ghost" href="index.asp">Cancelar</a>
          </div>
        </form>
      </div>

<!--#include file="includes/layout_foot.asp"-->
