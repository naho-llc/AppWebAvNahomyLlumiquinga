<%
'=========================================================================
' PoliMusic :: album_form.asp  (solo administradores)
' Crear / editar un album (TBL_ALBUM) y su tracklist (TBL_TRACKLIST).
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call RequireAdmin()

Dim action, albumId, fTitle, fYear, fImg, fSongs, errMsg
action  = LCase(Trim(Request.QueryString("action") & ""))
If action <> "edit" Then action = "new"
albumId = Trim(Request.QueryString("id") & "")
fTitle = Trim(Request.Form("title") & "")
fYear  = Trim(Request.Form("year") & "")
fImg   = Trim(Request.Form("image") & "")
fSongs = Request.Form("songs") & ""
errMsg = ""

Dim selSongs
Set selSongs = Server.CreateObject("Scripting.Dictionary")

If action = "edit" Then
    If Not IsNumeric(albumId) Then
        SetFlash "Album no encontrado.", "error"
        Response.Redirect "albums.asp"
        Response.End
    End If
    Dim rs
    Set rs = DbQuery("SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(CLng(albumId)))
    If rs.EOF Then
        rs.Close
        SetFlash "El album solicitado no existe.", "error"
        Response.Redirect "albums.asp"
        Response.End
    End If
    If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
        fTitle = Trim(rs("TITLE_ALBUM") & "")
        If Not IsNull(rs("RELEASE_YEAR")) Then fYear = CStr(rs("RELEASE_YEAR"))
        fImg   = Trim(rs("ALBUM_IMAGE_PATH") & "")
    End If
    rs.Close
    Set rs = Nothing

    Dim rs2
    Set rs2 = DbQuery("SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(CLng(albumId)))
    Do Until rs2.EOF
        selSongs.Add CStr(rs2("ID_SONG")), 1
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim errs, parts, k, yearVal
    errs = ""
    If fTitle = "" Then errs = errs & "<li>El titulo del album es obligatorio.</li>"
    If Len(fTitle) > 50 Then errs = errs & "<li>El titulo no puede exceder 50 caracteres.</li>"
    If fYear <> "" Then
        If Not IsNumeric(fYear) Then
            errs = errs & "<li>El ano de lanzamiento debe ser numerico.</li>"
        ElseIf CLng(fYear) < 1900 Or CLng(fYear) > 2100 Then
            errs = errs & "<li>El ano de lanzamiento debe estar entre 1900 y 2100.</li>"
        End If
    End If
    If Len(fImg) > 255 Then errs = errs & "<li>La ruta de imagen no puede exceder 255 caracteres.</li>"

    If errs = "" Then
        yearVal = Null
        If fYear <> "" Then yearVal = CLng(fYear)

        ' Reflejar casillas marcadas al re-renderizar
        parts = Split(fSongs, ",")
        For k = LBound(parts) To UBound(parts)
            If Trim(parts(k)) <> "" Then selSongs(CStr(CLng(parts(k)))) = 1
        Next

        On Error Resume Next
        conn.BeginTrans
        If Err.Number <> 0 Then
            errMsg = "No se pudo iniciar la transaccion: " & Err.Description
            Err.Clear
        Else
            If action = "edit" And IsNumeric(albumId) Then
                DbExec "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", _
                       Array(fTitle, yearVal, fImg, CLng(albumId))
            Else
                albumId = CStr(DbInsertReturningId("INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) OUTPUT INSERTED.ID_ALBUM VALUES (?, ?, ?)", Array(fTitle, yearVal, fImg)))
            End If

            If Err.Number = 0 And IsNumeric(albumId) Then
                DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(CLng(albumId))
                parts = Split(fSongs, ",")
                For k = LBound(parts) To UBound(parts)
                    If Trim(parts(k)) <> "" Then
                        DbExec "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(CLng(albumId), CLng(parts(k)))
                    End If
                Next
            End If

            If Err.Number <> 0 Then
                conn.RollbackTrans
                errMsg = "Error al guardar el album: " & Err.Description
                Err.Clear
            Else
                conn.CommitTrans
                On Error GoTo 0
                If action = "edit" Then
                    SetFlash "Album actualizado correctamente.", "ok"
                Else
                    SetFlash "Album creado correctamente.", "ok"
                End If
                Response.Redirect "albums.asp"
                Response.End
            End If
        End If
        On Error GoTo 0
    Else
        errMsg = "<ul style='margin:0;padding-left:18px'>" & errs & "</ul>"
    End If
End If

Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = IIf(action = "edit", "Editar album", "Nuevo album")
ACTIVE_NAV = "albums"
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title"><%= H(PAGE_TITLE) %></h2>
          <p class="section-sub">TBL_ALBUM &middot; TBL_TRACKLIST</p>
        </div>
        <div class="section-actions">
          <a class="btn btn-ghost" href="albums.asp">Volver a albumes</a>
        </div>
      </div>

      <% If errMsg <> "" Then %>
      <div class="alert alert-error" role="alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/></svg>
        <span><%= errMsg %></span>
      </div>
      <% End If %>

      <form method="post" action="album_form.asp?action=<%= action %><% If action = "edit" Then %>&amp;id=<%= albumId %><% End If %>" class="form-card" data-guard>
        <div class="form-grid">
          <input type="hidden" name="id" value="<%= Attr(albumId) %>">

          <div class="field">
            <label class="field-label" for="title">Titulo del album <span class="req">*</span></label>
            <input class="input" type="text" id="title" name="title" maxlength="50" required
                   value="<%= Attr(fTitle) %>" placeholder="Ej: Rock Album">
          </div>

          <div class="field">
            <label class="field-label" for="year">Ano de lanzamiento</label>
            <input class="input" type="number" id="year" name="year" min="1900" max="2100" step="1"
                   value="<%= Attr(fYear) %>" placeholder="Ej: 1976">
          </div>

          <div class="field">
            <label class="field-label" for="image">Ruta de la portada</label>
            <input class="input" type="text" id="image" name="image" maxlength="255" list="imgList"
                   value="<%= Attr(fImg) %>" placeholder="~/albumImages/album1.png">
            <datalist id="imgList">
            <%
            Dim imgs, n
            imgs = ListResourceFiles("PlaylistImages", "png")
            If Not IsEmpty(imgs) Then
                For n = LBound(imgs) To UBound(imgs)
            %>
              <option value="~/playlistImages/<%= Attr(imgs(n)) %>"></option>
            <%
                Next
            End If
            imgs = ListResourceFiles("PlaylistImages", "jpg")
            If Not IsEmpty(imgs) Then
                For n = LBound(imgs) To UBound(imgs)
            %>
              <option value="~/playlistImages/<%= Attr(imgs(n)) %>"></option>
            <%
                Next
            End If
            %>
            </datalist>
            <p class="field-hint">
              Formato <code>~/albumImages/album1.png</code>. Si el archivo no existe se muestra la portada por defecto.
              <% If fImg <> "" Then %>
                <% If ResourceUrl(fImg, "") <> "" Then %>
                  <span style="color:var(--ok)">&#10003; Imagen encontrada.</span>
                <% Else %>
                  <span style="color:var(--err)">&#9888; Imagen no encontrada; se usara la de por defecto.</span>
                <% End If %>
              <% End If %>
            </p>
          </div>

          <div class="field">
            <span class="field-label">Canciones del album (TBL_TRACKLIST)</span>
            <div class="check-grid">
            <%
            Dim songRS, sgId, sgName, sgAvail
            Set songRS = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH FROM TBL_SONG ORDER BY SONG_NAME", Null)
            Do Until songRS.EOF
                sgId   = songRS("ID_SONG")
                sgName = Trim(songRS("SONG_NAME") & "")
                sgAvail = IIf(SongUrl(songRS("SONG_PATH") & "") <> "", "", " (sin MP3)")
            %>
              <label class="check-item">
                <input type="checkbox" name="songs" value="<%= sgId %>" <%= IIf(selSongs.Exists(CStr(sgId)), "checked", "") %>>
                <span><%= H(sgName & sgAvail) %></span>
              </label>
            <%
                songRS.MoveNext
            Loop
            songRS.Close
            Set songRS = Nothing
            %>
            </div>
            <p class="field-hint">Marca las canciones que forman parte del album.</p>
          </div>

          <div class="form-actions">
            <button class="btn btn-primary" type="submit">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8"/><path d="M7 3v5h8"/></svg>
              Guardar album
            </button>
            <a class="btn btn-ghost" href="albums.asp">Cancelar</a>
          </div>
        </div>
      </form>

<!--#include file="includes/layout_foot.asp"-->
