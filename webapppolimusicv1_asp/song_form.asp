<%
'=========================================================================
' PoliMusic :: song_form.asp  (solo administradores)
' Crear / editar una cancion (TBL_SONG) y sus relaciones con artistas
' (TBL_PLAYER) y albumes (TBL_TRACKLIST).
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call RequireAdmin()

Dim action, songId, fName, fPath, fArtists, fAlbums, errMsg
action  = LCase(Trim(Request.QueryString("action") & ""))
If action <> "edit" Then action = "new"
songId  = Trim(Request.QueryString("id") & "")
fName   = Trim(Request.Form("name") & "")
fPath   = Trim(Request.Form("path") & "")
fArtists = Request.Form("artists") & ""
fAlbums  = Request.Form("albums") & ""
errMsg  = ""

'--- Cargar datos existentes en modo edicion ------------------------------
Dim selArtists, selAlbums
Set selArtists = Server.CreateObject("Scripting.Dictionary")
Set selAlbums  = Server.CreateObject("Scripting.Dictionary")

If action = "edit" Then
    If Not IsNumeric(songId) Then
        SetFlash "Cancion no encontrada.", "error"
        Response.Redirect "index.asp"
        Response.End
    End If
    Dim rs
    Set rs = DbQuery("SELECT SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(CLng(songId)))
    If rs.EOF Then
        rs.Close
        SetFlash "La cancion solicitada no existe.", "error"
        Response.Redirect "index.asp"
        Response.End
    End If
    If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
        fName = Trim(rs("SONG_NAME") & "")
        fPath = Trim(rs("SONG_PATH") & "")
    End If
    rs.Close
    Set rs = Nothing

    Dim rs2
    Set rs2 = DbQuery("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", Array(CLng(songId)))
    Do Until rs2.EOF
        selArtists.Add CStr(rs2("ID_ARTIST")), 1
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    Set rs2 = DbQuery("SELECT ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(CLng(songId)))
    Do Until rs2.EOF
        selAlbums.Add CStr(rs2("ID_ALBUM")), 1
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
End If

'--- Guardar ---------------------------------------------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim errs, parts, k
    errs = ""
    If fName = "" Then errs = errs & "<li>El nombre de la cancion es obligatorio.</li>"
    If Len(fName) > 50 Then errs = errs & "<li>El nombre no puede exceder 50 caracteres.</li>"
    If fPath = "" Then errs = errs & "<li>La ruta del archivo MP3 es obligatoria.</li>"
    If Len(fPath) > 255 Then errs = errs & "<li>La ruta no puede exceder 255 caracteres.</li>"

    If errs = "" Then
        ' Reflejar las casillas marcadas al re-renderizar
        parts = Split(fArtists, ",")
        For k = LBound(parts) To UBound(parts)
            If Trim(parts(k)) <> "" Then selArtists(CStr(CLng(parts(k)))) = 1
        Next
        parts = Split(fAlbums, ",")
        For k = LBound(parts) To UBound(parts)
            If Trim(parts(k)) <> "" Then selAlbums(CStr(CLng(parts(k)))) = 1
        Next

        On Error Resume Next
        conn.BeginTrans
        If Err.Number <> 0 Then
            errMsg = "No se pudo iniciar la transaccion: " & Err.Description
            Err.Clear
        Else
            If action = "edit" And IsNumeric(songId) Then
                DbExec "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ? WHERE ID_SONG = ?", Array(fName, fPath, CLng(songId))
            Else
                songId = CStr(DbInsertReturningId("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) OUTPUT INSERTED.ID_SONG VALUES (?, ?, 0)", Array(fName, fPath)))
            End If

            If Err.Number = 0 And IsNumeric(songId) Then
                ' Reconstruir relacion cancion-artista (TBL_PLAYER)
                DbExec "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(CLng(songId))
                parts = Split(fArtists, ",")
                For k = LBound(parts) To UBound(parts)
                    If Trim(parts(k)) <> "" Then
                        DbExec "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", Array(CLng(songId), CLng(parts(k)))
                    End If
                Next
                ' Reconstruir tracklist (TBL_TRACKLIST)
                DbExec "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(CLng(songId))
                parts = Split(fAlbums, ",")
                For k = LBound(parts) To UBound(parts)
                    If Trim(parts(k)) <> "" Then
                        DbExec "INSERT INTO TBL_TRACKLIST (ID_SONG, ID_ALBUM) VALUES (?, ?)", Array(CLng(songId), CLng(parts(k)))
                    End If
                Next
            End If

            If Err.Number <> 0 Then
                conn.RollbackTrans
                errMsg = "Error al guardar en la base de datos: " & Err.Description
                Err.Clear
            Else
                conn.CommitTrans
                On Error GoTo 0
                If action = "edit" Then
                    SetFlash "Cancion actualizada correctamente.", "ok"
                Else
                    SetFlash "Cancion creada correctamente.", "ok"
                End If
                Response.Redirect "index.asp"
                Response.End
            End If
        End If
        On Error GoTo 0
    Else
        errMsg = "<ul style='margin:0;padding-left:18px'>" & errs & "</ul>"
    End If
End If

Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = IIf(action = "edit", "Editar cancion", "Nueva cancion")
ACTIVE_NAV = "songs"
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title"><%= H(PAGE_TITLE) %></h2>
          <p class="section-sub">TBL_SONG &middot; TBL_PLAYER &middot; TBL_TRACKLIST</p>
        </div>
        <div class="section-actions">
          <a class="btn btn-ghost" href="index.asp">Volver al listado</a>
        </div>
      </div>

      <% If errMsg <> "" Then %>
      <div class="alert alert-error" role="alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/></svg>
        <span><%= errMsg %></span>
      </div>
      <% End If %>

      <form method="post" action="song_form.asp?action=<%= action %><% If action = "edit" Then %>&amp;id=<%= songId %><% End If %>" class="form-card" data-guard>
        <div class="form-grid">
          <input type="hidden" name="id" value="<%= Attr(songId) %>">

          <div class="field">
            <label class="field-label" for="name">Nombre de la cancion <span class="req">*</span></label>
            <input class="input" type="text" id="name" name="name" maxlength="50" required
                   value="<%= Attr(fName) %>" placeholder="Ej: Adventure">
          </div>

          <div class="field">
            <label class="field-label" for="path">Ruta del archivo MP3 <span class="req">*</span></label>
            <input class="input" type="text" id="path" name="path" maxlength="255" required list="mp3List"
                   value="<%= Attr(fPath) %>" placeholder="~/songFiles/bensound-adventure.mp3">
            <datalist id="mp3List">
            <%
            Dim mp3s, m
            mp3s = ListResourceFiles("SongFiles", "mp3")
            If Not IsEmpty(mp3s) Then
                For m = LBound(mp3s) To UBound(mp3s)
            %>
              <option value="~/songFiles/<%= Attr(mp3s(m)) %>"></option>
            <%
                Next
            End If
            %>
            </datalist>
            <p class="field-hint">
              Usa el formato <code>~/songFiles/archivo.mp3</code>. Las sugerencias provienen de la carpeta
              <b>SongFilesAndImages_resources/SongFiles</b>.
              <% If fPath <> "" Then %>
                <% If SongUrl(fPath) <> "" Then %>
                  <span style="color:var(--ok)">&#10003; Archivo encontrado en la carpeta de recursos.</span>
                <% Else %>
                  <span style="color:var(--err)">&#9888; El archivo no se encuentra localmente; se guardara de todas formas.</span>
                <% End If %>
              <% End If %>
            </p>
          </div>

          <div class="field">
            <span class="field-label">Artistas (TBL_PLAYER)</span>
            <div class="check-grid">
            <%
            Dim artRS, aid, aname
            Set artRS = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME", Null)
            Do Until artRS.EOF
                aid = artRS("ID_ARTIST")
                aname = Trim(artRS("ARTIST_NAME") & "")
                If Trim(artRS("ARTIST_LASTNAME") & "") <> "" Then aname = aname & " " & Trim(artRS("ARTIST_LASTNAME"))
            %>
              <label class="check-item">
                <input type="checkbox" name="artists" value="<%= aid %>" <%= IIf(selArtists.Exists(CStr(aid)), "checked", "") %>>
                <span><%= H(aname) %></span>
              </label>
            <%
                artRS.MoveNext
            Loop
            artRS.Close
            Set artRS = Nothing
            %>
            </div>
            <p class="field-hint">Opcional. Marca los artistas que interpretan la cancion.</p>
          </div>

          <div class="field">
            <span class="field-label">Albumes (TBL_TRACKLIST)</span>
            <div class="check-grid">
            <%
            Dim albRS, alid, altitle
            Set albRS = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Null)
            Do Until albRS.EOF
                alid = albRS("ID_ALBUM")
                altitle = Trim(albRS("TITLE_ALBUM") & "")
            %>
              <label class="check-item">
                <input type="checkbox" name="albums" value="<%= alid %>" <%= IIf(selAlbums.Exists(CStr(alid)), "checked", "") %>>
                <span><%= H(altitle) %></span>
              </label>
            <%
                albRS.MoveNext
            Loop
            albRS.Close
            Set albRS = Nothing
            %>
            </div>
            <p class="field-hint">Opcional. Marca los albumes que incluyen la cancion.</p>
          </div>

          <div class="form-actions">
            <button class="btn btn-primary" type="submit">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8"/><path d="M7 3v5h8"/></svg>
              Guardar cancion
            </button>
            <a class="btn btn-ghost" href="index.asp">Cancelar</a>
          </div>
        </div>
      </form>

<!--#include file="includes/layout_foot.asp"-->
