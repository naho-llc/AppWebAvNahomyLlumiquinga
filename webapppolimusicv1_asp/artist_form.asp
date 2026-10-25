<%
'=========================================================================
' PoliMusic :: artist_form.asp  (solo administradores)
' Crear / editar un artista (TBL_ARTIST).
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Call RequireAdmin()

Dim action, artistId, fName, fLast, fImg, errMsg
action   = LCase(Trim(Request.QueryString("action") & ""))
If action <> "edit" Then action = "new"
artistId = Trim(Request.QueryString("id") & "")
fName = Trim(Request.Form("name") & "")
fLast = Trim(Request.Form("lastname") & "")
fImg  = Trim(Request.Form("image") & "")
errMsg = ""

If action = "edit" Then
    If Not IsNumeric(artistId) Then
        SetFlash "Artista no encontrado.", "error"
        Response.Redirect "artists.asp"
        Response.End
    End If
    Dim rs
    Set rs = DbQuery("SELECT ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(CLng(artistId)))
    If rs.EOF Then
        rs.Close
        SetFlash "El artista solicitado no existe.", "error"
        Response.Redirect "artists.asp"
        Response.End
    End If
    If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
        fName = Trim(rs("ARTIST_NAME") & "")
        fLast = Trim(rs("ARTIST_LASTNAME") & "")
        fImg  = Trim(rs("ARTIST_IMAGE") & "")
    End If
    rs.Close
    Set rs = Nothing
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim errs
    errs = ""
    If fName = "" Then errs = errs & "<li>El nombre del artista es obligatorio.</li>"
    If Len(fName) > 100 Then errs = errs & "<li>El nombre no puede exceder 100 caracteres.</li>"
    If Len(fLast) > 100 Then errs = errs & "<li>El apellido no puede exceder 100 caracteres.</li>"
    If Len(fImg) > 255 Then errs = errs & "<li>La ruta de imagen no puede exceder 255 caracteres.</li>"

    If errs = "" Then
        On Error Resume Next
        If action = "edit" And IsNumeric(artistId) Then
            DbExec "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", _
                   Array(fName, fLast, fImg, CLng(artistId))
        Else
            DbExec "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)", _
                   Array(fName, fLast, fImg)
        End If
        If Err.Number <> 0 Then
            errMsg = "Error al guardar el artista: " & Err.Description
            Err.Clear
        Else
            On Error GoTo 0
            If action = "edit" Then
                SetFlash "Artista actualizado correctamente.", "ok"
            Else
                SetFlash "Artista creado correctamente.", "ok"
            End If
            Response.Redirect "artists.asp"
            Response.End
        End If
        On Error GoTo 0
    Else
        errMsg = "<ul style='margin:0;padding-left:18px'>" & errs & "</ul>"
    End If
End If

Dim PAGE_TITLE, ACTIVE_NAV
PAGE_TITLE = IIf(action = "edit", "Editar artista", "Nuevo artista")
ACTIVE_NAV = "artists"
%>
<!--#include file="includes/layout.asp"-->

      <div class="section-head">
        <div>
          <h2 class="section-title"><%= H(PAGE_TITLE) %></h2>
          <p class="section-sub">TBL_ARTIST</p>
        </div>
        <div class="section-actions">
          <a class="btn btn-ghost" href="artists.asp">Volver a artistas</a>
        </div>
      </div>

      <% If errMsg <> "" Then %>
      <div class="alert alert-error" role="alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/></svg>
        <span><%= errMsg %></span>
      </div>
      <% End If %>

      <form method="post" action="artist_form.asp?action=<%= action %><% If action = "edit" Then %>&amp;id=<%= artistId %><% End If %>" class="form-card" data-guard>
        <div class="form-grid">
          <input type="hidden" name="id" value="<%= Attr(artistId) %>">

          <div class="field">
            <label class="field-label" for="name">Nombre <span class="req">*</span></label>
            <input class="input" type="text" id="name" name="name" maxlength="100" required
                   value="<%= Attr(fName) %>" placeholder="Ej: Juan">
          </div>

          <div class="field">
            <label class="field-label" for="lastname">Apellido</label>
            <input class="input" type="text" id="lastname" name="lastname" maxlength="100"
                   value="<%= Attr(fLast) %>" placeholder="Ej: Gabriel">
          </div>

          <div class="field">
            <label class="field-label" for="image">Ruta de imagen</label>
            <input class="input" type="text" id="image" name="image" maxlength="255" list="imgList"
                   value="<%= Attr(fImg) %>" placeholder="~/ArtistImages/artist1.jpg">
            <datalist id="imgList">
            <%
            Dim imgs, n
            imgs = ListResourceFiles("UserImages", "jpg")
            If Not IsEmpty(imgs) Then
                For n = LBound(imgs) To UBound(imgs)
            %>
              <option value="~/userImages/<%= Attr(imgs(n)) %>"></option>
            <%
                Next
            End If
            imgs = ListResourceFiles("UserImages", "png")
            If Not IsEmpty(imgs) Then
                For n = LBound(imgs) To UBound(imgs)
            %>
              <option value="~/userImages/<%= Attr(imgs(n)) %>"></option>
            <%
                Next
            End If
            %>
            </datalist>
            <p class="field-hint">
              Formato <code>~/ArtistImages/archivo.jpg</code>. Si el archivo no existe se muestra una imagen por defecto.
              <% If fImg <> "" Then %>
                <% If ResourceUrl(fImg, "") <> "" Then %>
                  <span style="color:var(--ok)">&#10003; Imagen encontrada.</span>
                <% Else %>
                  <span style="color:var(--err)">&#9888; Imagen no encontrada; se usara la de por defecto.</span>
                <% End If %>
              <% End If %>
            </p>
          </div>

          <div class="form-actions">
            <button class="btn btn-primary" type="submit">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8"/><path d="M7 3v5h8"/></svg>
              Guardar artista
            </button>
            <a class="btn btn-ghost" href="artists.asp">Cancelar</a>
          </div>
        </div>
      </form>

<!--#include file="includes/layout_foot.asp"-->
