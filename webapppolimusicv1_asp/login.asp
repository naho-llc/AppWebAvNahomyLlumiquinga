<%
'=========================================================================
' PoliMusic :: login.asp
' Autenticacion de usuarios contra TBL_USER. Soporta passwords en texto
' plano y almacenados como hash SHA-1 hexadecimal.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<%
' Si ya hay sesion activa, entrar directamente.
If Trim(Session("USER_NAME") & "") <> "" Then
    Response.Redirect "index.asp"
    Response.End
End If

Dim errMsg, frmUser, frmPass
errMsg = ""
frmUser = Trim(Request.Form("username") & "")
frmPass = Trim(Request.Form("password") & "")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If frmUser = "" Or frmPass = "" Then
        errMsg = "Debes ingresar usuario y contrasena."
    ElseIf Len(frmPass) > 50 Then
        errMsg = "La contrasena no puede exceder 50 caracteres."
    Else
        Dim rs, dbPass, okAuth
        Set rs = DbQuery("SELECT ID_USER, USERNAME, PASSWORD, USER_TYPE, USER_PHOTO " & _
                         "FROM TBL_USER WHERE USERNAME = ?", Array(frmUser))
        If Not rs.EOF Then
            dbPass = rs("PASSWORD") & ""
            okAuth = False
            ' 1) Comparacion directa (password en texto plano)
            If StrComp(dbPass, frmPass, vbBinaryCompare) = 0 Then
                okAuth = True
            ' 2) Comparacion contra hash SHA-1 hexadecimal almacenado
            ElseIf UCase(dbPass) = SHA1(frmPass) Then
                okAuth = True
            End If
            If okAuth Then
                Session("USER_ID")    = rs("ID_USER")
                Session("USER_NAME")  = rs("USERNAME")
                Session("USER_TYPE")  = CInt(rs("USER_TYPE"))
                Session("USER_PHOTO") = rs("USER_PHOTO") & ""
                Session.Timeout      = 60
                rs.Close
                Set rs = Nothing
                SetFlash "Bienvenido a PoliMusic, " & frmUser & ".", "ok"
                Response.Redirect "index.asp"
                Response.End
            Else
                errMsg = "Usuario o contrasena incorrectos."
            End If
        Else
            errMsg = "El usuario no existe en la base de datos."
        End If
        If IsObject(rs) Then
            On Error Resume Next
            If rs.State = 1 Then rs.Close
            On Error GoTo 0
            Set rs = Nothing
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesion &middot; PoliMusic</title>
<link rel="icon" type="image/svg+xml" href="assets/img/favicon.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="css/styles.css">
</head>
<body class="login-page">
<div class="login-shell">

  <section class="login-brand">
    <img class="login-brand-logo" src="assets/img/logo.svg" alt="PoliMusic">
    <h1>PoliMusic</h1>
    <p class="login-tagline">Toda tu musica favorita, organizada y lista para sonar. Explora canciones, artistas y albumes desde un solo lugar.</p>
    <img class="vinyl" src="assets/img/vinyl.svg" alt="">
    <p class="login-brand-foot">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5.5" rx="7.5" ry="3"/><path d="M4.5 5.5v13c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3v-13"/><path d="M4.5 12c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3"/></svg>
      Conectado a BDD_PoliMusic
    </p>
  </section>

  <section class="login-panel">
    <div class="login-card">
      <h2>Iniciar sesion</h2>
      <p class="login-sub">Ingresa con tu cuenta de PoliMusic</p>

      <% If errMsg <> "" Then %>
      <div class="alert alert-error" role="alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/></svg>
        <span><%= H(errMsg) %></span>
      </div>
      <% End If %>

      <form method="post" action="login.asp" class="form-grid" data-guard autocomplete="on">
        <div class="field">
          <label class="field-label" for="username">Usuario</label>
          <input class="input" type="text" id="username" name="username" placeholder="tu.usuario"
                 value="<%= Attr(frmUser) %>" maxlength="50" required autofocus>
        </div>
        <div class="field">
          <label class="field-label" for="password">Contrasena</label>
          <input class="input" type="password" id="password" name="password" placeholder="********"
                 maxlength="50" required>
        </div>
        <button class="btn btn-primary btn-block" type="submit">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><path d="m10 17 5-5-5-5"/><path d="M15 12H3"/></svg>
          Entrar
        </button>
      </form>

      <div class="login-hint">
        <b>Usuarios de prueba</b><br>
        Usuario normal: <code>userNormal</code> &nbsp;&middot;&nbsp; Administrador: <code>userAdmin</code><br>
        Contrasena: <code>12345678</code>
      </div>
    </div>
  </section>
</div>
<script src="js/app.js"></script>
</body>
</html>
