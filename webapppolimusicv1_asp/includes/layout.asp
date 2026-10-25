<%
'=========================================================================
' PoliMusic :: includes/layout.asp
' Chrome comun de la aplicacion (sidebar + topbar + contenido).
' Espera las variables PAGE_TITLE y ACTIVE_NAV definidas por la pagina:
'   ACTIVE_NAV: "songs" | "artists" | "albums"
' Debe incluirse despues de helpers.asp / dbconn.asp / auth.asp.
' NOTA: sin "Option Explicit" (ver helpers.asp): la directiva solo puede
' aparecer una vez por unidad de compilacion.
'=========================================================================
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%= H(PAGE_TITLE) %> &middot; PoliMusic</title>
<meta name="description" content="PoliMusic - Reproductor y administracion musical.">
<link rel="icon" type="image/svg+xml" href="assets/img/favicon.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="css/styles.css">
</head>
<body>
<div class="app">

  <%'--- Barra lateral -----------------------------------------------------%>
  <aside class="sidebar">
    <a class="brand" href="index.asp">
      <img class="brand-mark" src="assets/img/logo.svg" alt="PoliMusic">
      <span class="brand-name">PoliMusic</span>
    </a>

    <nav class="nav" aria-label="Navegacion principal">
      <span class="nav-label">Menu</span>
      <a class="nav-item <% If ACTIVE_NAV = "songs" Then Response.Write "active" %>" href="index.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.8V20a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V9.8"/><path d="M9.5 21v-6h5v6"/></svg>
        <span>Inicio</span>
      </a>
      <a class="nav-item <% If ACTIVE_NAV = "artists" Then Response.Write "active" %>" href="artists.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4.5 21c1.4-3.6 4.2-5.4 7.5-5.4s6.1 1.8 7.5 5.4"/></svg>
        <span>Artistas</span>
      </a>
      <a class="nav-item <% If ACTIVE_NAV = "albums" Then Response.Write "active" %>" href="albums.asp">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="2.6"/><circle cx="12" cy="12" r="5.2" opacity=".45"/></svg>
        <span>Albumes</span>
      </a>

      <% If IsAdmin() Then %>
      <span class="nav-label">Administracion</span>
      <a class="nav-item nav-item-admin" href="song_form.asp?action=new">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
        <span>Nueva cancion</span>
      </a>
      <a class="nav-item nav-item-admin" href="artist_form.asp?action=new">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
        <span>Nuevo artista</span>
      </a>
      <a class="nav-item nav-item-admin" href="album_form.asp?action=new">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="M5 12h14"/></svg>
        <span>Nuevo album</span>
      </a>
      <% End If %>
    </nav>

    <div class="sidebar-user">
      <img class="avatar" src="<%= Attr(ResourceUrl(Session("USER_PHOTO"), DefaultAvatarUrl())) %>" alt="Foto de perfil">
      <div class="sidebar-user-info">
        <strong class="sidebar-user-name"><%= H(Session("USER_NAME")) %></strong>
        <span class="role-chip <% If IsAdmin() Then Response.Write "role-admin" %>"><%= H(RoleLabel()) %></span>
      </div>
      <a class="icon-btn" href="logout.asp" title="Cerrar sesion" aria-label="Cerrar sesion">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M15 17l5-5-5-5"/><path d="M20 12H9"/><path d="M12 4H6a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h6"/></svg>
      </a>
    </div>
  </aside>

  <%'--- Contenido principal ----------------------------------------------%>
  <main class="main">
    <header class="topbar">
      <div class="topbar-titles">
        <p class="topbar-hello">Bienvenido,</p>
        <h1 class="topbar-user"><%= H(Session("USER_NAME")) %></h1>
        <p class="topbar-role">Sesion iniciada como <b><%= H(RoleLabel()) %></b></p>
      </div>
      <div class="topbar-actions">
        <span class="pill pill-db" title="Base de datos conectada">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5.5" rx="7.5" ry="3"/><path d="M4.5 5.5v13c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3v-13"/><path d="M4.5 12c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3"/></svg>
          BDD_PoliMusic
        </span>
      </div>
    </header>

    <div class="content">
      <%'--- Mensaje flash (una sola vez) ---------------------------------%>
      <%
      Dim flashMsg, flashType
      flashMsg = Session("FLASH_MSG") & ""
      flashType = Session("FLASH_TYPE") & ""
      If flashMsg <> "" Then
      %>
      <div class="alert alert-<%= H(IIf(flashType = "error", "error", "ok")) %>" role="status">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><% If flashType = "error" Then %><circle cx="12" cy="12" r="9"/><path d="M12 8v4.5"/><circle cx="12" cy="16.2" r=".6" fill="currentColor"/><% Else %><circle cx="12" cy="12" r="9"/><path d="m8.5 12.5 2.4 2.4 4.6-5"/><% End If %></svg>
        <span><%= H(flashMsg) %></span>
        <button type="button" class="alert-close" aria-label="Cerrar">&times;</button>
      </div>
      <%
      End If
      Call ClearFlash()
      %>
