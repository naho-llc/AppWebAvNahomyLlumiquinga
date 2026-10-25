<%
'=========================================================================
' PoliMusic :: includes/auth.asp
' Guardia de sesion: exige usuario logueado y ofrece el guardia de admin.
' Debe incluirse despues de dbconn.asp y helpers.asp.
'=========================================================================
' NOTA: sin "Option Explicit" (ver helpers.asp).
'=========================================================================

If Trim(Session("USER_NAME") & "") = "" Then
    Response.Redirect "login.asp"
    Response.End
End If

' Redirige a index.asp si el usuario no es administrador.
Sub RequireAdmin()
    If Not IsAdmin() Then
        SetFlash "No tienes permisos para acceder a esa seccion.", "error"
        Response.Redirect "index.asp"
        Response.End
    End If
End Sub
%>
