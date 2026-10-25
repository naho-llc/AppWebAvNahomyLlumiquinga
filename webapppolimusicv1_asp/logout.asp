<%
'=========================================================================
' PoliMusic :: logout.asp
' Cierra la sesion actual (Session.Abandon destruye todas las variables
' de sesion, incluyendo usuario y rol) y regresa al login.
'=========================================================================
Option Explicit

Session.Abandon()

Response.Redirect "login.asp"
Response.End
%>
