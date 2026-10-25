<%
'=========================================================================
' PoliMusic :: default.asp
' Pagina por defecto: envia al listado o al login segun la sesion.
'=========================================================================
Option Explicit

If Trim(Session("USER_NAME") & "") <> "" Then
    Response.Redirect "index.asp"
    Response.End
Else
    Response.Redirect "login.asp"
    Response.End
End If
%>
