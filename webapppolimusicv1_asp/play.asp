<%
'=========================================================================
' PoliMusic :: play.asp
' Endpoint AJAX: incrementa el contador PLAYS de una cancion.
' Lo invoca js/player.js al empezar la reproduccion.
'=========================================================================
Option Explicit
%>
<!--#include file="includes/helpers.asp"-->
<!--#include file="includes/dbconn.asp"-->
<%
Dim songId
songId = Request.QueryString("id")

If Trim(Session("USER_NAME") & "") = "" Then
    Response.Write "NOAUTH"
ElseIf Not IsNumeric(songId) Then
    Response.Write "BADID"
Else
    DbExec "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(CLng(songId))
    Response.Write "OK"
End If
%>
