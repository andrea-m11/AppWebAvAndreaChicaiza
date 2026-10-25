<%@ CodePage = 65001 %>
<!--#include file="conexion.asp" -->
<!--#include file="seguridad.asp" -->
<%
'=====================================================
' reproducir.asp
' Endpoint AJAX: incrementa el contador PLAYS de una
' canción (TBL_SONG) cuando el usuario la reproduce.
' Devuelve JSON.
'=====================================================
Response.Charset = "UTF-8"

Dim id, conn
id = Request.Form("id") & ""

If Len(Session("ID_USER") & "") = 0 Then
    Response.Write "{""ok"":false,""error"":""sin sesion""}"
    Response.End
End If

If Not EsEntero(id) Then
    Response.Write "{""ok"":false,""error"":""id invalido""}"
    Response.End
End If

Set conn = AbrirConexion()
Ejecutar conn, "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(CLng(id))
conn.Close
Set conn = Nothing

Response.Write "{""ok"":true}"
%>
