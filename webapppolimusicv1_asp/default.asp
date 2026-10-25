<%@ CodePage = 65001 %>
<%
'=====================================================
' default.asp · Punto de entrada de PoliMusic
' Redirige al inicio si hay sesión, o al login.
'=====================================================

If Len(Session("ID_USER") & "") > 0 Then
    Response.Redirect "catalogo/home.asp"
    Response.End
Else
    Response.Redirect "auth/login.asp"
    Response.End
End If
%>
