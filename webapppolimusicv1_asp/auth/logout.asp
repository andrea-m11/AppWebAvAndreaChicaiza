<%@ CodePage = 65001 %>
<%
'=====================================================
' logout.asp · Cierre de sesión
' Abandona la sesión y regresa al login.
'=====================================================

Session.Abandon
Response.Redirect "login.asp"
Response.End
%>
