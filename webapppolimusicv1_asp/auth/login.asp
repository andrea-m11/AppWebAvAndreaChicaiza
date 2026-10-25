<%@ CodePage = 65001 %>
<!-- Los includes usan ruta virtual absoluta (/webapppolimusicv1_asp/includes/...)
     porque este IIS tiene deshabilitado "Habilitar rutas de acceso primario".
     Si renombras la carpeta de la aplicacion, actualiza estas rutas virtuales. -->
<!--#include virtual="/webapppolimusicv1_asp/includes/conexion.asp" -->
<!--#include virtual="/webapppolimusicv1_asp/includes/seguridad.asp" -->
<!--#include virtual="/webapppolimusicv1_asp/includes/rutas.asp" -->
<%
'=====================================================
' login.asp · Inicio de sesión de PoliMusic
' Valida las credenciales contra TBL_USER y guarda los
' datos del usuario en variables de sesión.
'=====================================================
Response.Charset = "UTF-8"

Dim errorMsg, usuario
errorMsg = ""
usuario = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    usuario = Trim(Request.Form("txtUsuario") & "")
    Dim clave
    clave = Request.Form("txtClave") & ""

    If Len(usuario) = 0 Or Len(clave) = 0 Then
        errorMsg = "Debe ingresar su usuario y contraseña."
    Else
        Dim conn, rs
        Set conn = AbrirConexion()
        ' Acepta contraseña directa o almacenada como hash SHA-1
        Set rs = Consulta(conn, "SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE USERNAME = ? AND (PASSWORD = ? OR PASSWORD = CONVERT(VARCHAR(40), HASHBYTES('SHA1', ?), 2))", Array(usuario, clave, clave))

        If Not rs.EOF Then
            ' Variables de sesión con los datos del usuario logueado
            Session("ID_USER") = CLng(rs("ID_USER").Value)
            Session("USERNAME") = CStr(rs("USERNAME").Value)
            Session("EMAIL") = CStr(rs("EMAIL").Value & "")
            Session("USER_TYPE") = CInt(rs("USER_TYPE").Value)
            Session("USER_PHOTO") = CStr(rs("USER_PHOTO").Value & "")
            Session.Timeout = 60

            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing

            Response.Redirect "../catalogo/home.asp"
            Response.End
        Else
            errorMsg = "Usuario o contraseña incorrectos. Verifique sus credenciales."
            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesión &middot; PoliMusic</title>
<link rel="stylesheet" href="<%=RaizVirtual()%>css/estilos.css">
</head>
<body class="login-body">
    <main class="login-card">
        <div class="login-logo">
            <svg viewBox="0 0 24 24" width="34" height="34" fill="currentColor" aria-hidden="true"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
        </div>
        <h1>Poli<em>Music</em></h1>
        <p class="login-lema">Tu música, tu estilo. Inicia sesión para continuar.</p>

        <% If Len(errorMsg) > 0 Then %>
        <div class="alert alert-err"><span class="alert-icon">&#9888;</span><div><%=Esc(errorMsg)%></div></div>
        <% End If %>

        <form method="post" action="login.asp" autocomplete="off">
            <div class="field">
                <label for="txtUsuario">Usuario</label>
                <input type="text" id="txtUsuario" name="txtUsuario" placeholder="Escribe tu usuario" value="<%=Esc(usuario)%>" required>
            </div>
            <div class="field">
                <label for="txtClave">Contraseña</label>
                <input type="password" id="txtClave" name="txtClave" placeholder="Escribe tu contraseña" required>
            </div>
            <button type="submit" class="btn btn-rosa" style="width:100%;padding:13px;">Ingresar a PoliMusic</button>
        </form>

        <div class="login-demo">
            <b>Usuarios de prueba (BDD_PoliMusic):</b><br>
            Usuario normal: <code>userNormal</code> / <code>12345678</code><br>
            Administrador: <code>userAdmin</code> / <code>12345678</code>
        </div>
    </main>
</body>
</html>
