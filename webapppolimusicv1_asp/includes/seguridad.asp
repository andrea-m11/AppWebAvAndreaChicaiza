<%
'=====================================================
' seguridad.asp
' Control de sesión, roles y utilidades de presentación
' Roles definidos en TBL_CATALOG (COD_CATALOG_PARENT='1'):
'   0 = USER (usuario normal)  -> CRUD de sus playlists
'   1 = ADMIN (administrador)  -> CRUD de canciones, artistas y álbumes
'=====================================================

'--- Escapa texto para evitar XSS ---
Function Esc(ByVal s)
    Dim t
    If IsNull(s) Then
        t = ""
    Else
        t = CStr(s)
    End If
    t = Server.HTMLEncode(t)
    t = Replace(t, "'", "&#39;")
    Esc = t
End Function

'--- ¿Hay sesión activa? ---
Function Logueado()
    Logueado = (Len(Session("ID_USER") & "") > 0)
End Function

'--- Ruta virtual absoluta de la raíz de la aplicación ---
' Devuelve p.ej. "/webapppolimusicv1_asp/". Funciona desde
' cualquier página y subcarpeta (auth/, catalogo/, admin/,
' playlists/) sin depender de la profundidad.
Function RaizVirtual()
    Dim sn, seg, app, p
    ' 1) Si la carpeta es una aplicación IIS, APPL_MD_PATH es exacto
    app = Request.ServerVariables("APPL_MD_PATH") & ""
    If Len(app) > 0 Then
        p = InStr(app, "/ROOT/")
        If p > 0 Then
            app = Mid(app, p + 6)
            If Len(app) > 0 Then
                If Left(app, 1) <> "/" Then app = "/" & app
                If Right(app, 1) <> "/" Then app = app & "/"
                RaizVirtual = app
                Exit Function
            End If
        End If
    End If
    ' 2) Heurística: la app vive en el primer segmento de la URL
    sn = Request.ServerVariables("SCRIPT_NAME") & ""
    If Left(sn, 1) = "/" Then sn = Mid(sn, 2)
    If Len(sn) = 0 Then
        RaizVirtual = "/"
        Exit Function
    End If
    seg = Split(sn, "/")
    Select Case LCase(seg(0))
        Case "auth", "catalogo", "admin", "playlists", "includes", "css", "js", "bd"
            ' El primer segmento ya es una carpeta interna: app en la raíz
            RaizVirtual = "/"
        Case Else
            If UBound(seg) >= 1 Then
                RaizVirtual = "/" & seg(0) & "/"
            Else
                RaizVirtual = "/"
            End If
    End Select
End Function

'--- Nombre del rol según el catálogo de la base de datos ---
Function NombreRol()
    Dim conn, rs, v
    NombreRol = "USER"
    If Not Logueado() Then Exit Function
    Set conn = AbrirConexion()
    v = Escalar(conn, "SELECT VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = '1' AND COD_CATALOG = ?", Array(CStr(Session("USER_TYPE"))))
    conn.Close
    Set conn = Nothing
    If Not IsNull(v) And Len(CStr(v)) > 0 Then NombreRol = CStr(v)
End Function

'--- Etiqueta del tipo de playlist según TBL_CATALOG (COD_CATALOG_PARENT='2') ---
Function EtiquetaTipoPlaylist(ByVal tipo)
    Dim conn, v
    Set conn = AbrirConexion()
    v = Escalar(conn, "SELECT VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = '2' AND COD_CATALOG = ?", Array(CStr(tipo)))
    conn.Close
    Set conn = Nothing
    If IsNull(v) Or Len(CStr(v)) = 0 Then
        EtiquetaTipoPlaylist = "Sin tipo"
    Else
        EtiquetaTipoPlaylist = CStr(v)
    End If
End Function

'--- Valida que un valor sea un número entero de 32 bits ---
Function EsEntero(ByVal v)
    EsEntero = False
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If IsArray(v) Then Exit Function
    Dim d
    d = CDbl(v)
    If d <> Int(d) Then Exit Function
    If d < -2147483648 Or d > 2147483647 Then Exit Function
    EsEntero = True
End Function

'--- Guarda un mensaje flash en sesión (ok / err) ---
Sub MensajeFlash(ByVal tipo, ByVal texto)
    If LCase(tipo) = "ok" Then
        Session("msg") = texto
    Else
        Session("err") = texto
    End If
End Sub

'--- Muestra y limpia los mensajes flash ---
Sub MostrarFlash()
    Dim html
    html = ""
    If Len(Session("msg") & "") > 0 Then
        html = html & "<div class='alert alert-ok'><span class='alert-icon'>&#10003;</span><div>" & Esc(Session("msg")) & "</div></div>"
        Session.Contents.Remove("msg")
    End If
    If Len(Session("err") & "") > 0 Then
        html = html & "<div class='alert alert-err'><span class='alert-icon'>&#9888;</span><div>" & Esc(Session("err")) & "</div></div>"
        Session.Contents.Remove("err")
    End If
    Response.Write html
End Sub

'--- Páginas que requieren sesión iniciada ---
Sub RequiereLogin()
    If Not Logueado() Then
        Response.Redirect RaizVirtual() & "auth/login.asp"
        Response.End
    End If
End Sub

'--- Páginas que requieren rol de administrador (USER_TYPE = 1) ---
Sub RequiereAdmin()
    If Not Logueado() Then
        Response.Redirect RaizVirtual() & "auth/login.asp"
        Response.End
    End If
    If CStr(Session("USER_TYPE") & "") <> "1" Then
        Session("err") = "Acceso denegado: esta sección es solo para administradores."
        Response.Redirect RaizVirtual() & "catalogo/home.asp"
        Response.End
    End If
End Sub
%>
