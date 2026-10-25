<%@ CodePage = 65001 %>
<!-- Los includes usan ruta virtual absoluta (/webapppolimusicv1_asp/includes/...)
     porque este IIS tiene deshabilitado "Habilitar rutas de acceso primario".
     Si renombras la carpeta de la aplicacion, actualiza estas rutas virtuales. -->
<!--#include virtual="/webapppolimusicv1_asp/includes/conexion.asp" -->
<!--#include virtual="/webapppolimusicv1_asp/includes/seguridad.asp" -->
<!--#include virtual="/webapppolimusicv1_asp/includes/rutas.asp" -->
<!--#include virtual="/webapppolimusicv1_asp/includes/layout.asp" -->
<%
'=====================================================
' artistas.asp · Listado de artistas (vista para todos)
' Tabla: TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME,
'        ARTIST_IMAGE). El CRUD es exclusivo del ADMIN.
'=====================================================

RequiereLogin()

Dim conn, rs, esAdmin
Set conn = AbrirConexion()
esAdmin = (CStr(Session("USER_TYPE") & "") = "1")
Set rs = Consulta(conn, "SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, (SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS CANCIONES FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME", Array())

Encabezado "Artistas", "artistas"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Cat&aacute;logo musical</div>
        <h2>Artistas <span class="acento">&middot; <% If esAdmin Then Response.Write "Gesti&oacute;n completa" Else Response.Write "cat&aacute;logo" %></span></h2>
        <p>Lista de artistas registrados en la base de datos. <% If esAdmin Then Response.Write "Puedes crear, editar y eliminar artistas." Else Response.Write "Explora los artistas del cat&aacute;logo." %></p>
    </div>
    <% If esAdmin Then %>
    <div class="page-actions">
        <a class="btn btn-rosa" href="../admin/artista_form.asp">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z"/></svg>
            Nuevo artista
        </a>
    </div>
    <% End If %>
</div>

<% If rs.EOF Then %>
<div class="empty">
    <svg viewBox="0 0 24 24" width="46" height="46" fill="currentColor"><path d="M12 12a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm0 2c-3.3 0-8 1.7-8 4v2h16v-2c0-2.3-4.7-4-8-4z"/></svg>
    <div><strong>No hay artistas registrados</strong><p><% If esAdmin Then Response.Write "Crea el primer artista con el bot&oacute;n &laquo;Nuevo artista&raquo;." Else Response.Write "Todav&iacute;a no se han registrado artistas." %></p></div>
</div>
<% Else %>

<div class="card-grid">
    <%
    Do While Not rs.EOF
        Dim nomCompleto
        nomCompleto = CStr(rs("ARTIST_NAME").Value)
        If Not IsNull(rs("ARTIST_LASTNAME").Value) And Len(CStr(rs("ARTIST_LASTNAME").Value)) > 0 Then
            nomCompleto = nomCompleto & " " & CStr(rs("ARTIST_LASTNAME").Value)
        End If
    %>
    <div class="media-card artist-card">
        <div class="media-cover">
            <img src="<%=Esc(RutaAsset(rs("ARTIST_IMAGE").Value, "usuario"))%>" alt="Imagen de <%=Esc(nomCompleto)%>">
            <div class="cover-fade"></div>
            <div class="artist-avatar-tag"><%=Esc(nomCompleto)%></div>
        </div>
        <div class="media-body">
            <div class="media-meta">
                <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
                <%=rs("CANCIONES").Value%> canci&oacute;n(es) asociada(s)
            </div>
            <% If esAdmin Then %>
            <div class="media-actions">
                <a class="btn btn-purple btn-sm" href="../admin/artista_form.asp?id=<%=rs("ID_ARTIST").Value%>">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M3 17.25V21h3.75L17.8 9.94l-3.75-3.75zM20.7 7.04a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75z"/></svg>
                    Editar
                </a>
                <form class="inline-form" method="post" action="../admin/artista_form.asp" onsubmit="return confirm('¿Eliminar al artista «<%=Esc(nomCompleto)%>»? También se quitará de las canciones asociadas.');">
                    <input type="hidden" name="id" value="<%=rs("ID_ARTIST").Value%>">
                    <input type="hidden" name="accion" value="eliminar">
                    <button class="btn btn-danger btn-sm" type="submit">
                        <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M6 7h12l-1 14H7zm3-3h6l1 2H8z"/></svg>
                        Eliminar
                    </button>
                </form>
            </div>
            <% End If %>
        </div>
    </div>
    <%
        rs.MoveNext
    Loop
    %>
</div>

<% End If %>

<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

Pie
%>
