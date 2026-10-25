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
' albumes.asp · Listado de álbumes (vista para todos)
' Tabla: TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR,
'        ALBUM_IMAGE_PATH). El CRUD es exclusivo del ADMIN.
' El tracklist (TBL_TRACKLIST) se ve en album_detalle.asp
'=====================================================

RequiereLogin()

Dim conn, rs, esAdmin
Set conn = AbrirConexion()
esAdmin = (CStr(Session("USER_TYPE") & "") = "1")
Set rs = Consulta(conn, "SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, (SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS CANCIONES FROM TBL_ALBUM al ORDER BY al.TITLE_ALBUM", Array())

Encabezado "Álbumes", "albumes"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Cat&aacute;logo musical</div>
        <h2>&Aacute;lbumes <span class="acento">&middot; <% If esAdmin Then Response.Write "Gesti&oacute;n completa" Else Response.Write "cat&aacute;logo" %></span></h2>
        <p>Lista de &aacute;lbumes registrados. <% If esAdmin Then Response.Write "Puedes editar sus datos o administrar las canciones que contiene (tracklist)." Else Response.Write "Explora los &aacute;lbumes y las canciones de cada tracklist." %></p>
    </div>
    <% If esAdmin Then %>
    <div class="page-actions">
        <a class="btn btn-rosa" href="../admin/album_form.asp">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z"/></svg>
            Nuevo álbum
        </a>
    </div>
    <% End If %>
</div>

<% If rs.EOF Then %>
<div class="empty">
    <svg viewBox="0 0 24 24" width="46" height="46" fill="currentColor"><path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm0 14a4 4 0 1 1 4-4 4 4 0 0 1-4 4zm0-6a2 2 0 1 0 2-2 2 2 0 0 0-2 2z"/></svg>
    <div><strong>No hay álbumes registrados</strong><p><% If esAdmin Then Response.Write "Crea el primer &aacute;lbum con el bot&oacute;n &laquo;Nuevo &aacute;lbum&raquo;." Else Response.Write "Todav&iacute;a no se han registrado &aacute;lbumes." %></p></div>
</div>
<% Else %>

<div class="card-grid">
    <%
    Do While Not rs.EOF
        Dim anio
        If IsNull(rs("RELEASE_YEAR").Value) Then anio = "Año no definido" Else anio = "Año " & CStr(rs("RELEASE_YEAR").Value)
    %>
    <div class="media-card">
        <div class="media-cover">
            <img src="<%=Esc(RutaAsset(rs("ALBUM_IMAGE_PATH").Value, "album"))%>" alt="Portada de <%=Esc(rs("TITLE_ALBUM").Value)%>">
            <div class="cover-fade"></div>
            <span class="media-badge"><%=anio%></span>
        </div>
        <div class="media-body">
            <div class="media-title"><%=Esc(rs("TITLE_ALBUM").Value)%></div>
            <div class="media-meta">
                <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
                <%=rs("CANCIONES").Value%> canci&oacute;n(es) en el tracklist
            </div>
            <div class="media-actions">
                <a class="btn btn-rosa btn-sm" href="album_detalle.asp?id=<%=rs("ID_ALBUM").Value%>">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M4 4h2v14h12v2H4zm4 2h10a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2zm2 2v6h8V8z"/></svg>
                    Tracklist
                </a>
                <% If esAdmin Then %>
                <a class="btn btn-purple btn-sm" href="../admin/album_form.asp?id=<%=rs("ID_ALBUM").Value%>">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M3 17.25V21h3.75L17.8 9.94l-3.75-3.75zM20.7 7.04a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75z"/></svg>
                    Editar
                </a>
                <form class="inline-form" method="post" action="../admin/album_form.asp" onsubmit="return confirm('¿Eliminar el álbum «<%=Esc(rs("TITLE_ALBUM").Value)%>»? También se quitarán sus canciones del tracklist.');">
                    <input type="hidden" name="id" value="<%=rs("ID_ALBUM").Value%>">
                    <input type="hidden" name="accion" value="eliminar">
                    <button class="btn btn-danger btn-sm" type="submit">
                        <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M6 7h12l-1 14H7zm3-3h6l1 2H8z"/></svg>
                        Eliminar
                    </button>
                </form>
                <% End If %>
            </div>
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
