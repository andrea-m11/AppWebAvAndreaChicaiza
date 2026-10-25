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
' playlists.asp · Playlists
' Usuario normal (USER_TYPE=0): CRUD de SUS playlists.
' Administrador (USER_TYPE=1): visualiza todas (lectura).
' Tabla: TBL_PLAYLIST + TBL_PLAYLIST_SONG
'=====================================================

RequiereLogin()

Dim conn, rs, esAdmin
Set conn = AbrirConexion()
esAdmin = (CStr(Session("USER_TYPE") & "") = "1")

If esAdmin Then
    Set rs = Consulta(conn, "SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS CANCIONES FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER ORDER BY p.ID_PLAYLIST", Array())
Else
    Set rs = Consulta(conn, "SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS CANCIONES FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER WHERE p.ID_USER = ? ORDER BY p.ID_PLAYLIST", Array(CLng(Session("ID_USER"))))
End If

Dim tituloPagina
If esAdmin Then
    tituloPagina = "Playlists de usuarios"
Else
    tituloPagina = "Mis playlists"
End If

Encabezado tituloPagina, "playlists"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Biblioteca personal</div>
        <h2><% If esAdmin Then Response.Write "Playlists <span class='acento'>&middot; todos los usuarios</span>" Else Response.Write "Mis playlists <span class='acento'>&middot; CRUD</span>" %></h2>
        <p><% If esAdmin Then Response.Write "Vista de solo lectura de las listas de reproducción creadas por los usuarios." Else Response.Write "Crea, edita y elimina tus listas de reproducción, y administra las canciones que contienen." %></p>
    </div>
    <% If Not esAdmin Then %>
    <div class="page-actions">
        <a class="btn btn-rosa" href="playlist_form.asp">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z"/></svg>
            Nueva playlist
        </a>
    </div>
    <% End If %>
</div>

<% If rs.EOF Then %>
<div class="empty">
    <svg viewBox="0 0 24 24" width="46" height="46" fill="currentColor"><path d="M4 4h2v14h12v2H4zm4 2h10a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2zm2 2v6h8V8z"/></svg>
    <div>
        <strong><% If esAdmin Then Response.Write "No hay playlists registradas" Else Response.Write "Todavía no tienes playlists" %></strong>
        <p><% If esAdmin Then Response.Write "Los usuarios aún no han creado listas de reproducción." Else Response.Write "Crea tu primera playlist con el botón «Nueva playlist»." %></p>
    </div>
</div>
<% Else %>

<div class="card-grid">
    <%
    Do While Not rs.EOF
        Dim etiqueta, fechaTxt, esPrivada
        etiqueta = EtiquetaTipoPlaylist(rs("TYPE").Value)
        esPrivada = (CLng(rs("TYPE").Value) = 1)
        If IsNull(rs("CREATION_DATE").Value) Then
            fechaTxt = "Sin fecha"
        Else
            fechaTxt = Day(rs("CREATION_DATE").Value) & "/" & Month(rs("CREATION_DATE").Value) & "/" & Year(rs("CREATION_DATE").Value)
        End If
    %>
    <div class="media-card">
        <div class="media-cover">
            <img src="<%=Esc(RutaAsset(rs("PHOTO").Value, "playlist"))%>" alt="Portada de <%=Esc(rs("TITLE").Value)%>">
            <div class="cover-fade"></div>
            <span class="media-badge<% If esPrivada Then Response.Write " privado" %>"><%=Esc(etiqueta)%></span>
        </div>
        <div class="media-body">
            <div class="media-title"><%=Esc(rs("TITLE").Value)%></div>
            <div class="media-sub">Creada por <%=Esc(rs("USERNAME").Value)%> &middot; <%=Esc(fechaTxt)%></div>
            <div class="media-meta">
                <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
                <%=rs("CANCIONES").Value%> canci&oacute;n(es)
            </div>
            <div class="media-actions">
                <a class="btn btn-rosa btn-sm" href="playlist_detalle.asp?id=<%=rs("ID_PLAYLIST").Value%>">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 4.5C7 4.5 2.7 6.9 1 10c1.7 3.1 6 5.5 11 5.5s9.3-2.4 11-5.5c-1.7-3.1-6-5.5-11-5.5zm0 9.5a4 4 0 1 1 4-4 4 4 0 0 1-4 4z"/></svg>
                    Ver
                </a>
                <% If Not esAdmin Then %>
                <a class="btn btn-purple btn-sm" href="playlist_form.asp?id=<%=rs("ID_PLAYLIST").Value%>">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M3 17.25V21h3.75L17.8 9.94l-3.75-3.75zM20.7 7.04a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75z"/></svg>
                    Editar
                </a>
                <form class="inline-form" method="post" action="playlist_form.asp" onsubmit="return confirm('¿Eliminar la playlist «<%=Esc(rs("TITLE").Value)%>»?');">
                    <input type="hidden" name="id" value="<%=rs("ID_PLAYLIST").Value%>">
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
