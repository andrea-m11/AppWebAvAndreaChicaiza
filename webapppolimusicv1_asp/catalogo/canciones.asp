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
' canciones.asp · Song List (Listado de Canciones)
' Muestra todas las canciones de TBL_SONG con su
' artista (TBL_PLAYER) y álbum (TBL_TRACKLIST).
' El administrador puede además gestionarlas (CRUD).
'=====================================================

RequiereLogin()

Dim conn, rs, esAdmin
Set conn = AbrirConexion()
esAdmin = (CStr(Session("USER_TYPE") & "") = "1")

Set rs = Consulta(conn, "SELECT ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS, 0) AS PLAYS FROM TBL_SONG ORDER BY ID_SONG", Array())

Encabezado "Listado de canciones", "canciones"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Catálogo musical</div>
        <h2>Song List <span class="acento">&middot; Listado de Canciones</span></h2>
        <p>A continuación se presenta el listado completo de canciones registradas en la base de datos. Pulsa el botón rosado para reproducir.</p>
    </div>
    <% If esAdmin Then %>
    <div class="page-actions">
        <a class="btn btn-rosa" href="../admin/cancion_form.asp">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z"/></svg>
            Nueva canción
        </a>
    </div>
    <% End If %>
</div>

<% If rs.EOF Then %>
<div class="empty">
    <svg viewBox="0 0 24 24" width="46" height="46" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
    <div><strong>No hay canciones registradas</strong><p>Todavía no se han ingresado canciones en la base de datos.</p></div>
</div>
<% Else %>

<div class="song-list">
    <div class="song-list-head">
        <div>#</div>
        <div>Portada</div>
        <div>Título / Artista</div>
        <div class="col-album">Álbum</div>
        <div>Reproducciones</div>
        <div style="text-align:right"><% If esAdmin Then Response.Write "Acciones" Else Response.Write "Reproducir" %></div>
    </div>

    <%
    Dim i, portada, artistas, albumes, rsAlb
    i = 0
    Do While Not rs.EOF
        i = i + 1
        portada = PortadaCancion(conn, CLng(rs("ID_SONG").Value))
        artistas = ArtistasCancion(conn, CLng(rs("ID_SONG").Value))

        ' Álbumes que contienen la canción
        albumes = ""
        Set rsAlb = Consulta(conn, "SELECT al.TITLE_ALBUM, al.RELEASE_YEAR FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = ? ORDER BY al.TITLE_ALBUM", Array(CLng(rs("ID_SONG").Value)))
        Do While Not rsAlb.EOF
            If Len(albumes) > 0 Then albumes = albumes & ", "
            albumes = albumes & CStr(rsAlb("TITLE_ALBUM").Value)
            If Not IsNull(rsAlb("RELEASE_YEAR").Value) Then albumes = albumes & " (" & CStr(rsAlb("RELEASE_YEAR").Value) & ")"
            rsAlb.MoveNext
        Loop
        rsAlb.Close
        Set rsAlb = Nothing
        If Len(albumes) = 0 Then albumes = "Sin álbum"
    %>
    <div class="song-row" style="animation-delay:<%= (i - 1) * 60 %>ms"
         data-id="<%=rs("ID_SONG").Value%>"
         data-name="<%=Esc(rs("SONG_NAME").Value)%>"
         data-artist="<%=Esc(artistas)%>"
         data-src="<%=Esc(RutaAsset(rs("SONG_PATH").Value, "cancion"))%>"
         data-cover="<%=Esc(portada)%>">
        <div class="song-index">
            <%=i%>
            <button class="row-play" type="button" title="Reproducir <%=Esc(rs("SONG_NAME").Value)%>" aria-label="Reproducir">
                <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M7 4v16l13-8z"/></svg>
            </button>
        </div>
        <div class="song-cover"><img src="<%=Esc(portada)%>" alt="Portada de <%=Esc(rs("SONG_NAME").Value)%>"></div>
        <div class="song-info">
            <div class="song-name"><%=Esc(rs("SONG_NAME").Value)%></div>
            <div class="song-artists"><%=Esc(artistas)%></div>
        </div>
        <div class="song-album"><%=Esc(albumes)%></div>
        <div class="song-plays">
            <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M12 3a9 9 0 1 0 9 9 9 9 0 0 0-9-9zm0 16a7 7 0 1 1 7-7 7 7 0 0 1-7 7z"/><path d="M12 7a5 5 0 1 0 5 5 5 5 0 0 0-5-5zm0 8a3 3 0 1 1 3-3 3 3 0 0 1-3 3z"/></svg>
            <span class="plays-num" data-plays="<%=rs("PLAYS").Value%>"><%=rs("PLAYS").Value%></span>
            <span class="eq" aria-hidden="true"><span></span><span></span><span></span><span></span></span>
        </div>
        <div class="song-actions">
            <% If esAdmin Then %>
            <a class="icon-btn" href="../admin/cancion_form.asp?id=<%=rs("ID_SONG").Value%>" title="Editar canción">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M3 17.25V21h3.75L17.8 9.94l-3.75-3.75zM20.7 7.04a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75z"/></svg>
            </a>
            <form class="inline-form" method="post" action="../admin/cancion_form.asp" onsubmit="return confirm('¿Eliminar la canción «<%=Esc(rs("SONG_NAME").Value)%>»? Esta acción no se puede deshacer.');">
                <input type="hidden" name="id" value="<%=rs("ID_SONG").Value%>">
                <input type="hidden" name="accion" value="eliminar">
                <button class="icon-btn danger" type="submit" title="Eliminar canción">
                    <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M6 7h12l-1 14H7zm3-3h6l1 2H8z"/></svg>
                </button>
            </form>
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
