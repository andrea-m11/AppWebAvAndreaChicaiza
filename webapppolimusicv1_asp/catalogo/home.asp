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
' home.asp · Panel principal
' Muestra el mensaje de bienvenida con el nombre del
' usuario guardado en la variable de sesión.
'=====================================================

RequiereLogin()

Dim conn, nCanciones, nArtistas, nAlbumes, nPlaylists
Set conn = AbrirConexion()
nCanciones = Escalar(conn, "SELECT COUNT(*) FROM TBL_SONG", Array())
nArtistas = Escalar(conn, "SELECT COUNT(*) FROM TBL_ARTIST", Array())
nAlbumes = Escalar(conn, "SELECT COUNT(*) FROM TBL_ALBUM", Array())
If CStr(Session("USER_TYPE") & "") = "1" Then
    nPlaylists = Escalar(conn, "SELECT COUNT(*) FROM TBL_PLAYLIST", Array())
Else
    nPlaylists = Escalar(conn, "SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CLng(Session("ID_USER"))))
End If
conn.Close
Set conn = Nothing

Encabezado "Panel principal", "home"

MostrarFlash
%>

<section class="hero">
    <div class="eyebrow">PoliMusic &middot; plataforma de streaming</div>
    <h2>&iexcl;Bienvenido, <span class="acento"><%=Esc(Session("USERNAME"))%></span>!</h2>
    <p>Explora el cat&aacute;logo completo de canciones, descubre artistas y &aacute;lbumes, arma tus listas de reproducci&oacute;n y disfruta de la m&uacute;sica con el reproductor de PoliMusic.</p>
    <div class="hero-actions">
        <a class="btn btn-rosa" href="canciones.asp">
            <svg viewBox="0 0 24 24" width="17" height="17" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
            Escuchar canciones
        </a>
        <a class="btn btn-purple" href="../playlists/playlists.asp">
            <svg viewBox="0 0 24 24" width="17" height="17" fill="currentColor"><path d="M4 4h2v14h12v2H4zm4 2h10a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2zm2 2v6h8V8z"/></svg>
            <% If CStr(Session("USER_TYPE") & "") = "1" Then Response.Write "Ver playlists" Else Response.Write "Mis playlists" %>
        </a>
        <a class="btn btn-ghost" href="artistas.asp"><% If CStr(Session("USER_TYPE") & "") = "1" Then Response.Write "Gestionar artistas" Else Response.Write "Ver artistas" %></a>
    </div>
</section>

<div class="stat-grid">
    <div class="stat-card">
        <div class="stat-ico">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
        </div>
        <div class="stat-num"><%=nCanciones%></div>
        <div class="stat-label">Canciones en el catálogo</div>
    </div>
    <div class="stat-card">
        <div class="stat-ico">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M12 12a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm0 2c-3.3 0-8 1.7-8 4v2h16v-2c0-2.3-4.7-4-8-4z"/></svg>
        </div>
        <div class="stat-num"><%=nArtistas%></div>
        <div class="stat-label">Artistas registrados</div>
    </div>
    <div class="stat-card">
        <div class="stat-ico">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm0 14a4 4 0 1 1 4-4 4 4 0 0 1-4 4zm0-6a2 2 0 1 0 2-2 2 2 0 0 0-2 2z"/></svg>
        </div>
        <div class="stat-num"><%=nAlbumes%></div>
        <div class="stat-label">Álbumes disponibles</div>
    </div>
    <div class="stat-card">
        <div class="stat-ico">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M4 4h2v14h12v2H4zm4 2h10a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2zm2 2v6h8V8z"/></svg>
        </div>
        <div class="stat-num"><%=nPlaylists%></div>
        <div class="stat-label"><% If CStr(Session("USER_TYPE") & "") = "1" Then Response.Write "Playlists de usuarios" Else Response.Write "Mis playlists" %></div>
    </div>
</div>

<div class="section-title">
    <h3>&iquest;Qué quieres hacer hoy?</h3>
</div>

<div class="card-grid">
    <a class="media-card" href="canciones.asp">
        <div class="media-cover">
            <img src="<%=RaizVirtual()%>SongFilesAndImages_resources/PlaylistImages/playlist1.png" alt="Listado de canciones">
            <div class="cover-fade"></div>
        </div>
        <div class="media-body">
            <div class="media-title">Song List</div>
            <div class="media-sub">Listado de canciones</div>
            <div class="media-meta">Explora y reproduce el catálogo completo</div>
        </div>
    </a>
    <a class="media-card" href="../playlists/playlists.asp">
        <div class="media-cover">
            <img src="<%=RaizVirtual()%>SongFilesAndImages_resources/PlaylistImages/playlist2.png" alt="Playlists">
            <div class="cover-fade"></div>
        </div>
        <div class="media-body">
            <div class="media-title">Playlists</div>
            <div class="media-sub">Crea y administra tus listas</div>
            <div class="media-meta">Agrupa tus canciones favoritas</div>
        </div>
    </a>
    <a class="media-card" href="albumes.asp">
        <div class="media-cover">
            <img src="<%=RaizVirtual()%>SongFilesAndImages_resources/PlaylistImages/playlist3.png" alt="Álbumes">
            <div class="cover-fade"></div>
        </div>
        <div class="media-body">
            <div class="media-title">&Aacute;lbumes</div>
            <div class="media-sub"><% If CStr(Session("USER_TYPE") & "") = "1" Then Response.Write "Gestiona álbumes y su tracklist" Else Response.Write "Explora álbumes y su tracklist" %></div>
            <div class="media-meta"><% If CStr(Session("USER_TYPE") & "") = "1" Then Response.Write "Crear, editar y organizar canciones" Else Response.Write "Descubre las canciones de cada álbum" %></div>
        </div>
    </a>
</div>

<%
Pie
%>
