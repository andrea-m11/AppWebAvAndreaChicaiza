<%
'=====================================================
' layout.asp
' Estructura visual común: barra lateral, barra superior
' con mensaje de bienvenida (lee Session("USERNAME")),
' contenido y reproductor global fijo.
' Los enlaces usan RaizVirtual() para funcionar desde
' cualquier subcarpeta (auth/, catalogo/, admin/, playlists/).
'=====================================================

Sub Encabezado(ByVal titulo, ByVal activo)
    Response.Charset = "UTF-8"
    Dim foto
    foto = RutaAsset(Session("USER_PHOTO") & "", "usuario")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%=Esc(titulo)%> &middot; PoliMusic</title>
<link rel="stylesheet" href="<%=RaizVirtual()%>css/estilos.css">
</head>
<body>
<div class="app">
    <aside class="sidebar">
        <a class="brand" href="<%=RaizVirtual()%>catalogo/home.asp">
            <span class="brand-note">
                <svg viewBox="0 0 24 24" width="22" height="22" fill="currentColor" aria-hidden="true"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
            </span>
            <span class="brand-text">Poli<em>Music</em></span>
        </a>
        <nav class="nav">
            <a href="<%=RaizVirtual()%>catalogo/home.asp" class="nav-item <% If activo = "home" Then Response.Write "activo" %>">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M3 10.5 12 3l9 7.5V21h-6v-6H9v6H3z"/></svg>
                <span>Inicio</span>
            </a>
            <a href="<%=RaizVirtual()%>catalogo/canciones.asp" class="nav-item <% If activo = "canciones" Then Response.Write "activo" %>">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
                <span>Canciones</span>
            </a>
            <a href="<%=RaizVirtual()%>catalogo/artistas.asp" class="nav-item <% If activo = "artistas" Then Response.Write "activo" %>">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 12a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm0 2c-3.3 0-8 1.7-8 4v2h16v-2c0-2.3-4.7-4-8-4z"/></svg>
                <span>Artistas</span>
            </a>
            <a href="<%=RaizVirtual()%>catalogo/albumes.asp" class="nav-item <% If activo = "albumes" Then Response.Write "activo" %>">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm0 14a4 4 0 1 1 4-4 4 4 0 0 1-4 4zm0-6a2 2 0 1 0 2-2 2 2 0 0 0-2 2z"/></svg>
                <span>&Aacute;lbumes</span>
            </a>
            <a href="<%=RaizVirtual()%>playlists/playlists.asp" class="nav-item <% If activo = "playlists" Then Response.Write "activo" %>">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M4 4h2v14h12v2H4zm4 2h10a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2zm2 2v6h8V8z"/></svg>
                <span><% If CStr(Session("USER_TYPE") & "") = "1" Then Response.Write "Playlists" Else Response.Write "Mis Playlists" %></span>
            </a>
        </nav>
    </aside>

    <div class="main">
        <header class="topbar">
            <div class="welcome">
                <h1 class="welcome-text">&iexcl;Bienvenido, <strong><%=Esc(Session("USERNAME"))%></strong>!</h1>
                <p class="welcome-sub"><%=Esc(titulo)%></p>
            </div>
            <div class="topbar-right">
                <div class="user-chip">
                    <img class="user-avatar" src="<%=Esc(foto)%>" alt="Foto de <%=Esc(Session("USERNAME"))%>">
                    <span class="user-name"><%=Esc(Session("USERNAME"))%></span>
                </div>
                <a class="btn btn-logout" href="<%=RaizVirtual()%>auth/logout.asp" title="Cerrar sesión">
                    <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M10 17v-2h4V9h-4V7H6v10zm6-10V3l6 5-6 5V9h-2v6h2v3l6-5z"/></svg>
                    <span>Salir</span>
                </a>
            </div>
        </header>
        <section class="content">
<%
End Sub

Sub Pie()
%>
        </section>
    </div>
</div>

<div class="player" id="player">
    <div class="player-track">
        <div class="player-cover" id="pCover">
            <img id="pImg" src="<%=RaizVirtual()%>SongFilesAndImages_resources/PlaylistImages/default.jpg" alt="Portada">
        </div>
        <div class="player-meta">
            <strong id="pTitle">Selecciona una canci&oacute;n</strong>
            <span id="pArtist">PoliMusic &middot; tu música, tu estilo</span>
        </div>
    </div>
    <div class="player-center">
        <div class="player-controls">
            <button class="pbtn" id="btnPrev" type="button" title="Anterior" aria-label="Anterior">
                <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M6 5h2v14H6zm12 0v14l-9-7z"/></svg>
            </button>
            <button class="pbtn pbtn-play" id="btnPlay" type="button" title="Reproducir / Pausar" aria-label="Reproducir">
                <svg id="icoPlay" viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M7 4v16l13-8z"/></svg>
                <svg id="icoPause" viewBox="0 0 24 24" width="24" height="24" fill="currentColor" style="display:none"><path d="M6 4h4v16H6zm8 0h4v16h-4z"/></svg>
            </button>
            <button class="pbtn" id="btnNext" type="button" title="Siguiente" aria-label="Siguiente">
                <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M16 5h2v14h-2zM6 5l9 7-9 7z"/></svg>
            </button>
        </div>
        <div class="player-seek">
            <span class="p-time" id="tCur">0:00</span>
            <input type="range" id="seek" min="0" max="1000" value="0" aria-label="Progreso">
            <span class="p-time" id="tDur">0:00</span>
        </div>
    </div>
    <div class="player-right">
        <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor" class="vol-ico"><path d="M4 9v6h4l5 4V5L8 9zm12.5 3a4.5 4.5 0 0 0-2-3.7v7.4a4.5 4.5 0 0 0 2-3.7z"/></svg>
        <input type="range" id="vol" min="0" max="100" value="80" aria-label="Volumen">
    </div>
    <audio id="audio" preload="none"></audio>
</div>

<script>var RUTA_BASE = "<%=RaizVirtual()%>";</script>
<script src="<%=RaizVirtual()%>js/app.js"></script>
</body>
</html>
<%
End Sub
%>
