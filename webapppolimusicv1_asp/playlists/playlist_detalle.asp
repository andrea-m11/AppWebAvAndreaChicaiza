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
' playlist_detalle.asp · Canciones de una playlist
' Administra TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG).
' El dueño puede agregar/quitar canciones; el
' administrador solo visualiza.
'=====================================================

RequiereLogin()

Dim accion, idPlaylist, idCancion
accion = Request.Form("accion") & Request.QueryString("accion") & ""
idPlaylist = Request.Form("id") & ""
If Len(idPlaylist) = 0 Then idPlaylist = Request.QueryString("id") & ""
idCancion = Request.Form("cancion") & ""
If Len(idCancion) = 0 Then idCancion = Request.QueryString("cancion") & ""

Dim conn, esAdmin, esDueno
Set conn = AbrirConexion()
esAdmin = (CStr(Session("USER_TYPE") & "") = "1")
esDueno = False

If Not EsEntero(idPlaylist) Then
    MensajeFlash "err", "Playlist no válida."
    conn.Close
    Set conn = Nothing
    Response.Redirect "playlists.asp"
    Response.End
End If

'-----------------------------------------------------
' Verificar existencia y permisos
'-----------------------------------------------------
Dim rsP, titulo, tipoP, fotoP, dueno, fechaP, usernameDueno
Set rsP = Consulta(conn, "SELECT p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, p.ID_USER, u.USERNAME FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER WHERE p.ID_PLAYLIST = ?", Array(CLng(idPlaylist)))

If rsP.EOF Then
    MensajeFlash "err", "La playlist solicitada no existe."
    rsP.Close
    conn.Close
    Set conn = Nothing
    Response.Redirect "playlists.asp"
    Response.End
End If

titulo = CStr(rsP("TITLE").Value)
tipoP = CLng(rsP("TYPE").Value)
fotoP = RutaAsset(rsP("PHOTO").Value, "playlist")
dueno = CLng(rsP("ID_USER").Value)
usernameDueno = CStr(rsP("USERNAME").Value)
If Not IsNull(rsP("CREATION_DATE").Value) Then
    fechaP = Day(rsP("CREATION_DATE").Value) & "/" & Month(rsP("CREATION_DATE").Value) & "/" & Year(rsP("CREATION_DATE").Value)
Else
    fechaP = "Sin fecha"
End If
rsP.Close
Set rsP = Nothing

esDueno = (dueno = CLng(Session("ID_USER")))

' Solo el dueño (o el administrador en modo lectura) puede ver la playlist
If Not esDueno And Not esAdmin Then
    MensajeFlash "err", "Solo puedes ver tus propias playlists."
    conn.Close
    Set conn = Nothing
    Response.Redirect "playlists.asp"
    Response.End
End If

'-----------------------------------------------------
' Acciones del dueño: agregar / quitar canciones
'-----------------------------------------------------
If accion = "agregar" And esDueno Then
    If EsEntero(idCancion) Then
        Dim yaEsta
        yaEsta = Escalar(conn, "SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(CLng(idPlaylist), CLng(idCancion)))
        If CDbl(yaEsta) = 0 Then
            Ejecutar conn, "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array(CLng(idPlaylist), CLng(idCancion))
            MensajeFlash "ok", "Canción agregada a la playlist."
        Else
            MensajeFlash "err", "La canción ya está en esta playlist."
        End If
    End If
    conn.Close
    Set conn = Nothing
    Response.Redirect "playlist_detalle.asp?id=" & CLng(idPlaylist)
    Response.End
End If

If accion = "quitar" And esDueno Then
    If EsEntero(idCancion) Then
        Ejecutar conn, "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(CLng(idPlaylist), CLng(idCancion))
        MensajeFlash "ok", "Canción removida de la playlist."
    End If
    conn.Close
    Set conn = Nothing
    Response.Redirect "playlist_detalle.asp?id=" & CLng(idPlaylist)
    Response.End
End If

'-----------------------------------------------------
' Canciones de la playlist y canciones disponibles
'-----------------------------------------------------
Dim rsEn, rsFuera
Set rsEn = Consulta(conn, "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS FROM TBL_PLAYLIST_SONG ps INNER JOIN TBL_SONG s ON s.ID_SONG = ps.ID_SONG WHERE ps.ID_PLAYLIST = ? ORDER BY s.SONG_NAME", Array(CLng(idPlaylist)))
Set rsFuera = Consulta(conn, "SELECT ID_SONG, SONG_NAME FROM TBL_SONG WHERE ID_SONG NOT IN (SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?) ORDER BY SONG_NAME", Array(CLng(idPlaylist)))

Encabezado "Playlist · " & titulo, "playlists"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Playlist</div>
        <h2>Canciones de la playlist</h2>
        <p>Reproduce la música directamente desde el reproductor de la parte inferior.</p>
    </div>
    <div class="page-actions">
        <a class="btn btn-ghost" href="playlists.asp">Volver a playlists</a>
    </div>
</div>

<div class="detail-head">
    <img class="detail-cover" src="<%=Esc(fotoP)%>" alt="Portada de <%=Esc(titulo)%>">
    <div class="detail-meta">
        <div class="eyebrow">Playlist</div>
        <h2><%=Esc(titulo)%></h2>
        <div class="detail-sub">Creada por <%=Esc(usernameDueno)%> &middot; <%=Esc(fechaP)%></div>
        <div class="detail-tags">
            <span class="tag-rosa"><%=Esc(EtiquetaTipoPlaylist(tipoP))%></span>
            <span class="tag">ID playlist: <%=CLng(idPlaylist)%></span>
            <% If esDueno And Not esAdmin Then %><span class="tag">Eres el due&ntilde;o</span><% End If %>
        </div>
    </div>
</div>

<% If rsEn.EOF Then %>
<div class="empty">
    <svg viewBox="0 0 24 24" width="46" height="46" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
    <div><strong>Playlist vacía</strong><p><% If esDueno Then Response.Write "Agrega canciones desde el panel de abajo." Else Response.Write "El dueño aún no ha agregado canciones." %></p></div>
</div>
<% Else %>

<div class="song-list">
    <div class="song-list-head">
        <div>#</div>
        <div>Portada</div>
        <div>T&iacute;tulo / Artista</div>
        <div class="col-album">&Aacute;lbum</div>
        <div>Reproducciones</div>
        <div style="text-align:right"><% If esDueno And Not esAdmin Then Response.Write "Acciones" Else Response.Write "Reproducir" %></div>
    </div>
    <%
    Dim i, portada, artistas, albumes, rsAlb2
    i = 0
    Do While Not rsEn.EOF
        i = i + 1
        portada = PortadaCancion(conn, CLng(rsEn("ID_SONG").Value))
        artistas = ArtistasCancion(conn, CLng(rsEn("ID_SONG").Value))

        albumes = ""
        Set rsAlb2 = Consulta(conn, "SELECT al.TITLE_ALBUM, al.RELEASE_YEAR FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = ? ORDER BY al.TITLE_ALBUM", Array(CLng(rsEn("ID_SONG").Value)))
        Do While Not rsAlb2.EOF
            If Len(albumes) > 0 Then albumes = albumes & ", "
            albumes = albumes & CStr(rsAlb2("TITLE_ALBUM").Value)
            If Not IsNull(rsAlb2("RELEASE_YEAR").Value) Then albumes = albumes & " (" & CStr(rsAlb2("RELEASE_YEAR").Value) & ")"
            rsAlb2.MoveNext
        Loop
        rsAlb2.Close
        Set rsAlb2 = Nothing
        If Len(albumes) = 0 Then albumes = "Sin álbum"
    %>
    <div class="song-row" style="animation-delay:<%= (i - 1) * 60 %>ms"
         data-id="<%=rsEn("ID_SONG").Value%>"
         data-name="<%=Esc(rsEn("SONG_NAME").Value)%>"
         data-artist="<%=Esc(artistas)%>"
         data-src="<%=Esc(RutaAsset(rsEn("SONG_PATH").Value, "cancion"))%>"
         data-cover="<%=Esc(portada)%>">
        <div class="song-index">
            <%=i%>
            <button class="row-play" type="button" title="Reproducir <%=Esc(rsEn("SONG_NAME").Value)%>" aria-label="Reproducir">
                <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M7 4v16l13-8z"/></svg>
            </button>
        </div>
        <div class="song-cover"><img src="<%=Esc(portada)%>" alt="Portada de <%=Esc(rsEn("SONG_NAME").Value)%>"></div>
        <div class="song-info">
            <div class="song-name"><%=Esc(rsEn("SONG_NAME").Value)%></div>
            <div class="song-artists"><%=Esc(artistas)%></div>
        </div>
        <div class="song-album"><%=Esc(albumes)%></div>
        <div class="song-plays">
            <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M12 3a9 9 0 1 0 9 9 9 9 0 0 0-9-9zm0 16a7 7 0 1 1 7-7 7 7 0 0 1-7 7z"/><path d="M12 7a5 5 0 1 0 5 5 5 5 0 0 0-5-5zm0 8a3 3 0 1 1 3-3 3 3 0 0 1-3 3z"/></svg>
            <span class="plays-num" data-plays="<%=rsEn("PLAYS").Value%>"><%=rsEn("PLAYS").Value%></span>
            <span class="eq" aria-hidden="true"><span></span><span></span><span></span><span></span></span>
        </div>
        <div class="song-actions">
            <% If esDueno And Not esAdmin Then %>
            <form class="inline-form" method="post" action="playlist_detalle.asp" onsubmit="return confirm('¿Quitar «<%=Esc(rsEn("SONG_NAME").Value)%>» de la playlist?');">
                <input type="hidden" name="accion" value="quitar">
                <input type="hidden" name="id" value="<%=CLng(idPlaylist)%>">
                <input type="hidden" name="cancion" value="<%=rsEn("ID_SONG").Value%>">
                <button class="btn btn-danger btn-sm" type="submit">Quitar</button>
            </form>
            <% End If %>
        </div>
    </div>
    <%
        rsEn.MoveNext
    Loop
    %>
</div>

<% End If %>

<% If esDueno And Not esAdmin Then %>
<div class="panel" style="margin-top:24px">
    <h3><span class="dot"></span>Agregar canciones a la playlist</h3>
    <% If rsFuera.EOF Then %>
        <p style="color:var(--texto-3);font-size:14px">No hay más canciones disponibles para agregar.</p>
    <% Else %>
        <div class="check-grid">
            <%
            Do While Not rsFuera.EOF
            %>
            <form class="inline-form" method="post" action="playlist_detalle.asp">
                <input type="hidden" name="accion" value="agregar">
                <input type="hidden" name="id" value="<%=CLng(idPlaylist)%>">
                <input type="hidden" name="cancion" value="<%=rsFuera("ID_SONG").Value%>">
                <button class="check-item" type="submit" style="width:100%;text-align:left">
                    <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z"/></svg>
                    <%=Esc(rsFuera("SONG_NAME").Value)%>
                </button>
            </form>
            <%
                rsFuera.MoveNext
            Loop
            %>
        </div>
    <% End If %>
</div>
<% End If %>

<%
rsEn.Close
rsFuera.Close
Set rsEn = Nothing
Set rsFuera = Nothing
conn.Close
Set conn = Nothing

Pie
%>
