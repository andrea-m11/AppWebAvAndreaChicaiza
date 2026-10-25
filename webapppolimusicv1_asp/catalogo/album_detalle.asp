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
' album_detalle.asp · Tracklist de un álbum
' Cualquier usuario logueado puede VER el tracklist;
' solo el ADMIN administra TBL_TRACKLIST.
'=====================================================

RequiereLogin()

Dim accion, idAlbum, idCancion
accion = Request.Form("accion") & Request.QueryString("accion") & ""
idAlbum = Request.Form("id") & ""
If Len(idAlbum) = 0 Then idAlbum = Request.QueryString("id") & ""
idCancion = Request.Form("cancion") & ""
If Len(idCancion) = 0 Then idCancion = Request.QueryString("cancion") & ""

Dim conn, esAdmin, titulo, anio, imagen
Set conn = AbrirConexion()
esAdmin = (CStr(Session("USER_TYPE") & "") = "1")

If Not EsEntero(idAlbum) Then
    MensajeFlash "err", "Álbum no válido."
    conn.Close
    Set conn = Nothing
    Response.Redirect "albumes.asp"
    Response.End
End If

'-----------------------------------------------------
' Acciones: agregar / quitar canción del tracklist
'-----------------------------------------------------
If accion = "agregar" And esAdmin Then
    If EsEntero(idCancion) Then
        ' Evitar duplicados
        Dim existe
        existe = Escalar(conn, "SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ? AND ID_SONG = ?", Array(CLng(idAlbum), CLng(idCancion)))
        If CDbl(existe) = 0 Then
            Ejecutar conn, "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(CLng(idAlbum), CLng(idCancion))
            MensajeFlash "ok", "Canción agregada al tracklist."
        Else
            MensajeFlash "err", "La canción ya pertenece a este álbum."
        End If
    End If
    conn.Close
    Set conn = Nothing
    Response.Redirect "album_detalle.asp?id=" & CLng(idAlbum)
    Response.End
End If

If accion = "quitar" And esAdmin Then
    If EsEntero(idCancion) Then
        Ejecutar conn, "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ? AND ID_SONG = ?", Array(CLng(idAlbum), CLng(idCancion))
        MensajeFlash "ok", "Canción removida del tracklist."
    End If
    conn.Close
    Set conn = Nothing
    Response.Redirect "album_detalle.asp?id=" & CLng(idAlbum)
    Response.End
End If

'-----------------------------------------------------
' Datos del álbum
'-----------------------------------------------------
Dim rsAlb
Set rsAlb = Consulta(conn, "SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(CLng(idAlbum)))
If rsAlb.EOF Then
    MensajeFlash "err", "El álbum solicitado no existe."
    rsAlb.Close
    conn.Close
    Set conn = Nothing
    Response.Redirect "albumes.asp"
    Response.End
End If

titulo = CStr(rsAlb("TITLE_ALBUM").Value)
If IsNull(rsAlb("RELEASE_YEAR").Value) Then anio = "Año no definido" Else anio = CStr(rsAlb("RELEASE_YEAR").Value)
imagen = RutaAsset(rsAlb("ALBUM_IMAGE_PATH").Value, "album")
rsAlb.Close
Set rsAlb = Nothing

'-----------------------------------------------------
' Canciones del tracklist y canciones disponibles
'-----------------------------------------------------
Dim rsEn, rsFuera
Set rsEn = Consulta(conn, "SELECT s.ID_SONG, s.SONG_NAME FROM TBL_TRACKLIST t INNER JOIN TBL_SONG s ON s.ID_SONG = t.ID_SONG WHERE t.ID_ALBUM = ? ORDER BY s.SONG_NAME", Array(CLng(idAlbum)))
Set rsFuera = Consulta(conn, "SELECT ID_SONG, SONG_NAME FROM TBL_SONG WHERE ID_SONG NOT IN (SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?) ORDER BY SONG_NAME", Array(CLng(idAlbum)))

Encabezado "Tracklist · " & titulo, "albumes"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">&Aacute;lbum &middot; tracklist</div>
        <h2>Canciones del &aacute;lbum</h2>
        <p><% If esAdmin Then Response.Write "Agregue o quite canciones del tracklist (tabla TBL_TRACKLIST)." Else Response.Write "Consulta las canciones que forman parte de este &aacute;lbum." %></p>
    </div>
    <div class="page-actions">
        <a class="btn btn-ghost" href="albumes.asp">Volver a álbumes</a>
    </div>
</div>

<div class="detail-head">
    <img class="detail-cover" src="<%=Esc(imagen)%>" alt="Portada de <%=Esc(titulo)%>">
    <div class="detail-meta">
        <div class="eyebrow">Álbum</div>
        <h2><%=Esc(titulo)%></h2>
        <div class="detail-sub">A&ntilde;o de lanzamiento: <%=Esc(anio)%></div>
        <div class="detail-tags">
            <span class="tag-rosa">Tracklist</span>
            <span class="tag">ID &aacute;lbum: <%=CLng(idAlbum)%></span>
        </div>
    </div>
</div>

<div class="panel">
    <h3><span class="dot"></span>Canciones en el álbum</h3>
    <% If rsEn.EOF Then %>
        <div class="empty" style="padding:36px 20px">
            <svg viewBox="0 0 24 24" width="38" height="38" fill="currentColor"><path d="M12 3v10.55A4 4 0 1 0 14 17V7h4V3h-6z"/></svg>
            <div><strong>Tracklist vacío</strong><p><% If esAdmin Then Response.Write "Agregue canciones desde el panel de abajo." Else Response.Write "El &aacute;lbum todav&iacute;a no tiene canciones." %></p></div>
        </div>
    <% Else %>
        <div class="song-list">
            <div class="song-list-head" style="grid-template-columns: <% If esAdmin Then Response.Write "56px minmax(0,1fr) 120px" Else Response.Write "56px minmax(0,1fr)" %>">
                <div>#</div>
                <div>Canci&oacute;n</div>
                <% If esAdmin Then %><div style="text-align:right">Acci&oacute;n</div><% End If %>
            </div>
            <%
            Dim n
            n = 0
            Do While Not rsEn.EOF
                n = n + 1
            %>
            <div class="song-row" style="grid-template-columns: <% If esAdmin Then Response.Write "56px minmax(0,1fr) 120px" Else Response.Write "56px minmax(0,1fr)" %>">
                <div class="song-index"><%=n%></div>
                <div class="song-info"><div class="song-name"><%=Esc(rsEn("SONG_NAME").Value)%></div></div>
                <% If esAdmin Then %>
                <div class="song-actions">
                    <form class="inline-form" method="post" action="album_detalle.asp" onsubmit="return confirm('¿Quitar «<%=Esc(rsEn("SONG_NAME").Value)%>» del tracklist?');">
                        <input type="hidden" name="accion" value="quitar">
                        <input type="hidden" name="id" value="<%=CLng(idAlbum)%>">
                        <input type="hidden" name="cancion" value="<%=rsEn("ID_SONG").Value%>">
                        <button class="btn btn-danger btn-sm" type="submit">Quitar</button>
                    </form>
                </div>
                <% End If %>
            </div>
            <%
                rsEn.MoveNext
            Loop
            %>
        </div>
    <% End If %>
</div>

<div class="panel">
    <h3><span class="dot"></span>Agregar canciones al álbum</h3>
    <% If Not esAdmin Then %>
        <p style="color:var(--texto-3);font-size:14px">Solo el administrador puede modificar el tracklist del &aacute;lbum.</p>
    <% ElseIf rsFuera.EOF Then %>
        <p style="color:var(--texto-3);font-size:14px">Todas las canciones registradas ya pertenecen a este álbum.</p>
    <% Else %>
        <div class="check-grid">
            <%
            Do While Not rsFuera.EOF
            %>
            <form class="inline-form" method="post" action="album_detalle.asp">
                <input type="hidden" name="accion" value="agregar">
                <input type="hidden" name="id" value="<%=CLng(idAlbum)%>">
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

<%
rsEn.Close
rsFuera.Close
Set rsEn = Nothing
Set rsFuera = Nothing
conn.Close
Set conn = Nothing

Pie
%>
