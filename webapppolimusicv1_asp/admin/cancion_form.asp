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
' cancion_form.asp · CRUD de canciones (solo ADMIN)
' Campos de TBL_SONG: SONG_NAME, SONG_PATH, PLAYS
' Además administra las asociaciones:
'   TBL_PLAYER     (canción -> artistas)
'   TBL_TRACKLIST  (canción -> álbumes)
'=====================================================

RequiereAdmin()

Dim accion, idCancion
accion = Request.Form("accion") & ""
idCancion = Request.Form("id") & ""
If Len(idCancion) = 0 Then idCancion = Request.QueryString("id") & ""

Dim nombre, ruta, plays
nombre = ""
ruta = ""
plays = 0

Dim conn
Set conn = AbrirConexion()

'-----------------------------------------------------
' Acción: ELIMINAR
'-----------------------------------------------------
If accion = "eliminar" Then
    If EsEntero(idCancion) Then
        On Error Resume Next
        conn.BeginTrans
        Ejecutar conn, "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(CLng(idCancion))
        Ejecutar conn, "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(CLng(idCancion))
        Ejecutar conn, "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array(CLng(idCancion))
        Ejecutar conn, "DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(CLng(idCancion))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "No se pudo eliminar la canción."
        Else
            conn.CommitTrans
            On Error GoTo 0
            MensajeFlash "ok", "Canción eliminada correctamente."
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/canciones.asp"
        Response.End
    Else
        MensajeFlash "err", "Identificador de canción no válido."
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/canciones.asp"
        Response.End
    End If
End If

'-----------------------------------------------------
' Acción: GUARDAR (insertar o actualizar)
'-----------------------------------------------------
If accion = "guardar" Then
    nombre = Trim(Request.Form("txtNombre") & "")
    ruta = Trim(Request.Form("txtRuta") & "")
    plays = Trim(Request.Form("txtPlays") & "")
    If Len(plays) = 0 Then plays = "0"
    Dim artistasSel, albumesSel
    artistasSel = Request.Form("chkArtistas") & ""
    albumesSel = Request.Form("chkAlbumes") & ""

    If Len(nombre) = 0 Then
        MensajeFlash "err", "El nombre de la canción es obligatorio."
    ElseIf Len(ruta) = 0 Then
        MensajeFlash "err", "La ruta del archivo de audio es obligatoria."
    ElseIf Not EsEntero(plays) Then
        MensajeFlash "err", "El número de reproducciones debe ser un entero."
    Else
        Dim nuevoId
        On Error Resume Next
        conn.BeginTrans

        If EsEntero(idCancion) Then
            ' Actualizar
            Ejecutar conn, "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", Array(nombre, ruta, CLng(plays), CLng(idCancion))
            nuevoId = CLng(idCancion)
        Else
            ' Insertar
            Ejecutar conn, "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?)", Array(nombre, ruta, CLng(plays))
            nuevoId = CLng(Escalar(conn, "SELECT SCOPE_IDENTITY()", Array()))
        End If

        ' Refrescar asociaciones
        Ejecutar conn, "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(nuevoId)
        Ejecutar conn, "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(nuevoId)

        Dim lista, k
        If Len(artistasSel) > 0 Then
            lista = Split(artistasSel, ",")
            For k = 0 To UBound(lista)
                If EsEntero(Trim(lista(k))) Then
                    Ejecutar conn, "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", Array(nuevoId, CLng(Trim(lista(k))))
                End If
            Next
        End If

        If Len(albumesSel) > 0 Then
            lista = Split(albumesSel, ",")
            For k = 0 To UBound(lista)
                If EsEntero(Trim(lista(k))) Then
                    Ejecutar conn, "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(CLng(Trim(lista(k))), nuevoId)
                End If
            Next
        End If

        If Err.Number <> 0 Then
            conn.RollbackTrans
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "Ocurrió un error al guardar la canción."
        Else
            conn.CommitTrans
            On Error GoTo 0
            MensajeFlash "ok", "Canción guardada correctamente."
            conn.Close
            Set conn = Nothing
            Response.Redirect "../catalogo/canciones.asp"
            Response.End
        End If
    End If
End If

'-----------------------------------------------------
' Cargar datos actuales de la canción (edición)
'-----------------------------------------------------
Dim dictArtistas, dictAlbumes
Set dictArtistas = Server.CreateObject("Scripting.Dictionary")
Set dictAlbumes = Server.CreateObject("Scripting.Dictionary")

If EsEntero(idCancion) And accion <> "guardar" Then
    Dim rsC
    Set rsC = Consulta(conn, "SELECT SONG_NAME, SONG_PATH, ISNULL(PLAYS,0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(CLng(idCancion)))
    If Not rsC.EOF Then
        nombre = CStr(rsC("SONG_NAME").Value)
        ruta = CStr(rsC("SONG_PATH").Value)
        plays = CLng(rsC("PLAYS").Value)
    Else
        MensajeFlash "err", "La canción solicitada no existe."
        rsC.Close
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/canciones.asp"
        Response.End
    End If
    rsC.Close
    Set rsC = Nothing

    Dim rsAux, clave
    Set rsAux = Consulta(conn, "SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", Array(CLng(idCancion)))
    Do While Not rsAux.EOF
        clave = "A" & CStr(rsAux("ID_ARTIST").Value)
        If Not dictArtistas.Exists(clave) Then dictArtistas.Add clave, True
        rsAux.MoveNext
    Loop
    rsAux.Close
    Set rsAux = Nothing

    Set rsAux = Consulta(conn, "SELECT ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(CLng(idCancion)))
    Do While Not rsAux.EOF
        clave = "B" & CStr(rsAux("ID_ALBUM").Value)
        If Not dictAlbumes.Exists(clave) Then dictAlbumes.Add clave, True
        rsAux.MoveNext
    Loop
    rsAux.Close
    Set rsAux = Nothing
End If

Dim rsArtistas, rsAlbumes
Set rsArtistas = Consulta(conn, "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME", Array())
Set rsAlbumes = Consulta(conn, "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Array())

Dim editando, tituloAccion, idParaForm
editando = EsEntero(idCancion)
If editando Then
    tituloAccion = "Editar canción"
    idParaForm = CLng(idCancion)
Else
    tituloAccion = "Nueva canción"
    idParaForm = 0
End If

Encabezado tituloAccion, "canciones"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Administraci&oacute;n &middot; canciones</div>
        <h2><%=tituloAccion%></h2>
        <p>Complete los datos de la canci&oacute;n y asocie los artistas y &aacute;lbumes correspondientes.</p>
    </div>
    <div class="page-actions">
        <a class="btn btn-ghost" href="../catalogo/canciones.asp">Volver al listado</a>
    </div>
</div>

<form method="post" action="cancion_form.asp">
    <input type="hidden" name="accion" value="guardar">
    <input type="hidden" name="id" value="<%=idParaForm%>">

    <div class="form-panel">
        <div class="field">
            <label for="txtNombre">Nombre de la canci&oacute;n *</label>
            <input type="text" id="txtNombre" name="txtNombre" maxlength="50" value="<%=Esc(nombre)%>" placeholder="Ej: Adventure" required>
        </div>

        <div class="field">
            <label for="txtRuta">Ruta del archivo de audio *</label>
            <input type="text" id="txtRuta" name="txtRuta" maxlength="255" value="<%=Esc(ruta)%>" placeholder="~/songFiles/bensound-adventure.mp3" required>
            <div class="hint">Ruta registrada en la base de datos (TBL_SONG.SONG_PATH). Los archivos se buscan dentro de la carpeta SongFilesAndImages_resources.</div>
        </div>

        <div class="field">
            <label for="txtPlays">N&uacute;mero de reproducciones</label>
            <input type="number" id="txtPlays" name="txtPlays" min="0" step="1" value="<%=plays%>">
        </div>

        <div class="field">
            <label>Artistas que interpretan la canci&oacute;n</label>
            <div class="check-grid">
                <%
                Do While Not rsArtistas.EOF
                    Dim nomArt, selArt
                    nomArt = CStr(rsArtistas("ARTIST_NAME").Value)
                    If Not IsNull(rsArtistas("ARTIST_LASTNAME").Value) And Len(CStr(rsArtistas("ARTIST_LASTNAME").Value)) > 0 Then
                        nomArt = nomArt & " " & CStr(rsArtistas("ARTIST_LASTNAME").Value)
                    End If
                    selArt = dictArtistas.Exists("A" & CStr(rsArtistas("ID_ARTIST").Value))
                %>
                <label class="check-item<% If selArt Then Response.Write " sel" %>">
                    <input type="checkbox" name="chkArtistas" value="<%=rsArtistas("ID_ARTIST").Value%>"<% If selArt Then Response.Write " checked" %>>
                    <%=Esc(nomArt)%>
                </label>
                <%
                    rsArtistas.MoveNext
                Loop
                %>
            </div>
            <div class="hint">Se almacena en la tabla TBL_PLAYER.</div>
        </div>

        <div class="field">
            <label>&Aacute;lbumes que contienen la canci&oacute;n</label>
            <div class="check-grid">
                <%
                Do While Not rsAlbumes.EOF
                    Dim nomAlb, selAlb
                    nomAlb = CStr(rsAlbumes("TITLE_ALBUM").Value)
                    If Not IsNull(rsAlbumes("RELEASE_YEAR").Value) Then nomAlb = nomAlb & " (" & CStr(rsAlbumes("RELEASE_YEAR").Value) & ")"
                    selAlb = dictAlbumes.Exists("B" & CStr(rsAlbumes("ID_ALBUM").Value))
                %>
                <label class="check-item<% If selAlb Then Response.Write " sel" %>">
                    <input type="checkbox" name="chkAlbumes" value="<%=rsAlbumes("ID_ALBUM").Value%>"<% If selAlb Then Response.Write " checked" %>>
                    <%=Esc(nomAlb)%>
                </label>
                <%
                    rsAlbumes.MoveNext
                Loop
                %>
            </div>
            <div class="hint">Se almacena en la tabla TBL_TRACKLIST.</div>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-rosa">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"/></svg>
                Guardar canci&oacute;n
            </button>
            <a class="btn btn-ghost" href="../catalogo/canciones.asp">Cancelar</a>
        </div>
    </div>
</form>

<%
rsArtistas.Close
rsAlbumes.Close
Set rsArtistas = Nothing
Set rsAlbumes = Nothing
conn.Close
Set conn = Nothing

Pie
%>
