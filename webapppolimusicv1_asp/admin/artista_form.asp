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
' artista_form.asp · Formulario CRUD de artistas (ADMIN)
' Tabla: TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME,
'        ARTIST_IMAGE)
'=====================================================

RequiereAdmin()

Dim accion, idArtista
accion = Request.Form("accion") & ""
idArtista = Request.Form("id") & ""
If Len(idArtista) = 0 Then idArtista = Request.QueryString("id") & ""

Dim nombre, apellido, imagen
nombre = ""
apellido = ""
imagen = ""

Dim conn
Set conn = AbrirConexion()

'-----------------------------------------------------
' Acción: ELIMINAR
'-----------------------------------------------------
If accion = "eliminar" Then
    If EsEntero(idArtista) Then
        On Error Resume Next
        conn.BeginTrans
        ' Se quitan primero las asociaciones canción-artista
        Ejecutar conn, "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(CLng(idArtista))
        Ejecutar conn, "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(CLng(idArtista))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "No se pudo eliminar el artista."
        Else
            conn.CommitTrans
            On Error GoTo 0
            MensajeFlash "ok", "Artista eliminado correctamente."
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/artistas.asp"
        Response.End
    Else
        MensajeFlash "err", "Identificador de artista no válido."
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/artistas.asp"
        Response.End
    End If
End If

'-----------------------------------------------------
' Acción: GUARDAR
'-----------------------------------------------------
If accion = "guardar" Then
    nombre = Trim(Request.Form("txtNombre") & "")
    apellido = Trim(Request.Form("txtApellido") & "")
    imagen = Trim(Request.Form("txtImagen") & "")

    If Len(nombre) = 0 Then
        MensajeFlash "err", "El nombre del artista es obligatorio."
    Else
        On Error Resume Next
        If EsEntero(idArtista) Then
            Ejecutar conn, "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", Array(nombre, apellido, imagen, CLng(idArtista))
        Else
            Ejecutar conn, "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)", Array(nombre, apellido, imagen)
        End If
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "Ocurrió un error al guardar el artista."
        Else
            On Error GoTo 0
            MensajeFlash "ok", "Artista guardado correctamente."
            conn.Close
            Set conn = Nothing
            Response.Redirect "../catalogo/artistas.asp"
            Response.End
        End If
    End If
End If

'-----------------------------------------------------
' Cargar datos del artista (edición)
'-----------------------------------------------------
If EsEntero(idArtista) And accion <> "guardar" Then
    Dim rsC
    Set rsC = Consulta(conn, "SELECT ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(CLng(idArtista)))
    If Not rsC.EOF Then
        nombre = CStr(rsC("ARTIST_NAME").Value)
        apellido = CStr(rsC("ARTIST_LASTNAME").Value & "")
        imagen = CStr(rsC("ARTIST_IMAGE").Value & "")
    Else
        MensajeFlash "err", "El artista solicitado no existe."
        rsC.Close
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/artistas.asp"
        Response.End
    End If
    rsC.Close
    Set rsC = Nothing
End If

Dim editando, tituloAccion, idParaForm
editando = EsEntero(idArtista)
If editando Then
    tituloAccion = "Editar artista"
    idParaForm = CLng(idArtista)
Else
    tituloAccion = "Nuevo artista"
    idParaForm = 0
End If

Encabezado tituloAccion, "artistas"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Administraci&oacute;n &middot; artistas</div>
        <h2><%=tituloAccion%></h2>
        <p>Complete la informaci&oacute;n del artista.</p>
    </div>
    <div class="page-actions">
        <a class="btn btn-ghost" href="../catalogo/artistas.asp">Volver al listado</a>
    </div>
</div>

<form method="post" action="artista_form.asp">
    <input type="hidden" name="accion" value="guardar">
    <input type="hidden" name="id" value="<%=idParaForm%>">

    <div class="form-panel">
        <div class="field">
            <label for="txtNombre">Nombre *</label>
            <input type="text" id="txtNombre" name="txtNombre" maxlength="100" value="<%=Esc(nombre)%>" placeholder="Ej: Elvis" required>
        </div>
        <div class="field">
            <label for="txtApellido">Apellido</label>
            <input type="text" id="txtApellido" name="txtApellido" maxlength="100" value="<%=Esc(apellido)%>" placeholder="Ej: Presley">
        </div>
        <div class="field">
            <label for="txtImagen">Ruta de la imagen</label>
            <input type="text" id="txtImagen" name="txtImagen" maxlength="255" value="<%=Esc(imagen)%>" placeholder="~/ArtistImages/artist1.jpg">
            <div class="hint">Ruta registrada en la base de datos (TBL_ARTIST.ARTIST_IMAGE). Si el archivo no existe se muestra una imagen de respaldo.</div>
        </div>
        <div class="form-actions">
            <button type="submit" class="btn btn-rosa">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"/></svg>
                Guardar artista
            </button>
            <a class="btn btn-ghost" href="../catalogo/artistas.asp">Cancelar</a>
        </div>
    </div>
</form>

<%
conn.Close
Set conn = Nothing

Pie
%>
