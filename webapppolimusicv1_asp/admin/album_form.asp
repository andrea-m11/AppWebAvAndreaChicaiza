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
' album_form.asp · Formulario CRUD de álbumes (ADMIN)
' Tabla: TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR,
'        ALBUM_IMAGE_PATH)
'=====================================================

RequiereAdmin()

Dim accion, idAlbum
accion = Request.Form("accion") & ""
idAlbum = Request.Form("id") & ""
If Len(idAlbum) = 0 Then idAlbum = Request.QueryString("id") & ""

Dim titulo, anio, imagen
titulo = ""
anio = ""
imagen = ""

Dim conn
Set conn = AbrirConexion()

'-----------------------------------------------------
' Acción: ELIMINAR
'-----------------------------------------------------
If accion = "eliminar" Then
    If EsEntero(idAlbum) Then
        On Error Resume Next
        conn.BeginTrans
        ' Se quitan primero las asociaciones álbum-canción
        Ejecutar conn, "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(CLng(idAlbum))
        Ejecutar conn, "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(CLng(idAlbum))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "No se pudo eliminar el álbum."
        Else
            conn.CommitTrans
            On Error GoTo 0
            MensajeFlash "ok", "Álbum eliminado correctamente."
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/albumes.asp"
        Response.End
    Else
        MensajeFlash "err", "Identificador de álbum no válido."
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/albumes.asp"
        Response.End
    End If
End If

'-----------------------------------------------------
' Acción: GUARDAR
'-----------------------------------------------------
If accion = "guardar" Then
    titulo = Trim(Request.Form("txtTitulo") & "")
    anio = Trim(Request.Form("txtAnio") & "")
    imagen = Trim(Request.Form("txtImagen") & "")

    If Len(titulo) = 0 Then
        MensajeFlash "err", "El título del álbum es obligatorio."
    ElseIf Len(anio) > 0 And Not EsEntero(anio) Then
        MensajeFlash "err", "El año de lanzamiento debe ser un número entero."
    Else
        Dim valAnio
        If Len(anio) > 0 Then valAnio = CLng(anio) Else valAnio = Null

        On Error Resume Next
        If EsEntero(idAlbum) Then
            Ejecutar conn, "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", Array(titulo, valAnio, imagen, CLng(idAlbum))
        Else
            Ejecutar conn, "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?)", Array(titulo, valAnio, imagen)
        End If
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "Ocurrió un error al guardar el álbum."
        Else
            On Error GoTo 0
            MensajeFlash "ok", "Álbum guardado correctamente."
            conn.Close
            Set conn = Nothing
            Response.Redirect "../catalogo/albumes.asp"
            Response.End
        End If
    End If
End If

'-----------------------------------------------------
' Cargar datos del álbum (edición)
'-----------------------------------------------------
If EsEntero(idAlbum) And accion <> "guardar" Then
    Dim rsC
    Set rsC = Consulta(conn, "SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(CLng(idAlbum)))
    If Not rsC.EOF Then
        titulo = CStr(rsC("TITLE_ALBUM").Value)
        If IsNull(rsC("RELEASE_YEAR").Value) Then anio = "" Else anio = CStr(rsC("RELEASE_YEAR").Value)
        imagen = CStr(rsC("ALBUM_IMAGE_PATH").Value & "")
    Else
        MensajeFlash "err", "El álbum solicitado no existe."
        rsC.Close
        conn.Close
        Set conn = Nothing
        Response.Redirect "../catalogo/albumes.asp"
        Response.End
    End If
    rsC.Close
    Set rsC = Nothing
End If

Dim editando, tituloAccion, idParaForm
editando = EsEntero(idAlbum)
If editando Then
    tituloAccion = "Editar álbum"
    idParaForm = CLng(idAlbum)
Else
    tituloAccion = "Nuevo álbum"
    idParaForm = 0
End If

Encabezado tituloAccion, "albumes"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Administraci&oacute;n &middot; &aacute;lbumes</div>
        <h2><%=tituloAccion%></h2>
        <p>Complete la informaci&oacute;n del álbum. Las canciones del tracklist se administran desde la opci&oacute;n &laquo;Tracklist&raquo;.</p>
    </div>
    <div class="page-actions">
        <a class="btn btn-ghost" href="../catalogo/albumes.asp">Volver al listado</a>
    </div>
</div>

<form method="post" action="album_form.asp">
    <input type="hidden" name="accion" value="guardar">
    <input type="hidden" name="id" value="<%=idParaForm%>">

    <div class="form-panel">
        <div class="field">
            <label for="txtTitulo">T&iacute;tulo del &aacute;lbum *</label>
            <input type="text" id="txtTitulo" name="txtTitulo" maxlength="50" value="<%=Esc(titulo)%>" placeholder="Ej: Rock Album" required>
        </div>
        <div class="field">
            <label for="txtAnio">A&ntilde;o de lanzamiento</label>
            <input type="number" id="txtAnio" name="txtAnio" min="1900" max="2100" step="1" value="<%=Esc(anio)%>" placeholder="Ej: 1976">
        </div>
        <div class="field">
            <label for="txtImagen">Ruta de la portada</label>
            <input type="text" id="txtImagen" name="txtImagen" maxlength="255" value="<%=Esc(imagen)%>" placeholder="~/albumImages/album1.png">
            <div class="hint">Ruta registrada en la base de datos (TBL_ALBUM.ALBUM_IMAGE_PATH). Si el archivo no existe se muestra una portada de respaldo.</div>
        </div>
        <div class="form-actions">
            <button type="submit" class="btn btn-rosa">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"/></svg>
                Guardar álbum
            </button>
            <a class="btn btn-ghost" href="../catalogo/albumes.asp">Cancelar</a>
        </div>
    </div>
</form>

<%
conn.Close
Set conn = Nothing

Pie
%>
