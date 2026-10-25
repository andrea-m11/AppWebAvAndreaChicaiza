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
' playlist_form.asp · CRUD de playlists (usuario normal)
' Tabla: TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE,
'        TYPE, PHOTO)
' El usuario solo puede modificar SUS propias playlists.
'=====================================================

If CStr(Session("USER_TYPE") & "") = "1" Then
    Session("err") = "Los administradores no modifican playlists; solo las visualizan."
    Response.Redirect "playlists.asp"
    Response.End
End If

Dim accion, idPlaylist
accion = Request.Form("accion") & ""
idPlaylist = Request.Form("id") & ""
If Len(idPlaylist) = 0 Then idPlaylist = Request.QueryString("id") & ""

Dim titulo, tipo, foto, fecha
titulo = ""
tipo = 0
foto = ""
fecha = Date()

Dim conn
Set conn = AbrirConexion()

'-----------------------------------------------------
' Acción: ELIMINAR (solo si pertenece al usuario)
'-----------------------------------------------------
If accion = "eliminar" Then
    If EsEntero(idPlaylist) Then
        Dim dueno
        dueno = Escalar(conn, "SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(CLng(idPlaylist)))
        If IsNull(dueno) Or CLng(dueno) <> CLng(Session("ID_USER")) Then
            MensajeFlash "err", "Solo puedes eliminar tus propias playlists."
            conn.Close
            Set conn = Nothing
            Response.Redirect "playlists.asp"
            Response.End
        End If
        On Error Resume Next
        conn.BeginTrans
        Ejecutar conn, "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(CLng(idPlaylist))
        Ejecutar conn, "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(CLng(idPlaylist))
        If Err.Number <> 0 Then
            conn.RollbackTrans
            Err.Clear
            On Error GoTo 0
            MensajeFlash "err", "No se pudo eliminar la playlist."
        Else
            conn.CommitTrans
            On Error GoTo 0
            MensajeFlash "ok", "Playlist eliminada correctamente."
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "playlists.asp"
        Response.End
    Else
        MensajeFlash "err", "Identificador de playlist no válido."
        conn.Close
        Set conn = Nothing
        Response.Redirect "playlists.asp"
        Response.End
    End If
End If

'-----------------------------------------------------
' Acción: GUARDAR
'-----------------------------------------------------
If accion = "guardar" Then
    titulo = Trim(Request.Form("txtTitulo") & "")
    tipo = Trim(Request.Form("selTipo") & "")
    foto = Trim(Request.Form("txtFoto") & "")
    Dim fechaTxt
    fechaTxt = Trim(Request.Form("txtFecha") & "")

    If Len(titulo) = 0 Then
        MensajeFlash "err", "El título de la playlist es obligatorio."
    ElseIf Not EsEntero(tipo) Then
        MensajeFlash "err", "Debe seleccionar un tipo de playlist válido."
    ElseIf Len(fechaTxt) = 0 Or Not IsDate(fechaTxt) Then
        MensajeFlash "err", "La fecha de creación no es válida."
    Else
        Dim fVal
        fVal = CDate(fechaTxt)
        If EsEntero(idPlaylist) Then
            Dim dueno2
            dueno2 = Escalar(conn, "SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(CLng(idPlaylist)))
            If IsNull(dueno2) Or CLng(dueno2) <> CLng(Session("ID_USER")) Then
                MensajeFlash "err", "Solo puedes editar tus propias playlists."
                conn.Close
                Set conn = Nothing
                Response.Redirect "playlists.asp"
                Response.End
            End If
            Ejecutar conn, "UPDATE TBL_PLAYLIST SET TITLE = ?, CREATION_DATE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?", Array(titulo, fVal, CLng(tipo), foto, CLng(idPlaylist))
        Else
            Ejecutar conn, "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (?, ?, ?, ?, ?)", Array(CLng(Session("ID_USER")), titulo, fVal, CLng(tipo), foto)
        End If
        MensajeFlash "ok", "Playlist guardada correctamente."
        conn.Close
        Set conn = Nothing
        Response.Redirect "playlists.asp"
        Response.End
    End If
End If

'-----------------------------------------------------
' Cargar datos de la playlist (edición)
'-----------------------------------------------------
If EsEntero(idPlaylist) And accion <> "guardar" Then
    Dim rsC
    Set rsC = Consulta(conn, "SELECT ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(CLng(idPlaylist)))
    If rsC.EOF Then
        MensajeFlash "err", "La playlist solicitada no existe."
        rsC.Close
        conn.Close
        Set conn = Nothing
        Response.Redirect "playlists.asp"
        Response.End
    End If
    If CLng(rsC("ID_USER").Value) <> CLng(Session("ID_USER")) Then
        MensajeFlash "err", "Solo puedes editar tus propias playlists."
        rsC.Close
        conn.Close
        Set conn = Nothing
        Response.Redirect "playlists.asp"
        Response.End
    End If
    titulo = CStr(rsC("TITLE").Value)
    tipo = CLng(rsC("TYPE").Value)
    foto = CStr(rsC("PHOTO").Value & "")
    If Not IsNull(rsC("CREATION_DATE").Value) Then
        fecha = rsC("CREATION_DATE").Value
    End If
    rsC.Close
    Set rsC = Nothing
End If

Dim editando, tipoSel, tituloAccion, idParaForm
editando = EsEntero(idPlaylist)

' Valor seguro del tipo para el selector (evita CLng de cadena vacía)
tipoSel = 0
If EsEntero(tipo) Then tipoSel = CLng(tipo)

If editando Then
    tituloAccion = "Editar playlist"
    idParaForm = CLng(idPlaylist)
Else
    tituloAccion = "Nueva playlist"
    idParaForm = 0
End If

Dim fTxt
fTxt = Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "-" & Right("0" & Day(fecha), 2)

Encabezado tituloAccion, "playlists"

MostrarFlash
%>

<div class="page-head">
    <div>
        <div class="eyebrow">Mis playlists</div>
        <h2><%=tituloAccion%></h2>
        <p>Los datos se almacenan en la tabla TBL_PLAYLIST.</p>
    </div>
    <div class="page-actions">
        <a class="btn btn-ghost" href="playlists.asp">Volver a mis playlists</a>
    </div>
</div>

<form method="post" action="playlist_form.asp">
    <input type="hidden" name="accion" value="guardar">
    <input type="hidden" name="id" value="<%=idParaForm%>">

    <div class="form-panel">
        <div class="field">
            <label for="txtTitulo">T&iacute;tulo *</label>
            <input type="text" id="txtTitulo" name="txtTitulo" maxlength="50" value="<%=Esc(titulo)%>" placeholder="Ej: Para estudiar" required>
        </div>

        <div class="field">
            <label for="selTipo">Tipo de playlist *</label>
            <select id="selTipo" name="selTipo" required>
                <option value="">-- Seleccione --</option>
                <option value="0"<% If tipoSel = 0 Then Response.Write " selected" %>>P&uacute;blica</option>
                <option value="1"<% If tipoSel = 1 Then Response.Write " selected" %>>Privada</option>
            </select>
            <div class="hint">Valores definidos en TBL_CATALOG (0 = Public, 1 = Private).</div>
        </div>

        <div class="field">
            <label for="txtFecha">Fecha de creaci&oacute;n *</label>
            <input type="date" id="txtFecha" name="txtFecha" value="<%=fTxt%>" required>
        </div>

        <div class="field">
            <label for="txtFoto">Ruta de la portada</label>
            <input type="text" id="txtFoto" name="txtFoto" maxlength="255" value="<%=Esc(foto)%>" placeholder="~/playlistImages/playlist1.png">
            <div class="hint">Ruta registrada en la base de datos (TBL_PLAYLIST.PHOTO). Si el archivo no existe se muestra una imagen de respaldo.</div>
        </div>

        <div class="form-actions">
            <button type="submit" class="btn btn-rosa">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"/></svg>
                Guardar playlist
            </button>
            <a class="btn btn-ghost" href="playlists.asp">Cancelar</a>
        </div>
    </div>
</form>

<%
conn.Close
Set conn = Nothing

Pie
%>
