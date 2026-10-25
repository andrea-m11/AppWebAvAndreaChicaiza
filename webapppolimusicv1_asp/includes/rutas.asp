<%
'=====================================================
' rutas.asp
' Resuelve las rutas guardadas en la base de datos
' (formato '~/songFiles/x.mp3') a archivos reales dentro
' de la carpeta SongFilesAndImages_resources/.
' Si el archivo no existe se usa una imagen/canción de
' respaldo para que la aplicación nunca muestre enlaces rojos.
'=====================================================

'--- Convierte una ruta de la BD a una ruta web válida ---
Function RutaAsset(ByVal rutaDB, ByVal tipo)
    Dim limpia, carpeta, archivo, base, candidata, fso
    tipo = LCase(CStr(tipo))

    rutaDB = CStr(rutaDB & "")
    limpia = Replace(Replace(Trim(rutaDB), "~/", ""), "\", "/")
    If Len(limpia) = 0 Then limpia = "desconocido"

    carpeta = ""
    archivo = limpia
    If InStr(limpia, "/") > 0 Then
        carpeta = LCase(Left(limpia, InStr(limpia, "/") - 1))
        archivo = Mid(limpia, InStr(limpia, "/") + 1)
    End If

    ' Mapeo de carpetas registradas en la BD a la carpeta real de recursos
    Select Case carpeta
        Case "songfiles"
            base = "SongFilesAndImages_resources/SongFiles/"
        Case "userimages"
            base = "SongFilesAndImages_resources/UserImages/"
        Case "playlistimages"
            base = "SongFilesAndImages_resources/PlaylistImages/"
        Case Else
            base = ""
    End Select

    If Len(base) > 0 And Len(archivo) > 0 Then
        candidata = base & archivo
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(Server.MapPath(RaizVirtual() & candidata)) Then
            RutaAsset = RaizVirtual() & candidata
            Exit Function
        End If
    End If

    ' Respaldos cuando el archivo no existe en disco
    Select Case tipo
        Case "cancion"
            RutaAsset = RaizVirtual() & "SongFilesAndImages_resources/SongFiles/bensound-adventure.mp3"
        Case "album"
            RutaAsset = RaizVirtual() & "SongFilesAndImages_resources/PlaylistImages/default.jpg"
        Case Else
            RutaAsset = RaizVirtual() & "SongFilesAndImages_resources/UserImages/default.jpg"
    End Select
End Function

'--- Portada de una canción: toma la imagen del primer álbum que la contiene ---
Function PortadaCancion(ByVal conn, ByVal idCancion)
    Dim ruta
    ruta = Escalar(conn, "SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM WHERE t.ID_SONG = ?", Array(CLng(idCancion)))
    If IsNull(ruta) Or Len(CStr(ruta & "")) = 0 Then
        PortadaCancion = RaizVirtual() & "SongFilesAndImages_resources/PlaylistImages/default.jpg"
    Else
        PortadaCancion = RutaAsset(ruta, "album")
    End If
End Function

'--- Artistas de una canción (cadena separada por comas) ---
Function ArtistasCancion(ByVal conn, ByVal idCancion)
    Dim rs, lista, primero
    lista = ""
    primero = True
    Set rs = Consulta(conn, "SELECT a.ARTIST_NAME, a.ARTIST_LASTNAME FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = ? ORDER BY a.ARTIST_NAME", Array(CLng(idCancion)))
    Do While Not rs.EOF
        If Not primero Then lista = lista & ", "
        lista = lista & CStr(rs("ARTIST_NAME").Value)
        If Not IsNull(rs("ARTIST_LASTNAME").Value) And Len(CStr(rs("ARTIST_LASTNAME").Value)) > 0 Then
            lista = lista & " " & CStr(rs("ARTIST_LASTNAME").Value)
        End If
        primero = False
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If Len(lista) = 0 Then lista = "Artista desconocido"
    ArtistasCancion = lista
End Function
%>
