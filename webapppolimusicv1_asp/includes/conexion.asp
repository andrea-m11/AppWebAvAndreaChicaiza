<%
'=====================================================
' conexion.asp
' Capa de acceso a datos para PoliMusic (ASP Clásico + VBScript)
' Base de datos: BDD_PoliMusic en SQL Server localhost
' Credenciales: usr_polimusic / Politecnica1 (script BDD_PoliMusic.sql)
'
' API de la capa (igual en todas las páginas):
'   Consulta(conn, sql, args)  -> ADODB.Recordset de solo lectura
'   Ejecutar(conn, sql, args)  -> número de filas afectadas
'   Escalar (conn, sql, args)  -> valor único (o Null)
'
' Los marcadores "?" de la sentencia SQL se sustituyen por
' literales escapados: cadenas con comillas duplicadas, enteros
' sin comillas, fechas en formato ISO 120 y NULL explícito.
' Así no se usa ADODB.Command ni el enlace de parámetros del
' proveedor (origen del error 80040e21 con SQLOLEDB en IIS) y
' se mantiene la protección contra inyección de SQL.
'=====================================================

'--- Constantes ADO ---
Const AD_USE_CLIENT    = 3
Const AD_OPEN_STATIC   = 3
Const AD_LOCK_READONLY = 1

'--- Cadenas de conexión: se prueban en este orden ---
Const CADENA_MSOLEDB  = "Provider=MSOLEDBSQL;Server=localhost;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;Encrypt=no;TrustServerCertificate=yes;"
Const CADENA_SQLNCLI  = "Provider=SQLNCLI11;Server=localhost;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;Encrypt=no;TrustServerCertificate=yes;"
Const CADENA_SQLOLEDB = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=BDD_PoliMusic;User ID=usr_polimusic;Password=Politecnica1;"

Function AbrirConexion()
    Dim conn, intento, ok, lista, i
    ok = False
    lista = Array(CADENA_MSOLEDB, CADENA_SQLNCLI, CADENA_SQLOLEDB)
    For i = 0 To UBound(lista)
        Err.Clear
        On Error Resume Next
        Set intento = Server.CreateObject("ADODB.Connection")
        intento.CommandTimeout = 30
        intento.Open lista(i)
        If Err.Number = 0 Then
            Set conn = intento
            ok = True
        Else
            Set intento = Nothing
        End If
        On Error GoTo 0
        If ok Then
            Exit For
        End If
    Next
    If Not ok Then
        Response.Write "<div style='font-family:Segoe UI,Arial;background:#12121c;color:#ff8fa3;padding:24px;border:1px solid #ff6b8a;border-radius:12px;margin:40px auto;max-width:640px'>"
        Response.Write "<h2 style='margin:0 0 8px'>No se pudo conectar a la base de datos</h2>"
        Response.Write "<p style='color:#b9b3d6'>Verifique que SQL Server esté ejecutándose en <b>localhost</b>, que la base <b>BDD_PoliMusic</b> exista (ejecute el script BDD_PoliMusic.sql) y que el login <b>usr_polimusic</b> tenga acceso.</p>"
        Response.Write "<p style='color:#8b85a8;font-size:12px'>Detalle técnico: " & Server.HTMLEncode(Err.Description) & "</p>"
        Response.Write "</div>"
        Response.End
    End If
    Set AbrirConexion = conn
End Function

'--- Convierte un valor VBScript en un literal SQL seguro ---
Function LiteralSQL(ByVal valor)
    Dim s, d
    If IsNull(valor) Or IsEmpty(valor) Then
        LiteralSQL = "NULL"
        Exit Function
    End If
    Select Case VarType(valor)
        Case vbInteger, vbLong, vbByte
            LiteralSQL = CStr(CLng(valor))
        Case vbBoolean
            If valor Then
                LiteralSQL = "1"
            Else
                LiteralSQL = "0"
            End If
        Case vbDate
            d = Year(valor) & "-" & Right("0" & Month(valor), 2) & "-" & Right("0" & Day(valor), 2)
            d = d & " " & Right("0" & Hour(valor), 2) & ":" & Right("0" & Minute(valor), 2) & ":" & Right("0" & Second(valor), 2)
            LiteralSQL = "CONVERT(datetime, '" & d & "', 120)"
        Case vbDouble, vbSingle, vbCurrency
            s = CStr(valor)
            s = Replace(s, ",", ".")
            LiteralSQL = s
        Case Else
            s = CStr(valor)
            LiteralSQL = "'" & Replace(s, "'", "''") & "'"
    End Select
End Function

'--- Reemplaza cada marcador "?" por su literal escapado ---
' Recorre la sentencia caracter a caracter para no tocar nunca
' un "?" que este dentro de un literal de texto ('...').
Function SQLConLiterales(ByVal sql, ByVal args)
    Dim vals, i, k, ch, sig, dentro, res
    If IsEmpty(args) Or IsNull(args) Then
        SQLConLiterales = sql
        Exit Function
    End If
    If IsArray(args) Then
        vals = args
    Else
        vals = Array(args)
    End If
    i = 0
    dentro = False
    res = ""
    k = 1
    Do While k <= Len(sql)
        ch = Mid(sql, k, 1)
        If ch = "'" Then
            sig = Mid(sql, k + 1, 1)
            If dentro And sig = "'" Then
                res = res & "''"
                k = k + 2
            Else
                dentro = Not dentro
                res = res & ch
                k = k + 1
            End If
        ElseIf ch = "?" And Not dentro Then
            If i <= UBound(vals) Then
                res = res & LiteralSQL(vals(i))
            Else
                res = res & "NULL"
            End If
            i = i + 1
            k = k + 1
        Else
            res = res & ch
            k = k + 1
        End If
    Loop
    SQLConLiterales = res
End Function

'--- Ejecuta una consulta y devuelve el Recordset ---
' args: matriz de valores para los marcadores (?). Pasar Array() si no hay.
Function Consulta(ByVal conn, ByVal sql, ByVal args)
    Dim rs
    sql = SQLConLiterales(sql, args)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = AD_USE_CLIENT
    rs.Open sql, conn, AD_OPEN_STATIC, AD_LOCK_READONLY
    Set Consulta = rs
End Function

'--- Ejecuta INSERT/UPDATE/DELETE y devuelve las filas afectadas ---
' args: matriz de valores para los marcadores (?). Pasar Array() si no hay.
Function Ejecutar(ByVal conn, ByVal sql, ByVal args)
    Dim afectados
    sql = SQLConLiterales(sql, args)
    conn.Execute sql, afectados
    Ejecutar = afectados
End Function

'--- Devuelve un único valor (primera columna de la primera fila) ---
' args: matriz de valores para los marcadores (?). Pasar Array() si no hay.
Function Escalar(ByVal conn, ByVal sql, ByVal args)
    Dim rs
    Set rs = Consulta(conn, sql, args)
    If rs.EOF Then
        Escalar = Null
    Else
        If IsNull(rs(0).Value) Then
            Escalar = Null
        Else
            Escalar = rs(0).Value
        End If
    End If
    rs.Close
    Set rs = Nothing
End Function
%>
