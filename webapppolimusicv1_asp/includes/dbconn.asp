<%
'=========================================================================
' PoliMusic :: includes/dbconn.asp
' Conexion central ADO a SQL Server + helpers de acceso a datos
' parametrizados (previenen inyeccion SQL).
' ASP Clasico / VBScript - Proyecto BDD_PoliMusic
'
' NOTA: este include usa la funcion H() definida en helpers.asp.
' Al incluir en las paginas respetar el orden:
'   helpers.asp  ->  dbconn.asp  ->  auth.asp
'=========================================================================
' NOTA: sin "Option Explicit" (ver helpers.asp): la directiva solo puede
' aparecer una vez por unidad de compilacion.
'=========================================================================

'--- Configuracion de conexion -------------------------------------------
Const DB_SERVER   = "localhost,14333"
Const DB_NAME     = "BDD_PoliMusic"
Const DB_USER     = "sa"
Const DB_PASSWORD = "Politecnica1"

'--- Constantes ADO (equivalente adovbs.inc) ------------------------------
Const adCmdText        = 1
Const adCmdStoredProc  = 4
Const adParamInput     = 1
Const adVarChar        = 200
Const adInteger        = 3
Const adDouble         = 5
Const adBoolean        = 11
Const adDBTimeStamp    = 135
Const adExecuteNoRecords = 128
Const adStateOpen      = 1

Dim conn
Dim dbLastError

'-------------------------------------------------------------------------
' Abre la conexion probando varios proveedores OLE DB disponibles.
'-------------------------------------------------------------------------
Sub OpenDB()
    Dim providers, i, cs, connected
    providers = Array("SQLOLEDB", "SQLNCLI11", "MSOLEDBSQL")
    connected = False
    dbLastError = ""
    For i = 0 To UBound(providers)
        cs = "Provider=" & providers(i) & _
             ";Data Source=" & DB_SERVER & _
             ";Initial Catalog=" & DB_NAME & _
             ";User ID=" & DB_USER & _
             ";Password=" & DB_PASSWORD & ";"
        On Error Resume Next
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.CommandTimeout = 30
        conn.Open cs
        If Err.Number = 0 Then
            connected = True
            Err.Clear
            On Error GoTo 0
            Exit For
        End If
        dbLastError = dbLastError & "[" & providers(i) & "] " & Err.Description & " "
        Err.Clear
        Set conn = Nothing
        On Error GoTo 0
    Next
    If Not connected Then Call ShowDbError()
End Sub

'-------------------------------------------------------------------------
' Ejecuta una sentencia parametrizada y devuelve el Recordset abierto.
'   params: matriz 1-D de valores (se tipan por VarType) o Null
'-------------------------------------------------------------------------
Function DbQuery(sql, params)
    Dim cmd, rs, i, ptype, psize
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        If IsArray(params) Then
            For i = LBound(params) To UBound(params)
                ptype = AdoTypeOf(params(i))
                psize = ParamSize(params(i), ptype)
                .Parameters.Append .CreateParameter("p" & i, ptype, adParamInput, psize, params(i))
            Next
        End If
        Set rs = .Execute
    End With
    Set DbQuery = rs
End Function

'-------------------------------------------------------------------------
' Ejecuta una sentencia de accion (INSERT/UPDATE/DELETE) parametrizada.
' Devuelve la cantidad de filas afectadas.
'-------------------------------------------------------------------------
Function DbExec(sql, params)
    Dim cmd, affected
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        If IsArray(params) Then
            Dim i, ptype, psize
            For i = LBound(params) To UBound(params)
                ptype = AdoTypeOf(params(i))
                psize = ParamSize(params(i), ptype)
                .Parameters.Append .CreateParameter("p" & i, ptype, adParamInput, psize, params(i))
            Next
        End If
        .Execute affected, , adExecuteNoRecords
    End With
    DbExec = affected
End Function

'-------------------------------------------------------------------------
' Ejecuta un INSERT con clausula OUTPUT INSERTED.<campo> y devuelve el
' valor generado (identidad) como Long. Null si no se obtuvo nada.
' Necesario porque SQLOLEDB ejecuta las sentencias parametrizadas en un
' ambito hijo (sp_executesql), donde SCOPE_IDENTITY() devuelve NULL.
'-------------------------------------------------------------------------
Function DbInsertReturningId(sql, params)
    Dim cmd, rs
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        If IsArray(params) Then
            Dim i, ptype, psize
            For i = LBound(params) To UBound(params)
                ptype = AdoTypeOf(params(i))
                psize = ParamSize(params(i), ptype)
                .Parameters.Append .CreateParameter("p" & i, ptype, adParamInput, psize, params(i))
            Next
        End If
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        DbInsertReturningId = CLng(rs.Fields(0).Value)
    Else
        DbInsertReturningId = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

'-------------------------------------------------------------------------
' Devuelve el primer campo de la primera fila (o Null).
'-------------------------------------------------------------------------
Function DbScalar(sql, params)
    Dim rs
    Set rs = DbQuery(sql, params)
    If Not rs.EOF Then
        DbScalar = rs.Fields(0).Value
    Else
        DbScalar = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

'-------------------------------------------------------------------------
' Devuelve una matriz 2-D [campo, fila] con todo el resultado,
' o Empty cuando no hay filas.
'-------------------------------------------------------------------------
Function DbRows(sql, params)
    Dim rs
    Set rs = DbQuery(sql, params)
    If rs.EOF Then
        DbRows = Empty
    Else
        DbRows = rs.GetRows()
    End If
    rs.Close
    Set rs = Nothing
End Function

'-------------------------------------------------------------------------
' Tamano de buffer para un parametro, tolerando valores Null
' (Len(Null & "") provoca "Uso no valido de Null").
'-------------------------------------------------------------------------
Private Function ParamSize(v, ptype)
    If IsNull(v) Then
        ParamSize = 1
    ElseIf ptype = adVarChar Then
        ParamSize = Len(v & "") + 1
    Else
        ParamSize = 8
    End If
End Function

'-------------------------------------------------------------------------
' Tipa un valor VBScript al tipo ADO correspondiente.
'-------------------------------------------------------------------------
Private Function AdoTypeOf(v)
    Select Case VarType(v)
        Case vbString:        AdoTypeOf = adVarChar
        Case vbInteger:       AdoTypeOf = adInteger
        Case vbLong:          AdoTypeOf = adInteger
        Case vbSingle:        AdoTypeOf = adDouble
        Case vbDouble:        AdoTypeOf = adDouble
        Case vbBoolean:       AdoTypeOf = adBoolean
        Case vbDate:          AdoTypeOf = adDBTimeStamp
        Case vbEmpty:         AdoTypeOf = adVarChar
        Case vbNull:          AdoTypeOf = adVarChar
        Case Else:            AdoTypeOf = adVarChar
    End Select
End Function

'-------------------------------------------------------------------------
' Muestra un error amigable si la base de datos no esta disponible.
'-------------------------------------------------------------------------
Sub ShowDbError()
    Response.Charset = "UTF-8"
    Response.Write "<!DOCTYPE html><html lang='es'><head><meta charset='utf-8'>" & _
        "<title>Error de conexion</title>" & _
        "<link rel='stylesheet' href='css/styles.css'></head>" & _
        "<body class='fatal-page'><div class='fatal-box'>" & _
        "<h1>No se pudo conectar con la base de datos</h1>" & _
        "<p>Verifique que el servicio SQL Server este activo en <b>" & H(DB_SERVER) & _
        "</b> y que la base <b>" & H(DB_NAME) & "</b> acepte las credenciales configuradas.</p>" & _
        "<p class='fatal-detail'>" & H(dbLastError) & "</p>" & _
        "<p><a class='btn btn-primary' href='login.asp'>Reintentar</a></p>" & _
        "</div></body></html>"
    Response.End
End Sub

'--- Abrir la conexion al procesar este include ---------------------------
Call OpenDB()
%>
