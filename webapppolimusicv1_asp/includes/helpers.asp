<%
'=========================================================================
' PoliMusic :: includes/helpers.asp
' Utilidades compartidas: encoding, resolucion de rutas de recursos,
' mensajes flash, manejo de sesion y SHA-1 puro en VBScript.
' ASP Clasico / VBScript - Proyecto BDD_PoliMusic
'=========================================================================
' NOTA: no declarar "Option Explicit" aqui. En ASP clasico todos los
' includes se concatenan en una misma unidad de compilacion y la directiva
' solo puede aparecer una vez; la declara cada pagina antes de incluir.
'=========================================================================

'--- Carpeta local donde viven los recursos oficiales ---------------------
Const RES_FOLDER = "SongFilesAndImages_resources"

'=========================================================================
' Encoding de salida
'=========================================================================
Function H(s)
    H = Server.HTMLEncode(s & "")
End Function

Function Attr(s)
    Attr = Server.HTMLEncode(s & "")
End Function

' IIf no existe en VBScript
Function IIf(cond, a, b)
    If cond Then IIf = a Else IIf = b
End Function

'=========================================================================
' Sesion
'=========================================================================
Function CurrentUserId()
    Dim t
    t = Session("USER_ID") & ""
    If IsNumeric(t) Then CurrentUserId = CLng(t) Else CurrentUserId = 0
End Function

Function CurrentUserName()
    CurrentUserName = Session("USER_NAME") & ""
End Function

Function IsAdmin()
    Dim t
    t = Session("USER_TYPE") & ""
    If IsNumeric(t) Then IsAdmin = (CInt(t) = 1) Else IsAdmin = False
End Function

' El nombre del rol se toma de TBL_CATALOG (COD_CATALOG_PARENT='1'):
'   COD_CATALOG '0' = USER  /  '1' = ADMIN
Function RoleName()
    If IsAdmin() Then RoleName = "ADMIN" Else RoleName = "USER"
End Function

Function RoleLabel()
    If IsAdmin() Then RoleLabel = "Administrador" Else RoleLabel = "Usuario normal"
End Function

'=========================================================================
' Mensajes flash (se guardan en Session y se muestran una sola vez)
'=========================================================================
Sub SetFlash(msg, ftype)
    Session("FLASH_MSG") = msg
    Session("FLASH_TYPE") = ftype
End Sub

Sub ClearFlash()
    On Error Resume Next
    Session.Contents.Remove "FLASH_MSG"
    Session.Contents.Remove "FLASH_TYPE"
    On Error GoTo 0
End Sub

'=========================================================================
' Sistema de archivos
'=========================================================================
Function FileExists(physPath)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(physPath)
    Set fso = Nothing
End Function

' Lista los nombres de archivo de una subcarpeta de recursos filtrados
' por extension. Devuelve una matriz (posiblemente vacia).
Function ListResourceFiles(subfolder, ext)
    Dim fso, folder, f, files
    files = Empty
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(Server.MapPath(RES_FOLDER & "/" & subfolder))
    If Err.Number = 0 Then
        For Each f In folder.Files
            If LCase(fso.GetExtensionName(f.Name)) = LCase(ext) Then
                If IsEmpty(files) Then
                    files = Array(f.Name)
                Else
                    ReDim Preserve files(UBound(files) + 1)
                    files(UBound(files)) = f.Name
                End If
            End If
        Next
    Else
        Err.Clear
    End If
    Set folder = Nothing
    Set fso = Nothing
    On Error GoTo 0
    ListResourceFiles = files
End Function

'=========================================================================
' Resolucion de rutas
' La base guarda rutas estilo ASP.NET: ~/songFiles/x.mp3
' Se traducen a la carpeta local de recursos y, si el archivo fisico
' no existe, se usa una imagen por defecto (fallbackUrl).
'=========================================================================
Function ResourceUrl(dbPath, fallbackUrl)
    Dim raw, parts, folder, fname, url
    raw = Trim(dbPath & "")
    If raw = "" Then
        ResourceUrl = fallbackUrl
        Exit Function
    End If
    If LCase(Left(raw, 2)) = "~/" Then raw = Mid(raw, 3)
    parts = Split(raw, "/")
    folder = parts(0)
    fname = parts(UBound(parts))
    Select Case LCase(folder)
        Case "songfiles":      folder = "SongFiles"
        Case "playlistimages": folder = "PlaylistImages"
        Case "userimages":     folder = "UserImages"
        Case "albumimages":    folder = "AlbumImages"
        Case "artistimages":   folder = "ArtistImages"
        Case Else:             folder = parts(0)
    End Select
    url = RES_FOLDER & "/" & folder & "/" & fname
    If FileExists(Server.MapPath(url)) Then
        ResourceUrl = url
    Else
        ResourceUrl = fallbackUrl
    End If
End Function

' URL del audio de una cancion. Devuelve "" si el MP3 no esta disponible
' localmente (el reproductor muestra el estado "no disponible").
Function SongUrl(dbPath)
    Dim raw, parts, folder, fname, url
    raw = Trim(dbPath & "")
    If raw = "" Then
        SongUrl = ""
        Exit Function
    End If
    If LCase(Left(raw, 2)) = "~/" Then raw = Mid(raw, 3)
    parts = Split(raw, "/")
    folder = parts(0)
    fname = parts(UBound(parts))
    Select Case LCase(folder)
        Case "songfiles": folder = "SongFiles"
        Case Else:        folder = parts(0)
    End Select
    url = RES_FOLDER & "/" & folder & "/" & fname
    If FileExists(Server.MapPath(url)) Then
        SongUrl = url
    Else
        SongUrl = ""
    End If
End Function

Function DefaultCoverUrl()
    DefaultCoverUrl = "assets/img/default-cover.svg"
End Function

Function DefaultAvatarUrl()
    DefaultAvatarUrl = "assets/img/default-avatar.svg"
End Function

'=========================================================================
' Mapas de relaciones para el listado de canciones
'   LoadArtistMap -> Dictionary ID_SONG => "Artista 1, Artista 2"
'   LoadAlbumMap  -> Dictionary ID_SONG => Array(titulo, imagenUrl)
'=========================================================================
Function LoadArtistMap()
    Dim d, rs, key, nm
    Set d = Server.CreateObject("Scripting.Dictionary")
    Set rs = DbQuery("SELECT p.ID_SONG, a.ARTIST_NAME, a.ARTIST_LASTNAME " & _
                     "FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
                     "ORDER BY p.ID_SONG, a.ID_ARTIST", Null)
    Do Until rs.EOF
        key = CStr(rs("ID_SONG"))
        nm = Trim(rs("ARTIST_NAME") & "")
        If Trim(rs("ARTIST_LASTNAME") & "") <> "" Then nm = nm & " " & Trim(rs("ARTIST_LASTNAME"))
        If d.Exists(key) Then
            d(key) = d(key) & ", " & nm
        Else
            d.Add key, nm
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set LoadArtistMap = d
End Function

Function LoadAlbumMap()
    Dim d, rs, key, title, img
    Set d = Server.CreateObject("Scripting.Dictionary")
    Set rs = DbQuery("SELECT t.ID_SONG, al.TITLE_ALBUM, al.ALBUM_IMAGE_PATH " & _
                     "FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
                     "ORDER BY t.ID_SONG, al.ID_ALBUM", Null)
    Do Until rs.EOF
        key = CStr(rs("ID_SONG"))
        title = Trim(rs("TITLE_ALBUM") & "")
        img = ResourceUrl(rs("ALBUM_IMAGE_PATH") & "", DefaultCoverUrl())
        If d.Exists(key) Then
            If title <> "" Then
                If d(key)(0) = "" Then d(key)(0) = title
            End If
        Else
            d.Add key, Array(title, img)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set LoadAlbumMap = d
End Function

'=========================================================================
' SHA-1 puro en VBScript (sin componentes externos).
' Necesario porque algunas cuentas guardan el hash SHA-1 hexadecimal
' del password (p.ej. username1 en el script de la base).
' Devuelve 40 caracteres hexadecimales en mayusculas.
'=========================================================================
'-------------------------------------------------------------------------
' Convierte texto Unicode a una cadena de bytes ANSI (Latin-1).
' StrConv no existe en VBScript (es de VBA), por eso se implementa aqui
' con AscW/ChrB. Caracteres fuera de Latin-1 se reemplazan por '?'.
'-------------------------------------------------------------------------
Private Function TextToAnsiBytes(s)
    Dim i, n, cp, out
    out = ""
    n = Len(s)
    For i = 1 To n
        cp = AscW(Mid(s, i, 1))
        If cp > 255 Then cp = 63   ' '?'
        out = out & ChrB(cp)
    Next
    TextToAnsiBytes = out
End Function

Function SHA1(sText)
    Dim h0, h1, h2, h3, h4
    h0 = 1732584193   ' 0x67452301
    h1 = 4023233417   ' 0xEFCDAB89
    h2 = 2562383102   ' 0x98BADCFE
    h3 = 271733878    ' 0x10325476
    h4 = 3285377520   ' 0xC3D2E1F0

    '--- mensaje a bytes (ANSI) + padding PKCS#7-style de SHA-1 -----------
    Dim msg, ml, needZeros, bits, k
    msg = TextToAnsiBytes(sText)
    ml = LenB(msg)
    needZeros = 55 - (ml - 64 * Int(ml / 64)) + 64      ' equivalente a (55 - ml Mod 64 + 64) Mod 64
    needZeros = needZeros - 64 * Int(needZeros / 64)    ' sin usar Mod (desborda con valores grandes)
    msg = msg & ChrB(&H80)
    Do While needZeros > 0
        msg = msg & ChrB(0)
        needZeros = needZeros - 1
    Loop
    bits = ml * 8
    For k = 1 To 4
        msg = msg & ChrB(0)   ' 32 bits altos del largo (siempre 0 aqui)
    Next
    For k = 3 To 0 Step -1
        msg = msg & ChrB(Int(bits / (256 ^ k)) - 256 * Int(bits / (256 ^ (k + 1))))
    Next

    '--- procesar bloques de 64 bytes -------------------------------------
    Dim w(79), chunk, j, a, b, c, d, e, f, kk, temp
    For chunk = 1 To LenB(msg) Step 64
        For j = 0 To 15
            w(j) = CLng(AscB(MidB(msg, chunk + j * 4, 1)))     * 16777216 + _
                   CLng(AscB(MidB(msg, chunk + j * 4 + 1, 1))) * 65536 + _
                   CLng(AscB(MidB(msg, chunk + j * 4 + 2, 1))) * 256 + _
                   CLng(AscB(MidB(msg, chunk + j * 4 + 3, 1)))
        Next
        For j = 16 To 79
            w(j) = Rotl32(Xor32(Xor32(w(j - 3), w(j - 8)), Xor32(w(j - 14), w(j - 16))), 1)
        Next
        a = h0: b = h1: c = h2: d = h3: e = h4
        For j = 0 To 79
            If j < 20 Then
                f = Or32(And32(b, c), And32(Not32(b), d))
                kk = 1518500249   ' 0x5A827999
            ElseIf j < 40 Then
                f = Xor32(Xor32(b, c), d)
                kk = 1859775393   ' 0x6ED9EBA1
            ElseIf j < 60 Then
                f = Or32(Or32(And32(b, c), And32(b, d)), And32(c, d))
                kk = 2400959708   ' 0x8F1BBCDC
            Else
                f = Xor32(Xor32(b, c), d)
                kk = 3395469782   ' 0xCA62C1D6
            End If
            temp = Add32(Add32(Add32(Rotl32(a, 5), f), Add32(e, kk)), w(j))
            e = d: d = c: c = Rotl32(b, 30): b = a: a = temp
        Next
        h0 = Add32(h0, a): h1 = Add32(h1, b): h2 = Add32(h2, c)
        h3 = Add32(h3, d): h4 = Add32(h4, e)
    Next

    SHA1 = Hex32(h0) & Hex32(h1) & Hex32(h2) & Hex32(h3) & Hex32(h4)
End Function

'--- aritmetica de 32 bits sin signo sobre doubles -----------------------
' NOTA: el operador Mod de VBScript coerciona sus operandos a Long y
' desborda (error 6) con valores >= 2^31, por eso en todas las funciones
' de abajo se extraen digitos con restas/Int en lugar de Mod.
Private Function Add32(a, b)
    Dim s
    s = CDbl(a) + CDbl(b)
    Add32 = s - Int(s / 4294967296) * 4294967296
End Function

' Devuelve el bit i (0..31) del valor de 32 bits v (rango 0..2^32-1).
Private Function Bit32(v, i)
    Bit32 = Int(v / (2 ^ i)) - 2 * Int(v / (2 ^ (i + 1)))
End Function

Private Function Rotl32(x, n)
    Dim hi, lo, div
    div = 2 ^ (32 - n)
    hi = Int(x / div)
    lo = x - Int(x / div) * div
    Rotl32 = Add32(lo * (2 ^ n), hi)
End Function

Private Function And32(a, b)
    Dim r, i, bit
    r = 0
    For i = 0 To 31
        bit = 2 ^ i
        If Bit32(a, i) = 1 And Bit32(b, i) = 1 Then r = r + bit
    Next
    And32 = r
End Function

Private Function Or32(a, b)
    Dim r, i, bit
    r = 0
    For i = 0 To 31
        bit = 2 ^ i
        If Bit32(a, i) = 1 Or Bit32(b, i) = 1 Then r = r + bit
    Next
    Or32 = r
End Function

Private Function Xor32(a, b)
    Dim r, i, bit
    r = 0
    For i = 0 To 31
        bit = 2 ^ i
        If Bit32(a, i) <> Bit32(b, i) Then r = r + bit
    Next
    Xor32 = r
End Function

Private Function Not32(a)
    Not32 = 4294967295 - CDbl(a)
End Function

Private Function Hex32(v)
    Dim i, dd, s
    s = ""
    For i = 7 To 0 Step -1
        dd = Int(v / (16 ^ i)) - 16 * Int(v / (16 ^ (i + 1)))
        If dd < 10 Then
            s = s & Chr(48 + dd)
        Else
            s = s & Chr(55 + dd)
        End If
    Next
    Hex32 = s
End Function
%>
