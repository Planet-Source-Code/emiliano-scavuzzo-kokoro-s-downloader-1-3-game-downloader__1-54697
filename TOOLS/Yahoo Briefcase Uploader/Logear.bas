Attribute VB_Name = "modLogear"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Public sRespuesta As String
Public sComandoMandar As String
Public sHost As String

'estados del logeo
Public Enum logEstados
    Logeando = 1
    LlendoASubir
    FrameDeSubir
End Enum

'identifica qué se debe hacer después de logearse
Public Enum Finalidad
    fnSubir = 1
    fnEspacioUsado
End Enum

Public leEstado As logEstados 'situación del logeo
Public leHastaDonde As logEstados 'hasta que posición debe llegar
Public fnFinalidad As Finalidad 'qué hacer después de logearse

'función que se llama cuando se quiere logear
Public Sub Logear(ByRef Socket As CSocketMaster, ByVal Usuario As String, ByVal Password As String, ByVal HastaDonde As logEstados, ByVal Finalidad As Finalidad)
bUsarProxy = False
leEstado = Logeando
leHastaDonde = HastaDonde
fnFinalidad = Finalidad
Cookies.BorrarTodo
CreaComandoMandar Usuario, Password
Socket.Connect sHost, 80
End Sub

'Esta función se llama después de logearse para encontrar
'el Crumb de la página para subir, y la guarda en la
'variable global sCrumb
Public Sub ObtenerCrumb()
Dim Inicio As Long
Dim Final As Long
Inicio = InStr(1, sRespuesta, ".crumb", vbTextCompare)
Inicio = InStr(Inicio, sRespuesta, "value=", vbTextCompare)
Inicio = Inicio + Len("value=" + Chr(34))
Final = InStr(Inicio, sRespuesta, Chr(34), vbTextCompare)
sCrumb = Mid(sRespuesta, Inicio, Final - Inicio)
End Sub


Public Sub CreaComandoMandar(Optional ByVal Usuario As Variant, Optional ByVal Password As Variant)
Select Case leEstado
Case Logeando
If Not IsMissing(Usuario) Then
    sHost = "login.yahoo.com"
    Dim Datos As String
    Datos = ".fUpdate=1" + "&"
    Datos = Datos + ".tries=1" + "&"
    Datos = Datos + ".done=" + Codificador("http://f1.pg.briefcase.yahoo.com/") + "&"
    Datos = Datos + ".src=bc" + "&"
    Datos = Datos + ".intl=us" + "&"
    Datos = Datos + "login=" + Usuario + "&"
    Datos = Datos + "passwd=" + Password
    sComandoMandar = "POST /config/login HTTP/1.0" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-shockwave-flash, */*" + vbCrLf
    sComandoMandar = sComandoMandar + "Referer: http://briefcase.yahoo.com/" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept-Language: es-mx" + vbCrLf
    sComandoMandar = sComandoMandar + "Content-Type: application/x-www-form-urlencoded" + vbCrLf
    sComandoMandar = sComandoMandar + "User-Agent: Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)" + vbCrLf
    sComandoMandar = sComandoMandar + "Host: " + sHost + vbCrLf
    sComandoMandar = sComandoMandar + "Content-Length: " + Format(str(Len(Datos)), "#") + vbCrLf + vbCrLf
    sComandoMandar = sComandoMandar + Datos
Else

    sHost = SacarHost(GetLocationDeRespuesta)
    Cookies.ExtraeCookies (sRespuesta)
    sComandoMandar = "GET " + QuitarHost(GetLocationDeRespuesta) + " HTTP/1.0" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-shockwave-flash, */*" + vbCrLf
    sComandoMandar = sComandoMandar + "Referer: http://briefcase.yahoo.com/" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept-Language: es-mx" + vbCrLf
    sComandoMandar = sComandoMandar + "Content-Type: text/html" + vbCrLf
    sComandoMandar = sComandoMandar + "User-Agent: Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)" + vbCrLf
    sComandoMandar = sComandoMandar + "Cookie: " + Cookies.CookiesEnLinea + vbCrLf
    sComandoMandar = sComandoMandar + "Host: " + sHost + vbCrLf + vbCrLf
End If

Case LlendoASubir
    sHost = SacarHost(LeeLinkSubirDeRespuesta)
    sComandoMandar = "GET " + QuitarHost(LeeLinkSubirDeRespuesta) + " HTTP/1.0" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-shockwave-flash, */*" + vbCrLf
    sComandoMandar = sComandoMandar + "Referer: http://briefcase.yahoo.com/" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept-Language: es-mx" + vbCrLf
    sComandoMandar = sComandoMandar + "Content-Type: text/html" + vbCrLf
    sComandoMandar = sComandoMandar + "User-Agent: Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)" + vbCrLf
    sComandoMandar = sComandoMandar + "Cookie: " + Cookies.CookiesEnLinea + vbCrLf
    sComandoMandar = sComandoMandar + "Host: " + sHost + vbCrLf + vbCrLf

Case FrameDeSubir
    sHost = SacarHost(LeeLinkFrameDeRespuesta)
    sComandoMandar = "GET " + QuitarHost(LeeLinkFrameDeRespuesta) + " HTTP/1.0" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-shockwave-flash, */*" + vbCrLf
    sComandoMandar = sComandoMandar + "Referer: http://briefcase.yahoo.com/" + vbCrLf
    sComandoMandar = sComandoMandar + "Accept-Language: es-mx" + vbCrLf
    sComandoMandar = sComandoMandar + "Content-Type: text/html" + vbCrLf
    sComandoMandar = sComandoMandar + "User-Agent: Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)" + vbCrLf
    sComandoMandar = sComandoMandar + "Cookie: " + Cookies.CookiesEnLinea + vbCrLf
    sComandoMandar = sComandoMandar + "Host: " + sHost + vbCrLf + vbCrLf

End Select
End Sub

'Devuelve el valor de la linea de Location. Si no tiene
'linea Location devuelve ""
Public Function GetLocationDeRespuesta() As String
Dim PosIn As Long
Dim PosFi As Long
GetLocationDeRespuesta = ""
PosIn = InStr(1, sRespuesta, "Location:", vbTextCompare)
If PosIn = 0 Then Exit Function
PosIn = PosIn + Len("Location: ")
PosFi = InStr(PosIn, sRespuesta, vbCrLf, vbTextCompare)
GetLocationDeRespuesta = Mid(sRespuesta, PosIn, PosFi - PosIn)
End Function

'transforma los datos del formulario
Function Codificador(ByVal Cadena As String) As String
Dim Inicio As Integer
Dim Fin As Integer
Dim Caracter As Integer
Fin = Len(Cadena)

For Inicio = 1 To Fin
Caracter = Asc(Mid(Cadena, Inicio, 1))

If ((Caracter >= 65) And (Caracter <= 90)) Or ((Caracter >= 97) And (Caracter <= 122)) Or ((Caracter >= 48) And (Caracter <= 57)) Or (Caracter = 46) Then 'si es una letra o un número o '.'
    Codificador = Codificador + Chr(Caracter)
Else 'si no es una letra o un número o '.'
    If (Caracter = 32) Then 'si es un espacio
        Codificador = Codificador + "+"
    Else 'si no es un espacio
        If (Caracter > 15) Then   'si es menor que 15 (F) le pone un 0 adelante
            Codificador = Codificador + "%" + Hex(Caracter)
        Else
            Codificador = Codificador + "%0" + Hex(Caracter)
        End If
    End If
End If

Next
End Function

'devuelve el la parte del host de una direccion web
'ej: 'http://www.dreamers.com/' => 'dreamers.com'
Function SacarHost(ByVal Direccion As String) As String
If Left(Direccion, 7) = "http://" Then Direccion = Mid(Direccion, 8, Len(Direccion) - 7)
If Left(Direccion, 4) = "www." Then Direccion = Mid(Direccion, 5, Len(Direccion) - 4)
Dim Inicio As Integer
Inicio = InStr(1, Direccion, "/", vbTextCompare)
If Inicio <> 0 Then Direccion = Left(Direccion, Inicio - 1)
SacarHost = Direccion
End Function

'saca la parte del archivo que se usa despues del
'comando GET para bajar archivos SI NO SE USA PROXY
'ej 'http://www.choto.com/info/arc.txt' => '/info/arc.txt'
Function QuitarHost(ByVal Texto As String) As String

If bUsarProxy = True Then
    QuitarHost = Texto
    Exit Function
End If

If Left(Texto, 7) = "http://" Then Texto = Right(Texto, Len(Texto) - 7)
Dim Inicio As Integer
Inicio = InStr(1, Texto, "/", vbTextCompare)
If Inicio = 0 Or Inicio = Null Then
    QuitarHost = "/"
Else
    QuitarHost = Right(Texto, Len(Texto) - Inicio + 1)
End If
End Function

'devuelve la string del link de la página para subir
'archivos que se obtiene de la respuesta después de logearse
Public Function LeeLinkSubirDeRespuesta() As String

Dim Ini As Long
Dim Fin As Long

Ini = InStr(1, sRespuesta, "My Documents", vbBinaryCompare)
'if ini = 0 or ini= null then 'insetas codigo de error al logear
Ini = InStr(Ini, sRespuesta, "href", vbTextCompare)
Ini = Ini + Len("href=" + Chr(34))
Fin = InStr(Ini, sRespuesta, Chr(34), vbTextCompare)

LeeLinkSubirDeRespuesta = Mid(sRespuesta, Ini, Fin - Ini)
End Function

'devuelve la string del link de la página con las casillas
'para subir archivos
Public Function LeeLinkFrameDeRespuesta()
Dim Ini As Long
Dim Fin As Long

Ini = InStr(1, sRespuesta, "src=", vbTextCompare)
Ini = Ini + Len("src=" + Chr(34))
Fin = InStr(Ini, sRespuesta, Chr(34), vbTextCompare)

LeeLinkFrameDeRespuesta = Mid(sRespuesta, Ini, Fin - Ini)
End Function

'develve el espacio libre en MB de una cuenta
Public Sub EspacioUsado()
Dim Inicio As Long
Dim Final As Long
Inicio = InStr(1, sRespuesta, "Using", vbTextCompare)
Inicio = Inicio + Len("Using ")
Final = InStr(Inicio, sRespuesta, " of", vbTextCompare)
MsgBox Mid(sRespuesta, Inicio, Final - Inicio)
HabilitaParaSubir
sRespuesta = ""
End Sub
