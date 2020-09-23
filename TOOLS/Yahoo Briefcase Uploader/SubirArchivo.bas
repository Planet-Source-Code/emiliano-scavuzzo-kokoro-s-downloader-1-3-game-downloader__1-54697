Attribute VB_Name = "SubirArchivo"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Public Const sHostSubir As String = "edit.briefcase.yahoo.com"
Public Const BUFFSIZE As Long = 8192

'buffers donde se guardan los datos del archivo
'temporalmente antes de ser enviado
Public Buffer As String * BUFFSIZE
Public Buffer2 As String

Public ArchivoSubId As Integer 'número del archivo (handler) que se está subiendo

Public Enum SubirAcciones
    MandandoParteAArchivo
    MandandoArchivo
    MandandoParteDArchivo
    Terminando
End Enum

Public EstadoEnvio As SubirAcciones

'Public SubidoParteAA As Boolean 'si ya se subió la parte antes del archivo


'devuelve el boundary con dos "-"
Function InsB() As String
InsB = "--" + sBoundary
End Function

'devuelve el Content-Disposicion del header
Function InsC(ByVal Name As String, Optional ByVal Argumento As String) As String
InsC = "Content-Disposition: form-data; name=" + Chr(34) + Name + Chr(34)
If IsMissing(Argumento) Then Exit Function
InsC = InsC + Argumento
End Function

Function Codificador(ByVal Cadena As String) As String
Dim Fin As Long
Dim Inicio As Long
Dim Caracter As Integer

Fin = Len(Cadena)

For Inicio = 1 To Fin
Caracter = Asc(Mid(Cadena, Inicio, 1))

If ((Caracter >= 65) And (Caracter <= 90)) Or ((Caracter >= 97) And (Caracter <= 122)) Or ((Caracter >= 48) And (Caracter <= 57)) Or (Caracter = 46) Or (Caracter = 47) Then  'si es una letra o un número o '.' o '/'
    Codificador = Codificador + Chr(Caracter)
Else 'si no es una letra o un número o '.' o '/'
    If (Caracter = 32) Then 'si es un espacio
        Codificador = Codificador + "+"
    Else 'si no es un espacio
        If (Caracter > 15) Then   'si es menor que 15 (F) le pone un 0 adelante
            Codificador = Codificador + "%" + LCase(Hex(Caracter))
        Else
            Codificador = Codificador + "%0" + LCase(Hex(Caracter))
        End If
    End If
End If
Next
End Function


'crea la parte que viene antes del archivo
Function CreaParteAArchivo() As String
CreaParteAArchivo = InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".briefcaseID") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + UsuarioAct + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".action") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "upload" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".src") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "bc" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".done") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "http://briefcase.yahoo.com/bc/" + UsuarioAct + "/lst?&.dir=" + Codificador(GetDato(IndiceSubiendo, keyDirectorio)) + "&.src=bc&.view=l" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".addlink") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".fnm") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".singleAdd") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".isIE") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".ocxUpload") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "http://" + sHostSubir + "/edit/" + UsuarioAct + "/upload_thru_ocx?.dir=" + Codificador(GetDato(IndiceSubiendo, keyDirectorio)) + "&&.action=upload&.furl=http%3a//briefcase.yahoo.com/bc/" + UsuarioAct + "%3fa%26.src=bc&.src=bc&.addlink=http%3a//" + sHostSubir + "/edit/" + UsuarioAct + "/add_mlink_form%3f.dir=" + Codificador(Codificador(GetDato(IndiceSubiendo, keyDirectorio))) + "%26.action=addlink%26.done=http%253a//briefcase.yahoo.com/bc/" + UsuarioAct + "/lst%253f%2526.dir=" + Codificador(Codificador(Codificador(GetDato(IndiceSubiendo, keyDirectorio)))) + "%2526.src=bc%2526.view=l&.done=http%3a//briefcase.yahoo.com/bc/" + UsuarioAct + "/lst%3f%26.dir=" + Codificador(Codificador(GetDato(IndiceSubiendo, keyDirectorio))) + "%26.src=bc%26.view=l&.ocxPath=http%3a//briefcase.yahoo.com/ocx" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".albUrl") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "http://briefcase.yahoo.com/bc/" + UsuarioAct + "/lst?&.dir=" + Codificador(GetDato(IndiceSubiendo, keyDirectorio)) + "&.src=bc&.view=l" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".dir") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + GetDato(IndiceSubiendo, keyDirectorio) + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".showUTLink") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".uType") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "2" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".crumb") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + sCrumb + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".drs") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "400" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".hires") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "y" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC(".muplform") + vbCrLf + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "y" + vbCrLf

CreaParteAArchivo = CreaParteAArchivo + InsB + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + InsC("file0", "; filename=" + Chr(34) + SacaParteArchivo(GetDato(IndiceSubiendo, keyPath)) + Chr(34)) + vbCrLf
CreaParteAArchivo = CreaParteAArchivo + "Content-Type: text/plain" + vbCrLf + vbCrLf
End Function

'crea la parte que viene después del archivo
Function CreaParteDArchivo() As String
CreaParteDArchivo = vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".dnm0") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC("file1", "; filename=" + Chr(34) + SacaParteArchivo(GetDato(IndiceSubiendo, keyPath)) + Chr(34)) + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".dnm1") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC("file2", "; filename=" + Chr(34) + Chr(34)) + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".dnm2") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC("file3", "; filename=" + Chr(34) + Chr(34)) + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".dnm3") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC("file4", "; filename=" + Chr(34) + Chr(34)) + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".dnm4") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC("file5", "; filename=" + Chr(34) + Chr(34)) + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".dnm5") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + InsC(".upload") + vbCrLf + vbCrLf
CreaParteDArchivo = CreaParteDArchivo + "Upload" + vbCrLf

CreaParteDArchivo = CreaParteDArchivo + InsB + "--" + vbCrLf
End Function

Function CreaCabecera() As String
CreaCabecera = "POST " + "http://" + sHost + "/edit/" + UsuarioAct + "/process_bcmultipart_form HTTP/1.0" + vbCrLf
CreaCabecera = CreaCabecera + "User-Agent: Opera/6.05 (Windows ME; U)  [en]" + vbCrLf
CreaCabecera = CreaCabecera + "Host: " + sHost + vbCrLf
CreaCabecera = CreaCabecera + "Accept: text/html, image/png, image/jpeg, image/gif, image/x-xbitmap, */*" + vbCrLf
CreaCabecera = CreaCabecera + "Accept-Language: en" + vbCrLf
CreaCabecera = CreaCabecera + "Accept-Charset: windows-1252;q=1.0, utf-8;q=1.0, utf-16;q=1.0, iso-8859-1;q=0.6, *;q=0.1" + vbCrLf
CreaCabecera = CreaCabecera + "Accept-Encoding: deflate, gzip, x-gzip, identity, *;q=0" + vbCrLf
CreaCabecera = CreaCabecera + "Referer: http://edit.briefcase.yahoo.com/edit/colimante/fupload_form?.dir=/My+Documents&.src=bc&.action=upload&.done=http%3a//briefcase.yahoo.com/bc/colimante/lst%3f%26.dir=/My%2bDocuments%26.src=bc%26.view=l&.mesg=" + vbCrLf
CreaCabecera = CreaCabecera + "Cookie: " + Cookies.CookiesEnLinea + vbCrLf
CreaCabecera = CreaCabecera + "Cookie2: $Version=" + Chr(34) + "1" + Chr(34) + vbCrLf
CreaCabecera = CreaCabecera + "Content-length: " + Trim(str(Len(CreaParteAArchivo) + LOF(ArchivoSubId) + Len(CreaParteDArchivo))) + vbCrLf
CreaCabecera = CreaCabecera + "Content-Type: multipart/form-data; boundary=" + sBoundary + vbCrLf
CreaCabecera = CreaCabecera + vbCrLf
End Function
