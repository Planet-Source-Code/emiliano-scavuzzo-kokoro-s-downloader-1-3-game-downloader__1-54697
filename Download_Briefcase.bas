Attribute VB_Name = "Download_Briefcase"
'Estructura del tipo de download Briefcase
' LN_FIN+1)   Cantidad de particiones en la que
'                 está dividido cada paquete
' LN_FIN+2) Tamaño en bytes de cada particion
' LN_FIN+3) String a buscar para obtener el nombre verdadero
' LN_FIN+4) Comienzo de los paquetes. Estructura:
'           "nombre.xxx, http://direcion"
' Despues vienen las particiones con la estructura:
            'http://direcion'
' Despues otro paquete de la estructura de LN_FIN+4
' NOTA: el último paquete es menor o igual al resto
'       pero nunca mayor. El resto de los paquetes
'       tienen todos el mismo tamaño
' NOTA2: el tamaño de las particiones es el tamaño del
'        paquete dividido la cantidad de particiones,
'        menos la última partición del último paquete
'        que es lo anterior más tam_paq mod cant_part

Option Explicit

Private Const LN_CANT_PAR   As Byte = LN_FIN + 1
Private Const LN_TAM_PAR    As Byte = LN_FIN + 2
Private Const LN_STR_BUS    As Byte = LN_FIN + 3
Private Const LN_COM_PAQ    As Byte = LN_FIN + 4

Private TamanoPaquete As Double 'tamaño del paquete que se está bajando
Private CantidadBajada As Double 'cantidad de bytes bajados del paquete que se está bajando

Private NroDePaqABajar As Integer 'nro de paquetes del juego seleccionados para bajar
Private NroDeParticiones As Integer 'nro de particiones de cada paquete

Private Type TypeParticion
    Direccion As String
End Type

Private Type TypePaquete
    Indice As Integer   'nro en la lista de paquetes
    Tamano As Double      'tamaño del paquete
    Nombre As String    'nombre del archivo del paquete
    Particion() As TypeParticion
End Type

Private Type TypePartBajando
    Paq As Integer
    Part As Integer
End Type

Private Paquete() As TypePaquete 'todos los datos de los paquetes
Private PAB As TypePartBajando 'particion que se está bajando

Private RespuestaInfo As String 'respuesta que devuelve el Winsock_BriefcaseInfo

Private Winsock_Brief_Down As Boolean
Private JuegoPath As String

Private HAND_ARCHIVO As Integer 'handler del archivo del PAB

Private SockInfo As CSocketMaster
Private SockDown As CSocketMaster

'lee los paquetes de la info y los pone en la lista
Public Sub BriefcaseLeepaquetes()
Dim CantPaquetes As Integer
Dim CantParticiones As Integer
Dim Cont As Integer
CantPaquetes = Val(Trim(LeeLinea(Info, LN_NRO_PAQ)))
CantParticiones = Val(Trim(LeeLinea(Info, LN_CANT_PAR)))

For Cont = 0 To CantPaquetes - 1
    Principal.listPaquetes.AddItem "Paquete" + Str(Cont + 1) + ":" + " " + LeeNombrePaquete(Cont * CantParticiones + LN_COM_PAQ), Cont
Next Cont
End Sub

'devuelve el nombre de archivo que se usa para el paquete
'en una linea de la info especificada
Private Function LeeNombrePaquete(ByVal Linea As Integer) As String
Dim PosComa As Integer
LeeNombrePaquete = Trim(LeeLinea(Info, Linea))
PosComa = InStr(1, LeeNombrePaquete, ",")
LeeNombrePaquete = Trim(Left(LeeNombrePaquete, PosComa - 1))
End Function

'función principal que empieza el download
Public Sub BriefcaseDownload()
NroDePaqABajar = Principal.listPaquetes.SelCount
NroDeParticiones = Val(Trim(LeeLinea(Info, LN_CANT_PAR)))
JuegoPath = DownloadsPath + Trim(LeeLinea(Info, LN_NOMBRE)) + "\"
ConstruyeArray
LimpiaEstado
ComienzaBajarPaquete 1
End Sub

'construye el array que contiene las direcciones
'de los juegos, los índices, los tamaños y los nombres
Private Sub ConstruyeArray()
ReDim Paquete(1 To NroDePaqABajar) As TypePaquete
Dim Contador1 As Integer
Dim Contador2 As Integer
For Contador1 = 1 To NroDePaqABajar
    Paquete(Contador1).Indice = IndiceDelSeleccionado(Contador1)
    Paquete(Contador1).Tamano = TamanoDePaqNro(Paquete(Contador1).Indice + 1)
    Paquete(Contador1).Nombre = NombreDePaqNro(Paquete(Contador1).Indice + 1)
    ReDim Paquete(Contador1).Particion(1 To NroDeParticiones)
        For Contador2 = 1 To NroDeParticiones
            Paquete(Contador1).Particion(Contador2).Direccion = DireccionDePaqNro(Paquete(Contador1).Indice + 1, Contador2)
        Next Contador2
Next Contador1
End Sub

'devuelve el índice del elemento selecionado
'número Numero
Private Function IndiceDelSeleccionado(ByVal Numero As Integer) As Integer
Dim ContIndices As Integer
Dim ContSelec As Integer

ContIndices = 0
ContSelec = 0
Do Until (ContSelec = Numero)
    If Principal.listPaquetes.Selected(ContIndices) Then ContSelec = ContSelec + 1
    ContIndices = ContIndices + 1
Loop
IndiceDelSeleccionado = ContIndices - 1
End Function

'devuelve el tamaño del paquete nro Numero
Public Function TamanoDePaqNro(ByVal Numero As Integer) As Double
Dim TamTotal As Double
Dim TamParticion As Double
Dim NroPaquetes As Integer
TamTotal = Val(Trim(LeeLinea(Info, LN_TAM_TOT)))
TamParticion = Val(Trim(LeeLinea(Info, LN_TAM_PAR)))
NroPaquetes = Val(Trim(LeeLinea(Info, LN_NRO_PAQ)))

If NroPaquetes = Numero Then
    TamanoDePaqNro = TamTotal - (TamParticion * NroDeParticiones * (NroPaquetes - 1))
Else
    TamanoDePaqNro = TamParticion * NroDeParticiones
End If
End Function

'devuelve la direccion de la partición Particion del
'paquete Paquete
Private Function DireccionDePaqNro(ByVal Paquete As Integer, ByVal Particion As Integer) As String
If Particion = 1 Then
    Dim Inicio As Integer
    DireccionDePaqNro = Trim(LeeLinea(Info, LN_COM_PAQ + ((Paquete - 1) * NroDeParticiones)))
    Inicio = InStr(1, DireccionDePaqNro, ",", vbTextCompare)
    DireccionDePaqNro = Trim(Right(DireccionDePaqNro, Len(DireccionDePaqNro) - Inicio))
Else
    DireccionDePaqNro = Trim(LeeLinea(Info, LN_COM_PAQ + ((Paquete - 1) * NroDeParticiones) + Particion - 1))
End If
End Function

'devuelve el nombre del paquete nro Numero
Private Function NombreDePaqNro(ByVal Numero As Integer) As String
NombreDePaqNro = LeeNombrePaquete((Numero - 1) * NroDeParticiones + LN_COM_PAQ)
End Function

'devuelve el tamaño del paquete que se está bajando
Public Function BriefcaseTamanoPaquete() As Double
BriefcaseTamanoPaquete = TamanoPaquete
End Function

'develve la cantidad de bytes bajados del paquete
'que se está bajando
Public Function BriefcaseCantidadBajada() As Double
BriefcaseCantidadBajada = CantidadBajada
End Function

'comienza a bajar un paquete nuevo indicado por NroPaq
Private Sub ComienzaBajarPaquete(ByVal NroPaq As Integer)
PAB.Paq = NroPaq
PAB.Part = 1
TamanoPaquete = Paquete(PAB.Paq).Tamano
CantidadBajada = 0
SombreaPaquete (Paquete(PAB.Paq).Indice)
If CreaPAB <> 0 Then Exit Sub
Principal.Label14.Caption = "Preparando datos..."
If bUsarProxy = True Then
    Winsock_BriefcaseInfoG.Connect strProxy, lngPuerto
Else
    Winsock_BriefcaseInfoG.Connect SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion), 80
End If
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_BriefcaseInfo_Local_Connect()
Dim Comando_Mandar As String
RespuestaInfo = ""
Comando_Mandar = "GET " + SacarParteArchivo(Paquete(PAB.Paq).Particion(PAB.Part).Direccion) + " HTTP/1.0" + vbCrLf
Comando_Mandar = Comando_Mandar + "Accept: " + WebAccept + vbCrLf
Comando_Mandar = Comando_Mandar + "Referer: " + SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion) + vbCrLf
Comando_Mandar = Comando_Mandar + "User-Agent: " + WebUserAgent + vbCrLf
Comando_Mandar = Comando_Mandar + "Host: " + SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion) + vbCrLf
Comando_Mandar = Comando_Mandar + vbCrLf
Winsock_BriefcaseInfoG.SendData Comando_Mandar
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_BriefcaseInfo_Local_DataArrival(ByVal bytesTotal As Long)
Dim Chunk As String
Winsock_BriefcaseInfoG.GetData Chunk, vbString
RespuestaInfo = RespuestaInfo + Chunk
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_BriefcaseInfo_Local_Close()
On Error GoTo ErrorHandler
Winsock_BriefcaseInfoG.CloseSck
SacarDireccionDeInfo
BajarPAB
Exit Sub
ErrorHandler:
    If HAND_ARCHIVO <> 0 Then Close HAND_ARCHIVO
    HAND_ARCHIVO = 0
    MessageBox Principal.hwnd, "Error al tratar de obtener datos del paquete.", "Error", MB_ICONERROR
    Terminar_Download_Acciones
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_BriefcaseInfo_Local_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock_BriefcaseInfoG.CloseSck
Select Case (Number)
Case 11001:
    MessageBox Principal.hwnd, "No se puede establecer enlace. Compruebe su conexión a Internet.", "Error", MB_ICONERROR
Case 10060:
    MessageBox Principal.hwnd, "Ha pasado el tiempo de espera. Intente más tarde.", "Error", MB_ICONERROR
Case Else
    MessageBox Principal.hwnd, Description, "Error", MB_ICONERROR
End Select
Close HAND_ARCHIVO
HAND_ARCHIVO = 0
Terminar_Download_Acciones
End Sub

'saca la direccion real del archivo de RespuestaInfo
'y lo guarda en Paquete(pab.Paq).Particion(pab.Part).Direccion
Private Function SacarDireccionDeInfo() As String
Dim Inicio As Integer
Dim Fin As Integer

Inicio = PosStringAvanzada(Trim(LeeLinea(Info, LN_STR_BUS)), "href=", RespuestaInfo)
If Inicio = 0 Then Err.Raise 1

Inicio = InStr(Inicio, RespuestaInfo, Chr(34))
If Inicio = 0 Or Inicio = Null Then Err.Raise 1
Inicio = Inicio + 1
Fin = InStr(Inicio, RespuestaInfo, Chr(34))
If Fin = 0 Or Fin = Null Then Err.Raise 1
Fin = Fin - 1

SacarDireccionDeInfo = Mid(RespuestaInfo, Inicio, Fin - Inicio + 1)
If SacarDireccionDeInfo = "" Then Err.Raise 1
Paquete(PAB.Paq).Particion(PAB.Part).Direccion = SacarDireccionDeInfo
RespuestaInfo = "" 'resetea la variable para liberar memoria

End Function

'esta función se llama después de que se obtuvo la
'direccion real
Private Sub BajarPAB()
'si es la primera partición
If PAB.Part = 1 Then Principal.Label14.Caption = "Conectando al server..."
If bUsarProxy = True Then
    Winsock_Briefcase_DownG.Connect strProxy, lngPuerto
Else
    Winsock_Briefcase_DownG.Connect SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion), 80
End If
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_Briefcase_Down_Local_Connect()
Dim Comando_Mandar As String
Winsock_Brief_Down = False
If PAB.Part = 1 Then Principal.Label14.Caption = "Conectado!"
Comando_Mandar = "GET " + SacarParteArchivo(Paquete(PAB.Paq).Particion(PAB.Part).Direccion) + " HTTP/1.0" + vbCrLf
Comando_Mandar = Comando_Mandar + "Accept: " + WebAccept + vbCrLf
Comando_Mandar = Comando_Mandar + "Referer: " + SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion) + vbCrLf
Comando_Mandar = Comando_Mandar + "User-Agent: " + WebUserAgent + vbCrLf
Comando_Mandar = Comando_Mandar + "Host: " + SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion) + vbCrLf
Comando_Mandar = Comando_Mandar + vbCrLf
Winsock_Briefcase_DownG.SendData Comando_Mandar
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_Briefcase_Down_Local_DataArrival(ByVal bytesTotal As Long)
Principal.Label14.Caption = "Bajando " + Paquete(PAB.Paq).Nombre
If ActivoTimerEstado = False And PAB.Part = 1 Then ArrancarTimerEstado

Dim Chunk As String
Winsock_Briefcase_DownG.GetData Chunk, vbString

'si no paso la cabecera
If Winsock_Brief_Down = False Then
    Dim Split As Long
    Split = InStr(1, Chunk, vbCrLf + vbCrLf)
    If Split = 0 Or Split = Null Then Exit Sub 'si no llegó la cabecera sale
    Winsock_Brief_Down = True
    Chunk = Right(Chunk, Len(Chunk) - Split - 3)
End If
'si paso la cabecera
CantidadBajada = CantidadBajada + Len(Chunk)
Put #HAND_ARCHIVO, LOF(HAND_ARCHIVO) + 1, Chunk
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_Briefcase_Down_Local_Close()
Winsock_Briefcase_DownG.CloseSck
ParticionBajada
End Sub

'funcion clonada del winsock de Principal
Public Sub Winsock_Briefcase_Down_Local_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock_Briefcase_DownG.CloseSck
Select Case (Number)
Case 11001:
    MessageBox Principal.hwnd, "No se puede establecer enlace. Compruebe su conexión a Internet.", "Error", MB_ICONERROR
Case 10060:
    MessageBox Principal.hwnd, "Ha pasado el tiempo de espera. Intente más tarde.", "Error", MB_ICONERROR
Case Else
    MessageBox Principal.hwnd, Description, "Error", MB_ICONERROR
End Select
Close HAND_ARCHIVO
HAND_ARCHIVO = 0
Terminar_Download_Acciones
End Sub

'crea el archivo PAB en la carpeta downloads
'lo borra si ya existía y guarda el handler
'en HAND_ARCHIVO
'si ocurre un error al tratar de crear el paquete
'termina el download y devuelve 1, sino devuelve 0
Private Function CreaPAB() As Byte
CreaPAB = 0
CreaDirJuegoSiNoExiste
Dim PaquetePath As String
PaquetePath = JuegoPath + Paquete(PAB.Paq).Nombre

On Error GoTo ErrorAlAcceder
If Dir(PaquetePath, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = Paquete(PAB.Paq).Nombre Then SetAttr PaquetePath, vbNormal: Kill PaquetePath
HAND_ARCHIVO = FreeFile
Open PaquetePath For Binary Lock Read Write As HAND_ARCHIVO

Exit Function
ErrorAlAcceder:
    CreaPAB = 1
    BriefcaseTerminaDownload
    MessageBox Principal.hwnd, "Error al intentar acceder al archivo " + PaquetePath, "Error", MB_ICONERROR
    Terminar_Download_Acciones
End Function

'esta función se llama cuando se terminó de bajar
'una partición cualquiera
Private Sub ParticionBajada()
On Error GoTo ErrorHandler
If LOF(HAND_ARCHIVO) <> TamanoEsperado(PAB) Then Err.Raise 1

If PAB.Part = NroDeParticiones Then 'si se terminó de bajar el paquete
    Close HAND_ARCHIVO
    HAND_ARCHIVO = 0
    PararTimerEstado
    LimpiaEstado
    MarcaPaquete (Paquete(PAB.Paq).Indice)
    If PAB.Paq = NroDePaqABajar Then 'si terminó de bajar todos los paquetes
        JuegoBajado
    Else 'si no terminó de bajar todos paquete
        ComienzaBajarPaquete (PAB.Paq + 1)
    End If
Else 'si no terminó de bajar paquete
    PAB.Part = PAB.Part + 1
    If bUsarProxy = True Then
        Winsock_BriefcaseInfoG.Connect strProxy, lngPuerto
    Else
        Winsock_BriefcaseInfoG.Connect SacarHost(Paquete(PAB.Paq).Particion(PAB.Part).Direccion), 80
    End If
End If
Exit Sub
ErrorHandler:
    If HAND_ARCHIVO <> 0 Then Close HAND_ARCHIVO
    HAND_ARCHIVO = 0
    MessageBox Principal.hwnd, "Error al intentar descargar el paquete.", "Error", MB_ICONERROR
    Terminar_Download_Acciones
End Sub

Private Sub JuegoBajado()
MessageBox Principal.hwnd, "Juego bajado correctamente.", "Todo OK", MB_ICONINFORMATION
Terminar_Download_Acciones
End Sub

'termina el download
Public Sub BriefcaseTerminaDownload()
Winsock_BriefcaseInfoG.CloseSck
Winsock_Briefcase_DownG.CloseSck
If HAND_ARCHIVO <> 0 Then Close HAND_ARCHIVO: HAND_ARCHIVO = 0
End Sub

'produce un error si encuentra algún error en la info
'sobre los paquetes
Public Sub BriefcaseCheckPaquetes()
If Not IsNumeric(Trim(LeeLinea(Info, LN_CANT_PAR))) Then Err.Raise 1
If Not IsNumeric(Trim(LeeLinea(Info, LN_TAM_PAR))) Then Err.Raise 1
If Trim(LeeLinea(Info, LN_STR_BUS)) = "" Then Err.Raise 1

Dim CantPaq As Integer
Dim CantPart As Integer
Dim TamPart As Double
Dim TamTotal As Double
TamPart = Val(Trim(LeeLinea(Info, LN_TAM_PAR)))
CantPaq = Val(Trim(LeeLinea(Info, LN_NRO_PAQ)))
CantPart = Val(Trim(LeeLinea(Info, LN_CANT_PAR)))
TamTotal = Val(Trim(LeeLinea(Info, LN_TAM_TOT)))
If TamPart * CantPart * (CantPaq - 1) > TamTotal Then Err.Raise 1
If TamTotal - (TamPart * CantPart * (CantPaq - 1)) > (TamPart * CantPart) Then Err.Raise 1

Dim Cont As Integer
Dim Inicio As Integer
Dim Linea As String

For Cont = 0 To CantPaq * CantPart - 1
    
    Linea = Trim(LeeLinea(Info, LN_COM_PAQ + Cont))
    
    If Cont Mod CantPart = 0 Then 'si es una linea con nombre
        If CuantasVecesEsta(",", Linea) <> 1 Then Err.Raise 1
        Inicio = InStr(1, Linea, ",")
        If Right(Linea, Len(Linea) - Inicio) = "" Then Err.Raise 1
    Else
        If Linea = "" Then Err.Raise 1
        If CuantasVecesEsta(",", Linea) <> 0 Then Err.Raise 1
    End If
    
Next Cont
End Sub

'esta función devuelve el tamaño que devería tener un
'paquete si el PAB se bajó correctamente
Private Function TamanoEsperado(PABB As TypePartBajando) As Double
If PABB.Part = NroDeParticiones Then
    TamanoEsperado = Paquete(PABB.Paq).Tamano
Else
    Dim TamDesconocido As Double
    TamDesconocido = Int(Paquete(PABB.Paq).Tamano / NroDeParticiones)
    TamanoEsperado = PABB.Part * TamDesconocido
End If
End Function
