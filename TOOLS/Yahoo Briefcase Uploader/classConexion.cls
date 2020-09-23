VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (ByVal lprasdialextensions As Long, ByVal lpszPhonebook As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByVal dwNotifierType As Long, lpvNotifier As Any, ByRef lphRasConn As Long) As Long
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpCb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As String, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)


Private Const RAS_MaxEntryName = 256
Private Const RAS_MaxDeviceName = 128
Private Const RAS_MaxDeviceType = 16
Private Const RAS95_MaxEntryName = 256
Private Const RAS_MaxPhoneNumber = 128
Private Const RAS_RASCONNSIZE = 412
Private Const RAS_MaxCallbackNumber = RAS_MaxPhoneNumber



Private Const RAS_NOTIFY_HWND = &HFFFFFFFF
Private Const ERROR_INVALID_HANDLE = 6
Private Const RASCS_Connected = &H2000
Private Const RASCS_Disconnected = &H2001

Private Const UNLEN = 256
Private Const PWLEN = 256
Private Const DNLEN = 12

Private Type RASDIALPARAMS
   dwSize As Long ' 1052
   szEntryName(RAS95_MaxEntryName) As Byte
   szPhoneNumber(RAS_MaxPhoneNumber) As Byte
   szCallbackNumber(RAS_MaxCallbackNumber) As Byte
   szUserName(UNLEN) As Byte
   szPassword(PWLEN) As Byte
   szDomain(DNLEN) As Byte
End Type

Private Type RASCONN
   dwSize As Long
   hRasConn As Long
   szEntryName(RAS_MaxEntryName) As Byte
   szDeviceType(RAS_MaxDeviceType) As Byte
   szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Private Type RASCONNSTATUS
    dwSize As Long
    RASCONNSTATE As Long
    dwError As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Public Enum eEstadoConexion
    DESCONECTADO = 0
    CONECTADO
    CONECTANDO
End Enum

'tipos de cuentas
Private Type TypeCuentaInternet
    Usuario As String
    Password As String
End Type

Private Type TypeConexion
    Cuentas() As TypeCuentaInternet
    NumeroTel As String
End Type

Private Enum TypeISP
    NINGUNA = 0
    INFOVIA
    ARNET
End Enum

Private ISPActual As TypeISP
Private CuentaActual As Integer
Private RDP As RASDIALPARAMS
Private ConexionHandle As Long
Private Internet(1 To 2) As TypeConexion

Private Sub Class_Initialize()
'inicializa las cuentas de internet
ISPActual = NINGUNA
CuentaActual = 0
Internet(ARNET).NumeroTel = "06102227638"
Internet(INFOVIA).NumeroTel = "06102223826"
ReDim Internet(ARNET).Cuentas(0)
ReDim Internet(INFOVIA).Cuentas(0)
AgregarCuenta INFOVIA, "myaccount1@infovia", "password1"
AgregarCuenta INFOVIA, "myaccount2@infovia", "password2"
AgregarCuenta ARNET, "myaccount3", "password3"
AgregarCuenta ARNET, "myaccount4", "password4"
AgregarCuenta ARNET, "myaccount5", "passowrd5"

'inicializa RDP
ChangeBytes "", RDP.szEntryName
ChangeBytes "*", RDP.szDomain
RDP.dwSize = 1052
End Sub

'Changes a Visual Basic unicode string to an byte array
'Returns True if it truncates str
Private Function ChangeBytes(ByVal str As String, Bytes() As Byte) As Boolean
    Dim lenBs As Long 'length of the byte array
    Dim lenStr As Long 'length of the string
    lenBs = UBound(Bytes) - LBound(Bytes)
    lenStr = LenB(StrConv(str, vbFromUnicode))
    If lenBs > lenStr Then
        CopyMemory Bytes(0), str, lenStr
        ZeroMemory Bytes(lenStr), lenBs - lenStr
    ElseIf lenBs = lenStr Then
        CopyMemory Bytes(0), str, lenStr
    Else
        CopyMemory Bytes(0), str, lenBs 'Queda truncado
        ChangeBytes = True
    End If
End Function

Private Function AgregarCuenta(ByVal ISP As TypeISP, ByVal Usuario As String, ByVal Password As String)
Dim Cant As Integer
Cant = UBound(Internet(ISP).Cuentas)
ReDim Preserve Internet(ISP).Cuentas(0 To Cant + 1)
Internet(ISP).Cuentas(Cant + 1).Usuario = Usuario
Internet(ISP).Cuentas(Cant + 1).Password = Password
End Function

'se conecta a internet intentando con todas las
'cuentas y devuelve TRUE si lo logra
Public Function Conectar(Optional DesconectarPrimero As Boolean = True) As Boolean
Conectar = False

If DesconectarPrimero Then DesconectarTodo

PreparaRDP ARNET, 1
ISPActual = ARNET
CuentaActual = 1

Dim Respuesta As Long

Do
    Respuesta = RasDial(0&, 0&, RDP, RAS_NOTIFY_HWND, Subidor.hWnd, ConexionHandle)
    If Respuesta <> 0 Then Exit Function 'no se pudo conectar
    Do
        DoEvents
    Loop Until EstadoConexion <> CONECTANDO
    If EstadoConexion = CONECTADO Then  'se conectó
        Conectar = True
    Else 'no se conectó
        Desconectar
        If CuentaActual < UBound(Internet(ISPActual).Cuentas) Then
            CuentaActual = CuentaActual + 1
            PreparaRDP ISPActual, CuentaActual
        Else 'si no quedan cuentas del ISP actual
            If ISPActual < UBound(Internet) Then
                ISPActual = ISPActual + 1
                CuentaActual = 1
                PreparaRDP ISPActual, CuentaActual
            Else 'si no quedan cuentas de ningún ISP
                Exit Do
            End If
        End If
    End If
Loop Until Conectar
End Function

'encuentra todas las conexiones y las desconecta
Public Sub DesconectarTodo()
Dim CodigoDevuelto As Long
Dim Cont As Long, lpRasConn(255) As RASCONN, lpCb As Long
Dim lpcConnections As Long, hRasConn As Long
'Set the structure's size
lpRasConn(0).dwSize = RAS_RASCONNSIZE
lpCb = RAS_MaxEntryName * lpRasConn(0).dwSize
lpcConnections = 0
'Enumerate all the available connections
CodigoDevuelto = RasEnumConnections(lpRasConn(0), lpCb, lpcConnections)

If CodigoDevuelto = 0 Then
    For Cont = 0 To lpcConnections - 1
        DoEvents
        hRasConn = lpRasConn(Cont).hRasConn
        CodigoDevuelto = RasHangUp(ByVal hRasConn)
    Next Cont
End If

Dim RasEstado As RASCONNSTATUS
Dim Desconectados As Boolean 'indica si todas las conexiones están colgadas
RasEstado.dwSize = 160
    Do
        Desconectados = True
        For Cont = 0 To lpcConnections - 1
            CodigoDevuelto = RasGetConnectStatus(lpRasConn(Cont).hRasConn, RasEstado)
            If CodigoDevuelto <> ERROR_INVALID_HANDLE Then Desconectados = False
            DoEvents
        Next Cont
    Loop Until Desconectados
End Sub

'desconecta la conexion del handle ConexionHandle
Public Sub Desconectar()
Dim CodigoDevuelto As Long

RasHangUp ConexionHandle

Dim RasEstado As RASCONNSTATUS
RasEstado.dwSize = 160
    Do
        CodigoDevuelto = RasGetConnectStatus(ConexionHandle, RasEstado)
        DoEvents
    Loop Until CodigoDevuelto = ERROR_INVALID_HANDLE
End Sub
Private Sub PreparaRDP(ByVal ISP As TypeISP, ByVal Conexion As Integer)
ChangeBytes Internet(ISP).NumeroTel, RDP.szPhoneNumber
ChangeBytes Internet(ISP).Cuentas(Conexion).Usuario, RDP.szUserName
ChangeBytes Internet(ISP).Cuentas(Conexion).Password, RDP.szPassword
End Sub

'se conecta con otro ISP y devuelve TRUE si lo logra
Public Function ConectarConOtroISP() As Boolean
Desconectar
ConectarConOtroISP = False

If ISPActual < UBound(Internet) Then
    ISPActual = ISPActual + 1
Else
    ISPActual = 1
End If
CuentaActual = 1

Do
    RasDial 0&, 0&, RDP, RAS_NOTIFY_HWND, Subidor.hWnd, ConexionHandle
    Do
        DoEvents
    Loop Until EstadoConexion <> CONECTANDO
    If EstadoConexion = CONECTADO Then  'se conectó
        ConectarConOtroISP = True
    Else 'no se conectó
        Desconectar
        If CuentaActual < UBound(Internet(ISPActual).Cuentas) Then
            CuentaActual = CuentaActual + 1
            PreparaRDP ISPActual, CuentaActual
        Else 'si no quedan cuentas del ISP actual
            If ISPActual < UBound(Internet) Then
                ISPActual = ISPActual + 1
                CuentaActual = 1
                PreparaRDP ISPActual, CuentaActual
            Else 'si no quedan cuentas de ningún ISP
                Exit Do
            End If
        End If
    End If
Loop Until ConectarConOtroISP
End Function

'devuelve el estado de la conexión
Public Function EstadoConexion() As eEstadoConexion
Dim Respuesta As Long
Dim RasEstado As RASCONNSTATUS
RasEstado.dwSize = 160
Respuesta = RasGetConnectStatus(ConexionHandle, RasEstado)
If Respuesta <> 0 Then
    EstadoConexion = DESCONECTADO
Else
    If RasEstado.RASCONNSTATE = RASCS_Connected Then
        EstadoConexion = CONECTADO
    ElseIf RasEstado.RASCONNSTATE = RASCS_Disconnected Then
        EstadoConexion = DESCONECTADO
    Else
        EstadoConexion = CONECTANDO
    End If
End If
End Function
