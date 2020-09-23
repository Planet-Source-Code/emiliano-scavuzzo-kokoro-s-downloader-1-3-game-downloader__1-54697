Attribute VB_Name = "VariablesGlobales"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Public IndiceSubiendo As Integer  'indice del elemento que se está subiendo
Public Transcurrido As Date

Public Const keyArchivo As String = "File"
Public Const keyEstado As String = "Estate"
Public Const keyIntento As String = "Attempt"
Public Const keyEnviado As String = "Sent"
Public Const keyTamTotal As String = "Total Size"
Public Const keyTiempo As String = "Time"
Public Const keyVelocidad As String = "Speed"
Public Const keyUsuario As String = "User"
Public Const keyPassword As String = "Password"
Public Const keyDirectorio As String = "Folder"
Public Const keyPath As String = "Path"
Public Const keyTamArchivo As String = "File Size"

'variables para cambiar tamaño
Public DirPorLargo As Long
Public FilePorLargo As Long
Public EspFijosHor As Integer 'espacio emtre dir1 y file1 + márgenes
Public EspEntreDirFile As Integer 'espacio entre dir1 y file1
Public MargenesLista As Integer 'margenes de la lista
Public EspFijosVer As Integer 'espacio emtre dir1 y lista + márgenes
Public DirPorAlto As Long
Public ListaPorAlto As Long
Public EspEntreDirLista As Long 'espacio entre dir1 y lista

Public UsuarioAct As String    'usuario actual
Public PasswordAct As String   'password actual

Public AUTOMATICO As Boolean   'indica si esta en modo automatico

Public Cookies As New classCookies
Public sCrumb As String     'uno de los datos necesario para subir
Public sBoundary As String
Public bUsarProxy As Boolean

Private Type TypeVelocidad
    Cantidad(1 To 10) As Long
    CantCargadas As Byte
    UltimaCantidad As Long
End Type

Public Velocidad As TypeVelocidad 'datos para sacar la velocidad

'cada cuantos segundos se verifica si la conexión esta muy lenta
Public Const LENT_FRECUENCIA As Long = 60 * 5
'al menos cuantos segundos debería tardar en subir 5 MB a
'la velocidad actual para considerar que va muy lento
Public Const LENT_MAX_TIEMPO As Long = 60 * 35
