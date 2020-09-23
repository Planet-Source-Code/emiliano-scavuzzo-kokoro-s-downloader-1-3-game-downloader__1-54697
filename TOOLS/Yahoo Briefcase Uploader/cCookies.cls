VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private cCookies() As classCookieDato
Private iCantidad As Integer

Private Sub Class_Initialize()
BorrarTodo
End Sub

Private Sub Class_Terminate()
Erase cCookies
End Sub

Property Get Cantidad() As Integer
Cantidad = iCantidad
End Property

Public Sub BorrarTodo()
Erase cCookies
iCantidad = 0
End Sub

Public Sub Agregar(Optional Nombre As Variant, Optional Cookie As Variant)
AgregarEspacioParaCookie
If Not IsMissing(Nombre) Then
    cCookies(Cantidad).Nombre = Nombre
End If

If Not IsMissing(Cookie) Then
    cCookies(Cantidad).Cookie = Cookie
End If
End Sub

Public Sub Quitar(Indice As Variant)
If Indice > Cantidad Then
    Err.Raise 1, , "The index is out of range"
    Exit Sub
End If
Dim Cont As Integer
For Cont = Indice To Cantidad - 1
    Set cCookies(Cont) = cCookies(Cont + 1)
Next
QuitarEspacioParaCookie
End Sub

Private Sub AgregarEspacioParaCookie()
ReDim Preserve cCookies(1 To iCantidad + 1)
Set cCookies(iCantidad + 1) = New classCookieDato
iCantidad = iCantidad + 1
End Sub

Private Sub QuitarEspacioParaCookie()
ReDim Preserve cCookies(1 To iCantidad - 1)
iCantidad = iCantidad - 1
End Sub

'extrae la cookies de una string
Public Sub ExtraeCookies(ByRef Datos As String)
Dim PosIn As Long
Dim PosFi As Long
PosIn = 1
PosIn = InStr(PosIn, Datos, "Set-Cookie:", vbTextCompare)

Do Until PosIn = 0
    AgregarEspacioParaCookie
    PosIn = PosIn + Len("Set-Cookie: ")
    PosFi = InStr(PosIn, Datos, "=", vbTextCompare)
    Me.Cookies(Cantidad).Nombre = Mid(Datos, PosIn, PosFi - PosIn)
    
    PosIn = PosFi + 1
    PosFi = InStr(PosIn, Datos, ";", vbTextCompare)
    Me.Cookies(Cantidad).Cookie = Mid(Datos, PosIn, PosFi - PosIn)
    
    PosIn = InStr(PosFi, Datos, "Set-Cookie:", vbTextCompare)
Loop
End Sub

Property Get Cookies(Indice) As classCookieDato
If Indice > Cantidad Then
    Err.Raise 2, , "The index is out of range"
    Exit Property
End If
Set Cookies = cCookies(Indice)
End Property

Property Get CookiesEnLinea()
If Cantidad = 0 Then
    CookiesEnLinea = ""
Else
    Dim Cont As Integer
    Dim Temp As String
    For Cont = 1 To Cantidad
        Temp = Temp + Cookies(Cont).Nombre + "=" + Cookies(Cont).Cookie + "; "
    Next Cont
    CookiesEnLinea = Left(Temp, Len(Temp) - 2)
End If
End Property

'Encuentra el índice de la cookie con el nombre dado.
'Si no la encuentra devuelve 0.
Property Get EncuentraC(ByVal Nombre As String) As Integer
Dim Cont As Integer
Dim Indice As Integer
Cont = 1
Indice = 0
Do Until (Cont > Cantidad) Or (Indice <> 0)
    If Cookies(Cont).Nombre = Nombre Then Indice = Cont
    Cont = Cont + 1
Loop
EncuentraC = Indice
End Property
