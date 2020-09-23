VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classCookieDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private sNombre As String
Private sCookie As String

Property Get Nombre() As String
Nombre = sNombre
End Property

Property Let Nombre(ByVal Nombre As String)
sNombre = Nombre
End Property

Property Get Cookie() As String
Cookie = sCookie
End Property

Property Let Cookie(ByVal Cookie As String)
sCookie = Cookie
End Property
