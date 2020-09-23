Attribute VB_Name = "ComMensajes"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Public Enum eTipoMensaje
    ArranqueAutomatico
    Desconectando
    Apagando
End Enum

Public TipoMensaje As eTipoMensaje
Public ApretadoCancelar As Boolean


Public Sub MostrarMensaje(ByVal Tipo As eTipoMensaje)
TipoMensaje = Tipo
Mensajes.Show vbModal, Subidor
Unload Mensajes
End Sub
