Attribute VB_Name = "Boundary"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit
Private Const CARACTERES As Integer = 10


Private Function CaracterAlAzar() As String
Dim Ascii As Integer
Ascii = Int(3 * Rnd + 1)

Select Case Ascii
    Case 1
        Ascii = Int((57 - 48 + 1) * Rnd + 48)
    Case 2
        Ascii = Int((90 - 65 + 1) * Rnd + 65)
    Case 3
        Ascii = Int((122 - 97 + 1) * Rnd + 97)
End Select

CaracterAlAzar = Chr(Ascii)
End Function

Public Function CreaBoundary() As String
Do
    CreaBoundary = "----------"
    Dim Cont As Integer
    For Cont = 1 To CARACTERES
        CreaBoundary = CreaBoundary + CaracterAlAzar
    Next Cont

Loop Until Not EstaTextoEnArchivo(GetDato(IndiceSubiendo, keyPath), CreaBoundary, 10000)
End Function
