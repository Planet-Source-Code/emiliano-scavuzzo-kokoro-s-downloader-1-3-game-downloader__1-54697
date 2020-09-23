Attribute VB_Name = "BuscarEnArchivo"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Public Function EstaTextoEnArchivo(ByVal Path As String, ByVal Texto As String, ByVal ChunkSize As Long) As Boolean
EstaTextoEnArchivo = False
Dim NroArchivo As Integer
Dim Indice As Long
Dim Contador As Integer
Dim Terminado As Boolean
Dim Chunk$

'el chunk tiene que ser por lo menos del tamaño del texto
'a buscar
If ChunkSize < Len(Texto) Then ChunkSize = Len(Texto)

NroArchivo = FreeFile
Open Path For Binary As #NroArchivo

If LOF(NroArchivo) < ChunkSize Then
    Chunk$ = Space(LOF(NroArchivo))
    Get #NroArchivo, 1, Chunk$
    EstaTextoEnArchivo = EstaTextoEnChunk(Chunk, Texto)

Else
    Terminado = False
    Indice = 1
    Contador = 0
    Do
        'DoEvents

        If (Indice + ChunkSize) > LOF(NroArchivo) Then
            Chunk$ = Space(LOF(NroArchivo) - Indice + 1)
            Terminado = True
        Else
            Chunk$ = Space(ChunkSize)
        End If
        
        Get #NroArchivo, Indice, Chunk$
        EstaTextoEnArchivo = EstaTextoEnChunk(Chunk, Texto)

        Indice = Indice + Len(Chunk$) - Len(Texto) + 1

    Loop Until Terminado Or (EstaTextoEnArchivo = True)

End If

Close #NroArchivo
End Function

Private Function EstaTextoEnChunk(ByRef Chunk As Variant, ByRef Texto As String) As Boolean
Dim Pos As Integer
Pos = InStr(1, Chunk, Texto, vbBinaryCompare)
If Pos = Null Or Pos = 0 Then
    EstaTextoEnChunk = False
Else
    EstaTextoEnChunk = True
End If
End Function
