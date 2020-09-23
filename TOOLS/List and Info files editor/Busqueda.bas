Attribute VB_Name = "Busqueda"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Public strBus As String
Public mMetodo As VbCompareMethod
Public bArriba As Boolean

'Devuelve la posición de la string strCorta en StrLarga
'después de la posición lPos
'Si no la encuentra devuelve 0
Public Function BuscarDespuesDe(ByRef strCorta, ByRef strLarga, ByVal lPos As Long) As Long
Dim PosTemp As Long
PosTemp = lPos
Do
    PosTemp = InStr(PosTemp + 1, strLarga, strCorta, mMetodo)
Loop Until (PosTemp > lPos) Or (PosTemp = 0) Or (PosTemp = Null)
BuscarDespuesDe = PosTemp
End Function

'Devuelve la posición de la string strCorta en StrLarga
'antes de la posición lPos
'Si no la encuentra devuelve 0
Public Function BuscarAntesDe(ByRef strCorta, ByRef strLarga, ByVal lPos As Long) As Long
Dim PosTemp As Long
Dim PosAnt As Long
PosTemp = 0
Do
    PosAnt = PosTemp
    PosTemp = InStr(PosTemp + 1, strLarga, strCorta, mMetodo)
Loop Until (PosTemp + Len(strBus) > lPos + 1) Or (PosTemp = 0) Or (PosTemp = Null)

If PosAnt <= lPos Then
    BuscarAntesDe = PosAnt
Else
    BuscarAntesDe = 0
End If
End Function

Public Sub BuscarSiguiente()
Dim lPos As Long
Dim txtTexto As TextBox
If Form1.mnuAjuste.Checked Then
    Set txtTexto = Form1.Text2
Else
    Set txtTexto = Form1.Text1
End If
lPos = txtTexto.SelStart
If bArriba Then
    lPos = BuscarAntesDe(strBus, txtTexto, lPos)
Else
    lPos = BuscarDespuesDe(strBus, txtTexto, lPos + txtTexto.SelLength)
End If

If lPos = 0 Then
    MsgBox "Cannot find """ + Left(strBus, 31) + """", vbOKOnly + vbInformation, "Criptonita"
Else
    txtTexto.SelStart = lPos - 1
    txtTexto.SelLength = Len(strBus)
End If
End Sub
