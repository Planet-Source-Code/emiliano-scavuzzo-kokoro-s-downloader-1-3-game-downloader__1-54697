Attribute VB_Name = "AdvanceSearch"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

Private Type typeStringSearch
    Text As String
    Repeat As Integer
End Type

Private Type typePositionString
    PreviousTo() As typeStringSearch
    LaterThan() As typeStringSearch
End Type

'Find a string between the A(?) and the B(?)
'
'*ie:
'A(2)="AfterThis", A(1)="AndThis", B(2)="BeforeThis"
'find the string strShort in strLong that comes after
'two strings "AfterThis" and one string "AndThis" and
'before two strings "BeforeThis".
'
'*If it finds it (once or more) returns the position <> 0.
'*If you don't pass any B(?) string then returns the first
'ocurrence. If you pass at least one B(?) string returns
'the last ocurrence before the B string.
'*The order of the A and B strings doesn't affect the result.
'First it takes care of the A strings and then the B strings.
'*If an error occurs (the A and B strings MUST be there) or
'in case it can't find strShort string, it returns 0.
Public Function PosAdvanceSearch(ByVal strStringPosition As String, strShort As String, strLong As String) As Long
PosAdvanceSearch = 0
Dim psPosition As typePositionString
Dim intCount As Integer
Dim intRepetitions As Integer
Dim lngStart As Long
Dim lngEnd As Long

BuildPosition psPosition, strStringPosition

lngStart = 0
For intCount = 1 To UBound(psPosition.LaterThan)
    For intRepetitions = 1 To psPosition.LaterThan(intCount).Repeat
        lngStart = InStr(lngStart + 1, strLong, psPosition.LaterThan(intCount).Text, vbBinaryCompare)
        If lngStart = 0 Or lngStart = Null Then Exit Function
    Next intRepetitions
Next intCount

Dim lngIndex As Long
Dim lngPreviousIndex As Long
lngEnd = Len(strLong)

For intCount = 1 To UBound(psPosition.PreviousTo)
    For intRepetitions = 1 To psPosition.PreviousTo(intCount).Repeat
        lngIndex = lngStart
        Do
            lngPreviousIndex = lngIndex
            lngIndex = InStr(lngIndex + 1, strLong, psPosition.PreviousTo(intCount).Text, vbBinaryCompare)
        Loop Until lngIndex = 0 Or lngIndex = Null Or lngIndex >= lngEnd
        If lngPreviousIndex = lngIndex + 1 Or lngPreviousIndex <= lngStart Then
            Exit Function 'it couldn't find the string after lngStart and before lngEnd
        Else
            lngEnd = lngPreviousIndex
        End If
    Next intRepetitions
Next intCount

If UBound(psPosition.PreviousTo) = 0 Then
    PosAdvanceSearch = InStr(lngStart, strLong, strShort, vbBinaryCompare)
Else
    lngIndex = lngStart
    Do
        lngPreviousIndex = lngIndex
        lngIndex = InStr(lngIndex + 1, strLong, strShort, vbBinaryCompare)
    Loop Until lngIndex = 0 Or lngIndex = Null Or lngIndex >= lngEnd

    If lngPreviousIndex = lngStart Then
        PosAdvanceSearch = 0
    Else
        PosAdvanceSearch = lngPreviousIndex
    End If
End If
End Function

'Build array with the string location.
'If it finds an error it raises error 1.
Private Sub BuildPosition(ByRef psPosition As typePositionString, ByVal strStringPosition As String)
ReDim psPosition.LaterThan(0)
ReDim psPosition.PreviousTo(0)
Dim strData As String
Dim intNumberData As Integer
Dim intCount As Integer

intNumberData = CountHowManyTimes(",", strStringPosition) + 1
For intCount = 1 To intNumberData

    Dim intStart As Integer
    Dim intEnd As Integer
    Dim strLetter As String
    Dim intRepetitions As Integer
    
    strData = Trim(DataFromLine(strStringPosition, intCount))
    If Len(strData) < Len("A(?)='?'") Then Err.Raise 1   'largo mínimo
    strLetter = UCase(Left(strData, 1))
    intStart = Len("A(")
    If Mid(strData, intStart, 1) <> "(" Then Err.Raise 1
    intStart = Len("A(?)")
    intEnd = InStr(intStart, strData, ")", vbTextCompare)
    If intEnd = 0 Or intEnd = Null Then Err.Raise 1
    intStart = Len("A(?")
    If Not IsNumeric(Mid(strData, intStart, intEnd - intStart)) Then Err.Raise 1
    intRepetitions = Val(Mid(strData, intStart, intEnd - intStart))
    If intRepetitions <= 0 Then Err.Raise 1
    intStart = intEnd + 1
    If Mid(strData, intStart, 2) <> "=" + Chr(34) Then Err.Raise 1
    intStart = intStart + Len("=" + Chr(34))
    intEnd = InStr(intStart, strData, Chr(34), vbTextCompare)
    If intEnd = 0 Or intEnd = Null Then Err.Raise 1
    If intEnd <> Len(strData) Then Err.Raise 1 'if the last " isn't at the end of the data
    strData = Decode(Mid(strData, intStart, intEnd - intStart))
    If Len(strData) = 0 Then Err.Raise 1
    
    Dim intTopIndex As Integer
    Select Case strLetter
        Case "A"
            intTopIndex = UBound(psPosition.LaterThan)
            ReDim Preserve psPosition.LaterThan(0 To intTopIndex + 1)
            psPosition.LaterThan(intTopIndex + 1).Repeat = intRepetitions
            psPosition.LaterThan(intTopIndex + 1).Text = strData
        Case "B"
            intTopIndex = UBound(psPosition.PreviousTo)
            ReDim Preserve psPosition.PreviousTo(0 To intTopIndex + 1)
            psPosition.PreviousTo(intTopIndex + 1).Repeat = intRepetitions
            psPosition.PreviousTo(intTopIndex + 1).Text = strData
        Case Else
            Err.Raise 1
    End Select
Next intCount
End Sub
