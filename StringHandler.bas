Attribute VB_Name = "StringHandler"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

Private Const MULT As Integer = 55 'encryption changer

'Receives a string with a number (bytes) and returns
'a string with that number converted to KB/MB.
Public Function Bytes2KB_MB(ByVal strBytes As String) As String
Dim dblSize As Double
If Not IsNumeric(strBytes) Then Err.Raise 1
dblSize = Int(Val(strBytes))
If dblSize >= 1048576 Then 'if it's > 0 = 1 MB
    If (((dblSize / 1048576) * 100) Mod 100) = 0 Then 'if we have precise MBs or X,00XXX
    Bytes2KB_MB = Format(dblSize / 1048576, "#") + " MB"
    Else 'if we don't have precise MBs neither X,00XXX
    Bytes2KB_MB = Format(dblSize / 1048576, "#####0.##") + " MB"
    End If
Else 'if it's < than 1 MB
    If (((dblSize / 1024) * 100) Mod 100) = 0 Then  'if we have precise KBs or X,00XXX
    Bytes2KB_MB = Format(dblSize / 1024, "#") + " KB"
    Else 'if we don't have precise KBs neither X,00XXX
    Bytes2KB_MB = Format(dblSize / 1024, "#####0.##") + " KB"
    End If
End If
End Function

'Returns a backslash (\) if the application path isn't
'the root (if it is root App.Path = C:\, if it isn't
'App.Path = C:\Something).
Public Function BarIfIsNotRoot() As String
BarIfIsNotRoot = "\"
If Right(App.Path, 1) = "\" Then BarIfIsNotRoot = ""
End Function

'Returns the line specified from a string.
'If it reachs the end of the string without reaching the
'line, returns an empty string.
Public Function ReadLine(ByVal strText As String, intLine As Integer) As String
Dim intCount As Integer
Dim intStart As Integer
Dim intEnd As Integer
intStart = 1
For intCount = 1 To intLine - 1
    intStart = InStr(intStart, strText, vbCrLf, vbTextCompare)
    If intStart = 0 Or intStart = Null Then ReadLine = "": Exit Function
    intStart = intStart + 2
Next intCount

intEnd = InStr(intStart, strText, vbCrLf, vbTextCompare)
If intEnd = 0 Or intEnd = Null Then intEnd = Len(strText) + 1
ReadLine = Mid(strText, intStart, intEnd - intStart)
End Function

'Returns how many times strShort appears in strLong.
Public Function CountHowManyTimes(ByRef strShort As String, ByRef strLong As String) As Integer
CountHowManyTimes = 0
If strShort = "" Then Exit Function

Dim lngStart As Long
Dim lngIndex As Long
lngStart = 1
lngIndex = 1
Do
lngStart = lngIndex
CountHowManyTimes = CountHowManyTimes + 1
lngStart = InStr(lngStart, strLong, strShort, vbBinaryCompare)
lngIndex = lngStart + Len(strShort)
Loop Until (lngStart = 0 Or lngStart = Null)

CountHowManyTimes = CountHowManyTimes - 1
End Function

'This function receives a line with data strings separated
'by commas and returns the string number intData.
'If it finds an error raises error 1.
Public Function DataFromLine(ByVal strLine As String, intData As Integer) As String
If intData < 1 Then Err.Raise 1
Dim intNmrOfCommas As Integer
intNmrOfCommas = CountHowManyTimes(",", strLine)
If (intData - 1) > intNmrOfCommas Then Err.Raise 1

Dim lngStart As Long
Dim lngEnd As Long
lngStart = 1
If intData = 1 Then 'if it is the first data string
    If intNmrOfCommas = 0 Then 'if it is the first and last data string (there's only one data string)
        lngEnd = Len(strLine) + 1
    Else 'if it is the first
        lngEnd = InStr(lngStart, strLine, ",", vbTextCompare)
    End If
Else
    Dim intContador As Integer
    For intContador = 1 To intData - 1 'finds the comma previous to the data string
        lngStart = InStr(lngStart + 1, strLine, ",", vbTextCompare)
    Next intContador
    
    If intData - 1 = intNmrOfCommas Then  'if it is the last data string
        lngEnd = Len(strLine) + 1
    Else 'if it isn't the first neither the last data string
        lngEnd = InStr(lngStart + 1, strLine, ",", vbTextCompare)
    End If
    lngStart = lngStart + 1
End If
 
DataFromLine = Mid(strLine, lngStart, lngEnd - lngStart)
End Function

'This function lets the data strings have any character,
'including commas. It uses a method similar to the one
'used on CGI, a '%' followed by a set of two hexadecimal
'characters that represents the ASCII code.
' ie:  '%A1' -> 'A'
Public Function Decode(ByVal strText As String) As String
Dim intRepeat As Integer
intRepeat = CountHowManyTimes("%", strText)
If intRepeat = 0 Then Decode = strText: Exit Function

Dim intIndex As Integer
Dim intPreviousIndex As Integer
Dim intCounter As Integer
intIndex = 1
Decode = ""

For intCounter = 1 To intRepeat
    intPreviousIndex = intIndex
    intIndex = InStr(intIndex, strText, "%")
    Decode = Decode + Mid(strText, intPreviousIndex, intIndex - intPreviousIndex)
    'if after a '%' there are two hexadecimal characters
    If (Len(strText) >= intIndex + 2) And IsNumeric("&H" + Mid(strText, intIndex + 1, 2)) Then
        Decode = Decode + Chr("&H" + Mid(strText, intIndex + 1, 2))
        intIndex = intIndex + 3
    Else
        Decode = Decode + Mid(strText, intIndex, 1)
        intIndex = intIndex + 1
    End If
Next intCounter
If Len(strText) >= intIndex Then Decode = Decode + Right(strText, Len(strText) - intIndex + 1)
End Function

'Returns TRUE if the URL is valid.
Public Function IsURL(ByVal strURL As String) As Boolean
strURL = StrConv(strURL, vbUnicode)
IsURL = (IsValidURL(ByVal 0&, strURL, 0) = S_OK)
End Function

'Returns the host part from a web address.
'ie: 'http://www.dreamers.com/something.txt' => 'www.dreamers.com'
Public Function TakeHost(ByVal strAddress As String) As String
strAddress = Trim(strAddress)
If Left(strAddress, 7) = "http://" Then strAddress = Mid(strAddress, 6, Len(strAddress) - 5)
If Left(strAddress, 2) = "//" Then strAddress = Mid(strAddress, 3, Len(strAddress) - 2)

Dim intStart As Integer
intStart = InStr(1, strAddress, "/", vbTextCompare)
If intStart <> 0 Then strAddress = Left(strAddress, intStart - 1)
TakeHost = strAddress
End Function

'Receives seconds and returns a string in HH:MM:SS format.
Public Function SecondsToHours(ByVal dblSeconds As Double) As String
Dim dblSecondsPart As Double
Dim dblMinutesPart As Double
Dim dblHoursPart As Double
dblSecondsPart = dblSeconds Mod 60
dblSeconds = dblSeconds - dblSecondsPart
dblMinutesPart = (dblSeconds Mod 3600) / 60
dblSeconds = dblSeconds - (dblMinutesPart * 60)
dblHoursPart = dblSeconds / 3600
SecondsToHours = Format(Str(dblHoursPart), "00") + ":" + Format(Str(dblMinutesPart), "00") + ":" + Format(Str(dblSecondsPart), "00")
End Function

'Converst a string into a byte array.
Public Function STRING_TO_BYTES(strString As String, ByRef bytArray() As Byte)
    bytArray = StrConv(strString, vbFromUnicode)
End Function

'Converst a byte array into a string.
Private Function BYTES_TO_STRING(bytBytes() As Byte) As String
    BYTES_TO_STRING = bytBytes
    BYTES_TO_STRING = StrConv(BYTES_TO_STRING, vbUnicode)
End Function

'Unencrypy text.
Public Function UnEncryptText(ByRef strText As String, Optional ByVal Index As Variant) As String
Dim lngCounter As Long
Dim bytTextArray() As Byte
If IsMissing(Index) Then Index = 0

STRING_TO_BYTES strText, bytTextArray()

For lngCounter = 0 To UBound(bytTextArray())
    bytTextArray(lngCounter) = (((bytTextArray(lngCounter) - (lngCounter + Index + 1) * MULT) Mod 256) + 256) Mod 256
Next

UnEncryptText = BYTES_TO_STRING(bytTextArray())
End Function
