Attribute VB_Name = "ListSeeker"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

'*********** Game list file format ***********
' 1) 'FullGames GameList'  (see LIST_VERIFICATION_LINE)
' 2) Update file address
' 3-X) List
'
'Format of each game line:
''Icon', 'Game title', 'Code', 'Type', 'Size [MB/KB]', 'Style'
' NOTES: 1) Space after the comma is optional
'        2) Use a point to separate decimal part from
'           the game size
Option Explicit

'This sub is called when an error occurs when
'trying to access the list file.
Public Sub ErrorAccessingList()
CleanGameList
HideSearching
frmMain.lsvGameList.Visible = True
MessageBox frmMain.hwnd, "Error accessing the game list file.", "Error", MB_ICONERROR
EnableSearch
End Sub

'This sub is called when an error occurs when
'trying to read the list file.
Public Sub ErrorReadingList()
CleanGameList
HideSearching
frmMain.lsvGameList.Visible = True
MessageBox frmMain.hwnd, "The game information seems damaged. Try later", "Error", MB_ICONERROR
EnableSearch
End Sub

'Show the animated title
Public Sub ShowSearching()
With frmMain
    .picCircle(2).Picture = .imlMiscellaneous.ListImages(1).Picture
    .Timer_Circle(2).Interval = 1000
    .lblCheckSearch(2).Visible = True
    .picCircle(2).Visible = True
End With
End Sub

'Hide the animated title
Public Sub HideSearching()
With frmMain
    .picCircle(2).Visible = False
    .lblCheckSearch(2).Visible = False
    .Timer_Points(2).Enabled = False
    .Timer_Circle(2).Enabled = False
End With
End Sub

'Clean the game list
Public Sub CleanGameList()
With frmMain
    .lsvGameList.ListItems.Clear
    .lsvGameList.SortKey = 0
    .lsvGameList.SortOrder = lvwAscending
End With
End Sub

'Enable list and search button
Public Sub EnableSearch()
With frmMain
    .cmdSearch.Enabled = True
    .cmdCancelSearch.Enabled = False
    .lsvGameList.Enabled = True
End With
End Sub

'Disable list and search button
Public Sub DisableSearch()
With frmMain
    .cmdSearch.Enabled = False
    .cmdCancelSearch.Enabled = True
    .lsvGameList.Enabled = False
End With
End Sub

'If the list wasn't found
Public Sub ListNoFound()
HideSearching
frmMain.lsvGameList.Visible = True
MessageBox frmMain.hwnd, "Database not found. Try later.", "Error", MB_ICONERROR
EnableSearch
End Sub

'If the list was found
Public Sub ListFound()
Dim intFileHandle As Integer
intFileHandle = FreeFile

On Error GoTo ErrorOpening
Open m_strGameListPath For Input Lock Read Write As #intFileHandle

On Error GoTo ErrorReading
Dim strReadLine As String
Input #intFileHandle, strReadLine 'read the verification line
Input #intFileHandle, strReadLine 'read the update file address line

Do Until (EOF(intFileHandle))
Line Input #intFileHandle, strReadLine
AddToList (strReadLine)
Loop

Close #intFileHandle

If m_blnDontUseColors = False Then ShowListColors

HideSearching
frmMain.lsvGameList.Visible = True
EnableSearch

Exit Sub
ErrorOpening:
    Close #intFileHandle
    ErrorAccessingList
    Exit Sub
ErrorReading:
    Close #intFileHandle
    ErrorReadingList
    Exit Sub
End Sub

'This sub receives a complete line from the game list
'file and write every piece of information on the
'correct column.
'If it receives wrong data it raises an error
Private Sub AddToList(ByVal strLine As String)
Dim intLastLine As Integer
intLastLine = frmMain.lsvGameList.ListItems.Count + 1

Dim intIconIndex As Integer
Dim strData As String
'first data
strData = Trim(DataFromLine(strLine, 1))
If Not IsNumeric(strData) Then Err.Raise 1
intIconIndex = Val(strData)
If (intIconIndex < 1) Or (intIconIndex > frmMain.imlIcons.ListImages.Count) Then Err.Raise 1

'second data
strData = Trim(DataFromLine(strLine, 2))
If strData = "" Then Err.Raise 1
strData = Decode(strData)
frmMain.lsvGameList.ListItems.Add , , strData, , intIconIndex

'third data
strData = Trim(DataFromLine(strLine, 3))
If strData = "" Then Err.Raise 1
strData = Decode(strData)
frmMain.lsvGameList.ListItems(intLastLine).ListSubItems.Add , , strData

'fourth data
strData = Trim(DataFromLine(strLine, 4))
If strData = "" Then Err.Raise 1
strData = Decode(strData)
frmMain.lsvGameList.ListItems(intLastLine).ListSubItems.Add , , strData

'fifth data
strData = Trim(DataFromLine(strLine, 5))
If Right(strData, 3) <> " MB" And Right(strData, 3) <> " KB" Then Err.Raise 1
If Not IsNumeric(Left(strData, Len(strData) - 3)) Then Err.Raise 1
frmMain.lsvGameList.ListItems(intLastLine).ListSubItems.Add , , CStr(CDbl(Val(Trim(Left(strData, Len(strData) - 3))))) + Right(strData, 3)

'sixth data
strData = Trim(DataFromLine(strLine, 6))
If (Len(strData) <> 2) And (Len(strData) <> 4) Then Err.Raise 1
SetStyle strData, intLastLine
End Sub

'Read the style data and assign the colors to the lines
'or creates the array for the flashing style.
'On error it raises error 1.
Private Sub SetStyle(ByVal strStyle As String, ByVal intLine As Integer)
Dim strFirstL As String
Dim strSecondL As String
Dim blnOneBold As Boolean
Dim lngOneCol As Long
Dim strKey As String
Dim intLastStyle As Integer

strKey = "l" + Trim(Str(intLine)) 'create the key

strFirstL = Left(strStyle, 1) 'first letter
strSecondL = Mid(strStyle, 2, 1) 'second letter

If StrComp(strFirstL, "B", vbTextCompare) = 0 Then
    blnOneBold = True
Else 'if it isn't bold
    If StrComp(strFirstL, "H", vbTextCompare) <> 0 Then Err.Raise 1
    blnOneBold = False
End If

lngOneCol = LetterToColor(strSecondL)
Dim intLength As Integer
intLength = Len(strStyle)

Select Case intLength

Case 2  'simple style
    frmMain.lsvGameList.ListItems(intLine).Key = strKey
    intLastStyle = UBound(ColorList)
    ReDim Preserve ColorList(0 To intLastStyle + 1)
    With ColorList(intLastStyle + 1)
        .Key = strKey
        .Estilo.Bold = blnOneBold
        .Estilo.Color = lngOneCol
    End With

Case 4  'flashing style

    Dim strThirdL As String
    Dim strFourthL As String
    Dim blnTwoBold As Boolean
    Dim lngTwoCol As Long
    
    strThirdL = Mid(strStyle, 3, 1)
    strFourthL = Right(strStyle, 1)
    If StrComp(strThirdL, "B", vbTextCompare) = 0 Then
        blnTwoBold = True
    Else 'if it isn't bold
        If StrComp(strThirdL, "H", vbTextCompare) <> 0 Then Err.Raise 1
        blnTwoBold = False
    End If
    lngTwoCol = LetterToColor(strFourthL)
    
    frmMain.lsvGameList.ListItems(intLine).Key = strKey
    intLastStyle = UBound(FlashList)
    ReDim Preserve FlashList(0 To intLastStyle + 1)
    
    With FlashList(intLastStyle + 1)
        .Key = strKey
        .First.Bold = blnOneBold
        .First.Color = lngOneCol
        .Second.Bold = blnTwoBold
        .Second.Color = lngTwoCol
    End With
        
End Select

End Sub

'Convert a letter to a long value that represents the
'color. On error it raises error 1.
Private Function LetterToColor(ByVal strLetter As String) As Long
If Len(strLetter) <> 1 Then Err.Raise 1

Select Case strLetter
    Case "N", "n"
        LetterToColor = &H80000008 'relative black
    Case "Q", "q"
        LetterToColor = RGB(0, 0, 0) 'absolute black
    Case "Z", "z"
        LetterToColor = RGB(255, 255, 255) 'white
    Case "A", "a"
        LetterToColor = RGB(0, 0, 255) 'blue
    Case "V", "v"
        LetterToColor = RGB(0, 255, 0) 'green
    Case "R", "r"
        LetterToColor = RGB(255, 0, 0) 'red
    Case "M", "m"
        LetterToColor = RGB(255, 255, 0) 'yellow
    Case "C", "c"
        LetterToColor = RGB(0, 255, 255) 'cyan (light blue)
    Case "G", "g"
        LetterToColor = RGB(255, 0, 255) 'magenta (purple)
    Case "W", "w"
        LetterToColor = RGB(0, 0, 191) 'dark blue
    Case "E", "e"
        LetterToColor = RGB(128, 0, 0) 'dark red
    Case "T", "t"
        LetterToColor = RGB(0, 130, 0) 'dark green
    Case "Y", "y"
        LetterToColor = RGB(127, 13, 151) 'dark purple
    Case "U", "u"
        LetterToColor = RGB(226, 110, 24) 'orange
    Case "I", "i"
        LetterToColor = RGB(209, 113, 188) 'pink
    Case "O", "o"
        LetterToColor = RGB(160, 172, 169) 'light gray
    Case "P", "p"
        LetterToColor = RGB(125, 135, 133) 'dark gray
    Case "S", "s"
        LetterToColor = RGB(59, 229, 187) 'bluish green
    Case "D", "d"
        LetterToColor = RGB(255, 255, 128) 'canary yellow
    Case "F", "f"
        LetterToColor = RGB(172, 173, 53) 'light brown
    Case "J", "j"
        LetterToColor = RGB(125, 126, 52) 'dark brown
    Case "K", "k"
        LetterToColor = RGB(149, 190, 245) 'sky blue
    Case "L", "l"
        LetterToColor = RGB(221, 192, 173) 'skin color
    Case "X", "x"
        LetterToColor = RGB(243, 105, 13) 'strong orange
    Case Else
        Err.Raise 1
End Select
End Function

'This sub sorts the game list.
'First it moves the the last part (MB/KB) to the front
'and adds 5 ceroes to the number taking into account
'the digits on the front.
'ie '23.4 MB' => 'MB00023.4'
'Then it orders alphabetically in ascending or descending
'way and converts the string back.
Public Sub OrderBySize()
Static blnFourthAscend As Boolean

Dim intIndex As Integer
Dim strData As String

With frmMain.lsvGameList
    For intIndex = 1 To .ListItems.Count
        strData = .ListItems(intIndex).ListSubItems(3)
        strData = Right(strData, 2) + Format(CDbl(Left(strData, Len(strData) - 3)), "00000.##")
        .ListItems(intIndex).ListSubItems(3) = strData
    Next intIndex

    .SortKey = 3
    If blnFourthAscend = True Then
        .SortOrder = lvwDescending
        blnFourthAscend = False
    Else
        .SortOrder = lvwAscending
        blnFourthAscend = True
    End If

    .Sorted = True
    .Sorted = False

    For intIndex = 1 To .ListItems.Count
        strData = .ListItems(intIndex).ListSubItems(3)
        strData = CStr(CDbl(Right(strData, Len(strData) - 2))) + " " + Left(strData, 2)
        .ListItems(intIndex).ListSubItems(3) = strData
    Next intIndex
End With

End Sub
