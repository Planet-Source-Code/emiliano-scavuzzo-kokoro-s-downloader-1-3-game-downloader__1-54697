Attribute VB_Name = "Miscellaneous"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

'Enable the ok button and the code textbox and
'disable the cancel button.
Public Sub EnableForInfo()
With frmMain
    .cmdCancelInfo.Enabled = False
    .txtCode.Enabled = True
    .cmdOK.Enabled = True
End With
End Sub

'Disable the ok button and the code textbox and
'enable the cancel button.
Public Sub DisableForInfo()
With frmMain
    .txtCode.Enabled = False
    .cmdOK.Enabled = False
    .cmdCancelInfo.Enabled = True
End With
End Sub

'Disable buttons after cmdOk
Public Sub DisableOne()
With frmMain
    DisableForInfo
    .cmdSelectAll.Enabled = False
    .cmdDownload.Enabled = False
End With
End Sub

'Hide to be able to show "Checking....."
Public Sub HideUno()
With frmMain
    .Label2.Visible = False
    .Label3.Visible = False
    .Label4.Visible = False
    .Label5.Visible = False
    .Label6.Visible = False
    .lblName.Visible = False
    .lblType.Visible = False
    .lblSize.Visible = False
    .lblNumPack.Visible = False
    .lblRequiredSoft.Visible = False
    .picPreview.Visible = False
    .cmdNotes.Visible = False
End With
End Sub

'Show the circle and the label "Checking..."
Public Sub ShowChecking()
    With frmMain
    .picCircle(1).Picture = .imlMiscellaneous.ListImages(1).Picture
    .Timer_Circle(1).Interval = 1000
    .lblCheckSearch(1).Visible = True
    .picCircle(1).Visible = True
    End With
End Sub

'Hide the circle and the label "Checking..."
Public Sub HideChecking()
With frmMain
    .picCircle(1).Visible = False
    .lblCheckSearch(1).Visible = False
    .Timer_Points(1).Enabled = False
    .Timer_Circle(1).Enabled = False
End With
End Sub

'Show the 'Wrong code' label.
Public Sub ShowWrongCode()
frmMain.lblWrongCode.Visible = True
End Sub

'Hide the 'Wrong code' label.
Public Sub HideWrongCode()
frmMain.lblWrongCode.Visible = False
End Sub

'Clean the info frame.
Public Sub CleanInfoScreen()
With frmMain
    Set .picPreview.Picture = LoadPicture
    .lblName.Caption = "********"
    .lblType.Caption = "******"
    .lblSize.Caption = "*******"
    .lblNumPack.Caption = "********"
    .lblRequiredSoft.Caption = "********"
End With
End Sub

Public Sub ShowOne()
With frmMain
    .picPreview.Visible = True
    .Label2.Visible = True
    .Label3.Visible = True
    .Label4.Visible = True
    .Label5.Visible = True
    .Label6.Visible = True
    .lblName.Visible = True
    .lblType.Visible = True
    .lblSize.Visible = True
    .lblNumPack.Visible = True
    .lblRequiredSoft.Visible = True
End With
End Sub

'Loads the preview image in the preview picturebox.
'If something goes wrong set the picturebox blank.
Public Sub LoadPreview()
On Error GoTo ErrorHandler
frmMain.picPreview.Picture = LoadPicture(m_strPreviewPath)
Exit Sub
ErrorHandler:
    Debug.Print "Error loading preview"
    frmMain.picPreview.Picture = LoadPicture
End Sub

'Retrieves the byte that represents the kind of download.
Function KindOfDownload() As Byte
KindOfDownload = Val(Trim(ReadLine(m_strInfo, LN_DOWN_KIND)))
End Function

'Returns a number that represents which tab is visible.
' 1 = first
' 2 = second
Function WhichTab() As Byte
WhichTab = frmMain.TabStrip1.SelectedItem.Index
End Function

'Show the specified tab.
Public Sub ChangeToTab(ByVal bytTabIndex As Byte)
With frmMain
    Select Case bytTabIndex
        Case 1
            .frmDownload.ZOrder (0)
            .frmSearch.Visible = False
            .TabStrip1.Tabs.Item(1).Selected = True
            .frmDownload.Visible = True
            If SetThisFocus(.txtCode.hwnd) = 0 Then SetThisFocus .Picture3.hwnd

        Case 2
            .frmSearch.ZOrder (0)
            .frmDownload.Visible = False
            .TabStrip1.Tabs.Item(2).Selected = True
            .frmSearch.Visible = True
            If SetThisFocus(.lsvGameList.hwnd) = 0 Then SetThisFocus .Picture6.hwnd
    End Select
End With
End Sub

'This sub is called when the info file contains
'wrong information.
Public Sub ErroneousInfo()
CleanInfoScreen
HideWrongCode
HideChecking
ShowOne
MessageBox frmMain.hwnd, "The game information seems damaged. Try later", "Error", MB_ICONERROR
EnableForInfo
End Sub

'This sub is called when the version number in the update
'file doesn't match with this version.
Public Sub UpdateNeededForVersion()
If m_blnUpdateByInfo Then
    CleanInfoScreen
    ShowOne
    HideChecking
    MessageBox frmMain.hwnd, "There is a new version of Kokoro's Downloader." + vbCrLf + "Visit the official web site for more details.", "New version", MB_ICONINFORMATION
    EnableForInfo
    SetThisFocus frmMain.txtCode.hwnd
Else
    HideSearching
    frmMain.lsvGameList.Visible = True
    MessageBox frmMain.hwnd, "There is a new version of Kokoro's Downloader." + vbCrLf + "Visit the official web site for more details.", "New version", MB_ICONINFORMATION
    EnableSearch
End If
End Sub

'This sub is called when the kind of download in the info
'file isn't available on this version.
Public Sub UpdateNeededForThisGame()
CleanInfoScreen
HideChecking
MessageBox frmMain.hwnd, "To download this game you need to update Kokoro's Downloader." + vbCrLf + "Visit the official web site for more details.", "Error", MB_ICONINFORMATION
EnableForInfo
SetThisFocus frmMain.txtCode.hwnd
End Sub

'Checks the info for wrong data.
'If it finds an error it raises an error.
Public Sub CheckInfo()
Dim strNameTemp As String
strNameTemp = ReadLine(m_strInfo, LN_NAME)
If Trim(strNameTemp) = "" Then Err.Raise 1
If CountHowManyTimes(Chr(34), strNameTemp) Then Err.Raise 1
If CountHowManyTimes("*", strNameTemp) Then Err.Raise 1
If CountHowManyTimes("/", strNameTemp) Then Err.Raise 1
If CountHowManyTimes(":", strNameTemp) Then Err.Raise 1
If CountHowManyTimes("<", strNameTemp) Then Err.Raise 1
If CountHowManyTimes(">", strNameTemp) Then Err.Raise 1
If CountHowManyTimes("?", strNameTemp) Then Err.Raise 1
If CountHowManyTimes("\", strNameTemp) Then Err.Raise 1
If CountHowManyTimes("|", strNameTemp) Then Err.Raise 1
If Not IsURL(ReadLine(m_strInfo, LN_UPD_SRV)) Then Err.Raise 1
If Trim(ReadLine(m_strInfo, LN_TYPE)) = "" Then Err.Raise 1
If Not IsNumeric(Trim(ReadLine(m_strInfo, LN_TOT_SIZE))) Or Val(Trim(ReadLine(m_strInfo, LN_TOT_SIZE))) = 0 Then Err.Raise 1
If Not IsNumeric(Trim(ReadLine(m_strInfo, LN_PACK_NUM))) Or Val(Trim(ReadLine(m_strInfo, LN_PACK_NUM))) = 0 Then Err.Raise 1
If Trim(ReadLine(m_strInfo, LN_SOFTWARE)) = "" Then Err.Raise 1
If Not IsNumeric(Trim(ReadLine(m_strInfo, LN_DOWN_KIND))) Then Err.Raise 1
End Sub

'Center a window relatively to the main form and avoid
'placing the window out of the screen.
Public Sub CenterOnMain(ByRef frmCenter As Form)
frmCenter.Top = (frmMain.Height - frmCenter.Height) / 2 + frmMain.Top
frmCenter.Left = (frmMain.Width - frmCenter.Width) / 2 + frmMain.Left
Dim lngTopMaximum As Long
lngTopMaximum = Screen.Height - frmCenter.Height
If frmCenter.Top > lngTopMaximum Then frmCenter.Top = lngTopMaximum

Dim lngLeftMaximum As Long
lngLeftMaximum = Screen.Width - frmCenter.Width
If frmCenter.Left > lngLeftMaximum Then frmCenter.Left = lngLeftMaximum

If frmCenter.Left < 0 Then frmCenter.Left = 0
End Sub

'Change the color and bold style from a game list line
'through the key.
Public Sub ListColorKey(ByVal strKey As String, ByVal lngColor As Long, Optional blnBold As Boolean = False)
With frmMain.lsvGameList
    .ListItems(strKey).ForeColor = lngColor
    If IsMissing(blnBold) Then
        .ListItems(strKey).Bold = False
    Else
        .ListItems(strKey).Bold = blnBold
    End If
End With
End Sub

'Show the simple style and flashing styles from the list.
Public Sub ShowListColors()
'flashing style
If UBound(FlashList) > 0 Then frmMain.Timer_Styles.Enabled = True

'simple style
Dim intCounter As Integer
For intCounter = 1 To UBound(ColorList)
    With frmMain.lsvGameList.ListItems(ColorList(intCounter).Key)
        .Bold = ColorList(intCounter).Estilo.Bold
        .ForeColor = ColorList(intCounter).Estilo.Color
    End With
Next intCounter
End Sub

'Hide the simple style and flashing styles from the list.
Public Sub HideListColors()
Dim intCounter As Integer

'flashing style
frmMain.Timer_Styles.Enabled = False
For intCounter = 1 To UBound(FlashList)
    With frmMain.lsvGameList.ListItems(FlashList(intCounter).Key)
        .Bold = False
        .ForeColor = &H80000008
    End With
Next intCounter

'simple style
For intCounter = 1 To UBound(ColorList)
    With frmMain.lsvGameList.ListItems(ColorList(intCounter).Key)
        .Bold = False
        .ForeColor = &H80000008
    End With
Next intCounter
End Sub

'Returns the info file extension according to which
'server we are using (main/backup).
Public Function InfoExtension() As String
If m_blnInfoFirstTry = True Then
    InfoExtension = "." + m_strExtensionInfoMain
Else
    InfoExtension = "." + m_strExtensionInfoBkp
End If
End Function

'Returns the info server to use depending on own server
'option and if we tried the main server or not.
Public Function InfoServer() As String
If m_blnUseOwnServer Then
    InfoServer = m_strOwnServer
Else
    If m_blnInfoFirstTry = True Then
        InfoServer = m_strDatabase
    Else
        InfoServer = m_strDatabaseBkp
    End If
End If
End Function

'Returns the list server to use depending on own server
'option and if we tried the main server or not.
Public Function ListServer() As String
If m_blnUseOwnServer Then
    ListServer = m_strOwnServer
Else
    If m_blnListFirstTry = True Then
        ListServer = m_strDatabase
    Else
        ListServer = m_strDatabaseBkp
    End If
End If
End Function

'Returns the game list name according to which server
'we are using (main/backup).
Public Function GameListName() As String
If m_blnInfoFirstTry = True Then
    GameListName = m_strGameListNameMain
Else
    GameListName = m_strGameListNameBkp
End If
End Function
