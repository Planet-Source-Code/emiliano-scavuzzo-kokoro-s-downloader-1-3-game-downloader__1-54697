Attribute VB_Name = "Initiation"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

Public Sub Initiations()
m_strTempPath = App.Path + BarIfIsNotRoot + TEMP_FOLDER + "\" 'directorio Temp (no borrar)
m_strDownloadsPath = App.Path + BarIfIsNotRoot + DOWN_FOLDER + "\" 'directorio Download
m_strPreviewPath = m_strTempPath + PREVIEW_NAME
m_strGameListPath = m_strTempPath + GAME_LIST_LOCAL_NAME
Load_Configuration

'prepare the listview for the game list
With frmMain.lsvGameList.ColumnHeaders
.Add , , m_strHeadListColumn1, 2840, 0
.Add , , m_strHeadListColumn2, 800, 0
.Add , , m_strHeadListColumn3, 1650, 0
.Add , , m_strHeadListColumn4, 1100, 1
End With

'initiate checking and searching labels
frmMain.lblCheckSearch(1) = CHECKING_STRING
frmMain.lblCheckSearch(2) = SEARCHING_STRING

'initiate the link
    With frmMain.lblLink
     .Move 0, 0
     .ForeColor = COLOR_INACTIVE_LINK
    End With
    With frmMain
      .picLink.Move .picLink.Left, .picLink.Top, .lblLink.Width, .lblLink.Height
    End With
    
'initiate the style array
ReDim FlashList(0 To 0)
ReDim ColorList(0 To 0)

m_blnUpdateCheckNeeded = True
m_blnUpdating = False

'help tooltips
  With clsToolTip3
       .DelayTime = 100
       .VisibleTime = 30000
       .TipWidth = 180
       .BkColor = vbBlack
       .TxtColor = vbYellow
       .Style = ttStyleStandard
       .SetToolTipObj frmMain.Picture3.hwnd, "Write the game code in the text box. To get the codes you can go to the Search section or visit the web site " + WEB_ADDRESS
  End With
  
  With clsToolTip4
       .DelayTime = 100
       .VisibleTime = 30000
       .TipWidth = 205
       .BkColor = vbBlack
       .TxtColor = vbYellow
       .Style = ttStyleStandard
       .SetToolTipObj frmMain.Picture4.hwnd, "Info about the game. The required software can be downloaded from the web site " + WEB_ADDRESS
  End With

  With clsToolTip5
       .DelayTime = 100
       .VisibleTime = 30000
       .TipWidth = 250
       .BkColor = vbBlack
       .TxtColor = vbYellow
       .Style = ttStyleStandard
       .SetToolTipObj frmMain.Picture5.hwnd, "Check the packages you wish to download. The packages are saved in the folder Downloads."
  End With
     
  With clsToolTip6
       .DelayTime = 100
       .VisibleTime = 30000
       .TipWidth = 245
       .BkColor = vbBlack
       .TxtColor = vbYellow
       .Style = ttStyleStandard
       .SetToolTipObj frmMain.Picture6.hwnd, "Click the search button to look up for the database. The database could be temporarily out of order while updating. To get the same game list visit " + WEB_ADDRESS
  End With
  
End Sub

'draw the form to be round-shaped
Public Sub DrawForm()

Dim ptPointClient As POINTAPI 'first point on the client area
Dim rtRectWindow As RECT 'window corners position
Dim rtRecMain As RECT

ClientToScreen frmMain.hwnd, ptPointClient
GetWindowRect frmMain.hwnd, rtRectWindow

rtRecMain.Left = ptPointClient.x - rtRectWindow.Left
rtRecMain.Top = ptPointClient.y - rtRectWindow.Top
rtRecMain.Right = rtRecMain.Left + frmMain.ScaleWidth + 1
rtRecMain.Bottom = rtRecMain.Top + frmMain.ScaleHeight + 1

Dim lngWindowRegion As Long
lngWindowRegion = CreateRoundRectRgn(rtRecMain.Left, rtRecMain.Top, rtRecMain.Right, rtRecMain.Bottom, 15, 15)
SetWindowRgn frmMain.hwnd, lngWindowRegion, True
End Sub

'Erase the close, move and size from the system menu
Public Sub CreateMenu()
Dim lngMenuHandle As Long
   
lngMenuHandle = GetSystemMenu(frmMain.hwnd, 0&)
DeleteMenu lngMenuHandle, SC_MAXIMIZE, MF_BYCOMMAND
DeleteMenu lngMenuHandle, SC_SIZE, MF_BYCOMMAND
DeleteMenu lngMenuHandle, SC_MOVE, MF_BYCOMMAND
End Sub

Public Sub CenterWindowOnScreen()
Dim ptPointClient As POINTAPI 'first point on the client area
Dim rtRectWindow As RECT 'window corners position
Dim rtRecMain As RECT

ClientToScreen frmMain.hwnd, ptPointClient
GetWindowRect frmMain.hwnd, rtRectWindow

rtRecMain.Left = ptPointClient.x - rtRectWindow.Left
rtRecMain.Top = ptPointClient.y - rtRectWindow.Top
frmMain.Top = ((Screen.Height - frmMain.Height) / 2) - ((rtRecMain.Top * Screen.TwipsPerPixelY) / 2)
frmMain.Left = ((Screen.Width - frmMain.Width) / 2) - ((rtRecMain.Left * Screen.TwipsPerPixelX) / 2)
End Sub

'Load the configuration and saved data
Private Sub Load_Configuration()
Const SOFTWARE As String = "Software\"

If getdword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUseProxy) = 1 Then m_blnUseProxy = True
If getdword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoPreview) = 1 Then m_blnDontDownPrev = True
If getdword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUseOwnServer) = 1 Then m_blnUseOwnServer = True
If getdword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoColors) = 1 Then m_blnDontUseColors = True
If getdword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoBackup) = 1 Then m_blnDontUseBackup = True

'the proxy and port are loaded even if you don't use them (show gray textboxes)
m_strProxy = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regProxy))
m_lngPort = getdword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regPort)
'the own server is loaded even if you don't use it (show gray textbox)
m_strOwnServer = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regOwnServer))

If (m_strProxy = "") Or (m_lngPort < 0) Or (m_lngPort > 65535) Then
    m_strProxy = ""
    m_lngPort = 0
    m_blnUseProxy = False
End If

If m_strOwnServer = "" Then
    m_blnUseOwnServer = False
End If

m_strDatabase = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regDatabaseMain))
m_strDatabaseBkp = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regDatabaseBkp))
m_strUpdateServer = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUpdateServer))
m_strExtensionInfoMain = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regExtInfoMain))
m_strExtensionInfoBkp = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regExtInfoBkp))
m_strGameListNameMain = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regListNameMain))
m_strGameListNameBkp = Trim(getstring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regListNameBkp))

If Not IsURL(m_strDatabase) Then m_strDatabase = DATABASE_MAIN_DEFAULT
If Not IsURL(m_strDatabaseBkp) Then m_strDatabaseBkp = DATABASE_BKP_DEFAULT
If Not IsURL(m_strUpdateServer) Then m_strUpdateServer = UPDATE_SERVER_DEFAULT
If Trim(m_strExtensionInfoMain) = "" Then m_strExtensionInfoMain = EXTENSION_INFO_DEFAULT
If Trim(m_strExtensionInfoBkp) = "" Then m_strExtensionInfoBkp = EXTENSION_INFO_DEFAULT
If Trim(m_strGameListNameMain) = "" Then m_strGameListNameMain = GAME_LIST_NAME_DEFAULT
If Trim(m_strGameListNameBkp) = "" Then m_strGameListNameBkp = GAME_LIST_NAME_DEFAULT

End Sub
