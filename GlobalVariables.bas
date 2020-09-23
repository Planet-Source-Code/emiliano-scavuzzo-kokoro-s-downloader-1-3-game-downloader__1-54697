Attribute VB_Name = "GlobalVariables"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

'web page address
Public Const WEB_ADDRESS As String = "http://www.juegosenteros.cjb.net/"
'program name
Public Const APP_NAME As String = "Kokoro Downloader 1.3"
'program version
Public Const VERSION As Integer = 13

Public Const TEMP_FOLDER As String = "Temp" 'name for the temp folder
Public Const DOWN_FOLDER As String = "Downloads" 'name for the download folder

Public Const PREVIEW_NAME As String = "preview.jpg" 'name for the preview file
Public Const GAME_LIST_LOCAL_NAME As String = "list.fgl" 'local name for the game list

Public Const DATABASE_MAIN_DEFAULT As String = "http://fgdata.my100megs.com/"
Public m_strDatabase As String 'main server where the game info files are stored
Public Const DATABASE_BKP_DEFAULT As String = "http://fgbackup.250free.com/"
Public m_strDatabaseBkp As String 'backup server
Public Const UPDATE_SERVER_DEFAULT As String = "http://www.geocities.com/fgupdate/update.txt"
Public m_strUpdateServer As String 'update server
Public Const EXTENSION_INFO_DEFAULT As String = "fgi"
Public m_strExtensionInfoMain As String 'extension for the info files on the main server
Public m_strExtensionInfoBkp As String 'extension for the info files on the backup server
Public Const GAME_LIST_NAME_DEFAULT As String = "list.fgl"
Public m_strGameListNameMain As String 'name of the game list file on the main server
Public m_strGameListNameBkp As String 'name of the game list file on the backup server

'lsvGameList headers strings
Public Const m_strHeadListColumn1 As String = "Name"
Public Const m_strHeadListColumn2 As String = "Code"
Public Const m_strHeadListColumn3 As String = "Type"
Public Const m_strHeadListColumn4 As String = "Size"

'animated titles strings
Public Const CHECKING_STRING As String = "Checking"
Public Const SEARCHING_STRING As String = "Searching"

'Notes windows title
Public Const NOTES_CAPTION As String = "Notes for "

'string to check info integrity
Public Const INFO_VERIFICATION_LINE As String = "FullGames Info"
'string to check game list integrity
Public Const LIST_VERIFICATION_LINE As String = "FullGames GameList"
'string to check update integrity
Public Const UPDATE_VERIFICATION_LINE As String = "FullGames Update"

'preferences
Public m_blnUseProxy As Boolean
Public m_blnDontDownPrev As Boolean
Public m_blnDontUseColors As Boolean
Public m_blnDontUseBackup As Boolean
Public m_strProxy As String
Public m_lngPort As Long
Public m_blnUseOwnServer As Boolean
Public m_strOwnServer As String
'name of the registry data
Public Const m_regUseProxy As String = "UseProxy"
Public Const m_regProxy As String = "Proxy"
Public Const m_regPort As String = "Port"
Public Const m_regUseOwnServer As String = "UseOwnServer"
Public Const m_regOwnServer As String = "OwnServer"
Public Const m_regNoPreview As String = "NoPreview"
Public Const m_regNoColors As String = "NoListColors"
Public Const m_regNoBackup As String = "NoBackUpServer"
Public Const m_regDatabaseMain As String = "Dabase"
Public Const m_regDatabaseBkp As String = "DatabaseBkp"
Public Const m_regUpdateServer As String = "UpdateServer"
Public Const m_regExtInfoMain As String = "ExtInfoMain"
Public Const m_regExtInfoBkp As String = "ExtInfoBkp"
Public Const m_regListNameMain As String = "ListNameMain"
Public Const m_regListNameBkp As String = "ListNameBkp"

'kind of downloads (KODs) to download games
Public Const DOWN_BRIEFCASE As Byte = 1

'download class
Public m_clsDownload As cCommon

'Constants that are used with ReadLine to read a line
'from the game info.
Public Const LN_FIRST       As Byte = 1
Public Const LN_UPD_SRV     As Byte = 2
Public Const LN_PREV_ADDR   As Byte = 3
Public Const LN_NAME        As Byte = 4
Public Const LN_TYPE        As Byte = 5
Public Const LN_TOT_SIZE    As Byte = 6
Public Const LN_PACK_NUM    As Byte = 7
Public Const LN_SOFTWARE    As Byte = 8
Public Const LN_DOWN_KIND   As Byte = 9
Public Const LN_NOTES       As Byte = 10
Public Const LN_END         As Byte = 10 'end of the section

'update file lines
Public Const LN_VER         As Byte = 2
Public Const LN_MAINSRV     As Byte = 3
Public Const LN_BKPSRV      As Byte = 4

Public m_strTempPath As String  'path of temp folder
Public m_strDownloadsPath As String 'path of downloads folder
Public m_strPreviewPath As String 'path of preview file
Public m_strPreviewAddress As String 'preview file URL address
Public m_strGameListPath As String 'path of game list file

Public clsToolTip3 As New cToolTip 'help tooltips
Public clsToolTip4 As New cToolTip
Public clsToolTip5 As New cToolTip
Public clsToolTip6 As New cToolTip

Public m_strInfo As String 'game info
Public m_strUpdate As String 'update from update server

'Mouse position used to move window
Public m_pntMousePos As POINTAPI

'This variable stores the mouse button that was pressed
'before the program receives a double-click message, so
'we can know if it was a left or right button double-click.
Public m_intLastMouseButton As Integer

'link colors
Public Const COLOR_ACTIVE_LINK = vbRed
Public Const COLOR_INACTIVE_LINK = vbBlue

'if TRUE the windows is compacted
Public m_blnCompactWindow As Boolean
'if TRUE there is download in progress
Public m_blnIsDownloading As Boolean
'if TRUE we are trying the main server
Public m_blnInfoFirstTry As Boolean
'if TRUE we are trying the main server
Public m_blnListFirstTry As Boolean
'if TRUE that means we haven't searched for updates
Public m_blnUpdateCheckNeeded As Boolean
'If TRUE the control must be passed to the system info,
'if not the control must be passed to the list system.
Public m_blnUpdateByInfo As Boolean
'If TRUE the update system is working
Public m_blnUpdating As Boolean

Private Type TypeSpeed
    Quantity(1 To 10) As Long
    QuanLoaded As Byte
    LastQuantity As Long
End Type

Public Speed As TypeSpeed 'data to calculate speed


Private Type TypeStyle
    Bold As Boolean
    Color As Long
End Type

'flash style
Private Type FlashListType
    Key As String
    First As TypeStyle
    Second As TypeStyle
End Type

'simple style
Private Type ColorListType
    Key As String
    Estilo As TypeStyle
End Type

Public FlashList() As FlashListType 'array of flashing items
Public ColorList() As ColorListType 'array of simple items
