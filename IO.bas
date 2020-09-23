Attribute VB_Name = "IO"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

'Make temp folder if it doesn't exist
Public Sub MakeTempFoldIfDoesntExist()
On Error GoTo ErrorHandler
If TEMP_FOLDER <> Dir(Left(m_strTempPath, Len(m_strTempPath) - 1), vbDirectory) Then MkDir m_strTempPath
'In case the sub can't create the folder it goes on
'without raising an error
ErrorHandler:
End Sub

'Make downloads folder if it doesn't exist
Public Sub MakeDownFoldIfDoesntExist()
On Error GoTo ErrorHandler
If DOWN_FOLDER <> Dir(Left(m_strDownloadsPath, Len(m_strDownloadsPath) - 1), vbDirectory) Then MkDir m_strDownloadsPath
'In case the sub can't create the folder it goes on
'without raising an error
ErrorHandler:
End Sub

'Make the folder where the game is saved if it
'doesn't exist
Public Sub MakeGameFoldIfDoesntExist()
On Error GoTo ErrorHandler
MakeDownFoldIfDoesntExist
Dim strGameName As String
Dim strGamePath As String
strGameName = Trim(ReadLine(m_strInfo, LN_NAME))
strGamePath = m_strDownloadsPath + strGameName + "\"
If strGameName <> Dir(Left(strGamePath, Len(m_strDownloadsPath) - 1), vbDirectory) Then
    MkDir strGamePath
End If
'In case the sub can't create the folder it goes on
'without raising an error
ErrorHandler:
End Sub

'Create the preview file if it doesn't exist
'or clear it if it does.
Public Sub MakePreviewIfDoesntExist()
On Error GoTo ErrorHandler
MakeTempFoldIfDoesntExist
If Dir(m_strPreviewPath, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = PREVIEW_NAME Then SetAttr m_strPreviewPath, vbNormal
Dim intFileHandle As Integer
intFileHandle = FreeFile
Open m_strPreviewPath For Output As #intFileHandle
Close #intFileHandle
'In case the sub can't access the file it goes on
'without raising an error
ErrorHandler:
    Close #intFileHandle
End Sub

'Create the game list file if it doesn't exist
'or clear it if it does.
Public Sub MakeListIfDoesntExist()
On Error GoTo ErrorHandler
MakeTempFoldIfDoesntExist
If Dir(m_strGameListPath, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = GAME_LIST_LOCAL_NAME Then SetAttr m_strGameListPath, vbNormal
Dim intFileHandle As Integer
intFileHandle = FreeFile
Open m_strGameListPath For Output As #intFileHandle
Close #intFileHandle
'In case the sub can't access the file it goes on
'without raising an error
ErrorHandler:
    Close #intFileHandle
End Sub

