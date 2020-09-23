Attribute VB_Name = "Communication"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

'The functions in this module take care of the
'communication between frmMain and the kind of
'downloads (KODs) classes.

Option Explicit

'Build the class that communicate the graphic interface
'with any kind of download (KOD).
Public Function BuildDownloadClass() As Boolean
BuildDownloadClass = True
Select Case (KindOfDownload)
Case DOWN_BRIEFCASE
    Set m_clsDownload = New cDownload_Briefcase
Case Else
    BuildDownloadClass = False
    UpdateNeededForThisGame
End Select
End Function

'Start the timer that controls the speed and time.
Public Sub StartStateTimer()
frmMain.Timer_Speed.Enabled = True
End Sub


'Stop the timer that controls the speed and time.
Public Sub StopStateTimer()
frmMain.Timer_Speed.Enabled = False
End Sub

'Returns the speed timer status.
Public Function IsStateTimerActive() As Boolean
IsStateTimerActive = frmMain.Timer_Speed.Enabled
End Function

'Shadow a package on the list.
'NOTE: the argument is the index of the element on the
'list. The index for the first package is 0.
Public Sub ShadowPackage(ByVal intIndex As Integer)
frmMain.lstPackages.ListIndex = intIndex
End Sub

'Mark a package with a sign at least it was already
'marked.
'NOTE: the argument is the index of the element on the
'list. The index for the first package is 0.
Public Sub MarkPackage(ByVal intIndex As Integer)
With frmMain.lstPackages
    If Right(.List(intIndex), 1) <> "®" Then .List(intIndex) = .List(intIndex) + " ®"
End With
End Sub

'Clean the state data.
Public Sub CleanState()
With frmMain
    Speed.LastQuantity = 0
    Speed.QuanLoaded = 0
    .ProgBar.Value = 0
    .lblState.Caption = "State"
    .lblRate.Caption = "0 %"
    .lblBytesDown.Caption = "0 from 0 bytes"
    .lblElapsedTime.Caption = "Elapsed Time:              00:00:00"
    .lblEstimatedTime.Caption = "Estimated Time:           00:00:00"
    .lblSpeed.Caption = "Speed: 0 KBps"
End With
End Sub

'Perform the graphic actions to finish the download.
Public Sub Finish_Download_Actions()
StopStateTimer
CleanState
EnableForInfo
With frmMain
    .txtCode.Enabled = True
    .lstPackages.Enabled = True
    .cmdSelectAll.Enabled = True
    .cmdCancelDown.Enabled = False
    .cmdDownload.Enabled = True
    m_blnIsDownloading = False
End With
End Sub
