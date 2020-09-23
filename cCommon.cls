VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

'The functions in this module take care of the
'communication between frmMain and the kind of
'downloads (KODs) classes.

Option Explicit

'Call the function that checks the info for wrong data.
'If it finds erroneous data it must raise an error.
Public Sub CheckPackages()
'
End Sub

'Call the function in charge of filling the package list.
Public Sub ReadPackages()
'
End Sub

'Call the function in charge of downloading the packages.
Public Sub Download()
'
End Sub

'Call the function in charge of stopping the download.
Public Sub StopDownload()
'
End Sub

'This function returns the size of the package being
'downloaded.
Public Function PackageSize() As Double
'
End Function

'This functions returns the amount of downloaded bytes of
'the package being downloaded.
Public Function AmountDownloaded() As Double
'
End Function
