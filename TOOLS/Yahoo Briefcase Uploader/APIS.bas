Attribute VB_Name = "APIS"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************
Option Explicit

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const EWX_SHUTDOWN = 1
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8
