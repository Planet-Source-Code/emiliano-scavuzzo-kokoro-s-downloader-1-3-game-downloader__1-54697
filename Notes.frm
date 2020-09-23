VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notes"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "Notes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNotes 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Notes.frx":08CA
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

Private Sub Form_Load()
CenterOnMain Me
Load_Notes
End Sub

Private Sub Load_Notes()
Me.Caption = NOTES_CAPTION + (Trim(ReadLine(m_strInfo, LN_NAME)))
txtNotes.Text = Decode(Trim(ReadLine(m_strInfo, LN_NOTES)))
End Sub
