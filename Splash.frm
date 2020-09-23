VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplash 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      MousePointer    =   99  'Custom
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      Top             =   0
      Width           =   4035
      Begin VB.Timer tmrPause2 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   3600
         Top             =   3720
      End
      Begin VB.Timer tmrPause1 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   3120
         Top             =   3720
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   960
         Picture         =   "Splash.frx":682D
         Top             =   960
         Width           =   2250
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ver 1.3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ver 1.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1815
         TabIndex        =   2
         Top             =   3255
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSplash"
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
    'if the program is already running it closes this instance
    If App.PrevInstance = True Then End
    
    Dim WindowRegion As Long
        
    Me.Width = picSplash.Width
    Me.Height = picSplash.Height
    
    WindowRegion = MakeRegion(picSplash)
    SetWindowRgn Me.hwnd, WindowRegion, True
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    DoEvents
    tmrPause1.Enabled = True
    End Sub

Private Sub tmrPause1_Timer()
tmrPause1.Enabled = False
frmMain.Show
Me.SetFocus
tmrPause2.Enabled = True
End Sub

Private Sub tmrPause2_Timer()
tmrPause2.Enabled = False
Unload Me
End Sub
