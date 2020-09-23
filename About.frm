VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Kokoro's Downloader"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   615
   End
   Begin MSComctlLib.ImageList imlAbout 
      Left            =   720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   67
      ImageHeight     =   106
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "About.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.Label Label8 
         Caption         =   "All rights reserved"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Capsule Corp."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Copyright © 1999-2004"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "2004"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "capsule@argentina.com"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   1440
         X2              =   3720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         X1              =   1455
         X2              =   3735
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   2295
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2280
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Rosario, Argentina"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Coded by Emiliano Scavuzzo"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Kokoro's Downloader 1.3"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterOnMain Me
Picture1.Picture = imlAbout.ListImages(1).Picture
End Sub
