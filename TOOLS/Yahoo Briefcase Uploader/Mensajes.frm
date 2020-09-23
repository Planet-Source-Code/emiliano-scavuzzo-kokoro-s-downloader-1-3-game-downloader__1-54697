VERSION 5.00
Begin VB.Form Mensajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subidor"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "Mensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCuenta 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblTiempo 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private Sub cmdCancelar_Click()
ApretadoCancelar = True
Unload Me
End Sub

Private Sub Form_Load()
ApretadoCancelar = False

Select Case TipoMensaje
    Case ArranqueAutomatico
        Me.Caption = "Initiating uploader"
        lblMensaje.Caption = "Initiating uploader. Press any button to abort."
        lblTiempo.Caption = "30"
    Case Desconectando
        Me.Caption = "Disconectiong from Internet"
        lblMensaje.Caption = "Disconectiong from Internet. Press Cancel to abort."
        lblTiempo.Caption = "30"
    Case Apagando
        Me.Caption = "Shutting down"
        lblMensaje.Caption = "Shutting down. Press Cancel to abort."
        lblTiempo.Caption = "30"
End Select

tmrCuenta.Enabled = True
End Sub

Private Sub tmrCuenta_Timer()
If lblTiempo.Caption = "00" Then Unload Me
lblTiempo.Caption = Format((Val(lblTiempo.Caption) - 1), "00")
End Sub
