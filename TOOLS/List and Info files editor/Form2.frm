VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4900
      TabIndex        =   7
      Top             =   550
      Width           =   1420
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Search &Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   4900
      TabIndex        =   6
      Top             =   120
      Width           =   1420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   690
      Left            =   3040
      TabIndex        =   3
      Top             =   630
      Width           =   1740
      Begin VB.OptionButton opAbajo 
         Caption         =   "&Down"
         Height          =   255
         Left            =   870
         TabIndex        =   5
         Top             =   295
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opArriba 
         Caption         =   "&Up"
         Height          =   295
         Left            =   110
         TabIndex        =   4
         Top             =   295
         Width           =   745
      End
   End
   Begin VB.CheckBox chMayus 
      Caption         =   "Coincidir &mayúsculas y minúsculas"
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   1020
      Width           =   2960
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   295
      HideSelection   =   0   'False
      Left            =   1070
      TabIndex        =   1
      Top             =   160
      Width           =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "&Search:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   195
      Width           =   925
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private Sub cmdBuscar_Click()
strBus = txtBusqueda
If chMayus.Value Then
    mMetodo = vbBinaryCompare
Else
    mMetodo = vbTextCompare
End If
bArriba = opArriba.Value
BuscarSiguiente
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Form1.Left + (Screen.TwipsPerPixelX * 45)
Me.Top = Form1.Top + (Screen.TwipsPerPixelY * 119)
txtBusqueda = strBus
If mMetodo = 0 Then
    chMayus.Value = 0
Else
    chMayus.Value = 1
End If
If bArriba Then
    opArriba.Value = True
    opAbajo.Value = False
Else
    opArriba.Value = False
    opAbajo.Value = True
End If
txtBusqueda.SelStart = 0
txtBusqueda.SelLength = Len(txtBusqueda)
End Sub

Private Sub txtBusqueda_Change()
If Len(txtBusqueda) = 0 Then
    cmdBuscar.Enabled = False
Else
    cmdBuscar.Enabled = True
End If
End Sub
