VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   HelpContextID   =   50
   Icon            =   "Preferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmPref1 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtPort 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   5
         TabIndex        =   2
         Top             =   440
         Width           =   735
      End
      Begin VB.TextBox txtProxy 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   440
         Width           =   2055
      End
      Begin VB.CheckBox chkUseProxy 
         Caption         =   "Use Proxy"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "This option is used only when downloading from the web."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   780
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Port:"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   465
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Proxy:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   470
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   5415
      Begin VB.CheckBox chkUseOwnServer 
         Caption         =   "Use own server"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   160
         Width           =   1455
      End
      Begin VB.TextBox txtOwnServer 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   440
         Width           =   3855
      End
      Begin VB.Label Label9 
         Caption         =   "Server:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   465
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Use own server instead of default servers."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   795
         Width           =   3015
      End
   End
   Begin VB.Frame frmPref4 
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   5415
      Begin VB.CheckBox chkDontUseBackup 
         Caption         =   "Disable backup server"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   160
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Disable backup check server in case that the communication with the main server fails."
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   550
         Width           =   4935
      End
   End
   Begin VB.Frame frmPref3 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   5415
      Begin VB.CheckBox chkDontUseCol 
         Caption         =   "Disable colors on list"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   160
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Turn colors and flashes off for the game list."
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame frmPref2 
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   5415
      Begin VB.CheckBox chkDontDownPrev 
         Caption         =   "Don't download image"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Check this option if you don't wish to download a preview image for the game. This makes the check faster."
         Height          =   450
         Left            =   240
         TabIndex        =   13
         Top             =   435
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmPreferences"
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
Show_Preferences
End Sub

'change the color of the textbox
Private Sub chkUseOwnServer_Click()
If chkUseOwnServer.Value = 0 Then
    txtOwnServer.BackColor = &HC0C0C0
    txtOwnServer.Enabled = False
Else
    txtOwnServer.BackColor = &H80000005
    txtOwnServer.Enabled = True
End If
End Sub

'change the color of the text boxes
Private Sub chkUseProxy_Click()
If chkUseProxy.Value = 0 Then
    txtProxy.BackColor = &HC0C0C0
    txtPort.BackColor = &HC0C0C0
    txtProxy.Enabled = False
    txtPort.Enabled = False
Else
    txtProxy.BackColor = &H80000005
    txtPort.BackColor = &H80000005
    txtProxy.Enabled = True
    txtPort.Enabled = True
End If
End Sub

Private Sub cmdOK_Click()
'check for erros
If (chkUseProxy.Value = 1) Then
    If (Trim(txtProxy.Text) = "") Then
    MessageBox hwnd, "You must insert the proxy to use.", "Error", MB_ICONEXCLAMATION
    Exit Sub
    
    ElseIf (Trim(txtPort.Text) = "") Then
    MessageBox hwnd, "You must insert the port to use.", "Error", MB_ICONEXCLAMATION
    Exit Sub

    ElseIf (IsNumeric(txtPort.Text) = False) Or (Val(txtPort.Text) < 0) Or (Val(txtPort.Text) > 65535) Then
    MessageBox hwnd, "The proxy port doesn't seem correct.", "Error", MB_ICONEXCLAMATION
    Exit Sub
    
    End If
End If

If (chkUseOwnServer.Value = 1) Then
    If (Trim(txtOwnServer.Text) = "") Then
    MessageBox hwnd, "You must insert the own server to use.", "Error", MB_ICONEXCLAMATION
    Exit Sub
    End If
End If
'save data
Save_Preferences
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

'save preferences
Private Sub Save_Preferences()

Const SOFTWARE As String = "Software\"

If (chkUseProxy.Value = 1) Then
    m_strProxy = Trim(txtProxy.Text)
    m_lngPort = Abs(Int(Val(txtPort.Text)))
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUseProxy, 1)
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regProxy, m_strProxy)
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regPort, m_lngPort)
    m_blnUseProxy = True
Else
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUseProxy, 0)
    m_blnUseProxy = False
End If

If (chkUseOwnServer.Value = 1) Then
    m_strOwnServer = Trim(txtOwnServer.Text)
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUseOwnServer, 1)
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regOwnServer, m_strOwnServer)
    m_blnUseOwnServer = True
Else
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regUseOwnServer, 0)
    m_blnUseOwnServer = False
End If

If (chkDontDownPrev.Value = 1) Then
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoPreview, 1)
    m_blnDontDownPrev = True
Else
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoPreview, 0)
    m_blnDontDownPrev = False
End If

If (chkDontUseCol.Value = 1) Then
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoColors, 1)
    'if it was using colors, it removes them
    If m_blnDontUseColors = False Then HideListColors
    m_blnDontUseColors = True
Else
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoColors, 0)
    'if it wasn't using colors, it shows them
    If m_blnDontUseColors = True Then ShowListColors
    m_blnDontUseColors = False
End If

If (chkDontUseBackup.Value = 1) Then
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoBackup, 1)
    m_blnDontUseBackup = True
Else
    Call SaveDword(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regNoBackup, 0)
    m_blnDontUseBackup = False
End If
End Sub

'show preferences
Private Sub Show_Preferences()
If m_blnUseProxy Then
    chkUseProxy.Value = 1
    txtPort.Text = Format(Str(m_lngPort), "0")
Else
    txtPort.Text = Format(Str(m_lngPort), "#")
End If
txtProxy.Text = m_strProxy

If m_blnUseOwnServer Then
    chkUseOwnServer.Value = 1
End If
txtOwnServer.Text = m_strOwnServer

If m_blnDontDownPrev Then chkDontDownPrev.Value = 1
If m_blnDontUseColors Then chkDontUseCol.Value = 1
If m_blnDontUseBackup Then chkDontUseBackup.Value = 1
End Sub
