VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Untitled - Criptonita"
   ClientHeight    =   3795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComDialog1 
      Left            =   3480
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      HideSelection   =   0   'False
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   3210
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&File"
      Begin VB.Menu mnuNuevo 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuGuardarComo 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigPag 
         Caption         =   "&Page Se&tup..."
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&E&xit"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edit"
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "&De&lete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSeparador4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeleccionar 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "&Word Wrap"
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Search"
      Begin VB.Menu mnuBuscar_ 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuBuscSig 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_UNDO = &H304
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const CF_TEXT = 1
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Private Const sNombreProg As String = "Criptonita"
Private bModificado As Boolean
Private bGuardado As Boolean
Private sTitulo As String
Private sPath As String
Private ptTextBox As TextBox

Private Const MULT As Integer = 55 'hace que cambie la encriptación

Public Function BYTES_TO_STRING(bBytes() As Byte) As String
    BYTES_TO_STRING = bBytes
    BYTES_TO_STRING = StrConv(BYTES_TO_STRING, vbUnicode)
End Function

'Functions to convert between strings and byte arrays
Public Function STRING_TO_BYTES(sString As String) As Byte()
    STRING_TO_BYTES = StrConv(sString, vbFromUnicode)
End Function

Private Function Encriptar(ByRef Texto As String, Optional ByVal Indice As Variant) As String
Dim TextoArray() As Byte
If IsMissing(Indice) Then Indice = 0
TextoArray() = STRING_TO_BYTES(Texto)

Dim Contador As Long
For Contador = 0 To UBound(TextoArray())
    TextoArray(Contador) = (TextoArray(Contador) + (Contador + Indice + 1) * MULT) Mod 256
Next

Encriptar = BYTES_TO_STRING(TextoArray())
End Function

Private Function Desencriptar(ByRef Texto As String, Optional ByVal Indice As Variant) As String
Dim Contador As Long
Dim TextoArray() As Byte
If IsMissing(Indice) Then Indice = 0

TextoArray() = STRING_TO_BYTES(Texto)

For Contador = 0 To UBound(TextoArray())
    TextoArray(Contador) = (((TextoArray(Contador) - (Contador + Indice + 1) * MULT) Mod 256) + 256) Mod 256
Next

Desencriptar = BYTES_TO_STRING(TextoArray())
End Function

Private Function LeeArchivo(ByVal Path As String) As String
Dim HArchivo As Long
HArchivo = FreeFile
Open Path For Binary As #HArchivo
LeeArchivo = Space(LOF(HArchivo))
Get #HArchivo, , LeeArchivo
Close #HArchivo
End Function

Private Sub GuardaArchivo(ByVal Path As String, ByRef Texto As String)
Dim HArchivo As Long
HArchivo = FreeFile
Open Path For Output As #HArchivo
Close #HArchivo
HArchivo = FreeFile
Open Path For Binary As #HArchivo
Put HArchivo, , Texto
Close HArchivo
End Sub

Private Sub Form_Load()
Set ptTextBox = Text1
Text1.Left = 0
Text2.Left = 0
Text1.Top = 0
Text2.Top = 0
sTitulo = "Untitled"
If Command <> "" Then
    sTitulo = Dir(Command)
    If sTitulo <> "" Then
        sPath = Command
        ptTextBox.Text = Desencriptar(LeeArchivo(sPath))
        Me.Caption = sTitulo + " - Criptonita"
        bGuardado = True
        bModificado = False
    Else
        sTitulo = "Untitled"
        MsgBox "File not found", vbCritical, "Error"
    End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Select Case SalidaPedida
    Case 0
        Cancel = True
    Case 2
        If bGuardado Then
            GuardaArchivo sPath, Encriptar(ptTextBox.Text)
        Else
            MuestraDiaGuardar
        End If
End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Or Me.WindowState = 2 Then
ptTextBox.Height = Me.Height - 675
ptTextBox.Width = Me.Width - 120
End If
End Sub

Private Sub mnuAbrir_Click()
Select Case SalidaPedida
    Case 1
            MuestraDiaAbrir
    Case 2
        If bGuardado Then
            GuardaArchivo sPath, Encriptar(ptTextBox.Text)
            bModificado = False
            MuestraDiaAbrir
        Else
            If MuestraDiaGuardar Then MuestraDiaAbrir
        End If
End Select
End Sub

Private Sub mnuAjuste_Click()
If mnuAjuste.Checked Then
    Set ptTextBox = Text1
    Text1 = Text2
    Text2 = ""
    Text1.Height = Text2.Height
    Text1.Width = Text2.Width
    Text1.Visible = True
    Text2.Visible = False
    mnuAjuste.Checked = False
Else
    Set ptTextBox = Text2
    Text2 = Text1
    Text1 = ""
    Text2.Height = Text1.Height
    Text2.Width = Text1.Width
    Text2.Visible = True
    Text1.Visible = False
    mnuAjuste.Checked = True
End If
ptTextBox.SelStart = 0
ptTextBox.SelLength = 0
End Sub

Private Sub mnuBuscar__Click()
Form2.Show vbModeless, Me
End Sub

Private Sub mnuBuscSig_Click()
If Len(strBus) = 0 Then
    Form2.Show vbModeless, Me
Else
    BuscarSiguiente
End If
End Sub

Private Sub mnuConfigPag_Click()
ComDialog1.ShowPrinter
End Sub

Private Sub mnuEdicion_Click()
If ptTextBox.SelLength > 0 Then
    mnuCortar.Enabled = True
    mnuCopiar.Enabled = True
    mnuEliminar.Enabled = True
Else
    mnuCortar.Enabled = False
    mnuCopiar.Enabled = False
    mnuEliminar.Enabled = False
End If
If IsClipboardFormatAvailable(CF_TEXT) Then
    mnuPegar.Enabled = True
Else
    mnuPegar.Enabled = False
End If
End Sub

Private Sub mnuCopiar_Click()
SendMessage ptTextBox.hwnd, WM_COPY, 0&, 0&
End Sub

Private Sub mnuCortar_Click()
SendMessage ptTextBox.hwnd, WM_CUT, 0&, 0&
End Sub

Private Sub mnuDeshacer_Click()
SendMessage ptTextBox.hwnd, WM_UNDO, 0&, 0&
End Sub

Private Sub mnuEliminar_Click()
SendMessage ptTextBox.hwnd, WM_CLEAR, 0&, 0&
End Sub

Private Sub mnuGuardar_Click()
If bGuardado Then
    GuardaArchivo sPath, Encriptar(ptTextBox.Text)
    bModificado = False
Else
    MuestraDiaGuardar
End If
End Sub

Private Sub mnuGuardarComo_Click()
MuestraDiaGuardar
End Sub

Private Function MuestraDiaGuardar() As Byte
On Error GoTo BotonCancelar
ComDialog1.Filter = "Info Documents (*.fgi)|*.fgi|List Documents (*.fgl)|*.fgl|All Files (*.*)|*.*"
ComDialog1.FileName = sTitulo
ComDialog1.DefaultExt = "fgi"
ComDialog1.Flags = 2 Or &H800
ComDialog1.CancelError = True
ComDialog1.ShowSave
sPath = ComDialog1.FileName
sTitulo = ComDialog1.FileTitle
bGuardado = True
bModificado = False
GuardaArchivo sPath, Encriptar(ptTextBox.Text)
Me.Caption = sTitulo + " - Criptonita"
MuestraDiaGuardar = 1
Exit Function
BotonCancelar:
    MuestraDiaGuardar = 0
End Function

Private Function MuestraDiaAbrir() As Byte
On Error GoTo BotonCancelar
ComDialog1.Filter = "Info Documents (*.fgi)|*.fgi|List Documents (*.fgl)|*.fgl|All Files (*.*)|*.*"
ComDialog1.FileName = ""
ComDialog1.DefaultExt = "fgi"
ComDialog1.Flags = &H1000 Or &H800
ComDialog1.CancelError = True
ComDialog1.ShowOpen
sPath = ComDialog1.FileName
sTitulo = ComDialog1.FileTitle
ptTextBox.Text = Desencriptar(LeeArchivo(sPath))
Me.Caption = sTitulo + " - Criptonita"
bGuardado = True
bModificado = False
MuestraDiaAbrir = 1
Exit Function
BotonCancelar:
    MuestraDiaAbrir = 0
End Function

Private Sub mnuImprimir_Click()
On Error GoTo Error_Handler

Printer.Print ptTextBox

Exit Sub
Error_Handler:
    MsgBox "Error printing", vbCritical, "Error"
End Sub

Private Sub mnuNuevo_Click()
Select Case SalidaPedida
    Case 1
        CreaNuevo
    Case 2
        If bGuardado Then
            GuardaArchivo sPath, Encriptar(ptTextBox.Text)
            CreaNuevo
        Else
             If MuestraDiaGuardar Then CreaNuevo
        End If
End Select
End Sub

Private Sub CreaNuevo()
ptTextBox.Text = ""
sTitulo = "Untitled"
sPath = ""
bGuardado = False
bModificado = False
Me.Caption = sTitulo + " - " + sNombreProg
mnuDeshacer.Enabled = False
End Sub

Private Sub mnuPegar_Click()
SendMessage ptTextBox.hwnd, WM_PASTE, 0&, 0&
End Sub

Private Sub mnuSalir_Click()
Unload Me
End Sub

'devuelve 0 si apreta cancelar, 1 si apreta no
'o el archivo ya está guardado, y 2 si apreta si
Private Function SalidaPedida() As Byte
If bModificado And ptTextBox.Text <> "" Then
    Dim Resul As Integer
    Dim sMensaje As String
    sMensaje = "The text in the "
    If bGuardado Then
        sMensaje = sMensaje + sPath
    Else
        sMensaje = sMensaje + sTitulo
    End If
    sMensaje = sMensaje + " file has canged." + vbCrLf + vbCrLf + "Do you want to save the changes?"
    Resul = MsgBox(sMensaje, vbYesNoCancel + vbExclamation, sNombreProg)
    Select Case Resul
        Case vbYes
            SalidaPedida = 2
        Case vbNo
            SalidaPedida = 1
        Case vbCancel
            SalidaPedida = 0
     End Select
Else
    SalidaPedida = 1
End If
End Function

Private Sub mnuSeleccionar_Click()
ptTextBox.SelStart = 0
ptTextBox.SelLength = Len(ptTextBox.Text)
End Sub

Private Sub Text1_Change()
bModificado = True
mnuDeshacer.Enabled = True
End Sub

Private Sub Text2_Change()
bModificado = True
mnuDeshacer.Enabled = True
End Sub
