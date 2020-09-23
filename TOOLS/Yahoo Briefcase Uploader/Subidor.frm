VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Subidor 
   Caption         =   "Yahoo Briefcase Uploader"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "Subidor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmComandos 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   7095
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Export list"
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Import list"
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdEspacioLibre 
         Caption         =   "Used space"
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDirectorio 
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Text            =   "/My Documents"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuitarTodos 
         Caption         =   "Remove all"
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Remove"
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Add"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSubir 
         Caption         =   "Upload"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDirectorio 
         Alignment       =   1  'Right Justify
         Caption         =   "Folder:"
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblUsuario 
         Alignment       =   1  'Right Justify
         Caption         =   "User:"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Timer tmrTiempo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   4800
   End
   Begin VB.Timer tmrTranscurrido 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   4800
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   3960
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin MSComctlLib.ListView lvLista 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Subidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

Private WithEvents sckSubir As CSocketMaster
Attribute sckSubir.VB_VarHelpID = -1
Private WithEvents sckLogear As CSocketMaster
Attribute sckLogear.VB_VarHelpID = -1

Private Conexion As New classConexion
Private BotonDelMouse As Integer 'para controlar la listview

Private Sub cmdAgregar_Click()
If EstanLosDatos Then
    Dim Cont As Integer
    For Cont = 0 To File1.ListCount - 1
        If File1.Selected(Cont) Then AgregarALista File1.List(Cont), CorregirPath(File1.Path) + File1.List(Cont), txtUsuario.Text, txtPassword.Text, txtDirectorio.Text
    Next Cont
Else
    MsgBox "USER/PASSWORD/FOLDER error", vbCritical, "Error"
End If
End Sub


'devuelve verdadero si todos los datos necesarios para
'agregar un archivo a la lista están puestos
Private Function EstanLosDatos() As Boolean
EstanLosDatos = False
If Trim(txtUsuario.Text) <> "" And Trim(txtPassword.Text) <> "" And Trim(txtDirectorio.Text) <> "" Then EstanLosDatos = True
End Function

Private Sub cmdCancelar_Click()
AUTOMATICO = False
AbortarUpload
End Sub

Private Sub AbortarUpload()
Debug.Print "AbortUpload"
sckSubir.CloseSck
sckLogear.CloseSck
tmrTranscurrido.Enabled = False
Close #ArchivoSubId
HabilitaParaSubir
End Sub

Private Sub cmdEspacioLibre_Click()
Dim Usuario As String
Dim Password As String
Usuario = Trim(txtUsuario.Text)
Password = Trim(txtPassword.Text)
DeshabilitaParaSubir
If Usuario = "" Or Password = "" Then
    MsgBox "USER/PASSWORD error", vbCritical, "Error"
Else
    Logear sckLogear, Usuario, Password, Logeando, fnEspacioUsado
End If
End Sub

Private Sub cmdExportar_Click()
If MsgBox("If you export the list you will overwrite the current one." + vbCrLf + _
          "Export anyway?", vbQuestion + vbYesNo, "Export") = vbYes Then
    ExportaLista
End If
End Sub

Private Sub cmdImportar_Click()
ImportaLista
End Sub

Private Sub cmdQuitar_Click()
If Not lvLista.SelectedItem Is Nothing Then
    lvLista.ListItems.Remove (lvLista.SelectedItem.Index)
End If
If Not lvLista.SelectedItem Is Nothing Then
    lvLista.ListItems(lvLista.SelectedItem.Index).Selected = True
End If
End Sub

Private Sub cmdQuitarTodos_Click()
Dim Cont As Integer
For Cont = 1 To lvLista.ListItems.Count
    lvLista.ListItems.Remove (1)
Next Cont
End Sub

Private Sub cmdSubir_Click()
DeshabilitaParaSubir
If lvLista.ListItems.Count = 0 Then
    MsgBox "There are no items on the list!", vbInformation, "Error"
    HabilitaParaSubir
    Exit Sub
End If

AUTOMATICO = False
PrepararLista
SubirArchivo GetNextIndex(True), True
End Sub

Private Sub SubirArchivo(ByVal Indice As Integer, ByVal ElPrimero As Boolean)
If Indice <> 0 Then 'quedan archivos por subir
    IndiceSubiendo = Indice
    SetDato IndiceSubiendo, keyEstado, "Uploading"
    SetDato IndiceSubiendo, keyIntento, str(Val(GetDato(IndiceSubiendo, keyIntento)) + 1)
    SetDato IndiceSubiendo, keyTiempo, "00:00:00"
    
    sBoundary = CreaBoundary
    
    ArchivoSubId = FreeFile
    Open GetDato(IndiceSubiendo, keyPath) For Random As #ArchivoSubId Len = BUFFSIZE
    
    Dim UsuarioAnt As String
    Dim PasswordAnt As String
    UsuarioAnt = UsuarioAct
    PasswordAnt = PasswordAct
    
    UsuarioAct = GetDato(IndiceSubiendo, keyUsuario)
    PasswordAct = GetDato(IndiceSubiendo, keyPassword)
    
    If ElPrimero Or UsuarioAct <> UsuarioAnt Then
        Velocidad.UltimaCantidad = 0
        Velocidad.CantCargadas = 0
        Logear sckLogear, UsuarioAct, PasswordAct, FrameDeSubir, fnSubir
    Else
        sckSubir.Connect sHost, 80
    End If
Else 'terminó de subir todos los archivos
    AbortarUpload
    If AUTOMATICO Then
        MostrarMensaje Desconectando
        If ApretadoCancelar = False Then
            Conexion.Desconectar
        Else
            End
        End If
        MostrarMensaje Apagando
        If ApretadoCancelar = False Then
            ApagarComputadora
        Else
            End
        End If
    End If
End If
End Sub

'resetea los datos de la lista que pudiesen haber cambiado
Private Sub PrepararLista()
Dim Cont As Integer
For Cont = 1 To lvLista.ListItems.Count
    SetDato Cont, keyEstado, "WAIT"
    SetDato Cont, keyIntento, "0"
    SetDato Cont, keyTiempo, "00:00:00"
    SetDato Cont, keyEnviado, "0"
    SetDato Cont, keyTamTotal, "0"
Next Cont
End Sub

Private Sub Dir1_Change()
File1.FileName = Dir1.Path
End Sub

Private Sub File1_DblClick()
If EstanLosDatos Then
    AgregarALista File1.FileName, CorregirPath(File1.Path) + File1.FileName, txtUsuario.Text, txtPassword.Text, txtDirectorio.Text
Else
    MsgBox "USER/PASSWORD/FOLDER error", vbCritical, "Error"
End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If EstanLosDatos Then
        Dim Cont As Integer
        For Cont = 0 To File1.ListCount - 1
            If File1.Selected(Cont) Then AgregarALista File1.List(Cont), CorregirPath(File1.Path) + File1.List(Cont), txtUsuario.Text, txtPassword.Text, txtDirectorio.Text
        Next Cont
    Else
        MsgBox "USER/PASSWORD/FOLDER error", vbCritical, "Error"
    End If
End If
End Sub

Private Sub Form_Load()
Set sckSubir = New CSocketMaster
Set sckLogear = New CSocketMaster

Inicializar
If Trim(Command) = "/a" Then
    Dim Hora As Date
    Hora = Now
    If TimeValue(Format(Hora, "HH:MM:SS")) >= TimeValue("4:00:00") Then
        If TimeValue(Format(Hora, "HH:MM:SS")) <= TimeValue("5:15:00") Then
            MostrarMensaje ArranqueAutomatico
            If ApretadoCancelar = False Then
                AUTOMATICO = True
                ImportaLista
                DeshabilitaParaSubir
                If GetNextIndex(True) <> 0 Then
                    tmrTiempo.Enabled = True
                    If Conexion.Conectar Then 'si se conectó
                        SubirArchivo GetNextIndex(True), True
                    Else 'si no se puede conectar apaga la PC
                        ApagarComputadora
                    End If
                Else
                    ApagarComputadora
                End If
            Else
                End
            End If
        Else
            End
        End If
    Else
        End
    End If
End If
End Sub

'Si se cierra el programa porque se está apagando el sistema
'crea el log. No funciona si el apagado es forzoso.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 2 Then CreaLog
End Sub

Private Sub Form_Resize()
Dim LargoUsable As Long 'el largo de la ventana menos los espacios entre las Dir1, File1 y los bordes
Dim EspEntDirCom As Long
LargoUsable = Me.Width - EspFijosHor
EspEntDirCom = frmComandos.Top - (Dir1.Top + Dir1.Height)

Dir1.Width = LargoUsable * DirPorLargo / 100
File1.Left = Dir1.Left + Dir1.Width + EspEntreDirFile
File1.Width = LargoUsable * FilePorLargo / 100

lvLista.Width = Me.Width - MargenesLista

Dim AltoUsable As Long
AltoUsable = Me.Height - EspFijosVer
If AltoUsable < 0 Then AltoUsable = 0
Dir1.Height = AltoUsable * DirPorAlto / 100
File1.Height = AltoUsable * DirPorAlto / 100
lvLista.Top = Dir1.Top + Dir1.Height + EspEntreDirLista
lvLista.Height = AltoUsable * ListaPorAlto / 100

frmComandos.Top = Dir1.Top + Dir1.Height + EspEntDirCom
frmComandos.Left = (Me.Width - frmComandos.Width) / 2
End Sub


'si se clickea fuera de los elementos deselecciona
'el elemento que estaba seleccionado
Private Sub lvLista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lvLista.SelectedItem = Nothing
End Sub


'devuelve el próximo index de archivo a bajar de acuerdo
'si es el primero y si está en modo automático
'si no quedan archivos devuelve 0
Private Function GetNextIndex(ElPrimero As Boolean) As Integer
Dim CantLista As Integer
CantLista = lvLista.ListItems.Count
If CantLista = 0 Then
    GetNextIndex = 0
    Exit Function
End If

If AUTOMATICO Then
    Dim Desde As Integer
    Dim Indice As Integer
    Desde = IndiceSubiendo + 1
    Indice = 0
    If ElPrimero Then Desde = 1
    Do Until (Desde > CantLista) Or (Indice <> 0)
        If GetDato(Desde, keyEstado) <> "DONE" Then Indice = Desde
        Desde = Desde + 1
    Loop
    GetNextIndex = Indice
Else

    If ElPrimero Then
        GetNextIndex = 1
    Else
        If CantLista > IndiceSubiendo Then
            GetNextIndex = IndiceSubiendo + 1
        Else
            GetNextIndex = 0
        End If
    End If
    
End If

End Function

Private Sub sckLogear_Connect()
sRespuesta = ""
sckLogear.SendData sComandoMandar
End Sub

Private Sub sckLogear_DataArrival(ByVal bytesTotal As Long)
Dim Chunk As String
sckLogear.GetData Chunk, vbString
sRespuesta = sRespuesta + Chunk
End Sub

Private Sub sckLogear_CloseSck()
sckLogear.CloseSck
Select Case leEstado

Case Logeando
    If GetLocationDeRespuesta = "" Then
        If leHastaDonde > leEstado Then
            leEstado = leEstado + 1
            CreaComandoMandar
            sckLogear.Connect sHost, 80
        Else
            TerminadoLogeo
        End If
    Else
        CreaComandoMandar
        sckLogear.Connect sHost, 80
    End If

Case LlendoASubir
    If leHastaDonde > leEstado Then
        leEstado = leEstado + 1
        CreaComandoMandar
        sckLogear.Connect sHost, 80
    Else
        TerminadoLogeo
    End If
Case FrameDeSubir
    TerminadoLogeo
End Select
End Sub

Private Sub sckLogear_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sckLogear.CloseSck
Debug.Print "sckLogear Error"
'insertar que pasa si hay error al logear
End Sub

Private Sub sckSubir_Connect()
Debug.Print "sckSubir Connect"
EstadoEnvio = MandandoParteAArchivo
Buffer = ""
Buffer2 = ""

SetDato IndiceSubiendo, keyTamTotal, Trim(str(Len(CreaParteAArchivo) + Len(CreaParteDArchivo) + Len(CreaCabecera) + LOF(ArchivoSubId)))

sckSubir.SendData CreaCabecera + CreaParteAArchivo
End Sub

Private Sub sckSubir_SendComplete()

If EstadoEnvio = MandandoParteAArchivo Then
    tmrTranscurrido.Enabled = True
    SetDato IndiceSubiendo, keyEnviado, Trim(str(Len(CreaCabecera) + Len(CreaParteAArchivo)))
    EstadoEnvio = MandandoArchivo
End If

Select Case EstadoEnvio
Case MandandoArchivo
    Get #ArchivoSubId, , Buffer
    If ((Seek(ArchivoSubId) - 1) * BUFFSIZE) > LOF(ArchivoSubId) Then
        Buffer2 = Left(Buffer, Len(Buffer) - (((Seek(ArchivoSubId) - 1) * BUFFSIZE) - LOF(ArchivoSubId)))
        sckSubir.SendData Buffer2
    Else
        SetDato IndiceSubiendo, keyEnviado, Trim(str(Val(GetDato(IndiceSubiendo, keyEnviado)) + Len(Buffer)))
        sckSubir.SendData Buffer
    End If
    If EOF(ArchivoSubId) Or ((Seek(ArchivoSubId) - 1) * BUFFSIZE) = LOF(ArchivoSubId) Then EstadoEnvio = MandandoParteDArchivo

Case MandandoParteDArchivo
    Debug.Print "Sending ParteDArchivo"
    SetDato IndiceSubiendo, keyEnviado, Trim(str(Val(GetDato(IndiceSubiendo, keyEnviado)) + Len(Buffer2)))
    sckSubir.SendData CreaParteDArchivo
    EstadoEnvio = Terminando

Case Terminando
    Debug.Print "Finishing"
    SetDato IndiceSubiendo, keyEnviado, Trim(str(Val(GetDato(IndiceSubiendo, keyEnviado)) + Len(CreaParteDArchivo)))
    sckSubir.CloseSck
    Close #ArchivoSubId
    tmrTranscurrido.Enabled = False
    SetDato IndiceSubiendo, keyEstado, "DONE"
    CreaLog
    If AUTOMATICO Then ExportaLista
    SubirArchivo GetNextIndex(False), False

End Select
End Sub

Private Sub sckSubir_CloseSck()
Debug.Print "sckSubir Close"
sckSubir.CloseSck
ManejaErrorAlSubir
End Sub

Private Sub sckSubir_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print "sckSubir Error"
sckSubir.CloseSck
ManejaErrorAlSubir
End Sub

'controla que pasa si ocurre un error al intentar subir
Private Sub ManejaErrorAlSubir()
Debug.Print "ManejaErrorAlSubir"
Close #ArchivoSubId
CreaLog
tmrTranscurrido.Enabled = False
If AUTOMATICO And Conexion.EstadoConexion = DESCONECTADO Then Conexion.Conectar True
SubirArchivo IndiceSubiendo, False
End Sub

Private Sub sckSubir_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Debug.Print "BYTES SENT: " & bytesSent & " BYTES REMAINING: " & bytesRemaining
End Sub

'timer que controla la hora a la que se tiene que
'desconectar cuando está en modo automático
Private Sub tmrTiempo_Timer()
If Hour(Now) = 7 And Minute(Now) >= 50 Then
    tmrTiempo.Enabled = False
    AbortarUpload
    CreaLog
    MostrarMensaje Desconectando
    If ApretadoCancelar = False Then
        Conexion.DesconectarTodo
    Else
        End
    End If
    MostrarMensaje Apagando
    If ApretadoCancelar = False Then
        ApagarComputadora
    Else
        End
    End If
End If
End Sub

Private Sub tmrTranscurrido_Timer()
'guarda el tiempo en una variable de tipo DATE
Transcurrido = GetDato(IndiceSubiendo, keyTiempo)
'le suma un segundo y lo muestra en pantalla
Transcurrido = DateAdd("s", 1, Transcurrido)
SetDato IndiceSubiendo, keyTiempo, Format(Transcurrido, "hh:mm:ss")
'controla que no vaya muy lento
If AUTOMATICO Then ControlaLentitud

'velocidad
If Velocidad.CantCargadas < UBound(Velocidad.Cantidad) Then
    Velocidad.CantCargadas = Velocidad.CantCargadas + 1
Else
    Dim Contador1 As Byte
    For Contador1 = 1 To UBound(Velocidad.Cantidad) - 1
        Velocidad.Cantidad(Contador1) = Velocidad.Cantidad(Contador1 + 1)
    Next Contador1
End If
Velocidad.Cantidad(Velocidad.CantCargadas) = CantidadSubida - Velocidad.UltimaCantidad
Velocidad.UltimaCantidad = CantidadSubida

Dim Transferido As Long 'suma de bytes bajados en los últimos n segundos (n = Velocidad.CantCargadas)
Dim Contador2 As Byte
For Contador2 = 1 To Velocidad.CantCargadas
    Transferido = Transferido + Velocidad.Cantidad(Contador2)
Next
SetDato IndiceSubiendo, keyVelocidad, Format(Transferido / 1024 / Velocidad.CantCargadas, "0.00") + " KBps"
End Sub

'devuelve los bytes subidos hasta el momento pero en
'forma de número
Private Function CantidadSubida() As Double
CantidadSubida = Val(GetDato(IndiceSubiendo, keyEnviado))
End Function

'Controla que no vaya muy lento cada 5 min y se reconecta
'con otra cuenta si pasa
Private Sub ControlaLentitud()
Dim lTiempo As Long
'pasa a segundos
lTiempo = Second(Transcurrido) + (60 * Minute(Transcurrido)) + (3600 * Hour(Transcurrido))
If lTiempo <> 0 And lTiempo Mod (LENT_FRECUENCIA) = 0 Then 'cada 5 min
    If (CantidadSubida / lTiempo) < (5000000 / (LENT_MAX_TIEMPO)) Then
        Debug.Print "Transfer too slow"
        tmrTranscurrido.Enabled = False
        sckSubir.CloseSck
        Close #ArchivoSubId
        CreaLog
        Conexion.Desconectar
        Conexion.ConectarConOtroISP
        SetDato IndiceSubiendo, keyTiempo, "00:00:00"
        SubirArchivo IndiceSubiendo, True
    End If
End If
End Sub

'función que se llama cuando se termina el logeo
Private Sub TerminadoLogeo()
Select Case fnFinalidad
    Case fnSubir
        ObtenerCrumb
        sckSubir.Connect sHost, 80
    Case fnEspacioUsado
        EspacioUsado
End Select
End Sub
