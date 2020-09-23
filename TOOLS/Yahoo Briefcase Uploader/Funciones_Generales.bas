Attribute VB_Name = "FuncionesGenerales"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit

'devuelve la cantidad de veces que aparece la string Corta
'en la string Larga
Public Function CuantasVecesEsta(ByVal Corta As String, ByVal Larga As String) As Integer
CuantasVecesEsta = 0
If Corta = "" Then Exit Function

Dim Inicio As Integer
Dim Indice As Integer
Inicio = 1
Indice = 1
Do
Inicio = Indice
CuantasVecesEsta = CuantasVecesEsta + 1
Inicio = InStr(Inicio, Larga, Corta, vbBinaryCompare)
Indice = Inicio + Len(Corta)
Loop Until (Inicio = 0 Or Inicio = Null)

CuantasVecesEsta = CuantasVecesEsta - 1
End Function

'esta función lee una línea con datos separados por comas
'y devuelve el dato nro nDato
'si detecta un error devuelve un error 1
Public Function DatoDeLinea(ByVal Linea As String, nDato As Integer) As String
If nDato < 1 Then Err.Raise 1
Dim NroDeComas As Integer
NroDeComas = CuantasVecesEsta(",", Linea)
If (nDato - 1) > NroDeComas Then Err.Raise 1

Dim Inicio As Long
Dim Fin As Long
Inicio = 1
If nDato = 1 Then 'si es el primer dato
    Fin = InStr(Inicio, Linea, ",", vbTextCompare)
Else
    Dim Contador As Integer
    For Contador = 1 To nDato - 1 'busca la coma enterior al dato
        Inicio = InStr(Inicio + 1, Linea, ",", vbTextCompare)
    Next Contador
    
    If nDato - 1 = NroDeComas Then  'si es el último dato
        Fin = Len(Linea) + 1
    Else 'si no el primero ni el último dato
        Fin = InStr(Inicio + 1, Linea, ",", vbTextCompare)
    End If
    Inicio = Inicio + 1
End If
 
DatoDeLinea = Mid(Linea, Inicio, Fin - Inicio)
End Function

'se asegura que el path tenga '\' al final
Public Function CorregirPath(ByVal Path As String) As String
If Right(Path, 1) <> "\" Then Path = Path + "\"
CorregirPath = Path
End Function

'crea un log con el estado de la subida al momento
Public Sub CreaLog()
Dim Archivo As Integer
Archivo = FreeFile
Open CorregirPath(App.Path) + "log.txt" For Output As #Archivo
Dim Cont As Integer
For Cont = 1 To Subidor.lvLista.ListItems.Count
    Print #Archivo, GetDato(Cont, keyArchivo) + " " + _
    GetDato(Cont, keyEstado) + " " + _
    GetDato(Cont, keyIntento) + " " + _
    GetDato(Cont, keyEnviado) + " " + _
    GetDato(Cont, keyTamTotal) + " " + _
    GetDato(Cont, keyTiempo) + " " + _
    GetDato(Cont, keyUsuario) + " " + _
    GetDato(Cont, keyPassword) + " " + _
    GetDato(Cont, keyDirectorio) + " " + _
    GetDato(Cont, keyPath)
Next Cont
Close #Archivo
End Sub

'importa la lista de archivos a subir
Public Sub ImportaLista()
Dim Archivo As Integer
Dim Cantidad As Integer
Dim LineaLeida As String
Dim Datos() As String

Archivo = FreeFile
Open CorregirPath(App.Path) + "list.txt" For Input As Archivo
Do Until EOF(Archivo)
    Line Input #1, LineaLeida
    Cantidad = CuantasVecesEsta(",", LineaLeida)
    Cantidad = Cantidad + 1
    If Cantidad = 4 Or Cantidad = 5 Then
        ReDim Datos(1 To 5)
        Dim Cont As Integer
        For Cont = 1 To Cantidad
            Datos(Cont) = DatoDeLinea(LineaLeida, Cont)
        Next Cont
        If Cantidad = 4 Then Datos(5) = "/My Documents"
        AgregarALista SacaParteArchivo(Datos(1)), Datos(1), Datos(3), Datos(4), Datos(5), Datos(2)
    Else
    'insertar código si la lista tiene datos incorrectos
    End If
Loop

Close Archivo
End Sub

'exporta la lista de archivos a subir
Public Sub ExportaLista()
Dim Archivo As Integer
Dim Cont As Integer
Archivo = FreeFile
Open CorregirPath(App.Path) + "list.txt" For Output As Archivo
    For Cont = 1 To Subidor.lvLista.ListItems.Count
        Print #1, GetDato(Cont, keyPath) + ", " + GetDato(Cont, keyEstado) + ", " + GetDato(Cont, keyUsuario) + ", " + GetDato(Cont, keyPassword) + ", " + GetDato(Cont, keyDirectorio)
    Next Cont
Close Archivo
End Sub

'devuelve la string del elemento y la columna indicada
Public Function GetDato(ByVal Elemento As Integer, ByVal Key As String) As String
Dim Index As Integer
With Subidor.lvLista
    Index = .ColumnHeaders(Key).Index - 1
    If Index = 0 Then
        GetDato = .ListItems(Elemento).Text
    Else
        GetDato = .ListItems(Elemento).ListSubItems(Index).Text
    End If
End With
End Function

'pone la string en el elemento y la columna indicada
Public Sub SetDato(ByVal Elemento As Integer, ByVal Key As String, ByVal Dato As String)
Dim Index As Integer
With Subidor.lvLista
    Index = .ColumnHeaders(Key).Index - 1
    If Index = 0 Then
        .ListItems(Elemento).Text = Dato
    Else
        .ListItems(Elemento).ListSubItems(Index).Text = Dato
    End If
End With
End Sub

'devuelve el nombre del archivo de un path
Function SacaParteArchivo(ByVal Path As String) As String
SacaParteArchivo = Path
If InStr(1, Path, "\", vbTextCompare) = 0 Then Exit Function
Dim Posicion As Long
Posicion = 1
Do Until (Mid(Path, Len(Path) - Posicion, 1) = "\")
    Posicion = Posicion + 1
Loop
SacaParteArchivo = Right(Path, Posicion)
End Function

'agrega un elemento a la lista con los datos pasados
Public Sub AgregarALista(ByVal Nombre As String, ByVal Path As String, ByVal Usuario As String, ByVal Password As String, ByVal Directorio As String, Optional ByVal Estado As Variant)
Nombre = Trim(Nombre)
Path = Trim(Path)
Usuario = Trim(Usuario)
Password = Trim(Password)
Directorio = Trim(Directorio)

If Dir(Path) = "" Then Exit Sub 'si no existe el archivo sale

Dim UltimaLinea As Integer
With Subidor.lvLista.ListItems
    UltimaLinea = .Count + 1
    .Add
End With

Dim Cont As Integer
With Subidor.lvLista
For Cont = 1 To .ColumnHeaders.Count - 1
    .ListItems(UltimaLinea).ListSubItems.Add
Next Cont
End With

SetDato UltimaLinea, keyArchivo, Nombre
If IsMissing(Estado) Then
    SetDato UltimaLinea, keyEstado, "WAIT"
Else
    SetDato UltimaLinea, keyEstado, Trim(Estado)
End If
SetDato UltimaLinea, keyIntento, "0"
SetDato UltimaLinea, keyEnviado, "0"
SetDato UltimaLinea, keyTamTotal, "0"
SetDato UltimaLinea, keyTiempo, "00:00:00"
SetDato UltimaLinea, keyVelocidad, "0 KBps"
SetDato UltimaLinea, keyUsuario, Usuario
SetDato UltimaLinea, keyPassword, Password
SetDato UltimaLinea, keyDirectorio, Directorio
SetDato UltimaLinea, keyPath, Path
SetDato UltimaLinea, keyTamArchivo, str(FileLen(Path))
End Sub

Public Sub ApagarComputadora()
ExitWindowsEx EWX_SHUTDOWN Or EWX_FORCE Or EWX_POWEROFF, 0&
End Sub

Public Sub DeshabilitaParaSubir()
With Subidor
    .cmdSubir.Enabled = False
    .cmdAgregar.Enabled = False
    .cmdQuitar.Enabled = False
    .cmdQuitarTodos.Enabled = False
    .File1.Enabled = False
    .cmdEspacioLibre.Enabled = False
    .cmdImportar.Enabled = False
    .cmdExportar.Enabled = False
    .cmdCancelar.Enabled = True
End With
End Sub

Public Sub HabilitaParaSubir()
With Subidor
    .cmdSubir.Enabled = True
    .cmdAgregar.Enabled = True
    .cmdQuitar.Enabled = True
    .cmdQuitarTodos.Enabled = True
    .File1.Enabled = True
    .cmdEspacioLibre.Enabled = True
    .cmdImportar.Enabled = True
    .cmdExportar.Enabled = True
    .cmdCancelar.Enabled = False
End With
End Sub
