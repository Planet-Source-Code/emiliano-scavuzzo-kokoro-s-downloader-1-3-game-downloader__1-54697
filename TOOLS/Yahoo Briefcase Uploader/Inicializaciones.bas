Attribute VB_Name = "Inicializaciones"
'**************************************************************************************
'This code is a tool for Kokoro's Downloader
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit
Public Sub Inicializar()
Randomize 'inicializa la semilla

With Subidor
    'inicializa la listview
    With .lvLista.ColumnHeaders
        .Add , keyArchivo, "File", 1500, 0
        .Add , keyEstado, "State", 1050, 0
        .Add , keyIntento, "Attempt", 720, 0
        .Add , keyEnviado, "Sent", 1000, 0
        .Add , keyTamTotal, "Total Size", 1150, 0
        .Add , keyTiempo, "Time", 900, 0
        .Add , keyVelocidad, "Speed", 900, 0
        .Add , keyUsuario, "User", 1000, 0
        .Add , keyPassword, "Password", 1000, 0
        .Add , keyDirectorio, "Folder", 1400, 0
        .Add , keyPath, "Path", 2000, 0
        .Add , keyTamArchivo, "File Size", 1400, 0
    End With

    'inicializa el dir
    '.Dir1.Path = "c:\"

    'inicializa las variables que se usan para controlar
    'el tamaño cuando se cambia el tamaño de la ventana
    DirPorLargo = .Dir1.Width / (.Dir1.Width + .File1.Width) * 100
    FilePorLargo = .File1.Width / (.Dir1.Width + .File1.Width) * 100
    EspFijosHor = .Width - .Dir1.Width - .File1.Width
    EspEntreDirFile = .File1.Left - .Dir1.Left - .Dir1.Width
    MargenesLista = .Width - .lvLista.Width
    EspFijosVer = .Height - .Dir1.Height - .lvLista.Height
    DirPorAlto = .Dir1.Height / (.Dir1.Height + .lvLista.Height) * 100
    ListaPorAlto = .lvLista.Height / (.Dir1.Height + .lvLista.Height) * 100
    EspEntreDirLista = .lvLista.Top - .Dir1.Top - .Dir1.Height
End With

End Sub
