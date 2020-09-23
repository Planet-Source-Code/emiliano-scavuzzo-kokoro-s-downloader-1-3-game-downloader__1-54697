VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

'NOTE: If FileSize = 0 that means the size of the file
'      is unknown.

'==============================================================================
'EVENTS
'==============================================================================

Public Event Starting(ByVal FileSize As Long, ByVal Header As String, ByVal FileHandle As Integer)
Public Event DataArrival(ByVal bytesTotal As Long)
Attribute DataArrival.VB_Description = "Occurs when data has been received from the remote computer"
Public Event Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
Public Event Completed()
Attribute Completed.VB_Description = "Occurs when download is completed"

'==============================================================================
'CONSTANTS
'==============================================================================

Public Enum AccessConstants
    cdDirect = 0
    cdNamedProxy = 1
End Enum

Public Enum Download_Type
    dtToFile = 0
    dtToBuffer = 1
End Enum

Private Const ACCEPT_TOKEN As String = "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-shockwave-flash, */*"
Private Const USERAGENT_TOKEN As String = "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)"

'==============================================================================
'MEMBER VARIABLES
'==============================================================================

Private m_acAccess As AccessConstants       'uses proxy or direct connection?
Private m_strProxy As String                'proxy host
Private m_strURL As String                  'URL to download
Private m_strDestination As String          'path to destination file
Private m_lngProxyPort As Long              'proxy port
Private m_blnRedirDisabled As Boolean       'disable automatic redirection?
Private m_dtDownloadType As Download_Type   'uses file or buffer download?
Private m_blnUnEncrypt As Boolean           'unencrypt downloaded data?
Private m_blnDontCloseFile As Boolean       'close file when finished?

Private m_strHeader As String               'download header
Private m_blnHeaderArrived As Boolean       'has the header arrived?
Private m_blnResuming As Boolean            'are we resuming?
Private m_lngResumeFrom As Long             'part of the file already downloaded
Private m_intFileHandle As Integer          'file handle
Private m_lngFileSize As Long               'file total size
Private m_strBuffer As String               'buffer

'our socket
Private WithEvents cmSocket As CSocketMaster
Attribute cmSocket.VB_VarHelpID = -1

Private Sub Class_Terminate()
Set cmSocket = Nothing
End Sub

'==============================================================================
'PROPERTIES
'==============================================================================

Public Property Get Proxy() As String
Attribute Proxy.VB_Description = "Returns/Sets HTTP proxy"
Proxy = m_strProxy
End Property

Public Property Let Proxy(ByVal strProxy As String)
m_strProxy = Trim(strProxy)
End Property

Public Property Get ProxyPort() As Long
Attribute ProxyPort.VB_Description = "Returns/Sets proxy port"
ProxyPort = m_lngProxyPort
End Property

Public Property Let ProxyPort(ByVal lngProxyPort As Long)
m_lngProxyPort = lngProxyPort
End Property

Public Property Get AccessType() As AccessConstants
Attribute AccessType.VB_Description = "Returns/Sets the proxy behavior for this control's connections"
AccessType = m_acAccess
End Property

Public Property Let AccessType(ByVal acAccess As AccessConstants)
m_acAccess = acAccess
End Property

Public Property Get URL() As String
Attribute URL.VB_Description = "Returns/Sets file URL"
URL = m_strURL
End Property

Public Property Let URL(ByVal strURL As String)
m_strURL = Trim(strURL)
End Property

Public Property Get Destination() As String
Attribute Destination.VB_Description = "Returns/Sets full path where the file will be saved"
Destination = m_strDestination
End Property

Public Property Let Destination(ByVal strDestination As String)
m_strDestination = Trim(Destination)
End Property

Public Property Get DisableRedirection() As Boolean
Attribute DisableRedirection.VB_Description = "Returns/Sets if automatic redirection is disabled"
DisableRedirection = m_blnRedirDisabled
End Property

Public Property Let DisableRedirection(ByVal blnRedir As Boolean)
m_blnRedirDisabled = blnRedir
End Property

Public Property Get FileSize() As Long
Attribute FileSize.VB_Description = "Returns file size in bytes or zero if it is unknown"
FileSize = m_lngFileSize
End Property

Public Property Get DownloadType() As Download_Type
DownloadType = m_dtDownloadType
End Property

Public Property Let DownloadType(DownloadType As Download_Type)
If cmSocket Is Nothing Then
    m_dtDownloadType = DownloadType
Else
    Err.Raise sckInvalidOp, "CDownload.DownloadType", "Invalid operation at current state"
End If
End Property

Public Property Get UnEncrypt() As Boolean
UnEncrypt = m_blnUnEncrypt
End Property

Public Property Let UnEncrypt(UnEncrypt As Boolean)
If cmSocket Is Nothing Then
    m_blnUnEncrypt = UnEncrypt
Else
    Err.Raise sckInvalidOp, "CDownload.UnEncrypt", "Invalid operation at current state"
End If
End Property

Public Property Get DontCloseFile() As Boolean
DontCloseFile = m_blnDontCloseFile
End Property

Public Property Let DontCloseFile(CloseFileWhenFinish As Boolean)
m_blnDontCloseFile = CloseFileWhenFinish
End Property

Public Sub Download(Optional URL As Variant, Optional Destination As Variant)
Attribute Download.VB_Description = "Download file"
DownloadOrResume URL, Destination
End Sub

Public Sub ResumeDownload(Optional URL As Variant, Optional Destination As Variant)
m_blnResuming = True
DownloadOrResume URL, Destination
End Sub

Public Function GetBuffer() As String
GetBuffer = m_strBuffer
End Function

Public Sub Cancel()
Attribute Cancel.VB_Description = "Method used to cancel download"
Reset
End Sub

Private Sub DownloadOrResume(Optional URL As Variant, Optional Destination As Variant)
On Error GoTo Error_Handler
Set cmSocket = New CSocketMaster

If Not IsMissing(URL) Then
    Select Case varType(URL)
    Case vbString
        m_strURL = Trim(URL)
    Case Else
        Err.Raise sckUnsupported, "CDownload.DownloadOrResume", "Unsupported variant type."
    End Select
End If

If Not IsMissing(Destination) Then
    Select Case varType(Destination)
    Case vbInteger To vbLong
        m_intFileHandle = Destination
    Case vbString
        m_strDestination = Trim(Destination)
    Case Else
        Err.Raise sckUnsupported, "CDownload.DownloadOrResume", "Unsupported variant type."
    End Select
End If

If m_acAccess = cdDirect Then
    cmSocket.Connect GetHostFromURL(m_strURL), 80
Else
    cmSocket.Connect m_strProxy, m_lngProxyPort
End If

Exit Sub
Error_Handler:
    Dim lngError As Long
    Dim strDescription As String
    lngError = Err.Number
    strDescription = Err.Description
    Reset
    RaiseEvent Error(lngError, strDescription, False)
End Sub


Private Sub cmSocket_Connect()
On Error GoTo Error_Handler

If m_dtDownloadType = dtToFile Then
    m_strBuffer = ""
    If m_intFileHandle = 0 Then 'if user didn't pass file handle
        
        'does the destination file exists?
        If Dir(m_strDestination, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = GetFileFromPath(m_strDestination) Then
    
            If m_blnResuming Then 'are we resuming?
                m_lngResumeFrom = FileLen(m_strDestination)
            Else
                SetAttr m_strDestination, vbNormal
                Kill m_strDestination
            End If

        Else
            m_blnResuming = False
        End If
    
        m_intFileHandle = FreeFile
        Open m_strDestination For Binary Lock Read Write As #m_intFileHandle
        
    Else 'if user pass file handle
    
        If m_blnResuming Then m_lngResumeFrom = LOF(m_intFileHandle)
        
    End If
    
Else  'we use buffer download
    If m_blnResuming Then
        m_lngResumeFrom = Len(m_strBuffer)
    Else
        m_strBuffer = ""
    End If
End If


Dim strCommand As String

strCommand = "GET " + GetFileFromURL(m_strURL) + " HTTP/1.0" + vbCrLf
strCommand = strCommand + "Accept: " + ACCEPT_TOKEN + vbCrLf
strCommand = strCommand + "Referer: " + GetHostFromURL(m_strURL) + vbCrLf
strCommand = strCommand + "User-Agent: " + USERAGENT_TOKEN + vbCrLf

If m_blnResuming And m_lngResumeFrom > 0 Then 'are we resuming?
    strCommand = strCommand + "Range: bytes=" + Trim(Str(m_lngResumeFrom)) + "-" + vbCrLf
End If

strCommand = strCommand + "Host: " + GetHostFromURL(m_strURL) + vbCrLf
strCommand = strCommand + vbCrLf

cmSocket.SendData strCommand

Exit Sub
Error_Handler:
    Dim lngError As Long
    Dim strDescription As String
    lngError = Err.Number
    strDescription = Err.Description
    Reset
    RaiseEvent Error(lngError, strDescription, False)
End Sub

Private Sub cmSocket_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Error_Handler
Dim strChunk As String
cmSocket.GetData strChunk

'if header hasn't arrived
If m_blnHeaderArrived = False Then
    
    m_strHeader = m_strHeader & strChunk
    
    Dim lngSplit As Long
    lngSplit = InStr(1, m_strHeader, vbCrLf + vbCrLf)
    
    'has the header finished on this chunk?
    If lngSplit = 0 Or lngSplit = Null Then Exit Sub
    
    'yes! the header has finished
    m_blnHeaderArrived = True
    
    'maybe this chunk is half header and half file
    'we split the two
    strChunk = Right(m_strHeader, Len(m_strHeader) - lngSplit - 3)
    m_strHeader = Left(m_strHeader, lngSplit + 3)
    
    'is redirection enabled?
    If m_blnRedirDisabled = False Then
        Dim strLocation As String
        strLocation = GetVariableValue(m_strHeader, "Location")
        'does the header indicates a redirection?
        If strLocation <> "" Then
            Dim intTemp As Integer
            intTemp = m_intFileHandle 'save file handle
            Reset
            m_intFileHandle = intTemp 'restore file handle
            
            If Left(strLocation, 1) = "/" Then
                m_strURL = GetHostFromURL(m_strURL) + strLocation
            Else
                m_strURL = strLocation
            End If
            
            Download
            Exit Sub
        End If
    End If
    
    Dim strFileSize As String
    
    strFileSize = GetVariableValue(m_strHeader, "Content-Length")
    If strFileSize = "" Then
        m_lngFileSize = 0
    Else
        m_lngFileSize = Val(strFileSize)
    End If
    
    Dim strContentRange As String
    
    strContentRange = GetVariableValue(m_strHeader, "Content-Range")
    If m_blnResuming And Len(strContentRange) = 0 Then CleanAndReopenFile
    
    RaiseEvent Starting(m_lngFileSize, m_strHeader, m_intFileHandle)
End If

'if header has arrived

If m_dtDownloadType = dtToFile Then
    If m_blnUnEncrypt Then
        Put m_intFileHandle, LOF(m_intFileHandle) + 1, UnEncryptText(strChunk, LOF(m_intFileHandle))
    Else
        Put m_intFileHandle, LOF(m_intFileHandle) + 1, strChunk
    End If
Else
    If m_blnUnEncrypt Then
        m_strBuffer = m_strBuffer & UnEncryptText(strChunk, Len(m_strBuffer))
    Else
        m_strBuffer = m_strBuffer & strChunk
    End If
End If

RaiseEvent DataArrival(Len(strChunk))

Exit Sub
Error_Handler:
    Dim lngError As Long
    Dim strDescription As String
    lngError = Err.Number
    strDescription = Err.Description
    Reset
    RaiseEvent Error(lngError, strDescription, False)
End Sub

Private Sub cmSocket_CloseSck()

'some web pages don't have headers so we have to
'raise all the events that couldn't be raised while
'the file was downloading
If m_blnHeaderArrived = False Then
    
    Dim strData As String
    If m_blnUnEncrypt Then
        strData = UnEncryptText(m_strHeader)
    Else
        strData = m_strHeader
    End If
    m_strHeader = ""
    
    If m_blnResuming Then
        If m_dtDownloadType = dtToFile Then
            CleanAndReopenFile
        Else
            m_strBuffer = ""
        End If
    End If
    
    RaiseEvent Starting(Len(strData), "", m_intFileHandle)
    
    If m_dtDownloadType = dtToFile Then
        Put m_intFileHandle, LOF(m_intFileHandle) + 1, strData
    Else
        m_strBuffer = strData
    End If
    
    RaiseEvent DataArrival(Len(strData))
End If

Reset
RaiseEvent Completed
End Sub

'Ups! We got an error
Private Sub cmSocket_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Reset
RaiseEvent Error(Number, Description, True)
End Sub

Private Function CleanAndReopenFile()
Close #m_intFileHandle

m_intFileHandle = FreeFile
Open m_strDestination For Output Lock Read Write As #m_intFileHandle
Close #m_intFileHandle

m_intFileHandle = FreeFile
Open m_strDestination For Binary Lock Read Write As #m_intFileHandle
End Function

'returns the host from an URL
'ie: 'http://www.yahoo.com/file.txt' => 'www.yahoo.com'
Private Function GetHostFromURL(ByVal strURL As String) As String

strURL = Trim(strURL)
If Left(strURL, 7) = "http://" Then strURL = Mid(strURL, 6, Len(strURL) - 5)
If Left(strURL, 2) = "//" Then strURL = Mid(strURL, 3, Len(strURL) - 2)

Dim Init As Integer
Init = InStr(1, strURL, "/", vbTextCompare)

If Init <> 0 Then strURL = Left(strURL, Init - 1)
GetHostFromURL = strURL

End Function

'get the file part from an URL that goes after the
'GET command to download files IF IT IS NOT USING PROXY
'ie: 'http://www.yahoo.com/file.txt' => '/file.txt'
Private Function GetFileFromURL(ByVal strURL As String) As String

If m_acAccess = cdNamedProxy Then
    GetFileFromURL = strURL
    Exit Function
End If

If Left(strURL, 7) = "http://" Then strURL = Right(strURL, Len(strURL) - 7)
Dim Init As Integer
Init = InStr(1, strURL, "/", vbTextCompare)
If Init = 0 Or Init = Null Then
    GetFileFromURL = "/"
Else
    GetFileFromURL = Right(strURL, Len(strURL) - Init + 1)
End If
End Function

'get file part from a path
'ie: 'c:\folder\file.txt' => 'file.txt'
Private Function GetFileFromPath(ByVal strPath As String) As String
GetFileFromPath = strPath
If InStr(1, strPath, "\", vbTextCompare) = 0 Then Exit Function
Dim Position As Long
Position = 1
Do Until (Mid(strPath, Len(strPath) - Position, 1) = "\")
    Position = Position + 1
Loop
GetFileFromPath = Right(strPath, Position)
End Function

'get variable value from the header
Private Function GetVariableValue(ByRef strHeader As String, ByVal strVariable As String) As String
Dim Init As Long
Dim Last As Long

Init = InStr(1, strHeader, strVariable, vbTextCompare)

If Init = 0 Or Init = Null Then
    GetVariableValue = ""
    Exit Function
End If

Init = Init + Len(strVariable) + 1
Last = InStr(Init, strHeader, vbCrLf, vbTextCompare)


GetVariableValue = Trim(Mid(strHeader, Init, Last - Init))

End Function

'reset variables
Private Sub Reset()
Set cmSocket = Nothing
m_strHeader = ""
m_blnHeaderArrived = False
m_blnResuming = False
m_lngResumeFrom = 0
If m_intFileHandle <> 0 And Not m_blnDontCloseFile Then
    Close #m_intFileHandle
End If
m_intFileHandle = 0
m_lngFileSize = 0
End Sub
