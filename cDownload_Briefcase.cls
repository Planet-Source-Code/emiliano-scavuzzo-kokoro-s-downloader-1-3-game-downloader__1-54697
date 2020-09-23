VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDownload_Briefcase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

'*********** Briefcase kind of download format ***********
' LN_END+1) Number of partitions the package is
'           divided on.
' LN_END+2) Size in bytes of each partition.
' LN_END+3) String to search for the real
'           package address.
' LN_END+4) Package lines. Format:
'           'name.xxx, http://address'
' After this line comes the partitions lines (if any)
' with this format:
'            'http://address'
' Then another package line with the format of LN_END+4
'
' NOTE: the last package is smaller or equal to the
'       rest but never bigger. The rest of the packages
'       have equal size.
' NOTE2: the size of each partition is the package
'        size divided by the amount of partitions,
'        except for the last partition from the last
'        package that is the previous plus (package
'        size) mod (amount of partitions).

Option Explicit

Private Const LN_NMR_PAR    As Byte = LN_END + 1
Private Const LN_PAR_SIZE   As Byte = LN_END + 2
Private Const LN_SRCH_STR   As Byte = LN_END + 3
Private Const LN_PACK_START As Byte = LN_END + 4

Private m_dblPackageSize As Double 'size of the package being downloaded
Private m_dblAmountDownloaded As Double 'amount of downloaded bytes of the package being downloaded

Private m_intNmrOfPackToDown As Integer 'amount of packages selected by the user to be downloaded
Private m_intNmrOfPartitions As Integer 'number of partitions the package is divided on

Private Type TypePartition
    Address As String
End Type

Private Type TypePackage
    Index As Integer    'index in the package list
    Size As Double      'package size
    name As String      'package file name
    Partition() As TypePartition
End Type

Private Type TypeDownPart
    Pack As Integer
    Part As Integer
End Type

Private pqPackage() As TypePackage 'packages needed data
Private PTD As TypeDownPart '(partition to download) partition being downloaded

Private m_strResponseInfo As String 'response from cdInfo

Private m_strGamePath As String
Private m_strPackagePath As String
Private m_intFileHandle As Integer

Private WithEvents cdInfo As CDownload
Attribute cdInfo.VB_VarHelpID = -1
Private WithEvents cdDownload As CDownload
Attribute cdDownload.VB_VarHelpID = -1

Implements cCommon

Private Sub Class_Initialize()
Set cdInfo = New CDownload
Set cdDownload = New CDownload
cdInfo.DownloadType = dtToBuffer
cdDownload.DontCloseFile = True
End Sub

Private Sub Class_Terminate()
Set cdInfo = Nothing
Set cdDownload = Nothing
End Sub

'Returns the amount of downloaded bytes of
'the package being downloaded.
Public Function cCommon_AmountDownloaded() As Double
cCommon_AmountDownloaded = m_dblAmountDownloaded
End Function

'Checks the info for wrong data.
'If it finds erroneous data it raises an error.
Public Sub cCommon_CheckPackages()

If Not IsNumeric(Trim(ReadLine(m_strInfo, LN_NMR_PAR))) Then Err.Raise 1
If Not IsNumeric(Trim(ReadLine(m_strInfo, LN_PAR_SIZE))) Then Err.Raise 1
If Trim(ReadLine(m_strInfo, LN_SRCH_STR)) = "" Then Err.Raise 1

Dim intNmrPack As Integer
Dim intNmrPart As Integer
Dim dblPartSize As Double
Dim dblTotalSize As Double
dblPartSize = Val(Trim(ReadLine(m_strInfo, LN_PAR_SIZE)))
intNmrPack = Val(Trim(ReadLine(m_strInfo, LN_PACK_NUM)))
intNmrPart = Val(Trim(ReadLine(m_strInfo, LN_NMR_PAR)))
dblTotalSize = Val(Trim(ReadLine(m_strInfo, LN_TOT_SIZE)))
If dblPartSize * intNmrPart * (intNmrPack - 1) > dblTotalSize Then Err.Raise 1
If dblTotalSize - (dblPartSize * intNmrPart * (intNmrPack - 1)) > (dblPartSize * intNmrPart) Then Err.Raise 1

Dim intCount As Integer
Dim intStart As Integer
Dim strLine As String

For intCount = 0 To intNmrPack * intNmrPart - 1
    
    strLine = Trim(ReadLine(m_strInfo, LN_PACK_START + intCount))
    
    If intCount Mod intNmrPart = 0 Then 'if it's a line with a file name
        If CountHowManyTimes(",", strLine) <> 1 Then Err.Raise 1
        intStart = InStr(1, strLine, ",")
        If Right(strLine, Len(strLine) - intStart) = "" Then Err.Raise 1
    Else
        If strLine = "" Then Err.Raise 1
        If CountHowManyTimes(",", strLine) <> 0 Then Err.Raise 1
    End If
    
Next intCount

End Sub

'Main function that starts the download.
Public Sub cCommon_Download()
m_intNmrOfPackToDown = frmMain.lstPackages.SelCount
m_intNmrOfPartitions = Val(Trim(ReadLine(m_strInfo, LN_NMR_PAR)))
m_strGamePath = m_strDownloadsPath + Trim(ReadLine(m_strInfo, LN_NAME)) + "\"
BuildArray
CleanState
StartPackageDownload 1
End Sub

'Fill the package list.
Public Sub cCommon_ReadPackages()
Dim intNmrPackages As Integer
Dim intNmrPartitions As Integer
Dim intCount As Integer
intNmrPackages = Val(Trim(ReadLine(m_strInfo, LN_PACK_NUM)))
intNmrPartitions = Val(Trim(ReadLine(m_strInfo, LN_NMR_PAR)))

For intCount = 0 To intNmrPackages - 1
    frmMain.lstPackages.AddItem "Package" + Str(intCount + 1) + ":" + " " + ReadPackageName(intCount * intNmrPartitions + LN_PACK_START), intCount
Next intCount
End Sub

'This function returns the size of the package being
'downloaded.
Public Function cCommon_PackageSize() As Double
cCommon_PackageSize = m_dblPackageSize
End Function

'Stop the download.
Public Sub cCommon_StopDownload()
cdInfo.Cancel
cdDownload.Cancel
If m_intFileHandle <> 0 Then Close #m_intFileHandle
m_intFileHandle = 0
End Sub

'Read the file name that is used for a package, on the
'specified line.
Private Function ReadPackageName(ByVal intLine As Integer) As String
Dim intCommaPos As Integer
ReadPackageName = Trim(ReadLine(m_strInfo, intLine))
intCommaPos = InStr(1, ReadPackageName, ",")
ReadPackageName = Trim(Left(ReadPackageName, intCommaPos - 1))
End Function

'Build the array that contains the game file addresses,
'indexes, sizes, and names.
Private Sub BuildArray()
ReDim pqPackage(1 To m_intNmrOfPackToDown) As TypePackage
Dim intCounter1 As Integer
Dim intCounter2 As Integer
For intCounter1 = 1 To m_intNmrOfPackToDown
    pqPackage(intCounter1).Index = IndexOfSelected(intCounter1)
    pqPackage(intCounter1).Size = SizeOfPackNmr(pqPackage(intCounter1).Index + 1)
    pqPackage(intCounter1).name = NameOfPackNmr(pqPackage(intCounter1).Index + 1)
    ReDim pqPackage(intCounter1).Partition(1 To m_intNmrOfPartitions)
        For intCounter2 = 1 To m_intNmrOfPartitions
            pqPackage(intCounter1).Partition(intCounter2).Address = AddressOfPackNmr(pqPackage(intCounter1).Index + 1, intCounter2)
        Next intCounter2
Next intCounter1
End Sub

'Returns the index from the selected item number
'intNumber.
Private Function IndexOfSelected(ByVal intNumber As Integer) As Integer
Dim intCountIndexes As Integer
Dim intCountSelect As Integer

intCountIndexes = 0
intCountSelect = 0
Do Until (intCountSelect = intNumber)
    If frmMain.lstPackages.Selected(intCountIndexes) Then intCountSelect = intCountSelect + 1
    intCountIndexes = intCountIndexes + 1
Loop
IndexOfSelected = intCountIndexes - 1
End Function

'Returns the size of the package number intNumber.
Public Function SizeOfPackNmr(ByVal intNumber As Integer) As Double
Dim dblTotalSize As Double
Dim dblPartitionSize As Double
Dim intNmrPackages As Integer
dblTotalSize = Val(Trim(ReadLine(m_strInfo, LN_TOT_SIZE)))
dblPartitionSize = Val(Trim(ReadLine(m_strInfo, LN_PAR_SIZE)))
intNmrPackages = Val(Trim(ReadLine(m_strInfo, LN_PACK_NUM)))

If intNmrPackages = intNumber Then
    SizeOfPackNmr = dblTotalSize - (dblPartitionSize * m_intNmrOfPartitions * (intNmrPackages - 1))
Else
    SizeOfPackNmr = dblPartitionSize * m_intNmrOfPartitions
End If
End Function

'Returns the address of the partition intPartition
'from the package intPackage.
Private Function AddressOfPackNmr(ByVal intPackage As Integer, ByVal intParticion As Integer) As String
If intParticion = 1 Then
    Dim intStart As Integer
    AddressOfPackNmr = Trim(ReadLine(m_strInfo, LN_PACK_START + ((intPackage - 1) * m_intNmrOfPartitions)))
    intStart = InStr(1, AddressOfPackNmr, ",", vbTextCompare)
    AddressOfPackNmr = Trim(Right(AddressOfPackNmr, Len(AddressOfPackNmr) - intStart))
Else
    AddressOfPackNmr = Trim(ReadLine(m_strInfo, LN_PACK_START + ((intPackage - 1) * m_intNmrOfPartitions) + intParticion - 1))
End If
End Function

'Returns the name of the package number intNumber.
Private Function NameOfPackNmr(ByVal intNumber As Integer) As String
NameOfPackNmr = ReadPackageName((intNumber - 1) * m_intNmrOfPartitions + LN_PACK_START)
End Function

'Start to download a new package specified by intPackNmr.
Private Sub StartPackageDownload(ByVal intPackNmr As Integer)
PTD.Pack = intPackNmr
PTD.Part = 1
m_dblPackageSize = pqPackage(PTD.Pack).Size
m_dblAmountDownloaded = 0
ShadowPackage (pqPackage(PTD.Pack).Index)
MakePTDFolder
frmMain.lblState.Caption = "Preparing data..."
If m_blnUseProxy = True Then
    cdInfo.AccessType = cdNamedProxy
    cdInfo.Proxy = m_strProxy
    cdInfo.ProxyPort = m_lngPort
Else
    cdInfo.AccessType = cdDirect
End If

cdInfo.Download pqPackage(PTD.Pack).Partition(PTD.Part).Address
End Sub

Private Sub cdInfo_Starting(ByVal FileSize As Long, ByVal Header As String, ByVal FileHandle As Integer)
m_strResponseInfo = ""
End Sub

Private Sub cdInfo_Completed()
On Error GoTo ErrorHandler

m_strResponseInfo = cdInfo.GetBuffer
TakeAddressFromInfo
DownloadPTD
Exit Sub

ErrorHandler:
    If m_intFileHandle <> 0 Then Close #m_intFileHandle
    m_intFileHandle = 0
    MessageBox frmMain.hwnd, "Error trying to obtain package data.", "Error", MB_ICONERROR
    Finish_Download_Actions
End Sub

Private Sub cdInfo_Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
Select Case (Number)
Case 11001:
    MessageBox frmMain.hwnd, "Could not locate remote server. Check your internet connection.", "Error", MB_ICONERROR
Case 10060:
    MessageBox frmMain.hwnd, "The connection time-out has expired. Try later.", "Error", MB_ICONERROR
Case Else
    MessageBox frmMain.hwnd, Description, "Error", MB_ICONERROR
End Select
If m_intFileHandle <> 0 Then Close #m_intFileHandle
m_intFileHandle = 0
Finish_Download_Actions
End Sub

'Grabs the real file name from m_strResponseInfo and
'stores it in pqPackage(PTD.Pack).Partition(PTD.Part).Address
Private Function TakeAddressFromInfo() As String
Dim intStart As Integer
Dim intEnd As Integer

intStart = PosAdvanceSearch(Trim(ReadLine(m_strInfo, LN_SRCH_STR)), "href=", m_strResponseInfo)
If intStart = 0 Then Err.Raise 1

intStart = InStr(intStart, m_strResponseInfo, Chr(34))
If intStart = 0 Or intStart = Null Then Err.Raise 1
intStart = intStart + 1
intEnd = InStr(intStart, m_strResponseInfo, Chr(34))
If intEnd = 0 Or intEnd = Null Then Err.Raise 1
intEnd = intEnd - 1

TakeAddressFromInfo = Mid(m_strResponseInfo, intStart, intEnd - intStart + 1)
If TakeAddressFromInfo = "" Then Err.Raise 1
pqPackage(PTD.Pack).Partition(PTD.Part).Address = TakeAddressFromInfo
m_strResponseInfo = "" 'reset string to free memory

End Function

'This sub is called after we grab the file real address.
Private Sub DownloadPTD()

If m_blnUseProxy = True Then
    cdDownload.AccessType = cdNamedProxy
    cdDownload.Proxy = m_strProxy
    cdDownload.ProxyPort = m_lngPort
Else
    cdDownload.AccessType = cdDirect
End If

'if it's the first partition
If PTD.Part = 1 Then
    frmMain.lblState.Caption = "Connecting to server..."
    cdDownload.Download pqPackage(PTD.Pack).Partition(PTD.Part).Address, m_strPackagePath
Else
    cdDownload.Download pqPackage(PTD.Pack).Partition(PTD.Part).Address, m_intFileHandle
End If
End Sub

Private Sub cdDownload_Starting(ByVal FileSize As Long, ByVal Header As String, ByVal FileHandle As Integer)
m_intFileHandle = FileHandle
If PTD.Part = 1 Then frmMain.lblState.Caption = "Connected!"
End Sub

Private Sub cdDownload_DataArrival(ByVal bytesTotal As Long)
frmMain.lblState.Caption = "Downloading " + pqPackage(PTD.Pack).name
If IsStateTimerActive = False And PTD.Part = 1 Then StartStateTimer
m_dblAmountDownloaded = m_dblAmountDownloaded + bytesTotal
End Sub
Private Sub cdDownload_Completed()
PartitionDownloaded
End Sub

Private Sub cdDownload_Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
Select Case (Number)
Case 11001:
    MessageBox frmMain.hwnd, "Could not locate remote server. Check your internet connection.", "Error", MB_ICONERROR
Case 10060:
    MessageBox frmMain.hwnd, "The connection time-out has expired. Try later.", "Error", MB_ICONERROR
Case Else
    MessageBox frmMain.hwnd, Description, "Error", MB_ICONERROR
End Select
Finish_Download_Actions
If m_intFileHandle <> 0 Then Close #m_intFileHandle
m_intFileHandle = 0
End Sub

Private Sub MakePTDFolder()
MakeGameFoldIfDoesntExist
m_strPackagePath = m_strGamePath + pqPackage(PTD.Pack).name
End Sub

'This sub is called after any partition is downloaded.
Private Sub PartitionDownloaded()
On Error GoTo ErrorHandler
If LOF(m_intFileHandle) <> ExpectedSize(PTD) Then Err.Raise 1

If PTD.Part = m_intNmrOfPartitions Then 'if the package download is completed
    If m_intFileHandle <> 0 Then Close #m_intFileHandle
    m_intFileHandle = 0
    StopStateTimer
    CleanState
    MarkPackage (pqPackage(PTD.Pack).Index)
    If PTD.Pack = m_intNmrOfPackToDown Then 'if all the packages were downloaded
        GameDownloaded
    Else 'if not all the packages were downloaded
        StartPackageDownload (PTD.Pack + 1)
    End If

Else 'if the package download isn't completed
    PTD.Part = PTD.Part + 1
    If m_blnUseProxy = True Then
        cdInfo.AccessType = cdNamedProxy
        cdInfo.Proxy = m_strProxy
        cdInfo.ProxyPort = m_lngPort
    Else
        cdInfo.AccessType = cdDirect
    End If
    cdInfo.Download pqPackage(PTD.Pack).Partition(PTD.Part).Address
End If

Exit Sub
ErrorHandler:
    If m_intFileHandle <> 0 Then Close #m_intFileHandle
    m_intFileHandle = 0
    MessageBox frmMain.hwnd, "Error tryng to download package.", "Error", MB_ICONERROR
    Finish_Download_Actions
End Sub

Private Sub GameDownloaded()
MessageBox frmMain.hwnd, "Game downloaded correctly.", "All done", MB_ICONINFORMATION
Finish_Download_Actions
End Sub

'This function returns the size a package should
'have if the PTD was downloaded correctly
Private Function ExpectedSize(PTDD As TypeDownPart) As Double
If PTDD.Part = m_intNmrOfPartitions Then
    ExpectedSize = pqPackage(PTDD.Pack).Size
Else
    Dim dblUnknownSize As Double
    dblUnknownSize = Int(pqPackage(PTDD.Pack).Size / m_intNmrOfPartitions)
    ExpectedSize = PTDD.Part * dblUnknownSize
End If
End Function

