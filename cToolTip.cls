VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************************************
'This tooltip class was written by
'Kaustubh Zoal
'**************************************************************************************

Option Explicit
'Initialization of New ClassNames
Private Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Type tagINITCOMMONCONTROLSEX
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be initialized.
End Type

' ToolTip Styles
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40 ' comctl32.dll v5.8 require

Private Const CW_USEDEFAULT = &H80000000

Private Const WS_POPUP = &H80000000

Private Const WM_USER = &H400

' ToolTip Messages
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_ADDTOOL = (WM_USER + 4)
Private Const TTM_DELTOOL = (WM_USER + 5)
Private Const TTM_NEWTOOLRECT = (WM_USER + 6)
Private Const TTM_GETTOOLINFO = (WM_USER + 8)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private Const TTF_IDISHWND = &H1
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10


Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hwnd        As Long
    uId         As Long
    cRect       As RECT
    hinst       As Long
    lpszText    As String
End Type

Public Enum TTStyle
    ttStyleStandard = 1
    ttStyleBalloon = 2
End Enum

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private hTT As Long
'local variable(s) to hold property value(s)
Private mvarTipWidth As Long 'local copy
Private mvarDelayTime As Long 'local copy
Private mvarVisibleTime As Long 'local copy
Private mvarBkColor As Long 'local copy
Private mvarTxtColor As Long 'local copy
Private mvarStyle As TTStyle 'local copy

Public Property Let TxtColor(ByVal vData As Long)
  mvarTxtColor = vData
  SendMessageLong hTT, TTM_SETTIPTEXTCOLOR, vData, 0&
End Property

Public Property Get TxtColor() As Long
  TxtColor = mvarTxtColor
End Property

Public Property Let BkColor(ByVal vData As Long)
  mvarBkColor = vData
  SendMessageLong hTT, TTM_SETTIPBKCOLOR, vData, 0&
End Property

Public Property Get BkColor() As Long
  BkColor = mvarBkColor
End Property

Public Property Let VisibleTime(ByVal vData As Long)
   mvarVisibleTime = vData
   SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, vData
End Property

Public Property Get VisibleTime() As Long
  VisibleTime = mvarVisibleTime
End Property

Public Property Let DelayTime(ByVal vData As Long)
    mvarDelayTime = vData
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_INITIAL, vData
End Property

Public Property Get DelayTime() As Long
    DelayTime = mvarDelayTime
End Property

Public Property Let TipWidth(ByVal vData As Long)
   mvarTipWidth = vData
   SendMessageLong hTT, TTM_SETMAXTIPWIDTH, 0, vData
End Property

Public Property Get TipWidth() As Long
   TipWidth = mvarTipWidth
End Property

Public Property Let Style(ByVal vData As TTStyle)
   Dim lStyle As Long
   mvarStyle = vData
   If hTT Then
      lStyle = GetWindowLong(hTT, GWL_STYLE)
      If vData = ttStyleBalloon Then lStyle = lStyle Or TTS_BALLOON
      If vData = ttStyleStandard And (lStyle And ttStyleBalloon) Then lStyle = lStyle Xor TTS_BALLOON
      SetWindowLong hTT, GWL_STYLE, lStyle
   End If
End Property

Public Property Get Style() As TTStyle
   Style = mvarStyle
End Property

Private Sub InitComctl32(dwFlags As Long)
   Dim icc As tagINITCOMMONCONTROLSEX
   On Error GoTo Err_OldVersion
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   InitCommonControlsEx icc
   On Error GoTo 0
   Exit Sub
Err_OldVersion:
   InitCommonControls
End Sub

Public Sub SetToolTipObj(objHwnd As Long, sTipText As String, Optional bCenter As Boolean = False)
    Dim TI As TOOLINFO
    With TI
        .hwnd = objHwnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
        If bCenter Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uId = objHwnd
        .lpszText = sTipText
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub SetToolTipItem(objHwnd As Long, nItem As Long, lft As Long, tp As Long, rght As Long, btm As Long, sTipText As String, Optional bCenter As Boolean = False)
   Dim TI As TOOLINFO, RC As RECT
   RC.Bottom = btm
   RC.Left = lft
   RC.Right = rght
   RC.Top = tp
   With TI
        .cRect = RC
        .hwnd = objHwnd
        .uFlags = TTF_SUBCLASS
        If bCenter Then
            .uFlags = .uFlags Or TTF_CENTERTIP
        End If
        .uId = nItem
        .lpszText = sTipText
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_ADDTOOL, 0, TI
End Sub

Public Sub DelToolTip(objHwnd As Long, Optional nItem As Long = -1)
   Dim TI As TOOLINFO
   TI.hwnd = objHwnd
   TI.cbSize = Len(TI)
   If nItem < 0 Then TI.uId = objHwnd Else TI.uId = nItem
   SendMessage hTT, TTM_DELTOOL, 0, TI
End Sub

Public Sub AjustItemRect(objHwnd As Long, nItem As Long, lft As Long, tp As Long, rght As Long, btm As Long)
    Dim TI As TOOLINFO, RC As RECT
    With TI
        .hwnd = objHwnd
        .uId = nItem
        .cbSize = Len(TI)
    End With
    SendMessage hTT, TTM_GETTOOLINFO, 0&, TI
    RC.Bottom = btm
    RC.Left = lft
    RC.Right = rght
    RC.Top = tp
    TI.cRect = RC
    SendMessage hTT, TTM_NEWTOOLRECT, 0&, TI
End Sub

Private Sub Class_Initialize()
   InitComctl32 ICC_BAR_CLASSES
   hTT = CreateWindowEx(0, "tooltips_class32", 0&, TTS_NOPREFIX Or TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)
   Style = ttStyleStandard
   TipWidth = 300
   BkColor = &HEEFFFF
   TxtColor = vbBlack
   DelayTime = 500
   VisibleTime = 2000
End Sub

Private Sub Class_Terminate()
  If hTT Then DestroyWindow (hTT)
End Sub

