VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2130
   DefaultCancel   =   -1  'True
   ForeColor       =   &H000000FF&
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   142
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************
'This button control was written by
'Leo Barsukov
'**************************************************************************************

Option Explicit

'For drawing the caption
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Rect drawing
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Create/Delete brush
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'For drawing lines
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Misc
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Dim cColor As Long
'Center
Private Const DT_CENTERABS = &H65

'Default system colours
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

'Events
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Height      As Long                 'Width
Private Width       As Long                 'Height

Private CurrText    As String               'Current caption
Private CurrFont    As StdFont              'Current font

'Rects structures
Private RC          As RECT
Private RC2         As RECT
Private RC3         As RECT

Private LastButton  As Byte                 'Last button pressed
Private isEnabled   As Boolean              'Enabled or not

'Default system colors
Private cFace       As Long
Private cLight      As Long
Private cHighLight  As Long
Private cShadow     As Long
Private cDarkShadow As Long
Private cText       As Long

Private lastStat    As Byte                 'Last property
Private TE          As String               'Text


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
        If isEnabled Then
            RaiseEvent Click
            UserControl.Refresh
        End If
End Sub

'Single click
Private Sub UserControl_Click()
        RaiseEvent Click
        UserControl.Refresh
End Sub


'Double click
Private Sub UserControl_DblClick()
    
    If LastButton = 1 Then
        'Call the mousedown sub
        UserControl_MouseDown 1, 1, 1, 1
    End If
    
End Sub

Public Property Get ForeColor() As OLE_COLOR
ForeColor = cColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
cColor = New_ForeColor
PropertyChanged "ForeColor"
End Property

'Initialize
Private Sub UserControl_Initialize()

    LastButton = 1   'Lastbutton = right mouse button
    RC2.Left = 2
    RC2.Top = 2
    SetColors        'Get default colors
    
End Sub

'Initialize properties
Private Sub UserControl_InitProperties()

    CurrText = "Caption"                'Caption
    isEnabled = True                    'Enabled
    Set CurrFont = UserControl.Font     'Font
    
End Sub


'Mousedown
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    LastButton = Button     'Set lastbutton
    
    If Button <> 2 Then
        Redraw 2, False     'Redraw button
    End If
    'Raise mousedown event
    RaiseEvent MouseDown(Button, Shift, x, y)
    
End Sub


'Mousemove
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button < 2 Then
        If x < 0 Or y < 0 Or x > Width Or y > Height Then   'Not inside button
            Redraw 0, False                                 'Redraw
        ElseIf Button = 1 Then                              'Right click
            Redraw 2, False                                 'Redraw
        End If
    End If
    
    'Raise mousemove event
    RaiseEvent MouseMove(Button, Shift, x, y)
    
End Sub


'Mouseup
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then
        Redraw 0, False     'Redraw
    End If
    
    'Raise mousrup event
    RaiseEvent MouseUp(Button, Shift, x, y)
    
End Sub


'Property Get: Caption
Public Property Get Caption() As String
    Caption = CurrText      'Return caption
End Property


'Property Let: Caption
Public Property Let Caption(ByVal newValue As String)
    CurrText = newValue     'Set caption
    Redraw 0, True          'Redraw
    PropertyChanged "TX"    'Last property changed is text
End Property


'Property Get: Enabled
Public Property Get Enabled() As Boolean
    Enabled = isEnabled     'Set enabled/disabled
End Property


'Property Let: Enabled
Public Property Let Enabled(ByVal newValue As Boolean)
    isEnabled = newValue            'Set enabled/disabled
    Redraw 0, True                  'Redraw
    UserControl.Enabled = isEnabled 'Set enabled/disabled
    PropertyChanged "ENAB"          'Last property changed is enabled
End Property


'Property Get: Font
Public Property Get Font() As Font
    Set Font = CurrFont             'Return font
End Property


'Property Set: Font
Public Property Set Font(ByRef newFont As Font)
    Set CurrFont = newFont          'Set font
    Set UserControl.Font = CurrFont 'Set font
    Redraw 0, True                  'Redraw
    PropertyChanged "FONT"          'Last property changed is font
End Property


'Property Get: hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd         'Return hWnd
End Property


'Resize
Private Sub UserControl_Resize()
    
    'Renew dimension variables
    Height = UserControl.ScaleHeight
    Width = UserControl.ScaleWidth
    
    'Set rect1
    RC.Bottom = Height
    RC.Right = Width
    
    'Set rect 2
    RC2.Bottom = Height
    RC2.Right = Width
    
    'Set rect 3
    RC3.Left = 4
    RC3.Top = 4
    RC3.Right = Width - 4
    RC3.Bottom = Height - 4
    
    Redraw 0, True          'Redraw
    
End Sub


'Read Properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cColor = PropBag.ReadProperty("ForeColor", &H80000012)
    CurrText = PropBag.ReadProperty("TX", "")                       'Caption
    isEnabled = PropBag.ReadProperty("ENAB", True)                  'Enabled
    Set CurrFont = PropBag.ReadProperty("FONT", UserControl.Font)   'Font
    
    UserControl.Enabled = isEnabled     'Set enabled state
    Set UserControl.Font = CurrFont     'Set font
    
    SetColors       'Set colours
    Redraw 0, True  'Redraw

End Sub


'Write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("ForeColor", cColor, &H80000012)
    PropBag.WriteProperty "TX", CurrText    'Caption
    PropBag.WriteProperty "ENAB", isEnabled 'Enabled state
    PropBag.WriteProperty "FONT", CurrFont  'Font

End Sub


'Redraw
Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)

  Dim i               As Long
  Dim stepXP1         As Single
  Dim XPface          As Long

    'No errors
    If Height = 0 Then Exit Sub
    
    lastStat = curStat  'Set property
    TE = CurrText       'Caption

    With UserControl
        .Cls                                        'Clear control
        DrawRectangle 0, 0, Width, Height, cFace    'Draw button face
        
        If isEnabled = True Then            'If enabled
            SetTextColor .hDC, cText        'Set text colour
            
            If curStat = 0 Then             'If button is up
                
                 'Gradient step
                stepXP1 = 25 / Height
                'Shift color
                XPface = ShiftColor(cFace, &H30)
                'Draw gradient background
                For i = 1 To Height
                    DrawLine 0, i, Width, i, ShiftColor(XPface, -stepXP1 * i)
                Next
                'Set caption
                SetTextColor UserControl.hDC, cColor
                DrawText .hDC, CurrText, Len(CurrText), RC, DT_CENTERABS
                'Draw outline
                DrawLine 2, 0, Width - 2, 0, &H733C00
                DrawLine 2, Height - 1, Width - 2, Height - 1, &H733C00
                DrawLine 0, 2, 0, Height - 2, &H733C00
                DrawLine Width - 1, 2, Width - 1, Height - 2, &H733C00
                'Draw corners
                SetPixel UserControl.hDC, 1, 1, &H7B4D10
                SetPixel UserControl.hDC, 1, Height - 2, &H7B4D10
                SetPixel UserControl.hDC, Width - 2, 1, &H7B4D10
                SetPixel UserControl.hDC, Width - 2, Height - 2, &H7B4D10
                'Draw shadows
                DrawLine 2, Height - 2, Width - 2, Height - 2, ShiftColor(XPface, -&H30)
                DrawLine 1, Height - 3, Width - 2, Height - 3, ShiftColor(XPface, -&H20)
                DrawLine Width - 2, 2, Width - 2, Height - 2, ShiftColor(XPface, -&H24)
                DrawLine Width - 3, 3, Width - 3, Height - 3, ShiftColor(XPface, -&H18)
                'Draw highlights
                DrawLine 2, 1, Width - 2, 1, ShiftColor(XPface, &H10)
                DrawLine 1, 2, Width - 2, 2, ShiftColor(XPface, &HA)
                DrawLine 1, 2, 1, Height - 2, ShiftColor(XPface, -&H5)
                DrawLine 2, 3, 2, Height - 3, ShiftColor(XPface, -&HA)
            ElseIf curStat = 2 Then     'Button is down
                
                'Set gradient step
                stepXP1 = 15 / Height
                'Shift color
                XPface = ShiftColor(cFace, &H30)
                XPface = ShiftColor(XPface, -32)
                'Draw gradient background
                For i = 1 To Height
                    DrawLine 0, Height - i, Width, Height - i, ShiftColor(XPface, -stepXP1 * i)
                Next i
                'Draw caption
                SetTextColor .hDC, cColor
                DrawText .hDC, CurrText, Len(CurrText), RC2, DT_CENTERABS
                'Draw outline
                DrawLine 2, 0, Width - 2, 0, &H733C00
                DrawLine 2, Height - 1, Width - 2, Height - 1, &H733C00
                DrawLine 0, 2, 0, Height - 2, &H733C00
                DrawLine Width - 1, 2, Width - 1, Height - 2, &H733C00
                'Draw corners
                SetPixel UserControl.hDC, 1, 1, &H7B4D10
                SetPixel UserControl.hDC, 1, Height - 2, &H7B4D10
                SetPixel UserControl.hDC, Width - 2, 1, &H7B4D10
                SetPixel UserControl.hDC, Width - 2, Height - 2, &H7B4D10
                'Draw shadows
                DrawLine 2, Height - 2, Width - 2, Height - 2, ShiftColor(XPface, &H10)
                DrawLine 1, Height - 3, Width - 2, Height - 3, ShiftColor(XPface, &HA)
                DrawLine Width - 2, 2, Width - 2, Height - 2, ShiftColor(XPface, &H5)
                DrawLine Width - 3, 3, Width - 3, Height - 3, XPface
                'Draw highlights
                DrawLine 2, 1, Width - 2, 1, ShiftColor(XPface, -&H20)
                DrawLine 1, 2, Width - 2, 2, ShiftColor(XPface, -&H18)
                DrawLine 1, 2, 1, Height - 2, ShiftColor(XPface, -&H20)
                DrawLine 2, 2, 2, Height - 2, ShiftColor(XPface, -&H16)
            End If
        Else    'Disabled state
            
            'Shift color
            XPface = ShiftColor(cFace, &H30)
            'Draw button face
            DrawRectangle 0, 0, Width, Height, ShiftColor(XPface, -&H18)
            'Caption
            SetTextColor .hDC, ShiftColor(XPface, -&H68)
            DrawText .hDC, CurrText, Len(CurrText), RC, DT_CENTERABS
            'Draw outline
            DrawLine 2, 0, Width - 2, 0, ShiftColor(XPface, -&H54)
            DrawLine 2, Height - 1, Width - 2, Height - 1, ShiftColor(XPface, -&H54)
            DrawLine 0, 2, 0, Height - 2, ShiftColor(XPface, -&H54)
            DrawLine Width - 1, 2, Width - 1, Height - 2, ShiftColor(XPface, -&H54)
            'Draw corners
            SetPixel UserControl.hDC, 1, 1, ShiftColor(XPface, -&H48)
            SetPixel UserControl.hDC, 1, Height - 2, ShiftColor(XPface, -&H48)
            SetPixel UserControl.hDC, Width - 2, 1, ShiftColor(XPface, -&H48)
            SetPixel UserControl.hDC, Width - 2, Height - 2, ShiftColor(XPface, -&H48)
        End If
    End With
    
End Sub


'Draw rectangle
Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)

  Dim bRect As RECT
  Dim hBrush As Long
  Dim Ret As Long
    
    'Fill out rect
    bRect.Left = x
    bRect.Top = y
    bRect.Right = x + Width
    bRect.Bottom = y + Height
    
    'Create brush
    hBrush = CreateSolidBrush(Color)
    
    If OnlyBorder = False Then  'Just border
        Ret = FillRect(UserControl.hDC, bRect, hBrush)
    Else    'Fill whole rect
        Ret = FrameRect(UserControl.hDC, bRect, hBrush)
    End If
    
    'Delete brush
    Ret = DeleteObject(hBrush)
    
End Sub


'Draw line
Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

  Dim pt As POINTAPI

    UserControl.ForeColor = Color           'Set forecolor
    MoveToEx UserControl.hDC, X1, Y1, pt    'Move to X1/Y1
    LineTo UserControl.hDC, X2, Y2          'Draw line to X2/Y2
    
End Sub


'Set Colours
Private Sub SetColors()
    
    'Get system colours and save into variables
    cFace = RGB(200, 200, 255)
    
    '####################################
    '# cFace = GetSysColor(COLOR_BTNFACE)
    '####################################
    
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(COLOR_BTNTEXT)
    
End Sub


'Shift colors
Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long

  Dim Red As Long, Blue As Long, Green As Long
    
    'Shift blue
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * Value) \ &HC0)
    'Shift green
    Green = ((Color \ &H100) Mod &H100) + Value
    'Shift red
    Red = (Color And &HFF) + Value
    
    'Check red bounds
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'Check green bounds
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'Check blue bounds
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If
    
    'Return color
    ShiftColor = RGB(Red, Green, Blue)
  
End Function
