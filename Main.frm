VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   Caption         =   "Kokoro's Downloader 1.3     http://www.juegosenteros.cjb.net"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   FillColor       =   &H80000012&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":08CA
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   Begin VB.Frame frmDownload 
      BorderStyle     =   0  'None
      Height          =   7155
      Left            =   200
      TabIndex        =   1
      Top             =   720
      Width           =   6750
      Begin VB.PictureBox picLink 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   3375
         TabIndex        =   37
         Top             =   0
         Width           =   3375
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            Caption         =   "http://www.juegosenteros.cjb.net/"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            MouseIcon       =   "Main.frx":0F76
            MousePointer    =   99  'Custom
            TabIndex        =   38
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   2955
         End
      End
      Begin VB.Frame frmDown 
         Caption         =   "3rd Download"
         Height          =   2655
         Left            =   120
         TabIndex        =   13
         Top             =   4440
         Width           =   6495
         Begin Kokoro.Button cmdCancelDown 
            Height          =   375
            Left            =   5040
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            ForeColor       =   0
            TX              =   "Cancel"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Kokoro.Button cmdDownload 
            Height          =   375
            Left            =   3720
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            ForeColor       =   0
            TX              =   "Download"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Kokoro.Button cmdSelectAll 
            Height          =   375
            Left            =   960
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            ForeColor       =   0
            TX              =   "Select all"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3120
            Picture         =   "Main.frx":10C8
            ScaleHeight     =   300
            ScaleWidth      =   300
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   180
            Width           =   300
         End
         Begin MSComctlLib.ProgressBar ProgBar 
            Height          =   255
            Left            =   3540
            TabIndex        =   22
            Top             =   960
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.ListBox lstPackages 
            Height          =   1860
            ItemData        =   "Main.frx":1190
            Left            =   120
            List            =   "Main.frx":1192
            Style           =   1  'Checkbox
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label lblEstimatedTime 
            Caption         =   "Estimated Time:           00:00:00"
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label lblRate 
            Alignment       =   1  'Right Justify
            Caption         =   "0 %"
            Height          =   255
            Left            =   4680
            TabIndex        =   27
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Speed: 0 KBps"
            Height          =   255
            Left            =   3600
            TabIndex        =   26
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label lblElapsedTime 
            Caption         =   "Elapsed Time:              00:00:00"
            Height          =   255
            Left            =   3600
            TabIndex        =   25
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Label lblState 
            Caption         =   "State"
            Height          =   255
            Left            =   3600
            TabIndex        =   24
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label lblBytesDown 
            Alignment       =   2  'Center
            Caption         =   "0 from 0 bytes"
            Height          =   255
            Left            =   3520
            TabIndex        =   23
            Top             =   1560
            Width           =   2910
         End
         Begin VB.Line Line1 
            X1              =   3480
            X2              =   3480
            Y1              =   240
            Y2              =   2520
         End
      End
      Begin VB.Frame frmInfo 
         Caption         =   "2nd Info"
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6495
         Begin Kokoro.Button cmdNotes 
            Height          =   495
            Left            =   5760
            Top             =   1680
            Visible         =   0   'False
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            ForeColor       =   0
            TX              =   "Notes"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6120
            Picture         =   "Main.frx":1194
            ScaleHeight     =   300
            ScaleWidth      =   300
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   180
            Width           =   300
         End
         Begin VB.PictureBox picCircle 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   1680
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1150
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox picPreview 
            BackColor       =   &H8000000C&
            Height          =   2310
            Left            =   120
            ScaleHeight     =   2250
            ScaleWidth      =   3000
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   3060
         End
         Begin VB.Label lblCheckSearch 
            Caption         =   "Checking"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   17.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Index           =   1
            Left            =   2040
            TabIndex        =   20
            Top             =   1080
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.Label lblType 
            Caption         =   "******"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3960
            TabIndex        =   14
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Type:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   9
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSize 
            Caption         =   "*******"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3840
            TabIndex        =   16
            Top             =   1680
            UseMnemonic     =   0   'False
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Size:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   10
            Top             =   1680
            UseMnemonic     =   0   'False
            Width           =   615
         End
         Begin VB.Label lblWrongCode 
            Alignment       =   2  'Center
            Caption         =   "Wrong code or game not available"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   855
            Left            =   480
            TabIndex        =   7
            Top             =   1080
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.Label lblRequiredSoft 
            Caption         =   "********"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   400
            Left            =   1560
            TabIndex        =   18
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   4815
         End
         Begin VB.Label lblNumPack 
            Caption         =   "********"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   4800
            TabIndex        =   17
            Top             =   2280
            UseMnemonic     =   0   'False
            Width           =   615
         End
         Begin VB.Label lblName 
            Caption         =   "********"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   3960
            TabIndex        =   15
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "Required software:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2640
            UseMnemonic     =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Num of packages:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   2280
            UseMnemonic     =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   8
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   615
         End
      End
      Begin VB.Frame frmCode 
         Caption         =   "1st Insert Code"
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6495
         Begin Kokoro.Button cmdOK 
            Default         =   -1  'True
            Height          =   375
            Left            =   2520
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ForeColor       =   0
            TX              =   "OK"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Kokoro.Button cmdCancelInfo 
            Cancel          =   -1  'True
            Height          =   375
            Left            =   4080
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ForeColor       =   0
            TX              =   "Cancel"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6120
            Picture         =   "Main.frx":125C
            ScaleHeight     =   300
            ScaleWidth      =   300
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   180
            Width           =   300
         End
         Begin VB.TextBox txtCode 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   405
            Left            =   960
            MaxLength       =   5
            TabIndex        =   2
            Top             =   300
            Width           =   1365
         End
      End
   End
   Begin VB.Frame frmSearch 
      BorderStyle     =   0  'None
      Height          =   7125
      Left            =   150
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   6825
      Begin VB.PictureBox picCircle 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   1430
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2780
         Visible         =   0   'False
         Width           =   375
      End
      Begin Kokoro.Button cmdCancelSearch 
         Height          =   495
         Left            =   3120
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         ForeColor       =   0
         TX              =   "Cancel"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Kokoro.Button cmdSearch 
         Height          =   495
         Left            =   960
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         ForeColor       =   0
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5520
         Picture         =   "Main.frx":1324
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   32
         Top             =   240
         Width           =   300
      End
      Begin MSComctlLib.ListView lsvGameList 
         Height          =   6435
         Left            =   0
         TabIndex        =   28
         Top             =   645
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   11351
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblCheckSearch 
         Caption         =   "Searching"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Index           =   2
         Left            =   1875
         TabIndex        =   36
         Top             =   2580
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":22F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":26D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3284
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3634
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":39F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4568
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4930
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":50B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":548C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5850
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5C34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer_Link 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   3600
   End
   Begin VB.Timer Timer_Styles 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   6600
   End
   Begin Kokoro.Button cmdAbout 
      Height          =   255
      Left            =   5760
      Top             =   375
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ForeColor       =   16711680
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Kokoro.Button cmdPreferences 
      Height          =   255
      Left            =   4440
      Top             =   375
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ForeColor       =   16711680
      TX              =   "Preferences"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer_Points 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   80
      Left            =   0
      Top             =   6120
   End
   Begin VB.Timer Timer_Circle 
      Enabled         =   0   'False
      Index           =   2
      Left            =   0
      Top             =   5640
   End
   Begin VB.Timer Timer_Circle 
      Enabled         =   0   'False
      Index           =   1
      Left            =   0
      Top             =   4080
   End
   Begin VB.Timer Timer_Points 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   80
      Left            =   0
      Top             =   4560
   End
   Begin VB.Timer Timer_Speed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5040
   End
   Begin MSComctlLib.ImageList imlMiscellaneous 
      Left            =   0
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6418
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":67FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":7168
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":748C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":77B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":7ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":7E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":812C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8450
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8774
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8C24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7575
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   13361
      ShowTips        =   0   'False
      TabMinWidth     =   3682
      ImageList       =   "imlMiscellaneous"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Download"
            ImageVarType    =   2
            ImageIndex      =   13
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageVarType    =   2
            ImageIndex      =   14
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgHelp 
      Height          =   225
      Left            =   6045
      Picture         =   "Main.frx":913C
      ToolTipText     =   "Help"
      Top             =   45
      Width           =   225
   End
   Begin VB.Image imgRestore 
      Height          =   225
      Left            =   6600
      Picture         =   "Main.frx":944E
      ToolTipText     =   "Maximize/Restore"
      Top             =   45
      Width           =   225
   End
   Begin VB.Image imgTitleIcon 
      Height          =   240
      Left            =   90
      Picture         =   "Main.frx":9760
      Top             =   50
      Width           =   240
   End
   Begin VB.Image imgMinimize 
      Height          =   225
      Left            =   6360
      Picture         =   "Main.frx":9B78
      ToolTipText     =   "Minimize"
      Top             =   45
      Width           =   225
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   6840
      Picture         =   "Main.frx":9E8A
      ToolTipText     =   "Close"
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgTitleBar 
      Height          =   300
      Left            =   360
      Top             =   0
      Width           =   5670
   End
   Begin VB.Label lblTitleBar 
      BackStyle       =   0  'Transparent
      Caption         =   "Kokoro's Downloader 1.3    http://www.juegosenteros.cjb.net/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   30
      Width           =   5535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

'Name.......... Kokoro's Downloader
'Version....... 1.3
'Description... Game downloader
'Author........ Emiliano Scavuzzo <anshoku@yahoo.com>
'Date.......... June, 28th 2004

'Copyright (c) 2004 by Emiliano Scavuzzo
'Rosario, Argentina
'
'Credits:
'Button Control:    Leo Barsukov
'Tooltip class:     Kaustubh Zoal
'Translucent code:  ArthurW <arthurw@digitelone.com>


'*********** Info files format ***********
' 1) 'FullGames Info' (see INFO_VERIFICATION_LINE)
' 2) Update file address
' 2) Preview image address
' 3) Game name
' 4) Type of game
' 5) Game total size in bytes
' 6) Amount of packages
' 7) Required software
' 8) Kind of download (KOD)
' 9) Notes
' 10-X) KOD linesOption Explicit

Private WithEvents Download_List As CDownload
Attribute Download_List.VB_VarHelpID = -1
Private WithEvents Download_Update As CDownload
Attribute Download_Update.VB_VarHelpID = -1
Private WithEvents Download_Preview As CDownload
Attribute Download_Preview.VB_VarHelpID = -1
Private WithEvents Download_Info As CDownload
Attribute Download_Info.VB_VarHelpID = -1


Private Sub cmdOK_Click()
Start_Check
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModeless, Me
End Sub

Private Sub cmdNotes_Click()
Unload frmNotes
frmNotes.Show vbModeless, Me
End Sub

Private Sub Restore()
If m_blnCompactWindow Then
    frmInfo.Visible = True
    frmDown.Top = 2 * frmInfo.Top + frmInfo.Height - frmCode.Top - frmCode.Height
    frmDownload.Height = 477
    frmSearch.Height = 477
    TabStrip1.Height = 505
    lsvGameList.Height = 6435
    lblCheckSearch(2).Top = 2580
    picCircle(2).Top = 2780
    Me.Height = 8430
    m_blnCompactWindow = False
Else
    frmInfo.Visible = False
    frmDown.Top = frmInfo.Top
    frmDownload.Height = 260
    frmSearch.Height = 260
    TabStrip1.Height = 290
    lsvGameList.Height = 3300
    lblCheckSearch(2).Top = lblCheckSearch(2).Top - 800
    picCircle(2).Top = picCircle(2).Top - 800
    Me.Height = 5200
    
    m_blnCompactWindow = True
End If
DrawForm
End Sub

Private Sub Form_Load()
'initiate variables and others
Me.Height = Screen.TwipsPerPixelX * (TabStrip1.Top + TabStrip1.Height + 34)
DrawForm
CenterWindowOnScreen
CreateMenu
Set Download_List = New CDownload
Set Download_Update = New CDownload
Set Download_Preview = New CDownload
Set Download_Info = New CDownload
Download_Update.DownloadType = dtToBuffer
Download_Info.DownloadType = dtToBuffer
Download_List.UnEncrypt = True
Download_Info.UnEncrypt = True
Initiations
MakeTempFoldIfDoesntExist
MakeDownFoldIfDoesntExist
End Sub

'if it is downloading asks you if you want to close
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If m_blnIsDownloading = True Then
    Dim intCloseResp As Integer
    intCloseResp = MessageBox(hwnd, "A game is being downloaded." + vbCrLf + "¿Close anyway?", "Close program", MB_YESNO Or MB_ICONQUESTION Or MB_DEFBUTTON2)
    If intCloseResp = IDNO Then Cancel = True
End If
End Sub

Private Sub cmdCancelInfo_Click()
Stop_Check
End Sub

Private Sub cmdSearch_Click()
DisableSearch

Timer_Styles.Enabled = False
ReDim FlashList(0 To 0)
ReDim ColorList(0 To 0)

CleanGameList
MakeListIfDoesntExist

lsvGameList.Visible = False
ShowSearching
Timer_Points(2).Enabled = True
Timer_Circle(2).Enabled = True

m_blnListFirstTry = True
If m_blnUpdateCheckNeeded And m_blnUpdating = False And m_blnUseOwnServer = False Then
    m_blnUpdateByInfo = False
    ConnectUpdateSocket
    Exit Sub
End If
ConnectSearchSocket
End Sub

'Connect the socket to the game list server
Private Sub ConnectSearchSocket()
If m_blnUseProxy = True Then
    Download_List.AccessType = cdNamedProxy
    Download_List.Proxy = m_strProxy
    Download_List.ProxyPort = m_lngPort
Else
    Download_List.AccessType = cdDirect
End If
Download_List.Download TakeHost(ListServer) + "/" + GameListName, m_strGameListPath
End Sub

'Start code checking routine
Public Sub Start_Check()
With frmMain
    If .txtCode.Text = "" Then
        MessageBox .hwnd, "You must insert a game code!", "Error", MB_ICONEXCLAMATION
        SetThisFocus .txtCode.hwnd
        Exit Sub
    End If

    DisableOne
    HideUno
    HideWrongCode
    .lstPackages.Clear
    ShowChecking

    .Timer_Points(1).Enabled = True
    .Timer_Circle(1).Enabled = True
    
    m_blnInfoFirstTry = True
    If m_blnUpdateCheckNeeded And m_blnUpdating = False And m_blnUseOwnServer = False Then
        m_blnUpdateByInfo = True
        ConnectUpdateSocket
        Exit Sub
    End If
End With
ConnectCheckSocket
End Sub

'Connect the socket to the code checking server
Public Sub ConnectCheckSocket()
If m_blnUseProxy = True Then
    Download_Info.AccessType = cdNamedProxy
    Download_Info.Proxy = m_strProxy
    Download_Info.ProxyPort = m_lngPort
Else
    Download_Info.AccessType = cdDirect
End If
Download_Info.Download TakeHost(InfoServer) + "/" + LCase(txtCode) + InfoExtension
End Sub

'Start update seeking routine
Public Sub ConnectUpdateSocket()
Debug.Print "Looking for updates in "; m_strUpdateServer
m_blnUpdating = True
m_blnUpdateCheckNeeded = False
If m_blnUseProxy = True Then
    Download_Update.AccessType = cdNamedProxy
    Download_Update.Proxy = m_strProxy
    Download_Update.ProxyPort = m_lngPort
Else
    Download_Update.AccessType = cdDirect
End If
Download_Update.Download m_strUpdateServer
End Sub

'Stop checking routine and clean screen to let you
'insert a new code.
Public Sub Stop_Check()
Download_Update.Cancel
If m_blnUpdating And m_blnUpdateByInfo Then
    m_blnUpdating = False
    m_blnUpdateCheckNeeded = True
End If
Download_Info.Cancel
Download_Preview.Cancel
CleanInfoScreen
HideWrongCode
HideChecking
ShowOne
EnableForInfo
End Sub

Private Sub cmdCancelSearch_Click()
Download_Update.Cancel
If m_blnUpdating And Not m_blnUpdateByInfo Then
    m_blnUpdating = False
    m_blnUpdateCheckNeeded = True
End If
Download_List.Cancel
CleanGameList
HideSearching
frmMain.lsvGameList.Visible = True
EnableSearch
End Sub

Private Sub cmdDownload_Click()
cmdOK.Enabled = False
cmdCancelInfo.Enabled = False
txtCode.Enabled = False
lstPackages.Enabled = False
cmdSelectAll.Enabled = False
cmdDownload.Enabled = False
cmdCancelDown.Enabled = True
m_blnIsDownloading = True
m_clsDownload.Download
End Sub

Private Sub cmdCancelDown_Click()
Dim intRespStop As Integer
intRespStop = MessageBox(hwnd, "¿Do you wish to cancel current download?", "Stop download", MB_YESNO Or MB_ICONQUESTION Or MB_DEFBUTTON2)
If intRespStop = IDYES Then
    m_clsDownload.StopDownload
    Finish_Download_Actions
End If
End Sub

Private Sub cmdPreferences_Click()
frmPreferences.Show vbModeless, Me
End Sub

Private Sub cmdSelectAll_Click()
Dim intCounter As Integer
For intCounter = 0 To lstPackages.ListCount - 1
    lstPackages.Selected(intCounter) = True
Next intCounter
lstPackages.Refresh
End Sub

Private Sub imgHelp_Click()
ShellExecute Me.hwnd, "Open", App.Path + BarIfIsNotRoot + "kokoro.hlp", vbNullString, App.Path, 1
End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgHelp.Picture = imlMiscellaneous.ListImages(12).Picture
End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgHelp.Picture = imlMiscellaneous.ListImages(11).Picture
End Sub

Private Sub imgTitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
m_pntMousePos.x = x
m_pntMousePos.y = y
End Sub

Private Sub imgTitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Me.Left = Me.Left + x - m_pntMousePos.x
    Me.Top = Me.Top + y - m_pntMousePos.y
End If
End Sub

'Close cross
Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgClose.Picture = imlMiscellaneous.ListImages(4).Picture
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgClose.Picture = imlMiscellaneous.ListImages(3).Picture
End Sub

Private Sub imgMinimize_Click()
SendMessage hwnd, WM_SYSCOMMAND, SC_MINIMIZE, vbNull
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgMinimize.Picture = imlMiscellaneous.ListImages(6).Picture
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
imgMinimize.Picture = imlMiscellaneous.ListImages(5).Picture
End Sub

Private Sub imgRestore_Click()
If m_blnCompactWindow Then
    imgRestore.Picture = imlMiscellaneous.ListImages(7).Picture
Else
    imgRestore.Picture = imlMiscellaneous.ListImages(9).Picture
End If
Restore
End Sub

Private Sub imgRestore_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If m_blnCompactWindow Then
    imgRestore.Picture = imlMiscellaneous.ListImages(10).Picture
Else
    imgRestore.Picture = imlMiscellaneous.ListImages(8).Picture
End If
End Sub

Private Sub imgRestore_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If m_blnCompactWindow Then
    imgRestore.Picture = imlMiscellaneous.ListImages(9).Picture
Else
    imgRestore.Picture = imlMiscellaneous.ListImages(7).Picture
End If
End Sub

Private Sub lblLink_Click()
ShellExecute hwnd, "Open", WEB_ADDRESS, "", App.Path, 1
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
With lblLink
    If .ForeColor = COLOR_INACTIVE_LINK Then
        .ForeColor = COLOR_ACTIVE_LINK
        Timer_Link.Interval = 1
        Timer_Link.Enabled = True
    End If
End With
End Sub

'If you click outside the list unselect the elements
Private Sub lsvGameList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
m_intLastMouseButton = Button
lsvGameList.SelectedItem = Nothing
End Sub

Private Sub lsvGameList_DblClick()
If Not lsvGameList.SelectedItem Is Nothing And m_intLastMouseButton = vbLeftButton Then
    If m_blnIsDownloading = True Then
    MessageBox hwnd, "You are already downloading a game.", "Error", MB_ICONEXCLAMATION
    Else
    Stop_Check
    txtCode.Text = lsvGameList.ListItems(lsvGameList.SelectedItem.Index).ListSubItems(1)
    txtCode.SelStart = Len(txtCode.Text)
    ChangeToTab 1
    Start_Check
    End If
End If
End Sub

'If there isnt's any package selected disable
'download button
Private Sub lstPackages_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If lstPackages.SelCount = 0 Then
    cmdDownload.Enabled = False
Else
    cmdDownload.Enabled = True
End If
End Sub

'If there isnt's any package selected disable
'download button
Private Sub lstPackages_ItemCheck(Item As Integer)
If lstPackages.SelCount = 0 Then
    cmdDownload.Enabled = False
Else
    cmdDownload.Enabled = True
End If
End Sub

'The game info wasn't fount by Download_Info
Private Sub Game_Not_Found()
HideUno
HideChecking
ShowWrongCode

cmdSelectAll.Enabled = False
cmdDownload.Enabled = False

lstPackages.Clear
EnableForInfo
If WhichTab = 1 Then SetThisFocus txtCode.hwnd
End Sub

'The game info was fount by Download_Info
Private Sub Game_Found()
lstPackages.Clear
'check the don't download image option
If m_blnDontDownPrev = True Then
    Preview_Finished (False)
Else
    m_strPreviewAddress = Trim(ReadLine(m_strInfo, LN_PREV_ADDR))
    MakePreviewIfDoesntExist
    If m_blnUseProxy = True Then
        Download_Preview.AccessType = cdNamedProxy
        Download_Preview.Proxy = m_strProxy
        Download_Preview.ProxyPort = m_lngPort
    Else
        Download_Preview.AccessType = cdDirect
    End If
    Download_Preview.Download m_strPreviewAddress, m_strPreviewPath
End If
End Sub

'When the preview image is downloaded (or not) this sub
'is called
Private Sub Preview_Finished(ByVal blnPreviewFound As Boolean)
On Error GoTo ErrorInfo

'if the image was downloaded it shows it, if not it cleans the picturebox
If blnPreviewFound = True Then
    LoadPreview
Else
    frmMain.picPreview.Picture = LoadPicture
End If

CheckInfo

If m_blnUseOwnServer = False Then
    If m_strUpdateServer <> Trim(ReadLine(m_strInfo, LN_UPD_SRV)) Then
        m_strUpdateServer = Trim(ReadLine(m_strInfo, LN_UPD_SRV))
        Call savestring(HKEY_LOCAL_MACHINE, "Software\" + APP_NAME, m_regUpdateServer, m_strUpdateServer)
        m_blnUpdateCheckNeeded = True
    End If
End If

lblName.Caption = Trim(ReadLine(m_strInfo, LN_NAME))
lblType.Caption = Trim(ReadLine(m_strInfo, LN_TYPE))
lblSize.Caption = Bytes2KB_MB(Trim(ReadLine(m_strInfo, LN_TOT_SIZE)))
lblNumPack.Caption = Trim(ReadLine(m_strInfo, LN_PACK_NUM))
lblRequiredSoft.Caption = Trim(ReadLine(m_strInfo, LN_SOFTWARE))

HideChecking
ShowOne

If BuildDownloadClass = False Then Exit Sub

m_clsDownload.CheckPackages
m_clsDownload.ReadPackages

HideWrongCode
EnableForInfo

If Trim(ReadLine(m_strInfo, LN_NOTES)) <> "" Then cmdNotes.Visible = True

If WhichTab = 1 Then SetThisFocus txtCode.hwnd
If lstPackages.ListCount > 0 Then
    cmdSelectAll.Enabled = True
    lstPackages.ListIndex = 0
End If

Exit Sub
ErrorInfo:
    ErroneousInfo
End Sub

'The update was found by Download_Update
Private Sub Update_Found()
On Error GoTo ErrorHandler
Dim strLine As String
Const SOFTWARE As String = "Software\"

strLine = ReadLine(m_strUpdate, LN_VER)

If IsNumeric(strLine) Then
    If Int(strLine) > VERSION Then
        m_blnUpdateCheckNeeded = True
        m_blnUpdating = False
        UpdateNeededForVersion
        Exit Sub
    End If
End If

strLine = ReadLine(m_strUpdate, LN_MAINSRV)

If IsURL(DataFromLine(strLine, 1)) Then
    m_strDatabase = Trim(DataFromLine(strLine, 1))
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regDatabaseMain, m_strDatabase)
End If

If Trim(DataFromLine(strLine, 2)) <> "" Then
    m_strExtensionInfoMain = Trim(DataFromLine(strLine, 2))
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regExtInfoMain, m_strExtensionInfoMain)
End If

If Trim(DataFromLine(strLine, 3)) <> "" Then
    m_strGameListNameMain = Trim(DataFromLine(strLine, 3))
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regListNameMain, m_strGameListNameMain)
End If

strLine = ReadLine(m_strUpdate, LN_BKPSRV)

If IsURL(DataFromLine(strLine, 1)) Then
    m_strDatabaseBkp = Trim(DataFromLine(strLine, 1))
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regDatabaseBkp, m_strDatabaseBkp)
End If

If Trim(DataFromLine(strLine, 2)) <> "" Then
    m_strExtensionInfoBkp = Trim(DataFromLine(strLine, 2))
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regExtInfoBkp, m_strExtensionInfoBkp)
End If

If Trim(DataFromLine(strLine, 3)) <> "" Then
    m_strGameListNameBkp = Trim(DataFromLine(strLine, 3))
    Call savestring(HKEY_LOCAL_MACHINE, SOFTWARE + APP_NAME, m_regListNameBkp, m_strGameListNameBkp)
End If

m_strUpdate = ""
Update_Finished

Exit Sub
ErrorHandler:
    Update_Finished
    Exit Sub
End Sub

'When the update has finished this sub is called which
'returns the control to the system that asked for the
'update.
Private Sub Update_Finished()
m_blnUpdating = False
If m_blnUpdateByInfo Then
    ConnectCheckSocket
Else
    ConnectSearchSocket
End If
End Sub

'Control the game list sorting
Private Sub lsvGameList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

Static blnFirstAscen As Boolean
Static blnSecondAscen As Boolean
Static blnThirdAscen As Boolean

Select Case ColumnHeader

Case m_strHeadListColumn1
    lsvGameList.SortKey = 0
    If blnFirstAscen = True Then
        lsvGameList.SortOrder = lvwDescending
        blnFirstAscen = False
    Else
        lsvGameList.SortOrder = lvwAscending
        blnFirstAscen = True
    End If

Case m_strHeadListColumn2
    lsvGameList.SortKey = 1
    If blnSecondAscen = True Then
        lsvGameList.SortOrder = lvwDescending
        blnSecondAscen = False
    Else
        lsvGameList.SortOrder = lvwAscending
        blnSecondAscen = True
    End If

Case m_strHeadListColumn3
    lsvGameList.SortKey = 2
    If blnThirdAscen = True Then
        lsvGameList.SortOrder = lvwDescending
        blnThirdAscen = False
    Else
        lsvGameList.SortOrder = lvwAscending
        blnThirdAscen = True
    End If

Case m_strHeadListColumn4
    OrderBySize
    Exit Sub

End Select

lsvGameList.Sorted = True
lsvGameList.Sorted = False
End Sub

'Control of tabstrip
Private Sub TabStrip1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
ChangeToTab TabStrip1.SelectedItem.Index
End Sub

'Control flash colors styles
Private Sub Timer_Styles_Timer()
Static blnIsFirst As Boolean
Dim intCounter As Integer

If blnIsFirst = True Then
    For intCounter = 1 To UBound(FlashList)
        With FlashList(intCounter)
            ListColorKey .Key, .Second.Color, .Second.Bold
        End With
    Next intCounter
    blnIsFirst = False
Else
    For intCounter = 1 To UBound(FlashList)
        With FlashList(intCounter)
            ListColorKey .Key, .First.Color, .First.Bold
        End With
    Next intCounter
    blnIsFirst = True
End If

End Sub

Private Sub Timer_Link_Timer()
Dim pt As POINTAPI
Dim lngX As Long
Dim lngY As Long
  'check if the mouse pointer is over the link
With picLink
    GetCursorPos pt
    ScreenToClient .hwnd, pt
    lngX = pt.x * Screen.TwipsPerPixelX
    lngY = pt.y * Screen.TwipsPerPixelY
    If (lngX < 0) Or (lngX > .Width) Or (lngY < 0) Or (lngY > .Height) Then
        lblLink.ForeColor = COLOR_INACTIVE_LINK
        Timer_Link.Enabled = False
    End If
End With
End Sub

'Control dots
Private Sub Timer_Points_Timer(intIndex As Integer)
Dim strTitle As String
Select Case intIndex
    Case 1
        strTitle = CHECKING_STRING
    Case 2
        strTitle = SEARCHING_STRING
End Select
If lblCheckSearch(intIndex).Caption = strTitle + "............." Then lblCheckSearch(intIndex).Caption = strTitle
lblCheckSearch(intIndex).Caption = lblCheckSearch(intIndex).Caption + "."
End Sub

'Control circle
Private Sub Timer_Circle_Timer(intIndex As Integer)
If picCircle(intIndex).Picture = imlMiscellaneous.ListImages(1).Picture Then
    picCircle(intIndex).Picture = imlMiscellaneous.ListImages(2).Picture
    Timer_Circle(intIndex).Interval = 500
Else
    picCircle(intIndex).Picture = imlMiscellaneous.ListImages(1).Picture
    Timer_Circle(intIndex).Interval = 1000
End If
End Sub

'Timer that controls the state.
'NOTE: make sure PackageSize and AmountDownloaded return
'the right data before enabling this timer
Private Sub Timer_Speed_Timer()
'progress bar
ProgBar.Max = m_clsDownload.PackageSize
ProgBar.Value = m_clsDownload.AmountDownloaded
'porcentaje
If m_clsDownload.PackageSize <> 0 Then lblRate.Caption = Trim(Str(Int(m_clsDownload.AmountDownloaded * 100 / m_clsDownload.PackageSize))) + " %"
'amount downloaded
lblBytesDown.Caption = Trim(CStr(Format(m_clsDownload.AmountDownloaded, "###,###,##0"))) + " from " + Trim(CStr(Format(m_clsDownload.PackageSize, "###,###,##0"))) + " bytes"
'elapsed time
Dim datHours As Date
datHours = Format(Right(lblElapsedTime.Caption, 8), "hh:mm:ss")
lblElapsedTime.Caption = "Elapsed Time:              " + Format(DateAdd("s", 1, datHours), "hh:mm:ss")
'speed
If Speed.QuanLoaded < UBound(Speed.Quantity) Then
    Speed.QuanLoaded = Speed.QuanLoaded + 1
Else
    Dim bytCounter1 As Byte
    For bytCounter1 = 1 To UBound(Speed.Quantity) - 1
        Speed.Quantity(bytCounter1) = Speed.Quantity(bytCounter1 + 1)
    Next bytCounter1
End If
Speed.Quantity(Speed.QuanLoaded) = m_clsDownload.AmountDownloaded - Speed.LastQuantity
Speed.LastQuantity = m_clsDownload.AmountDownloaded

Dim lngTransferred As Long 'sum the bytes downloaded in the last n seconds (n = Speed.QuanLoaded)
Dim bytCounter2 As Byte
For bytCounter2 = 1 To Speed.QuanLoaded
    lngTransferred = lngTransferred + Speed.Quantity(bytCounter2)
Next
lblSpeed.Caption = "Speed: " + Format(lngTransferred / 1024 / Speed.QuanLoaded, "0.00") + " KBps"
'estimated time
Dim dblSecondsLeft As Double
If lngTransferred <> 0 And (lngTransferred / 1024 / Speed.QuanLoaded) >= 0.01 Then
    dblSecondsLeft = (m_clsDownload.PackageSize - m_clsDownload.AmountDownloaded) \ (lngTransferred / Speed.QuanLoaded)
    lblEstimatedTime.Caption = "Estimated Time:           " + SecondsToHours(dblSecondsLeft)
Else
    lblEstimatedTime.Caption = "Estimated Time:           ??:??:??"
End If
End Sub

Private Sub Download_Update_Starting(ByVal FileSize As Long, ByVal Header As String, ByVal FileHandle As Integer)
m_strUpdate = ""
End Sub

Private Sub Download_Update_Completed()
m_strUpdate = Download_Update.GetBuffer

Dim intCheck As Integer
intCheck = InStr(1, m_strUpdate, UPDATE_VERIFICATION_LINE, vbBinaryCompare)
If intCheck <> 0 Then
    Update_Found
Else
    Update_Finished
End If
End Sub

Private Sub Download_Update_Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
Update_Finished
End Sub

Private Sub Download_Info_Starting(ByVal FileSize As Long, ByVal Header As String, ByVal FileHandle As Integer)
m_strInfo = ""
End Sub

Private Sub Download_Info_Completed()
m_strInfo = Download_Info.GetBuffer

Dim intCheck As Integer
intCheck = InStr(1, m_strInfo, INFO_VERIFICATION_LINE, vbBinaryCompare)
If intCheck = 0 Or intCheck = Null Then
    If (m_blnDontUseBackup = True) Or (m_blnInfoFirstTry = False) Or (m_blnUseOwnServer = True) Then
        Game_Not_Found
    Else
        m_blnInfoFirstTry = False
        ConnectCheckSocket
    End If
Else
    Game_Found
End If

End Sub

Private Sub Download_Info_Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
If (m_blnDontUseBackup = True) Or (m_blnInfoFirstTry = False) Or (m_blnUseOwnServer = True) Then
    HideChecking
    CleanInfoScreen
    ShowOne
    Select Case (Number)
    Case 11001:
        MessageBox hwnd, "Could not locate remote server. Check your internet connection.", "Error", MB_ICONERROR
    Case 10060:
        MessageBox hwnd, "The connection time-out has expired. Try later.", "Error", MB_ICONERROR
    Case Else
        MessageBox hwnd, Description, "Error", MB_ICONERROR
    End Select
    EnableForInfo
    If WhichTab = 1 Then SetThisFocus txtCode.hwnd
Else 'if you enabled the backup server and it isn't the first try
    m_blnInfoFirstTry = False
    ConnectCheckSocket
End If
End Sub

Private Sub Download_Preview_Completed()
Preview_Finished (True)
End Sub

Private Sub Download_Preview_Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
Preview_Finished (False)
End Sub

Private Sub Download_List_Completed()
Dim intFileHandle As Integer
intFileHandle = FreeFile
Dim strReadLine As String

On Error GoTo ErrorOnOpen
Open m_strGameListPath For Input Lock Read Write As intFileHandle
On Error GoTo ErrorOnRead
Line Input #intFileHandle, strReadLine

If strReadLine = LIST_VERIFICATION_LINE Then 'if it found the file
    Line Input #intFileHandle, strReadLine
    Close intFileHandle
    If IsURL(strReadLine) Then 'if it is a valid URL
        If m_blnUseOwnServer = False Then
            If m_strUpdateServer <> Trim(strReadLine) Then
                m_strUpdateServer = Trim(strReadLine)
                Call savestring(HKEY_LOCAL_MACHINE, "Software\" + APP_NAME, m_regUpdateServer, m_strUpdateServer)
                m_blnUpdateCheckNeeded = True
            End If
        End If
        ListFound
    Else 'if it is an invalid URL it raises an error
        Err.Raise 1
    End If
Else 'if it didn't find the file
    Close intFileHandle
    If (m_blnDontUseBackup = True) Or (m_blnListFirstTry = False) Or (m_blnUseOwnServer = True) Then
        ListNoFound
    Else
        m_blnListFirstTry = False
        MakeListIfDoesntExist
        ConnectSearchSocket
    End If
End If

Exit Sub
ErrorOnOpen:
    Close intFileHandle
    ErrorAccessingList
    Exit Sub
ErrorOnRead:
    Close intFileHandle
    ErrorReadingList
    Exit Sub
End Sub


Private Sub Download_List_Error(ByVal Number As Integer, Description As String, SocketError As Boolean)
If SocketError Then

    If (m_blnDontUseBackup = True) Or (m_blnListFirstTry = False) Or (m_blnUseOwnServer = True) Then
        HideSearching
        frmMain.lsvGameList.Visible = True
        Select Case (Number)
        Case 11001:
            MessageBox hwnd, "Could not locate remote server. Check your internet connection.", "Error", MB_ICONERROR
        Case 10060:
            MessageBox hwnd, "The connection time-out has expired. Try later.", "Error", MB_ICONERROR
        Case Else
            MessageBox hwnd, Description, "Error", MB_ICONERROR
        End Select
    
        EnableSearch
    
    Else
        m_blnListFirstTry = False
        MakeListIfDoesntExist
        ConnectSearchSocket
    End If

Else

    CleanGameList
    HideSearching
    lsvGameList.Visible = True
    MessageBox Me.hwnd, Description, "Error", MB_ICONERROR
    EnableSearch
    
End If
End Sub
