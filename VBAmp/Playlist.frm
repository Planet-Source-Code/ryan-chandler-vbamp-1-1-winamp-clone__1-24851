VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Playlist 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   2715
   ClientTop       =   1170
   ClientWidth     =   5820
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2760
      Top             =   4575
   End
   Begin MSComctlLib.ListView SecondsList 
      Height          =   1020
      Left            =   1200
      TabIndex        =   50
      Top             =   5340
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seconds"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column2"
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.PictureBox ExitD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3960
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   49
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Exit2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3960
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   48
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ShadeD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3825
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   47
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Shade2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3825
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   46
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ListBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   3420
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   27
      Top             =   2925
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox ListMen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   3465
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   26
      Top             =   2925
      Visible         =   0   'False
      Width           =   330
      Begin VB.PictureBox ButtonLst 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   45
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonLst 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   44
         Top             =   270
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonLst 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox RemBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   600
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   19
      Top             =   2655
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox MisBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   1470
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   18
      Top             =   2925
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox SelBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   1035
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   17
      Top             =   2925
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox AddBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   165
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   16
      Top             =   2925
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox MisMen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   1515
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   14
      Top             =   2925
      Visible         =   0   'False
      Width           =   330
      Begin VB.PictureBox ButtonMis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   42
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonMis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonMis 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox SelMen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   1080
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   13
      Top             =   2925
      Visible         =   0   'False
      Width           =   330
      Begin VB.PictureBox ButtonSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   39
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   38
         Top             =   270
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox AddMen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   210
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   12
      Top             =   2925
      Visible         =   0   'False
      Width           =   330
      Begin VB.PictureBox ButtonAdd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   31
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonAdd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   30
         Top             =   270
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonAdd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox RemMen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   645
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   15
      Top             =   2655
      Visible         =   0   'False
      Width           =   330
      Begin VB.PictureBox ButtonRem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   36
         Top             =   810
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonRem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   35
         Top             =   540
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonRem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   34
         Top             =   270
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox ButtonRem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin MSComctlLib.ListView Playlist2 
      Height          =   1020
      Left            =   1200
      TabIndex        =   20
      Top             =   5025
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SongName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   1905
      EndProperty
   End
   Begin MSComDlg.CommonDialog M3UOpener 
      Left            =   4170
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      InitDir         =   "C:\Program Files\Napster\Music\"
   End
   Begin VB.PictureBox List 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3465
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   25
      Top             =   3465
      Width           =   330
   End
   Begin MSComctlLib.ListView Playlist5 
      Height          =   1020
      Left            =   -15
      TabIndex        =   24
      Top             =   5340
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SongName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   1905
      EndProperty
   End
   Begin MSComctlLib.ListView Playlist4 
      Height          =   1020
      Left            =   3630
      TabIndex        =   23
      Top             =   5025
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SongName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   1905
      EndProperty
   End
   Begin MSComctlLib.ListView Playlist3 
      Height          =   1020
      Left            =   2415
      TabIndex        =   22
      Top             =   5025
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SongName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   1905
      EndProperty
   End
   Begin VB.PictureBox Misc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1515
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   11
      Top             =   3465
      Width           =   330
   End
   Begin VB.PictureBox Select 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1080
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   10
      Top             =   3465
      Width           =   330
   End
   Begin VB.PictureBox Add 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   210
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   8
      Top             =   3465
      Width           =   330
   End
   Begin VB.PictureBox Remove 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   645
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   9
      Top             =   3465
      Width           =   330
   End
   Begin MSComctlLib.ListView Playlist 
      Height          =   1020
      Left            =   -15
      TabIndex        =   0
      Top             =   5025
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SongName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   1905
      EndProperty
   End
   Begin VB.PictureBox Exit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3960
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   6
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox Shade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3825
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   5
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox BotmBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      Picture         =   "Playlist.frx":0000
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   4
      Top             =   3345
      Width           =   4125
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3375
      Top             =   5235
   End
   Begin VB.PictureBox PLEdit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   90
      Picture         =   "Playlist.frx":0046
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   1
      Top             =   4245
      Visible         =   0   'False
      Width           =   4200
      Begin VB.PictureBox BASEPLEdit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2790
         Left            =   570
         Picture         =   "Playlist.frx":5CC0
         ScaleHeight     =   186
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   280
         TabIndex        =   7
         Top             =   -2775
         Visible         =   0   'False
         Width           =   4200
      End
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   0
      Picture         =   "Playlist.frx":B93A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   32
      Top             =   0
      Width           =   4125
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "Playlist.frx":B980
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   2
      Top             =   0
      Width           =   4125
   End
   Begin VB.PictureBox ListView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   180
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   244
      TabIndex        =   53
      Top             =   300
      Width           =   3660
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   68
         Top             =   45
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   67
         Top             =   240
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   66
         Top             =   435
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   65
         Top             =   630
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   64
         Top             =   825
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   63
         Top             =   2775
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   62
         Top             =   2580
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   61
         Top             =   2385
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   60
         Top             =   2190
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   59
         Top             =   1995
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   58
         Top             =   1800
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   57
         Top             =   1605
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   56
         Top             =   1020
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   55
         Top             =   1215
         Width           =   405
      End
      Begin VB.PictureBox Time 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   3195
         ScaleHeight     =   195
         ScaleWidth      =   405
         TabIndex        =   54
         Top             =   1410
         Width           =   405
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   83
         Top             =   45
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   69
         Top             =   2775
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   70
         Top             =   2580
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   71
         Top             =   2385
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   72
         Top             =   2190
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   73
         Top             =   1995
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   74
         Top             =   1800
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   75
         Top             =   1605
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   76
         Top             =   1410
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   77
         Top             =   1215
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   78
         Top             =   1020
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   79
         Top             =   825
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   80
         Top             =   630
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   81
         Top             =   435
         Width           =   3630
      End
      Begin VB.PictureBox Item 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   3630
         TabIndex        =   82
         Top             =   240
         Width           =   3630
      End
   End
   Begin VB.PictureBox PlaylistSlider 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   3750
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   21
      Top             =   300
      Width           =   375
   End
   Begin VB.PictureBox LeftBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3915
      Left            =   0
      Picture         =   "Playlist.frx":B9C6
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label TestLabel3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4305
      TabIndex        =   52
      Top             =   1905
      Width           =   45
   End
   Begin VB.Label TestLabel2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5430
      TabIndex        =   51
      Top             =   885
      Width           =   45
   End
   Begin VB.Label TestLabel 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4650
      TabIndex        =   28
      Top             =   1815
      Width           =   45
   End
End
Attribute VB_Name = "Playlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OrigX2, OrigY2

Function XTwips(Pixels) As Integer

    'convert pixels to twips
    XTwips = Pixels * Screen.TwipsPerPixelX

End Function

Function YTwips(Pixels) As Integer

    'convert pixels to twips
    YTwips = Pixels * Screen.TwipsPerPixelY

End Function

Private Sub Add_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Add_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Menus(True) = False Then
        AddMenu = True
    Else
        Menus False
    End If
End Sub

Private Sub AddBar_Click()
    If Menus(True) = True Then Menus (False)
End Sub

Private Sub AddBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Bar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    PlaylistKeyDown KeyCode

End Sub

Private Sub Bar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Menus(True) = False Then
    If Button = 1 Then
        PlayChoosing = True
        OrigX2 = X
        OrigY2 = Y
    End If
    Else
    Menus False
    End If

End Sub

Private Sub Bar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim XCaseVar As Integer, YCaseVar As Integer
    If PlayChoosing = True Then
        'there are some statements that i put as comments, because i couldn't find out how to get the height of the start menu so it doesn't snap to the bottom
        'these statements make the form snap to the sides of the screen
        XCaseVar = 1
        YCaseVar = 1
        
        'snapping to sides of screen
        If Me.Left + XTwips(X - OrigX2) < XTwips(10) Then
            If Me.Left + XTwips(X - OrigX2) > XTwips(-10) Then
            XCaseVar = 2
            End If
        End If
        If Me.Left + XTwips(X - OrigX2) + Me.Width > Screen.Width - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX2) + Me.Width < Screen.Width + XTwips(10) Then
            XCaseVar = 3
            End If
        End If
        If Me.Top + YTwips(Y - OrigY2) < YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY2) > YTwips(-10) Then
            YCaseVar = 2
            End If
        End If
        
        'snapping to main form
        If Me.Top + YTwips(Y - OrigY2) > VBAmpMain.Top - YTwips(10) - Me.Height And Me.Top + YTwips(Y - OrigY2) < VBAmpMain.Top + VBAmpMain.Height + YTwips(10) Then
        If Me.Left + XTwips(X - OrigX2) > VBAmpMain.Left + VBAmpMain.Width - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX2) < VBAmpMain.Left + VBAmpMain.Width + XTwips(10) Then
            XCaseVar = 4
            End If
        End If
        If Me.Left + XTwips(X - OrigX2) + Me.Width > VBAmpMain.Left - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX2) + Me.Width < VBAmpMain.Left + XTwips(10) Then
            XCaseVar = 5
            End If
        End If
        If Me.Left + XTwips(X - OrigX2) > VBAmpMain.Left - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX2) < VBAmpMain.Left + XTwips(10) Then
            XCaseVar = 6
            End If
        End If
        End If
        If Me.Left + XTwips(X - OrigX2) > VBAmpMain.Left - XTwips(10) - Me.Width And Me.Left + XTwips(X - OrigX2) < VBAmpMain.Left + VBAmpMain.Width + XTwips(10) Then
        If Me.Top + YTwips(Y - OrigY2) > VBAmpMain.Top + VBAmpMain.Height - YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY2) < VBAmpMain.Top + VBAmpMain.Height + YTwips(10) Then
            YCaseVar = 3
            End If
        End If
        If Me.Top + YTwips(Y - OrigY2) + Me.Height > VBAmpMain.Top - YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY2) + Me.Height < VBAmpMain.Top + YTwips(10) Then
            YCaseVar = 4
            End If
        End If
        If Me.Top + YTwips(Y - OrigY2) > VBAmpMain.Top - YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY2) < VBAmpMain.Top + YTwips(10) Then
            YCaseVar = 5
            End If
        End If
        End If
        'If Me.Top + YTwips(Y - OrigY) + Me.Height > Screen.Height - YTwips(10) And Me.Top + YTwips(Y - OrigY) + Me.Height < Screen.Height + YTwips(10) Then Me.Top = Screen.Height - Me.Height - StartMenVar
        'prevents flickering, because sometimes, for some reason the x and y's would fit more than one if statement(i think i might have made an error)
        XSnapped = True
        YSnapped = True
        Select Case XCaseVar
            Case 1
                Me.Left = Me.Left + XTwips(X - OrigX2)
                XSnapped = False
            Case 2
                Me.Left = 0
                XSnapped = False
            Case 3
                Me.Left = Screen.Width - Me.Width
                XSnapped = False
            Case 4
                Me.Left = VBAmpMain.Left + VBAmpMain.Width
            Case 5
                Me.Left = VBAmpMain.Left - Me.Width
            Case 6
                Me.Left = VBAmpMain.Left
        End Select
        Select Case YCaseVar
            Case 1
                Me.Top = Me.Top + YTwips(Y - OrigY2)
                YSnapped = False
            Case 2
                Me.Top = 0
                YSnapped = False
            Case 3
                Me.Top = VBAmpMain.Top + VBAmpMain.Height
            Case 4
                Me.Top = VBAmpMain.Top - Me.Height
            Case 5
                Me.Top = VBAmpMain.Top
        End Select
                
    End If

End Sub

Private Sub Bar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        PlayChoosing = False
        PlaylistSnapped = False
        If XSnapped = True Or YSnapped = True Then PlaylistSnapped = True
    End If

End Sub

Private Sub BotmBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub BotmBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub BotmBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonAdd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub ButtonAdd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error

    If Button = 1 Then
    AddMenu = False
    UpdateFormPlaylist
    If Index = 2 Then
    'add a file to playlist
    With VBAmpMain
    Extras.MP3Opener.ShowOpen
    If Playlist.ListItems.Count = 0 Then
        .Song.Caption = Extras.MP3Opener.FileName
        .SongText.Caption = "1. " & Mid(GetFileName(Extras.MP3Opener.FileName), 1, Len(GetFileName(Extras.MP3Opener.FileName)) - 4)
        NewMP3.FileName = Extras.MP3Opener.FileName
    End If
    AnotherMP3.FileName = Extras.MP3Opener.FileName
    Playlist.ListItems.Add , , Mid(GetFileName(Extras.MP3Opener.FileName), 1, Len(GetFileName(Extras.MP3Opener.FileName)) - 4)
    Playlist2.ListItems.Add , , Extras.MP3Opener.FileName
    SecondsList.ListItems.Add , , AnotherMP3.Seconds
    End With
    UpdatePlaylist
    UpdatePlaylistSlider False
    End If
    End If
    
Error:
End Sub

Private Sub ButtonLst_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub ButtonLst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error
    
    LstMenu = False
    UpdateFormPlaylist
    If Index = 2 Then
        M3UOpener.ShowOpen
        Me.Refresh
        VBAmpMain.Refresh
        OpenM3U M3UOpener.FileName
    ElseIf Index = 1 Then
        M3UOpener.ShowSave
        Me.Refresh
        VBAmpMain.Refresh
        SaveM3U M3UOpener.FileName
    End If
    
Error:
End Sub

Private Sub ButtonMis_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub ButtonMis_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = 0 Then
        PopMenu = True
        'make the extras form in the same position so you can get the popupmenu in the place you want it to be
        Extras.Left = Me.Left
        Extras.Top = Me.Top
        Extras.PopupMenu Extras.Sort, 2, 120, MisMen.Top - 38
        PopMenu = False
    ElseIf Index = 1 Then
        PopMenu = True
        Extras.Left = Me.Left
        Extras.Top = Me.Top
        Extras.PopupMenu Extras.FileInf, 2, 120, MisMen.Top - 20
        PopMenu = False
    End If
    MisMenu = False
    UpdateFormPlaylist

End Sub

Private Sub ButtonRem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub ButtonRem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error
    
    If Button = 1 Then
    RemMenu = False
    UpdateFormPlaylist
    If Index = 3 Then
    If SelectedIndex <> -1 Then
    Playlist.ListItems.Remove SelectedIndex + 1
    Playlist2.ListItems.Remove SelectedIndex + 1
    SecondsList.ListItems.Remove SelectedIndex + 1
    Item(SelectedIndex - SongNumber).BackColor = ListView.BackColor
    Time(SelectedIndex - SongNumber).BackColor = ListView.BackColor
    SelectedIndex = -1
    UpdatePlaylist
    End If
    ElseIf Index = 1 Then
    Playlist.ListItems.Clear
    Playlist2.ListItems.Clear
    If SelectedIndex <> -1 Then
    Item(SelectedIndex - SongNumber).BackColor = ListView.BackColor
    SelectedIndex = -1
    End If
    UpdatePlaylist
    UpdatePlaylistSlider False
    End If
    End If
    
Error:
End Sub

Private Sub ButtonSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub ButtonSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SelMenu = False
    UpdateFormPlaylist

End Sub

Private Sub Exit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Exit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Exit2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Exit2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Form_Load()
    
    'set colors
    UnSelectedForeColor = RGB(0, &HFF, 0)
    CurrentForeColor = RGB(&HFF, &HFF, &HFF)
    ListView.BackColor = RGB(0, 0, 0)
    SelectedBackColor = RGB(0, 0, &HC6)
    
    For I = Item.LBound To Item.UBound
        Item(I).BackColor = vbBlack
        Time(I).BackColor = vbBlack
    Next
    
    'size form
    Me.Height = VBAmpMain.YTwips(LeftBar.Height)
    Me.Width = VBAmpMain.XTwips(BotmBar.Width)
    
    MouseOverPlaylist
    UpdateFormPlaylist
    UpdatePlaylist
    UpdatePlaylistSlider False
    SelectedIndex = -1
    M3UOpener.Filter = "M3Us|*.m3u"

End Sub

Private Sub Item_DblClick(Index As Integer)

    If Index < Playlist.ListItems.Count Then
        Playlist.ListItems.Item(Index + SongNumber + 1).Selected = True
        VBAmpMain.OpenAFile Playlist2.ListItems.Item(Index + SongNumber + 1).Text, Index + SongNumber + 1 & ". " & Playlist.ListItems.Item(Index + SongNumber + 1).Text
        UpdatePlaylist
    End If

End Sub

Private Sub Item_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    PlaylistKeyDown KeyCode

End Sub

Private Sub Item_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Menus(True) = False Then
    'Item(SelectedIndex - SongNumber).BackStyle = 0
    If Index + SongNumber < Playlist.ListItems.Count Then
    If SelectedIndex - SongNumber <> Index Then
    'Item(Index).BackStyle = 1
    Item(SelectedItem - SongNumber - 1).ForeColor = CurrentForeColor
    End If
    SelectedIndex = Index + SongNumber
    End If
    If Index + SongNumber >= Playlist.ListItems.Count Then SelectedIndex = -1
    UpdatePlaylist
    Else
    Menus False
    End If

End Sub

Public Sub MouseDown()

    If Menus(True) = False Then
    PlayButtonDown = True
    UpdateFormPlaylist
    Else
    Menus False
    End If

End Sub

Public Sub MouseUp()

    PlayButtonDown = False
    UpdateFormPlaylist

End Sub

Private Sub LeftBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub LeftBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub LeftBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub List_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Menus(True) = False Then
        LstMenu = True
    Else
        Menus False
    End If
End Sub

Private Sub ListBar_Click()
    If Menus(True) = True Then Menus (False)
End Sub

Private Sub ListBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub ListView_KeyDown(KeyCode As Integer, Shift As Integer)

    PlaylistKeyDown KeyCode

End Sub

Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Menus False
    If Y > Item(VisibleItemCount - 1).Top + 12 Then
        SelectedIndex = -1
    Else
        SelectedIndex = Int((Y - 2) / 13) + SongNumber
    End If
    UpdatePlaylist
End Sub

Private Sub MisBar_Click()
    If Menus(True) = True Then Menus (False)
End Sub

Private Sub MisBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Misc_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Misc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And Menus(True) = False Then
        MisMenu = True
    Else
        Menus False
    End If

End Sub

Public Sub PlaylistKeyDown(KeyCode As Integer)
On Error GoTo Error
    If SelectedIndex < SongNumber Or SelectedIndex > SongNumber - 1 + Item.Count Then SelectedIndex = SongNumber - 1
    If SelectedIndex + 1 < Playlist.ListItems.Count And KeyCode = vbKeyDown Then
        SelectedIndex = SelectedIndex + 1
        If SelectedIndex > SongNumber - 1 + Item.Count Then SongNumber = SongNumber + 1
    End If
    If SelectedIndex - 1 >= 0 And KeyCode = vbKeyUp Then
        SelectedIndex = SelectedIndex - 1
        If SelectedIndex < SongNumber Then SongNumber = SongNumber - 1
    End If
    If KeyCode = vbKeyReturn Then
        Playlist.ListItems.Item(SelectedIndex + 1).Selected = True
        VBAmpMain.OpenAFile Playlist2.ListItems.Item(SelectedIndex + 1).Text, SelectedIndex + 1 & ". " & Playlist.ListItems.Item(SelectedIndex + 1).Text
    End If
    UpdatePlaylist
Error:
End Sub

Private Sub PlaylistSlider_KeyDown(KeyCode As Integer, Shift As Integer)
    
    PlaylistKeyDown KeyCode

End Sub

Private Sub PlaylistSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And X > 9 And X < 18 Then
    If Playlist.ListItems.Count <> VisibleItemCount Then
    If Y < Round((SongNumber / (Playlist.ListItems.Count - VisibleItemCount)) * PlaylistSlider.Height) Or Y > Round((SongNumber / (Playlist.ListItems.Count - VisibleItemCount)) * PlaylistSlider.Height) + 18 Then
        If Y >= 9 And Y <= PlaylistSlider.Height - 9 Then SongNewY = Y - 9
        If Y < 7 Then SongNewY = 0
        If Y > PlaylistSlider.Height - 9 Then SongNewY = PlaylistSlider.Height - 18
    End If
    End If
    ChoosingSongNumber = True
    UpdatePlaylistSlider True
    OrigY = Y
    OrigPosition = OrigY - SongNewY
    End If

End Sub

Private Sub PlaylistSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ChoosingSongNumber = True Then
    If SongNewY + (Y - OrigY) >= 0 And SongNewY + (Y - OrigY) <= PlaylistSlider.Height - 9 Then
        SongNewY = SongNewY + (Y - OrigY)
        OrigY = Y
    ElseIf SongNewY + (Y - OrigY) < 0 Then
        SongNewY = 0
        OrigY = OrigPosition
    ElseIf SongNewY + (Y - OrigY) > PlaylistSlider.Height - 9 Then
        SongNewY = PlaylistSlider.Height - 18
        OrigY = PlaylistSlider.Height - 18 + OrigPosition
    End If
    UpdatePlaylistSlider True
    End If

End Sub

Private Sub PlaylistSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then ChoosingSongNumber = False
    UpdatePlaylistSlider False

End Sub

Private Sub RemBar_Click()
    If Menus(True) = True Then Menus (False)
End Sub

Private Sub RemBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Remove_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Remove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Menus(True) = False Then
        RemMenu = True
    Else
        Menus False
    End If
End Sub

Private Sub SelBar_Click()
    If Menus(True) = True Then Menus (False)
End Sub

Private Sub SelBar_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Select_KeyDown(KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Select_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Menus(True) = False Then
        SelMenu = True
    Else
        Menus False
    End If
End Sub

Private Sub Shade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Shade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Shade2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Shade2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Time_DblClick(Index As Integer)
    Item_DblClick Index
End Sub

Private Sub Time_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PlaylistKeyDown KeyCode
End Sub

Private Sub Time_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Item_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()

    '500
    If Me.Visible = True Then
    MouseOverPlaylist
    UpdateFormPlaylist
    End If

End Sub

Private Sub Timer2_Timer()

    '100
    If Menus(True) = True Then
        MouseOverPlaylist
        UpdateFormPlaylist
    End If

End Sub
