VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form VBAmpMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBAmp"
   ClientHeight    =   2520
   ClientLeft      =   3885
   ClientTop       =   2280
   ClientWidth     =   4125
   Icon            =   "VBAmpMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "VBAmpMain.frx":08CA
   Picture         =   "VBAmpMain.frx":1194
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   2280
      Top             =   2055
   End
   Begin VB.PictureBox WinShadeOpen2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3225
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   78
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadeNext2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3090
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   77
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadeStop2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2940
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   76
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadePause2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2805
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   75
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadePlay2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2655
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   74
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadePrev2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2520
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   73
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox ButtonWinShadeTrue2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   72
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonWinShade2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   71
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonWinShadeTrueD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   70
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonWinShadeD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   69
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonWinShadeTrue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   68
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonMinD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3660
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   67
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonMin2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3660
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   66
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonMenuD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   90
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   62
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonMenu2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   90
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   65
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonExitD 
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
      TabIndex        =   63
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox ButtonExit2 
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
      TabIndex        =   64
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox BlankSec 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1170
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   61
      Top             =   390
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox BlankMin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   60
      Top             =   390
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox PLTrueD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3630
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   59
      Top             =   870
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PLD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3630
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   58
      Top             =   870
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PLTrue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3630
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   57
      Top             =   870
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ButtonRepeatTrueD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   56
      Top             =   1335
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox ButtonRepeatD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   55
      Top             =   1335
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox ButtonRepeatTrue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   54
      Top             =   1335
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox ButtonShuffleD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2460
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   51
      Top             =   1335
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox ButtonShuffleTrueD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2460
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   52
      Top             =   1335
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox ButtonShuffleTrue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2460
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   53
      Top             =   1335
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox ButtonShuffle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2460
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   12
      Top             =   1335
      Width           =   675
   End
   Begin VB.PictureBox ButtonOpenD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2040
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   50
      Top             =   1335
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox ButtonPrevD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   49
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ButtonPlayD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   585
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   48
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ButtonPauseD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   930
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   47
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ButtonStopD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1275
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   46
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox ButtonNextD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1620
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   45
      Top             =   1320
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox SliderWSPos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   3390
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   41
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox WinShadeOpen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3225
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   40
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadeNext 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3090
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   39
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadeStop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2940
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   38
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadePause 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2805
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   37
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadePlay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2655
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   36
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadePrev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2520
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   35
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox WinShadeTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   1875
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   34
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Stereo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3585
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   33
      Top             =   615
      Width           =   435
   End
   Begin VB.PictureBox Mono 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3180
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   32
      Top             =   615
      Width           =   435
   End
   Begin VB.PictureBox PL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3630
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   30
      Top             =   870
      Width           =   345
   End
   Begin VB.PictureBox EQ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3285
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   29
      Top             =   870
      Width           =   345
   End
   Begin VB.PictureBox KHZ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   28
      Top             =   585
      Width           =   240
   End
   Begin VB.PictureBox KBPS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1605
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   27
      Top             =   585
      Width           =   315
   End
   Begin VB.PictureBox Negative 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   570
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   26
      Top             =   480
      Width           =   75
   End
   Begin VB.PictureBox SliderPos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   240
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   248
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.PictureBox Indicator 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   390
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   24
      Top             =   420
      Width           =   135
   End
   Begin VB.PictureBox RemSec 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1170
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   23
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox RemMin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   22
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox ElapsedSec 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1170
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   20
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox ElapsedMin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   19
      Top             =   390
      Width           =   315
   End
   Begin VB.Timer Timer2 
      Interval        =   180
      Left            =   1395
      Top             =   2055
   End
   Begin VB.PictureBox ScrollText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1605
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   16
      Top             =   330
      Width           =   2400
   End
   Begin VB.PictureBox SliderBalance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2655
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   15
      Top             =   855
      Width           =   570
   End
   Begin VB.PictureBox SliderVolume 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1605
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   14
      Top             =   855
      Width           =   1020
   End
   Begin VB.PictureBox ButtonRepeat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   13
      Top             =   1335
      Width           =   420
   End
   Begin VB.PictureBox ButtonOpen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2040
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   11
      Top             =   1335
      Width           =   330
   End
   Begin VB.PictureBox ButtonNext 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1620
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   10
      Top             =   1320
      Width           =   330
   End
   Begin VB.PictureBox ButtonStop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1275
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   9
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox ButtonPause 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   930
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   8
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox ButtonPlay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   585
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   7
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox ButtonPrev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   6
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox ButtonWinShade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   5
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox ButtonMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   90
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   4
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox ButtonMin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3660
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   3
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox ButtonExit 
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
      TabIndex        =   2
      Top             =   45
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   960
      Top             =   2055
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1830
      Top             =   2055
   End
   Begin VB.PictureBox Volume 
      Height          =   135
      Left            =   1605
      ScaleHeight     =   75
      ScaleWidth      =   960
      TabIndex        =   43
      Top             =   885
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Balance 
      Height          =   135
      Left            =   2655
      ScaleHeight     =   75
      ScaleWidth      =   510
      TabIndex        =   44
      Top             =   885
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   0
      Picture         =   "VBAmpMain.frx":11DA
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   0
      Picture         =   "VBAmpMain.frx":9F0C
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   0
      Picture         =   "VBAmpMain.frx":12C3E
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   0
      Picture         =   "VBAmpMain.frx":1B970
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   0
      Width           =   4125
   End
   Begin VB.PictureBox BASEMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      Picture         =   "VBAmpMain.frx":246A2
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.PictureBox Main 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      Picture         =   "VBAmpMain.frx":27282
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   1
      Top             =   0
      Width           =   4125
   End
   Begin VB.PictureBox ElapsedRem 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   480
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   21
      Top             =   330
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Song 
      Height          =   255
      Left            =   930
      TabIndex        =   31
      Top             =   885
      Visible         =   0   'False
      Width           =   2460
   End
   Begin MediaPlayerCtl.MediaPlayer WinMedia 
      Height          =   540
      Left            =   645
      TabIndex        =   17
      Top             =   495
      Visible         =   0   'False
      Width           =   855
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -130
      WindowlessVideo =   0   'False
   End
   Begin VB.Label SongText 
      Height          =   225
      Left            =   720
      TabIndex        =   18
      Top             =   690
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "VBAmpMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function XTwips(Pixels) As Integer

    'convert pixels to twips
    XTwips = Pixels * Screen.TwipsPerPixelX

End Function

Function YTwips(Pixels) As Integer

    'convert pixels to twips
    YTwips = Pixels * Screen.TwipsPerPixelY

End Function

Private Sub Bar_DblClick(Index As Integer)

    ShadeD = False
    ButtonNum = 2
    DblClicked = True
    UpdatePictures

End Sub

Private Sub Bar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Choosing = True
        OrigX = X
        OrigY = Y
    End If
    If Button = 2 Then
        Extras.PopupMenu Extras.Options
    End If
    
End Sub

Private Sub Bar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XCaseVar As Integer, YCaseVar As Integer
    If Button = 1 And Choosing = True Then
        XCaseVar = 1
        YCaseVar = 1
        TopVar = Playlist.Top
        LeftVar = Playlist.Left
        'snapping to playlist
        If PlaylistSnapped = False Then
        If Me.Top + YTwips(Y - OrigY) > TopVar - YTwips(10) - Me.Height And Me.Top + YTwips(Y - OrigY2) < TopVar + Playlist.Height + YTwips(10) Then
        If Me.Left + XTwips(X - OrigX) > LeftVar + Playlist.Width - XTwips(10) And Me.Left + XTwips(X - OrigX) < LeftVar + Playlist.Width + XTwips(10) Then
            XCaseVar = 2
        End If
        If Me.Left + XTwips(X - OrigX) + Me.Width > LeftVar - XTwips(10) And Me.Left + XTwips(X - OrigX) + Me.Width < LeftVar + XTwips(10) Then
            XCaseVar = 3
        End If
        If Me.Left + XTwips(X - OrigX) > LeftVar - XTwips(10) And Me.Left + XTwips(X - OrigX) < LeftVar + XTwips(10) Then
            XCaseVar = 4
        End If
        End If
        If Me.Left + XTwips(X - OrigX) > LeftVar - XTwips(10) - Me.Width And Me.Left + XTwips(X - OrigX2) < LeftVar + Playlist.Width + XTwips(10) Then
        If Me.Top + YTwips(Y - OrigY) > TopVar + Playlist.Height - YTwips(10) And Me.Top + YTwips(Y - OrigY) < TopVar + Playlist.Height + YTwips(10) Then
            YCaseVar = 2
        End If
        If Me.Top + YTwips(Y - OrigY) + Me.Height > TopVar - YTwips(10) And Me.Top + YTwips(Y - OrigY) + Me.Height < TopVar + YTwips(10) Then
            YCaseVar = 3
        End If
        If Me.Top + YTwips(Y - OrigY) > TopVar - YTwips(10) And Me.Top + YTwips(Y - OrigY) < TopVar + YTwips(10) Then
            YCaseVar = 4
        End If
        End If
        End If
        
        XSnapped = True
        YSnapped = True
        Select Case XCaseVar
            Case 1
                Me.Left = Me.Left + XTwips(X - OrigX)
                If PlaylistSnapped = True Then Playlist.Left = Playlist.Left + XTwips(X - OrigX)
                If PlaylistSnapped = False Then XSnapped = False
                LeftVar = Playlist.Left
            Case 2
                Me.Left = LeftVar + Playlist.Width
            Case 3
                Me.Left = LeftVar - Me.Width
            Case 4
                Me.Left = LeftVar
        End Select
        Select Case YCaseVar
            Case 1
                Me.Top = Me.Top + YTwips(Y - OrigY)
                If PlaylistSnapped = True Then Playlist.Top = Playlist.Top + YTwips(Y - OrigY)
                If PlaylistSnapped = False Then YSnapped = False
                TopVar = Playlist.Top
            Case 2
                Me.Top = TopVar + Playlist.Height
            Case 3
                Me.Top = TopVar - Me.Height
            Case 4
                Me.Top = TopVar
        End Select
        Me.Refresh
        Playlist.Refresh
    End If

End Sub

Private Sub Bar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Choosing = False
        PlaylistSnapped = False
        If XSnapped = True Or YSnapped = True Then PlaylistSnapped = True
    End If

End Sub

Private Sub ButtonExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonExit2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonExit2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonMenu2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonMenu2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonMin2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonMin2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonRepeat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonRepeat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonRepeatTrue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonRepeatTrue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonShuffle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonShuffle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonShuffleTrue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonShuffleTrue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonWinShade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonWinShade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonWinShade2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonWinShade2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonWinShadeTrue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonWinShadeTrue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ButtonWinShadeTrue2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub ButtonWinShadeTrue2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Public Sub ChangeSkin(Dir, Optional BaseSkin As Boolean)
On Error Resume Next

    RestoreDefaults
    If BaseSkin = False Then
    If Right(Dir, 1) <> "\" Then Dir = Dir & "\"
        'load all the skin pictures
        Extras.Buttons = LoadPicture(Dir & "cbuttons.bmp")
        Main = LoadPicture(Dir & "main.bmp")
        Extras.PlayPaus = LoadPicture(Dir & "playpaus.bmp")
        Extras.Numbers = LoadPicture(Dir & "numbers.bmp")
        Negative.PaintPicture Extras.Numbers, 0, 0, , , 21, 6, 5, 1
        If FileExists(Dir & "nums_ex.bmp") = True Then
            Extras.Numbers = LoadPicture(Dir & "nums_ex.bmp")
            Negative.PaintPicture Extras.Numbers, 0, 0, , , 101, 6, 5, 1
        End If
        Extras.PosBar = LoadPicture(Dir & "posbar.bmp")
        Extras.Text = LoadPicture(Dir & "text.bmp")
        Extras.Shuffle = LoadPicture(Dir & "shufrep.bmp")
        Extras.Bars = LoadPicture(Dir & "titlebar.bmp")
        Extras.VolumePic = LoadPicture(Dir & "volume.bmp")
        'if there is no special slider for balance, then it uses the volume slider picture
        Extras.BalancePic = LoadPicture(Dir & "volume.bmp")
        Extras.BalancePic = LoadPicture(Dir & "balance.bmp")
        Extras.MonoSter = LoadPicture(Dir & "monoster.bmp")
        Playlist.PLEdit = LoadPicture(Dir & "pledit.bmp")
    End If
    
    If BaseSkin = False Then
        If FileExists(Dir & "pledit.txt") = True Then
            Open Dir & "pledit.txt" For Input As #1
            Dim String1 As String
            Line Input #1, String1
            Line Input #1, String1
                UnSelectedForeColor = RGB("&H" & Mid(String1, 9, 2), "&H" & Mid(String1, 11, 2), "&H" & Mid(String1, 13, 2))
            Line Input #1, String1
                CurrentForeColor = RGB("&H" & Mid(String1, 10, 2), "&H" & Mid(String1, 12, 2), "&H" & Mid(String1, 14, 2))
            Line Input #1, String1
                BackColor = RGB("&H" & Mid(String1, 11, 2), "&H" & Mid(String1, 13, 2), "&H" & Mid(String1, 15, 2))
            Line Input #1, String1
                SelectedBackColor = RGB("&H" & Mid(String1, 13, 2), "&H" & Mid(String1, 15, 2), "&H" & Mid(String1, 17, 2))
            Line Input #1, String1
            Playlist.ListView.BackColor = BackColor
            For I = Playlist.Item.LBound To Playlist.Item.UBound
                Playlist.Item(I).Font.Name = Mid(String1, 6, Len(String1) - 5)
                Playlist.Time(I).Font.Name = Playlist.Item(I).Font.Name
            Next
            Playlist.TestLabel.Font.Name = Playlist.Item(0).Font.Name
            Playlist.TestLabel2.Font.Name = Playlist.Item(0).Font.Name
            Playlist.TestLabel3.Font.Name = Playlist.Item(0).Font.Name
            Close #1
        End If
    End If
    'For I = Playlist.Item.LBound To Playlist.Item.UBound
        'If I <> SelectedIndex - SongNumber Or SelectedIndex >= Playlist.Playlist.ListItems.Count Or SelectedIndex = -1 Then Playlist.Item(I).BackStyle = 0
        'If I + 1 <> SelectedItem - SongNumber Then
            'Playlist.Item(I).ForeColor = UnSelectedForeColor
            'Playlist.Time(I).ForeColor = UnSelectedForeColor
            
        'If I + 1 = SelectedItem - SongNumber Then Playlist.Item(I).ForeColor = CurrentForeColor
    'Next
    DrawButtons
    UpdatePictures True
    UpdateFormPlaylist
    UpdateSliders
    UpdatePlaylist
    UpdatePlaylistSlider False
    
End Sub

Private Sub ElapsedMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Elapsed = False
    If Button = 1 Then MouseDown
End Sub

Private Sub ElapsedMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub ElapsedSec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Elapsed = False
    If Button = 1 Then MouseDown
End Sub

Private Sub ElapsedSec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub EQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub EQ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub


Private Sub Form_Load()
'On Error Resume Next
    
    Elapsed = True
    ShuffleBln = False
    RepeatBln = False
    PLBln = True
    PLTrue.Visible = True
    
    Me.Width = XTwips(Main.Width)
    Me.Height = YTwips(Main.Height)
    
    'don't start in winshade mode
    WinShade = False
    
    'set initial volume and balance positions
    VolNewX = SliderVolume.Width - 14
    BalanceNewX = (SliderBalance.Width - 14) / 2
    SongPosNewX = 0
    WSSongPosNewX = 0
    UpdatePictures True
    UpdateSliders
    
    'make a borderless form that shows in taskbar with icon
    SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) + WS_DLGFRAME
    
    'Set string to scroll
    SongText.Caption = "VBAmp 1.1: It Really Whips Winamp's Ass"
    
    'Common Dialog stuff
    Extras.MP3Opener.Filter = "MP3s|*.mp3"
    'this is my music directory, you can change it to your music directory to open your mp3's
    Extras.MP3Opener.InitDir = "C:\My Music\"

    'paint pictures that don't change
    Negative.PaintPicture Extras.Numbers, 0, 0, , , 21, 6, 5, 1
    Mono.PaintPicture Extras.MonoSter, 0, 0, , , 29, 12
    Stereo.PaintPicture Extras.MonoSter, 0, 0, , , 0, 12
    
    NumbertoPicture 0, ElapsedMin, True
    NumbertoPicture 0, ElapsedSec, True
    NumbertoPicture 0, RemMin, True
    NumbertoPicture 0, RemSec, True
    Indicator.PaintPicture Extras.PlayPaus, 0, 0, , , 18, 0
           
    Playlist.Top = Me.Top + Me.Height
    Playlist.Left = Me.Left
    Playlist.Visible = True
    PlaylistSnapped = True
    
    SliderWSPos.PaintPicture Extras.Bars, 0, 0, , , 0, 36
    
    'Draw Button Pictures
    DrawButtons
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Indicator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Indicator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Main_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If Y <= 20 Then
            Choosing = True
            OrigX = X
            OrigY = Y
        End If
        If Y > 20 Then MouseDown
        If X >= ElapsedRem.Left And X <= ElapsedRem.Left + ElapsedRem.Width And Y >= ElapsedRem.Top And Y <= ElapsedRem.Top + ElapsedRem.Height Then
        If Elapsed = False Then
            Elapsed = True
        ElseIf Elapsed = True Then
            Elapsed = False
        End If
        End If
    End If
    
End Sub

Private Sub Main_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim XCaseVar As Integer, YCaseVar As Integer
    If Button = 1 And Choosing = True Then
        'there are some statements that i put as comments, because i couldn't find out how to get the height of the start menu so it doesn't snap to the bottom
        'these statements make the form snap to the sides of the screen
        XCaseVar = 1
        YCaseVar = 1
        
        'snapping to sides of screen
        If Me.Left + XTwips(X - OrigX) < XTwips(10) Then
            If Me.Left + XTwips(X - OrigX) > XTwips(-10) Then
            XCaseVar = 2
            End If
        End If
        If Me.Left + XTwips(X - OrigX) + Me.Width > Screen.Width - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX) + Me.Width < Screen.Width + XTwips(10) Then
            XCaseVar = 3
            End If
        End If
        If Me.Top + YTwips(Y - OrigY) < YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY) > YTwips(-10) Then
            YCaseVar = 2
            End If
        End If
        
        'snapping to playlist
        If Me.Top + YTwips(Y - OrigY) > Playlist.Top - YTwips(10) - Me.Height And Me.Top + YTwips(Y - OrigY2) < Playlist.Top + Playlist.Height + YTwips(10) Then
        If Me.Left + XTwips(X - OrigX) > Playlist.Left + Playlist.Width - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX) < Playlist.Left + Playlist.Width + XTwips(10) Then
            XCaseVar = 4
            End If
        End If
        If Me.Left + XTwips(X - OrigX) + Me.Width > Playlist.Left - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX) + Me.Width < Playlist.Left + XTwips(10) Then
            XCaseVar = 5
            End If
        End If
        If Me.Left + XTwips(X - OrigX) > Playlist.Left - XTwips(10) Then
            If Me.Left + XTwips(X - OrigX) < Playlist.Left + XTwips(10) Then
            XCaseVar = 6
            End If
        End If
        End If
        If Me.Left + XTwips(X - OrigX) > Playlist.Left - XTwips(10) - Me.Width And Me.Left + XTwips(X - OrigX2) < Playlist.Left + Playlist.Width + XTwips(10) Then
        If Me.Top + YTwips(Y - OrigY) > Playlist.Top + Playlist.Height - YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY) < Playlist.Top + Playlist.Height + YTwips(10) Then
            YCaseVar = 3
            End If
        End If
        If Me.Top + YTwips(Y - OrigY) + Me.Height > Playlist.Top - YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY) + Me.Height < Playlist.Top + YTwips(10) Then
            YCaseVar = 4
            End If
        End If
        If Me.Top + YTwips(Y - OrigY) > Playlist.Top - YTwips(10) Then
            If Me.Top + YTwips(Y - OrigY) < Playlist.Top + YTwips(10) Then
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
                Me.Left = Me.Left + XTwips(X - OrigX)
                If PlaylistSnapped = True Then Playlist.Left = Playlist.Left + XTwips(X - OrigX)
                If PlaylistSnapped = False Then XSnapped = False
            Case 2
                If PlaylistSnapped = True Then
                TempDiff = Playlist.Left - Me.Left
                    Me.Left = 0
                    Playlist.Left = TempDiff
                End If
                If PlaylistSnapped = False Then
                    Me.Left = 0
                    XSnapped = False
                End If
            Case 3
                If PlaylistSnapped = True Then
                TempDiff = Me.Left - Playlist.Left
                    Me.Left = Screen.Width - Me.Width
                    Playlist.Left = Screen.Width - TempDiff
                End If
                If PlaylistSnapped = False Then
                    Me.Left = Screen.Width - Me.Width
                    XSnapped = False
                End If
            Case 4
                Me.Left = Playlist.Left + Playlist.Width
            Case 5
                Me.Left = Playlist.Left - Me.Width
            Case 6
                Me.Left = Playlist.Left
        End Select
        Select Case YCaseVar
            Case 1
                Me.Top = Me.Top + YTwips(Y - OrigY)
                If PlaylistSnapped = True Then Playlist.Top = Playlist.Top + YTwips(Y - OrigY)
                If PlaylistSnapped = False Then XSnapped = False
            Case 2
                If PlaylistSnapped = True Then
                TempDiff = Playlist.Top - Me.Top
                    Me.Top = 0
                    Playlist.Top = TempDiff
                End If
                If PlaylistSnapped = False Then
                    Me.Top = 0
                    XSnapped = False
                End If
            Case 3
                Me.Top = Playlist.Top + Playlist.Height
            Case 4
                Me.Top = Playlist.Top - Me.Height
            Case 5
                Me.Top = Playlist.Top
        End Select
    End If

End Sub

Private Sub Main_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Choosing = True Then
        Choosing = False
        PlaylistSnapped = False
        If XSnapped = True Or YSnapped = True Then PlaylistSnapped = True
        End If
        MouseUp
    End If
End Sub

Public Sub MouseDown()
    ButtonDown = True
    MouseOver
    UpdatePictures
    UpdatePictures
End Sub

Public Sub MouseUp()
    ButtonDown = False
    MouseOver
    UpdatePictures
    UpdatePictures
End Sub

Public Sub NextSong()
'On Error Resume Next
    Dim NextItem
    
    If ShuffleBln = False Then
        NextItem = SelectedItem + 1
        If SelectedItem + 1 <= Playlist.Playlist.ListItems.Count Then Playlist.Playlist.ListItems.Item(NextItem).Selected = True
        'if repeat is selected and the song being played is the last in the list then it goes to song number one
        If SelectedItem + 1 > Playlist.Playlist.ListItems.Count And RepeatBln = True And Playlist.Playlist.ListItems.Count > 0 Then
            Playlist.Playlist.ListItems.Item(1).Selected = True
            NextItem = 1
            SongNumber = 0
        End If
        'only adds the song to list of songs played if a new song is played
        If SelectedItem + 1 <= Playlist.Playlist.ListItems.Count Or RepeatBln = True Then Extras.SongList.AddItem SelectedItem
        If WinMedia.PlayState = mpPlaying Then OpenAFile Playlist.Playlist2.ListItems.Item(NextItem).Text, NextItem & ". " & Playlist.Playlist.ListItems.Item(NextItem).Text
        If WinMedia.PlayState <> mpPlaying Then OpenAFile Playlist.Playlist2.ListItems.Item(NextItem).Text, NextItem & ". " & Playlist.Playlist.ListItems.Item(NextItem).Text, True
        'if the new song is not visible in the playlist then it is made visible
        If SongNumber + VisibleItemCount < NextItem Then SongNumber = SongNumber + 1
        UpdatePlaylist
        UpdatePlaylistSlider False
    End If
    If ShuffleBln = True Then
        Extras.SongList.AddItem SelectedItem
        Randomize
        TemporaryNum = Int(Playlist.Playlist.ListItems.Count * Rnd + 1)
        Playlist.Playlist.ListItems.Item(TemporaryNum).Selected = True
        If SongNumber + VisibleItemCount < TemporaryNum Or SongNumber >= TemporaryNum Then
        If TemporaryNum - Int(VisibleItemCount / 2) - 1 >= 0 Then SongNumber = TemporaryNum - Int(VisibleItemCount / 2) - 1 Else SongNumber = 0
        End If
        UpdatePlaylist
        UpdatePlaylistSlider False
        If WinMedia.PlayState = mpPlaying Then OpenAFile Playlist.Playlist2.ListItems.Item(TemporaryNum).Text, TemporaryNum & ". " & Playlist.Playlist.ListItems.Item(TemporaryNum).Text
        If WinMedia.PlayState <> mpPlaying Then OpenAFile Playlist.Playlist2.ListItems.Item(TemporaryNum).Text, TemporaryNum & ". " & Playlist.Playlist.ListItems.Item(TemporaryNum).Text, True
    End If

End Sub

Public Sub OpenAFile(Optional File As String, Optional FileTitle As String, Optional DontPlay As Boolean)
On Error GoTo ErrorHandle
    
    UpdatePictures
    If File = "" Then
        Extras.MP3Opener.ShowOpen
        SelectedIndex = -1
        UpdatePlaylist
        File = Extras.MP3Opener.FileName
        NewMP3.FileName = File
        FileTitle = Mid(GetFileName(File), 1, Len(GetFileName(File)) - 4)
        Playlist.Playlist.ListItems.Clear
        Playlist.Playlist2.ListItems.Clear
        Playlist.SecondsList.ListItems.Clear
        Playlist.Playlist.ListItems.Add , , FileTitle
        Playlist.Playlist2.ListItems.Add , , File
        Playlist.Playlist.ListItems.Item(1).Selected = True
        Playlist.SecondsList.ListItems.Add , , NewMP3.Seconds
        UpdatePlaylist
        UpdatePlaylistSlider False
        FileTitle = "1. " & FileTitle
    End If
    Me.Refresh
    Playlist.Refresh
    WinMedia.Stop
    Song.Caption = File
    NewMP3.FileName = File
    SongText.Caption = FileTitle
    If DontPlay = False Then Play

ErrorHandle:
End Sub

Public Sub Pause()

    If WinMedia.PlayState = mpPlaying Then
        WinMedia.Pause
    ElseIf WinMedia.PlayState = mpPaused Then
        WinMedia.Play
    End If
    
End Sub

Public Sub Play()
'On Error GoTo Error
    
    If WinMedia.PlayState <> mpPaused Then
        WinMedia.FileName = Song.Caption
        WinMedia.Play
    Else
        WinMedia.Play
    End If
    UpdatePictures
    Exit Sub

Error:
   If FileExists(Song.Caption) = True Then MsgBox "Cannot Initialize Windows Media Player"
End Sub

Public Sub PrevSong()
On Error Resume Next

    If Extras.SongList.ListCount >= 1 Then
        Playlist.Playlist.ListItems.Item(Int(Extras.SongList.List(Extras.SongList.ListCount - 1))).Selected = True
        If SongNumber + VisibleItemCount < Int(Extras.SongList.List(Extras.SongList.ListCount - 1)) Or SongNumber >= Int(Extras.SongList.List(Extras.SongList.ListCount - 1)) Then
        If Int(Extras.SongList.List(Extras.SongList.ListCount - 1)) - Int(VisibleItemCount / 2) - 1 >= 0 Then SongNumber = Int(Extras.SongList.List(Extras.SongList.ListCount - 1)) - Int(VisibleItemCount / 2) - 1 Else SongNumber = 0
        End If
        UpdatePlaylist
        UpdatePlaylistSlider False
        If WinMedia.PlayState = mpPlaying Then OpenAFile Playlist.Playlist2.ListItems.Item(Int(Extras.SongList.List(Extras.SongList.ListCount - 1))).Text, Extras.SongList.List(Extras.SongList.ListCount - 1) & ". " & Playlist.Playlist.ListItems.Item(Int(Extras.SongList.List(Extras.SongList.ListCount - 1))).Text
        If WinMedia.PlayState <> mpPlaying Then OpenAFile Playlist.Playlist2.ListItems.Item(Int(Extras.SongList.List(Extras.SongList.ListCount - 1))).Text, Extras.SongList.List(Extras.SongList.ListCount - 1) & ". " & Playlist.Playlist.ListItems.Item(Int(Extras.SongList.List(Extras.SongList.ListCount - 1))).Text, True
        Extras.SongList.RemoveItem (Extras.SongList.ListCount - 1)
    End If

End Sub

Private Sub Mono_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Mono_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Negative_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Elapsed = True
    If Button = 1 Then MouseDown
End Sub

Private Sub Negative_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub PL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub PL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub PLTrue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub PLTrue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub RemMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Elapsed = True
    If Button = 1 Then MouseDown
End Sub

Private Sub RemMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub RemSec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Elapsed = True
    If Button = 1 Then MouseDown
End Sub

Private Sub RemSec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Public Sub RestoreDefaults()
'On Error Resume Next

    SavePicture Extras.BASEButtons, App.Path & "\BASEButtons.bmp"
    SavePicture BASEMain, App.Path & "\BASEMain.bmp"
    SavePicture Extras.BASEPlayPaus, App.Path & "\BASEPlayPaus.bmp"
    SavePicture Extras.BASENumbers, App.Path & "\BASENumbers.bmp"
    SavePicture Extras.BASEPosBar, App.Path & "\BASEPosBar.bmp"
    SavePicture Extras.BASEText, App.Path & "\BASEText.bmp"
    SavePicture Extras.BASEShuffle, App.Path & "\BASEShuffle.bmp"
    SavePicture Extras.BASEBars, App.Path & "\BASEBars.bmp"
    SavePicture Extras.BASEVolumePic, App.Path & "\BASEVolumePic.bmp"
    SavePicture Extras.BASEBalancePic, App.Path & "\BASEBalancePic.bmp"
    SavePicture Extras.BASEMonoSter, App.Path & "\BASEMonoSter.bmp"
    SavePicture Playlist.BASEPLEdit, App.Path & "\BASEPLEdit.bmp"
    UnSelectedForeColor = RGB(0, &HFF, 0)
    CurrentForeColor = RGB(&HFF, &HFF, &HFF)
    SelectedBackColor = RGB(0, 0, &HC6)
    Playlist.ListView.BackColor = RGB(0, 0, 0)
    Playlist.PLEdit = LoadPicture(App.Path & "\BASEPLEdit.bmp")
    Buttons = LoadPicture(App.Path & "\BASEButtons.bmp")
    'Main = LoadPicture(VBAmpMain.BASEMain.Picture)
    Main = LoadPicture(App.Path & "\BASEMain.bmp")
    PlayPaus = LoadPicture(App.Path & "\BASEPlayPaus.bmp")
    Numbers = LoadPicture(App.Path & "\BASENumbers.bmp")
    PosBar = LoadPicture(App.Path & "\BASEPosBar.bmp")
    Text = LoadPicture(App.Path & "\BASEText.bmp")
    Shuffle = LoadPicture(App.Path & "\BASEShuffle.bmp")
    Bars = LoadPicture(App.Path & "\BASEBars.bmp")
    VolumePic = LoadPicture(App.Path & "\BASEVolumePic.bmp")
    BalancePic = LoadPicture(App.Path & "\BASEBalancePic.bmp")
    MonoSter = LoadPicture(App.Path & "\BASEMonoSter.bmp")
    Kill App.Path & "\BASEPLEdit.bmp"
    Kill App.Path & "\BASEButtons.bmp"
    Kill App.Path & "\BASEMain.bmp"
    Kill App.Path & "\BASEPlayPaus.bmp"
    Kill App.Path & "\BASENumbers.bmp"
    Kill App.Path & "\BASEPosBar.bmp"
    Kill App.Path & "\BASEText.bmp"
    Kill App.Path & "\BASEShuffle.bmp"
    Kill App.Path & "\BASEBars.bmp"
    Kill App.Path & "\BASEVolumePic.bmp"
    Kill App.Path & "\BASEBalancePic.bmp"
    Kill App.Path & "\BASEMonoSter.bmp"
    For I = Playlist.Item.LBound To Playlist.Item.UBound
        Playlist.Item(I).Font.Name = "Arial"
        Playlist.Item(I).BackColor = Playlist.ListView.BackColor
        Playlist.Time(I).Font.Name = "Arial"
        Playlist.Time(I).BackColor = Playlist.ListView.BackColor
    Next
    Playlist.TestLabel3.Font.Name = "Arial"
    Playlist.TestLabel2.Font.Name = "Arial"
    Playlist.TestLabel.Font.Name = "Arial"

End Sub

Private Sub ScrollText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        ChoosingScrollText = True
        OrigX = X
    End If

End Sub

Private Sub ScrollText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ChoosingScrollText = True Then
        If (SongMidNum - Round((X - OrigX) / 5)) > Len(SongString) Then
            SongMidNum = 1
        ElseIf (SongMidNum - Round((X - OrigX) / 5)) < 1 Then
            SongMidNum = Len(SongString)
        Else
            SongMidNum = SongMidNum - Round((X - OrigX) / 5)
        End If
        OrigX = X
        
    End If

End Sub

Private Sub ScrollText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then ChoosingScrollText = False

End Sub

Private Sub SliderBalance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseOver
    If Button = 1 And ButtonOver.Balance = True Then
    If X < BalanceNewX Or X > BalanceNewX + 14 Then
        If X >= 7 And X < 18 Or X > 20 And X <= 31 Then BalanceNewX = X - 105
        If X < 7 Then BalanceNewX = 0
        If X > 31 Then BalanceNewX = 24
        If X >= 18 And X <= 20 Then BalanceNewX = 12
    End If
    ChoosingBalance = True
    UpdateSliders
    OrigX = X
    OrigPosition = OrigX - BalanceNewX
    End If

End Sub

Private Sub SliderBalance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ChoosingBalance = True Then
    If BalanceNewX + (X - OrigX) >= 0 And BalanceNewX + (X - OrigX) < 11 Then
        BalanceNewX = BalanceNewX + (X - OrigX)
        OrigX = X
    ElseIf BalanceNewX + (X - OrigX) > 13 And BalanceNewX + (X - OrigX) <= 24 Then
        BalanceNewX = BalanceNewX + (X - OrigX)
        OrigX = X
    ElseIf BalanceNewX + (X - OrigX) < 0 Then
        BalanceNewX = 0
        OrigX = OrigPosition
    ElseIf BalanceNewX + (X - OrigX) > 24 Then
        BalanceNewX = 24
        OrigX = 24 + OrigPosition
    ElseIf BalanceNewX + (X - OrigX) >= 11 And BalanceNewX + (X - OrigX) <= 13 Then
        BalanceNewX = 12
        OrigX = 12 + OrigPosition
    End If
    UpdateSliders
    End If

End Sub

Private Sub SliderBalance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then ChoosingBalance = False
    UpdateSliders

End Sub

Private Sub SliderPos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
    If WinMedia.PlayState = mpPaused Or WinMedia.PlayState = mpPlaying Then
    If X < SongPosX Or X > SongPosX + 29 Then
        If X >= 15 And X <= SliderPos.Width - 15 Then SongPosNewX = X - 15
        If X < 15 Then SongPosNewX = 0
        If X > SliderPos.Width - 15 Then SongPosNewX = 220
    End If
    ChoosingSongPos = True
    OrigX = X
    OrigPosition = OrigX - SongPosNewX
    UpdateSliders
    End If
    End If

End Sub

Private Sub SliderPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ChoosingSongPos = True Then
    If WinMedia.PlayState = mpStopped Then
        ChoosingSongPos = False
        Exit Sub
    End If
    If SongPosNewX + (X - OrigX) >= 0 And SongPosNewX + (X - OrigX) <= 219 Then
        SongPosNewX = SongPosNewX + (X - OrigX)
        OrigX = X
    ElseIf SongPosNewX + (X - OrigX) < 0 Then
        SongPosNewX = 0
        OrigX = OrigPosition
    ElseIf SongPosNewX + (X - OrigX) > 219 Then
        SongPosNewX = 219
        OrigX = 219 + OrigPosition
    End If
    End If

End Sub

Private Sub SliderPos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
    If SongPosNewX + (X - OrigX) >= 0 And SongPosNewX + (X - OrigX) <= 219 Then
        SongPosNewX = SongPosNewX + (X - OrigX)
        OrigX = X
    ElseIf SongPosNewX + (X - OrigX) < 0 Then
        SongPosNewX = 0
        OrigX = OrigPosition
    ElseIf SongPosNewX + (X - OrigX) > 219 Then
        SongPosNewX = 219
        OrigX = 219 + OrigPosition
    End If
    UpdateSliders
    WinMedia.CurrentPosition = Position
    ChoosingSongPos = False
    End If

End Sub

Private Sub SliderVolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseOver
    If Button = 1 And ButtonOver.Volume = True Then
    If X < VolNewX Or X > VolNewX + 14 Then
        If X >= 7 And X <= 61 Then VolNewX = X - 7
        If X < 7 Then VolNewX = 0
        If X > 61 Then VolNewX = 54
    End If
    ChoosingVolume = True
    UpdateSliders
    OrigX = X
    OrigPosition = OrigX - VolNewX
    End If
    

End Sub

Private Sub SliderVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ChoosingVolume = True Then
    If VolNewX + (X - OrigX) >= 0 And VolNewX + (X - OrigX) <= 54 Then
        VolNewX = VolNewX + (X - OrigX)
        OrigX = X
    ElseIf VolNewX + (X - OrigX) < 0 Then
        VolNewX = 0
        OrigX = OrigPosition
    ElseIf VolNewX + (X - OrigX) > 54 Then
        VolNewX = 54
        OrigX = 54 + OrigPosition
    End If
    UpdateSliders
    End If

End Sub

Private Sub SliderVolume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then ChoosingVolume = False
    UpdateSliders

End Sub
Public Sub StopMedia()
    
    WinMedia.Stop
    WinMedia.CurrentPosition = 0
    UpdatePictures
    
End Sub

Private Sub SliderWSPos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
    If WinMedia.PlayState = mpPlaying Or WinMedia.PlayState = mpPaused Then
    If X < WSSongPosX Or X > WSSongPosX + 3 Then
        If X >= 1 And X <= SliderWSPos.Width - 1 Then WSSongPosNewX = X - 1
        If X < 1 Then WSSongPosNewX = 0
        If X > SliderWSPos.Width - 1 Then WSSongPosNewX = 14
    End If
    ChoosingWSSongPos = True
    UpdateSliders
    End If
    End If

End Sub

Private Sub SliderWSPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ChoosingWSSongPos = True Then
    If X >= 1 And X <= SliderWSPos.Width - 1 Then
    WSSongPosNewX = X - 1
    End If
    If X < 1 Then WSSongPosNewX = 0
    If X > SliderWSPos.Width - 1 Then WSSongPosNewX = 14
    End If

End Sub

Private Sub SliderWSPos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    WinMedia.CurrentPosition = WSPosition
    ChoosingWSSongPos = False
    UpdateSliders

End Sub

Private Sub Stereo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseDown
End Sub

Private Sub Stereo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MouseUp
End Sub

Private Sub Timer1_Timer()

    '50
    If PLBln = False And Me.WindowState <> vbMinimized Then
        Playlist.Visible = False
    ElseIf Me.WindowState <> vbMinimized Then
        Playlist.Visible = True
    End If
    
    If WinMedia.CurrentPosition >= WinMedia.Duration - 1 And WinMedia.PlayState = mpPlaying Then NextSong
    
    If ChoosingScrollText = True Then
        If Len(SongString) > 30 Then
        If SongMidNum <> 1 Then SongString2 = Mid(SongString, SongMidNum, Len(SongString)) & Left(SongString, SongMidNum - 1)
        If SongMidNum = 1 Then SongString2 = Mid(SongString, SongMidNum, Len(SongString))
        End If
        StringToPicture SongString2, ScrollText, 4, 5, 30
    End If
    
    If WinMedia.PlayState <> mpPaused And WinShade = True Then WinShadeTime.Visible = True
    
    If WinMedia.PlayState = mpPaused And WinShade = True And ChoosingWSSongPos = False Then
        If PauseBln = True Then
            StringToPicture "   :", WinShadeTime, 3, 0, 6
        ElseIf PauseBln = False Then
            If Elapsed = False Then StringToPicture "-" & GetTime(WinMedia.Duration - WinMedia.CurrentPosition, True), WinShadeTime, 3, 0, 6
            If Elapsed = True Then StringToPicture " " & GetTime(WinMedia.CurrentPosition, True), WinShadeTime, 3, 0, 6
        End If
    End If
    If ChoosingSongPos = True Then UpdateSliders
    If ButtonDown = True Then
        MouseOver
        UpdatePictures
    End If
    
End Sub

Private Sub Timer2_Timer()
'On Error Resume Next

    '180
    'if a new song is selected set the new song to the songstring
    If SongText.Caption <> "VBAmp 1.1: It Really Whips Winamp's Ass" Then
    If Len(SongText.Caption) + 7 <= 30 And SongString <> SongText.Caption & " (" & GetTime(NewMP3.Seconds) & ")" Then
        SongString = SongText.Caption & " (" & GetTime(NewMP3.Seconds) & ")"
        SongMidNum = 1
    End If
    
    If Len(SongText.Caption) + 7 > 30 And SongString <> SongText.Caption & " (" & GetTime(NewMP3.Seconds) & ")  ***  " Then
        SongString = SongText.Caption & " (" & GetTime(NewMP3.Seconds) & ")  ***  "
        SongMidNum = 1
    End If
    End If
    If SongText.Caption = "VBAmp 1.1: It Really Whips Winamp's Ass" Then
    If Len(SongText.Caption) > 30 And SongString <> SongText.Caption & "  ***  " Then
        SongString = SongText.Caption & "  ***  "
        SongMidNum = 1
    End If
    End If
    
    'scroll songs that are longer than the display
    If ChoosingScrollText = False Then
    If Len(SongString) > 30 Then
        If SongMidNum <> 1 Then SongString2 = Mid(SongString, SongMidNum, Len(SongString)) & Left(SongString, SongMidNum - 1)
        If SongMidNum = 1 Then SongString2 = Mid(SongString, SongMidNum, Len(SongString))
        If ChoosingVolume = False And ChoosingBalance = False And ChoosingSongPos = False And ChoosingScrollText = False Then SongMidNum = SongMidNum + 1
        If SongMidNum > Len(SongString) Then SongMidNum = 1
    End If
    End If
    
    'don't scroll the songs that can fit in the display
    If Len(SongString) <= 30 Then
        SongString2 = SongString
    End If
    
    UpdateSliders

End Sub

Private Sub Timer3_Timer()
'On Error Resume Next
    '1000
    'make times blink if paused
    If WinMedia.PlayState = mpPaused Then
    If BlankMin.Visible = True Then
        BlankMin.Visible = False
        BlankSec.Visible = False
    ElseIf BlankMin.Visible = False Then
        BlankMin.Visible = True
        BlankSec.Visible = True
    End If
    Else
    BlankMin.Visible = False
    BlankSec.Visible = False
    End If
    
    If WinMedia.PlayState = mpPaused And WinShade = True And ChoosingWSSongPos = False Then
        If PauseBln = True Then
            PauseBln = False
        ElseIf PauseBln = False Then
            PauseBln = True
        End If
    End If

    If ChoosingSongPos = False Then
        If WinMedia.Duration <> 0 Then SongPosNewX = (WinMedia.CurrentPosition / WinMedia.Duration) * 220
    End If
    
    If ChoosingWSSongPos = False Then
        If WinMedia.Duration <> 0 Then WSSongPosNewX = (WinMedia.CurrentPosition / WinMedia.Duration) * 14
    End If

End Sub

Private Sub Timer4_Timer()
    
    '500
    MouseOver
    UpdatePictures

End Sub

Private Sub WinShadeNext_Click()
    NextSong
End Sub

Private Sub WinShadeNext2_Click()
    NextSong
End Sub

Private Sub WinShadeOpen_Click()
    OpenAFile
End Sub

Private Sub WinShadeOpen2_Click()
    OpenAFile
End Sub

Private Sub WinShadePause_Click()
    Pause
End Sub

Private Sub WinShadePause2_Click()
    Pause
End Sub

Private Sub WinShadePlay_Click()
    Play
End Sub

Private Sub WinShadePlay2_Click()
    Play
End Sub

Private Sub WinShadePrev_Click()
    PrevSong
End Sub

Private Sub WinShadePrev2_Click()
    PrevSong
End Sub

Private Sub WinShadeStop_Click()
    StopMedia
End Sub

Private Sub WinShadeStop2_Click()
    StopMedia
End Sub

Private Sub WinShadeTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        If Elapsed = True Then
            Elapsed = False
        ElseIf Elapsed = False Then
            Elapsed = True
        End If
        UpdatePictures
    End If

End Sub
