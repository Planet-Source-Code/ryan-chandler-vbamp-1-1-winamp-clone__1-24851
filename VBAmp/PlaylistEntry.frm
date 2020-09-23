VERSION 5.00
Begin VB.Form PlaylistEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Playlist Entry"
   ClientHeight    =   1245
   ClientLeft      =   3630
   ClientTop       =   4485
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2445
      TabIndex        =   5
      Top             =   855
      Width           =   1095
   End
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1215
      TabIndex        =   4
      Top             =   855
      Width           =   1095
   End
   Begin VB.TextBox NewEntry 
      Height          =   285
      Left            =   630
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox OldEntry 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   630
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   3975
   End
   Begin VB.Label NewLabel 
      Caption         =   "New"
      Height          =   240
      Left            =   195
      TabIndex        =   3
      Top             =   510
      Width           =   360
   End
   Begin VB.Label OldLabel 
      Caption         =   "Old"
      Height          =   240
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   285
   End
End
Attribute VB_Name = "PlaylistEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()

    VBAmpMain.Enabled = True
    Playlist.Enabled = True
    Unload Me

End Sub

Private Sub OK_Click()

    Playlist.Playlist2.ListItems.Item(SelectedIndex + 1).Text = NewEntry.Text
    VBAmpMain.Enabled = True
    Playlist.Enabled = True
    Unload Me

End Sub
