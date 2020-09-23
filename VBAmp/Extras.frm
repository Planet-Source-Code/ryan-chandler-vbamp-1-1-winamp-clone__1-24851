VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Extras 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extras"
   ClientHeight    =   4950
   ClientLeft      =   3870
   ClientTop       =   2490
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "Extras.frx":0000
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox VolumePic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   1050
      Picture         =   "Extras.frx":0046
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   0
      Top             =   15
      Width           =   1020
   End
   Begin VB.PictureBox BASEVolumePic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   1050
      Picture         =   "Extras.frx":2D6E
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   3
      Top             =   15
      Width           =   1020
   End
   Begin VB.ListBox SongList 
      Height          =   2400
      ItemData        =   "Extras.frx":5A96
      Left            =   4605
      List            =   "Extras.frx":5A98
      TabIndex        =   21
      Top             =   -15
      Width           =   2190
   End
   Begin VB.PictureBox BASEText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   15
      Picture         =   "Extras.frx":5A9A
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   20
      Top             =   2010
      Width           =   2325
   End
   Begin VB.PictureBox BASEPosBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      Picture         =   "Extras.frx":7984
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   19
      Top             =   1860
      Width           =   4605
   End
   Begin VB.PictureBox BASENumbers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2895
      Picture         =   "Extras.frx":8140
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   18
      Top             =   1320
      Width           =   1485
   End
   Begin VB.PictureBox BASEPlayPaus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2055
      Picture         =   "Extras.frx":893A
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   17
      Top             =   1695
      Width           =   630
   End
   Begin VB.PictureBox BASEMonoSter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2040
      Picture         =   "Extras.frx":8EB2
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   16
      Top             =   1320
      Width           =   870
   End
   Begin VB.PictureBox BASEButtons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      Picture         =   "Extras.frx":9774
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   15
      Top             =   1320
      Width           =   2040
   End
   Begin VB.PictureBox BASEBars 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   15
      Picture         =   "Extras.frx":AD00
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   14
      Top             =   15
      Width           =   5160
   End
   Begin VB.PictureBox MonoSter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2025
      Picture         =   "Extras.frx":EA9A
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   13
      Top             =   1320
      Width           =   870
   End
   Begin VB.PictureBox PlayPaus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2055
      Picture         =   "Extras.frx":F35C
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   12
      Top             =   1695
      Width           =   630
   End
   Begin VB.PictureBox Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   15
      Picture         =   "Extras.frx":F8D4
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   11
      Top             =   2010
      Width           =   2325
   End
   Begin VB.PictureBox Bars 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      Picture         =   "Extras.frx":117BE
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   10
      Top             =   0
      Width           =   5160
   End
   Begin VB.PictureBox Shuffle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2355
      Picture         =   "Extras.frx":15558
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   9
      Top             =   2025
      Width           =   1380
   End
   Begin VB.PictureBox Buttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      Picture         =   "Extras.frx":17040
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   8
      Top             =   1320
      Width           =   2040
   End
   Begin VB.PictureBox PosBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   15
      Picture         =   "Extras.frx":185CC
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   7
      Top             =   1860
      Width           =   4605
   End
   Begin VB.PictureBox Numbers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2895
      Picture         =   "Extras.frx":18D88
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   6
      Top             =   1320
      Width           =   1485
   End
   Begin VB.PictureBox BASEShuffle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2340
      Picture         =   "Extras.frx":19582
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   5
      Top             =   2040
      Width           =   1380
   End
   Begin VB.PictureBox Mouse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3750
      Picture         =   "Extras.frx":1B06A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   2025
      Width           =   480
   End
   Begin VB.PictureBox BASEBalancePic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   15
      Picture         =   "Extras.frx":1B374
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   2
      Top             =   15
      Width           =   1020
   End
   Begin VB.PictureBox BalancePic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   30
      Picture         =   "Extras.frx":1D7FC
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   1
      Top             =   15
      Width           =   1020
   End
   Begin MSComDlg.CommonDialog MP3Opener 
      Left            =   4365
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Skins 
         Caption         =   "Skins"
         Begin VB.Menu Base 
            Caption         =   "Base Skin"
         End
         Begin VB.Menu Change 
            Caption         =   "Change Skin"
         End
      End
   End
   Begin VB.Menu Sort 
      Caption         =   "Sort"
      Begin VB.Menu TitleSort 
         Caption         =   "Sort by Title"
      End
      Begin VB.Menu FileNameSort 
         Caption         =   "Sort by Filename"
      End
      Begin VB.Menu PathFileNameSort 
         Caption         =   "Sort by Path and Filename"
      End
      Begin VB.Menu none 
         Caption         =   "-"
      End
      Begin VB.Menu ReverseSort 
         Caption         =   "Reverse List"
      End
   End
   Begin VB.Menu FileInf 
      Caption         =   "FileInf"
      Begin VB.Menu Info 
         Caption         =   "File Info"
      End
      Begin VB.Menu PlayEntry 
         Caption         =   "Playlist Entry"
      End
   End
End
Attribute VB_Name = "Extras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Base_Click()

    VBAmpMain.ChangeSkin "", True

End Sub

Private Sub Change_Click()
On Error GoTo Error

    Load SkinBrowser
    SkinBrowser.Visible = True
    VBAmpMain.Enabled = False
    Playlist.Enabled = False

Error:
End Sub

Private Sub FileNameSort_Click()
On Error GoTo Error
    
    With Playlist
    .Playlist3.ListItems.Clear
    .Playlist4.ListItems.Clear
    .Playlist5.ListItems.Clear
    For I = 1 To .Playlist.ListItems.Count
        'moves the items to be sorted to playlist3
        .Playlist3.ListItems.Add(, , GetFileName(.Playlist2.ListItems.Item(I).Text)).SubItems(1) = I
        'copies playlist and playlist2
        .Playlist4.ListItems.Add , , .Playlist2.ListItems.Item(I).Text
        .Playlist5.ListItems.Add , , .Playlist.ListItems.Item(I).Text
        'copies first column of seconds list into second column of seconds list (didn't feel like making another listview)
        .SecondsList.ListItems.Item(I).SubItems(1) = .SecondsList.ListItems.Item(I).Text
    Next I
    'sort all the items in playlist3
    .Playlist3.SortOrder = lvwAscending
    .Playlist3.Sorted = True
    .Playlist3.Sorted = False
    For I = 1 To .Playlist.ListItems.Count
        'for setting the selectedindex after the for loop or else the selected index will keep on changing and come out incorrect
        Dim TempSelectedIndex
        If Int(.Playlist3.ListItems.Item(I).SubItems(1)) = SelectedItem Then .Playlist.ListItems.Item(I).Selected = True
        If Int(.Playlist3.ListItems.Item(I).SubItems(1)) - 1 = SelectedIndex Then TempSelectedIndex = I - 1
        .Playlist2.ListItems.Item(I).Text = .Playlist4.ListItems.Item(Int(.Playlist3.ListItems.Item(I).SubItems(1))).Text
        .Playlist.ListItems.Item(I).Text = .Playlist5.ListItems.Item(Int(.Playlist3.ListItems.Item(I).SubItems(1))).Text
        .SecondsList.ListItems.Item(I).Text = .SecondsList.ListItems.Item(Int(.Playlist3.ListItems.Item(I).SubItems(1))).SubItems(1)
    Next I
    SelectedIndex = TempSelectedIndex
    UpdatePlaylist
    End With

Error:
End Sub

Private Sub PathFileNameSort_Click()
On Error GoTo Error
    
    With Playlist
    .Playlist3.ListItems.Clear
    For I = 1 To .Playlist.ListItems.Count
        .Playlist2.ListItems.Item(I).SubItems(1) = I
        .Playlist3.ListItems.Add , , .Playlist.ListItems.Item(I).Text
        .SecondsList.ListItems.Item(I).SubItems(1) = .SecondsList.ListItems.Item(I).Text
    Next I
    .Playlist2.SortOrder = lvwAscending
    .Playlist2.Sorted = True
    .Playlist2.Sorted = False
    For I = 1 To .Playlist.ListItems.Count
        Dim TempSelectedIndex
        If Int(.Playlist2.ListItems.Item(I).SubItems(1)) = SelectedItem Then .Playlist.ListItems.Item(I).Selected = True
        If Int(.Playlist2.ListItems.Item(I).SubItems(1)) - 1 = SelectedIndex Then TempSelectedIndex = I - 1
        .Playlist.ListItems.Item(I).Text = .Playlist3.ListItems.Item(Int(.Playlist2.ListItems.Item(I).SubItems(1))).Text
        .SecondsList.ListItems.Item(I).Text = .SecondsList.ListItems.Item(Int(.Playlist2.ListItems.Item(I).SubItems(1))).SubItems(1)
    Next I
    SelectedIndex = TempSelectedIndex
    UpdatePlaylist
    End With

Error:
End Sub

Private Sub PlayEntry_Click()
On Error GoTo Error
    
    If SelectedIndex <> -1 Then
    Load PlaylistEntry
    PlaylistEntry.OldEntry = Playlist.Playlist2.ListItems.Item(SelectedIndex + 1).Text
    PlaylistEntry.NewEntry = PlaylistEntry.OldEntry
    PlaylistEntry.Visible = True
    VBAmpMain.Enabled = False
    Playlist.Enabled = False
    End If

Error:
End Sub

Private Sub ReverseSort_Click()
On Error GoTo Error

    With Playlist
    .Playlist3.ListItems.Clear
    .Playlist4.ListItems.Clear
    For I = 1 To .Playlist.ListItems.Count
        .Playlist3.ListItems.Add , , .Playlist.ListItems.Item(.Playlist.ListItems.Count - I + 1).Text
        .Playlist4.ListItems.Add , , .Playlist2.ListItems.Item(.Playlist.ListItems.Count - I + 1).Text
        .SecondsList.ListItems.Item(I).SubItems(1) = .SecondsList.ListItems.Item(.Playlist.ListItems.Count - I + 1).Text
    Next I
    For I = 1 To .Playlist.ListItems.Count
        .Playlist.ListItems.Item(I).Text = .Playlist3.ListItems.Item(I).Text
        .Playlist2.ListItems.Item(I).Text = .Playlist4.ListItems.Item(I).Text
        .SecondsList.ListItems.Item(I).Text = .SecondsList.ListItems.Item(I).SubItems(1)
    Next I
    SelectedIndex = .Playlist.ListItems.Count - SelectedIndex + 1
    .Playlist.ListItems.Item(SelectedItem).Selected = True
    UpdatePlaylist
    End With

Error:
End Sub

Private Sub TitleSort_Click()
On Error GoTo Error
    
    With Playlist
    .Playlist3.ListItems.Clear
    For I = 1 To .Playlist.ListItems.Count
        .Playlist.ListItems.Item(I).SubItems(1) = I
        .Playlist3.ListItems.Add , , .Playlist2.ListItems.Item(I).Text
        .SecondsList.ListItems.Item(I).SubItems(1) = .SecondsList.ListItems.Item(I).Text
    Next I
    .Playlist.SortOrder = lvwAscending
    .Playlist.Sorted = True
    .Playlist.Sorted = False
    For I = 1 To .Playlist.ListItems.Count
        Dim TempSelectedIndex
        If Int(.Playlist.ListItems.Item(I).SubItems(1)) = SelectedItem Then .Playlist.ListItems.Item(I).Selected = True
        If Int(.Playlist.ListItems.Item(I).SubItems(1)) - 1 = SelectedIndex Then TempSelectedIndex = I - 1
        .Playlist2.ListItems.Item(I).Text = .Playlist3.ListItems.Item(Int(.Playlist.ListItems.Item(I).SubItems(1))).Text
        .SecondsList.ListItems.Item(I).Text = .SecondsList.ListItems.Item(Int(.Playlist.ListItems.Item(I).SubItems(1))).SubItems(1)
    Next I
    SelectedIndex = TempSelectedIndex
    UpdatePlaylist
    End With

Error:
End Sub
