VERSION 5.00
Begin VB.Form SkinBrowser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Skins"
   ClientHeight    =   4530
   ClientLeft      =   3930
   ClientTop       =   2655
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ShowInTaskbar   =   0   'False
   Begin VB.Frame SkinsFrame 
      Caption         =   "Skins"
      Height          =   4500
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   4230
      Begin VB.TextBox Description 
         BackColor       =   &H80000004&
         Height          =   1095
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3075
         Width           =   3915
      End
      Begin VB.ListBox SkinBrowser 
         Height          =   2790
         ItemData        =   "SkinBrowser.frx":0000
         Left            =   150
         List            =   "SkinBrowser.frx":0002
         TabIndex        =   2
         Top             =   255
         Width           =   3915
      End
   End
   Begin VB.DirListBox Dir 
      Height          =   1215
      Left            =   1980
      TabIndex        =   0
      Top             =   930
      Width           =   1800
   End
End
Attribute VB_Name = "SkinBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir_Change()

    For I = 0 To Dir.ListCount
        If FileExists(Dir.List(I) & "\" & "main.bmp") = True Then
        SkinBrowser.AddItem GetFileName(Dir.List(I))
        End If
    Next I

End Sub

Private Sub Dir_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then Dir.Path = Dir.List(Dir.ListIndex)
        
End Sub

Private Sub Form_Load()
On Error Resume Next

    Dir.Path = "C:\Program Files\Winamp\Skins"

End Sub

Private Sub OK_Click()
On Error Resume Next

    VBAmpMain.ChangeSkin Dir.Path & "\" & SkinBrowser.List(SkinBrowser.ListIndex)
    VBAmpMain.Enabled = True
    VBAmpMain.Visible = True
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    VBAmpMain.Enabled = True
    Playlist.Enabled = True

End Sub

Private Sub SkinBrowser_Click()
On Error Resume Next
    Description.Text = ""
    If FileExists(Dir.Path & "\" & SkinBrowser.List(SkinBrowser.ListIndex) & "\Readme.txt") = True Then
    Open Dir.Path & "\" & SkinBrowser.List(SkinBrowser.ListIndex) & "\Readme.txt" For Input As #1
    Do Until EOF(1)
    Line Input #1, temporarystring
    If Description.Text = "" Then
        Description.Text = temporarystring
    Else
        Description.Text = Description.Text & vbCrLf & temporarystring
    End If
    Loop
    Close (1)
    End If
    VBAmpMain.ChangeSkin Dir.Path & "\" & SkinBrowser.List(SkinBrowser.ListIndex)
    


End Sub
