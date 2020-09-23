Attribute VB_Name = "VBAmpModule"
Type POINTAPI
    X As Long
    Y As Long
End Type

'functions for borderless forms that show in the taskbar
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'function to get the position of the mouse
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'function to get active window
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Integer
'functions used to get the top of the start menu
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Global FormOver As Boolean
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Global PlaylistSnapped As Boolean
'type for telling the program which button the mouse is over
Type OverType
    'for changing the pictures on the form
    Menu As Boolean
    Minimize As Boolean
    WinShade As Boolean
    Exit As Boolean
    Previous As Boolean
    Play As Boolean
    Pause As Boolean
    Stop As Boolean
    Next As Boolean
    Open As Boolean
    Shuffle As Boolean
    Repeat As Boolean
    PL As Boolean
    'for changing mouse icons
    SongScroll As Boolean
    SongPos As Boolean
    Volume As Boolean
    Balance As Boolean
    Bar As Boolean
    WSSongPos As Boolean
End Type

Type PlayOverType
    Exit As Boolean
    Shade As Boolean
    AddFile As Boolean
    AddDir As Boolean
    AddURL As Boolean
    RemMisc As Boolean
    RemAll As Boolean
    RemCrop As Boolean
    RemFile As Boolean
    SelInverse As Boolean
    SelZero As Boolean
    SelAll As Boolean
    MisSort As Boolean
    MisInfo As Boolean
    MisOptions As Boolean
    LstLoad As Boolean
    LstSave As Boolean
    LstNew As Boolean
End Type

'mouse position var
Global MousePos As POINTAPI
'mouseover buttons var
Global ButtonOver As OverType
Global PlayOver As PlayOverType

'for borderless forms
Global Const GWL_STYLE = (-16)
Global Const WS_DLGFRAME = &H400000

'if this is true then you are left clicking
Global ButtonDown As Boolean
Global PlayButtonDown As Boolean

'Vars for checking which form the focus is on
Global MainFocus As Boolean
Global MainFoc As Integer
Global PlaylistFocus As Boolean

'boolean to tell if you are in winshade mode or not
Global WinShade As Boolean

'the "D"'s stand for down; for checking to see if the mouse is down over the buttons
Global PlayExitD As Boolean
Global PlayShadeD As Boolean
Global ExitD As Boolean
Global MinD As Boolean
Global MenuD As Boolean
Global ShadeD As Boolean
Global PrevD As Boolean
Global PlayD As Boolean
Global PauseD As Boolean
Global StopD As Boolean
Global NextD As Boolean
Global OpenD As Boolean
'for the repeat and shuffle there are two booleans because there are four pictures for the buttons
Global RepeatD As Boolean
Global RepeatBln As Boolean
Global ShuffleD As Boolean
Global ShuffleBln As Boolean
'for eq and pl buttons it is the same as repeat and shuffle
Global PLD As Boolean
Global PLBln As Boolean


Global SongNewY
'Volume Vars
Global Volume As Integer
Global VolVar
Global VolNewX
'Balance Vars
Global Balance As Integer
Global BalanceVar
Global BalanceNewX
'PosBar Vars
Global Position As Integer
Global SongPosX
Global SongPosNewX
'WinShade PosBar Vars
Global WSPosition As Integer
Global WSSongPosX
Global WSSongPosNewX
'Vars used when moving things
Global Choosing As Boolean
Global PlayChoosing As Boolean
Global ChoosingVolume As Boolean
Global ChoosingBalance As Boolean
Global ChoosingSongPos As Boolean
Global ChoosingScrollText As Boolean
Global ChoosingWSSongPos As Boolean
Global ChoosingSongNumber As Boolean
Global SongNumber As Integer
Global OldSongNumber As Integer

Global OrigPosition
Global OrigX
Global OrigY
'Vars for position of mouse
Global PosX
Global PosY
'Vars used instead of me.left and me.top(saves time and space)
Global X
Global Y

Global TextString As String
Global TextString2 As String
Global TextString3 As String

Global VolumeString As String
Global BalanceString As String
Global SongPosString As String
Global WSSongPosString As String
Global SongString As String
Global SongString2 As String
Global SongMidNum

Global Minutes As Integer
Global Seconds As Integer
Global MinuteSec As String

Global Elapsed As Boolean
Global NewMP3 As New MP3Info
Global AnotherMP3 As New MP3Info
Global ShadeY As Integer
Global MainY As Integer
Global MovePlaylist As Boolean

Global PauseBln As Boolean
Global CurrentForeColor
Global UnSelectedForeColor
Global BackColor
Global SelectedItem As Integer
Global SelectedBackColor
Global SelectedIndex As Integer

Global XSnapped As Boolean
Global YSnapped As Boolean

Global AddMenu As Boolean
Global RemMenu As Boolean
Global SelMenu As Boolean
Global MisMenu As Boolean
Global LstMenu As Boolean

Global PopMenu As Boolean
Global ButtonNum As Integer
Global ButtonNum2 As Integer

Global Mins
Global RemMins
Global Secs
Global RemSecs

Global DblClicked As Boolean

Public Sub UpdatePictures(Optional Update As Boolean)
'On Error Resume Next

    With VBAmpMain
    'check to see which form the focus is on (if any)
    If .hwnd <> GetActiveWindow() Then
        MainFocus = False
        .ButtonExit2.Visible = True
        .ButtonMenu2.Visible = True
        .ButtonMin2.Visible = True
        If WinShade = False Then
            .ButtonWinShade2.Visible = True
            .Bar(2).Visible = True
            .Bar(3).Visible = False
        Else
            .Bar(3).Visible = True
        End If
        Choosing = False
        ChoosingVolume = False
        ChoosingBalance = False
        ChoosingSongPos = False
        ChoosingScrollText = False
        ChoosingWSSongPos = False
    Else
        If WinShade = False Then
            .Bar(0).Visible = True
            .Bar(1).Visible = False
        Else
            .Bar(1).Visible = True
        End If
        .Bar(3).Visible = False
        .Bar(2).Visible = False
        MainFocus = True
        .ButtonWinShadeTrue2.Visible = False
        .ButtonWinShade2.Visible = False
        .ButtonExit2.Visible = False
        .ButtonMenu2.Visible = False
        .ButtonMin2.Visible = False
    End If

    If ButtonNum = 2 Then
        If StopD = False And .ButtonStopD.Visible = True Then
            .ButtonStopD.Visible = False
            .StopMedia
        End If
        If PlayD = False And .ButtonPlayD.Visible = True Then
            .ButtonPlayD.Visible = False
            .Play
        End If
        If PauseD = False And .ButtonPauseD.Visible = True Then
            .ButtonPauseD.Visible = False
            .Pause
        End If
        If PrevD = False And .ButtonPrevD.Visible = True Then
            .ButtonPrevD.Visible = False
            .PrevSong
        End If
        If NextD = False And .ButtonNextD.Visible = True Then
            .ButtonNextD.Visible = False
            .NextSong
        End If
        If OpenD = False And .ButtonOpenD.Visible = True Then
            .ButtonOpenD.Visible = False
            .OpenAFile
        End If
        If ShuffleD = False Then
            If ShuffleBln = True Then
                .ButtonShuffleD.Visible = False
            Else
                .ButtonShuffleTrueD.Visible = False
            End If
        End If
        If RepeatD = False Then
            If RepeatBln = True Then
                .ButtonRepeatD.Visible = False
            Else
                .ButtonRepeatTrueD.Visible = False
            End If
        End If
        If PLD = False Then
            If PLBln = True Then
                .PLD.Visible = False
            Else
                .PLTrueD.Visible = False
            End If
        End If
        If ExitD = False And .ButtonExitD.Visible = True Then
            End
        End If
        If MinD = False And .ButtonMinD.Visible = True Then
            Playlist.Visible = False
            .WindowState = vbMinimized
            .ButtonMinD.Visible = False
        End If
        If ShadeD = False Then
            If .ButtonWinShadeD.Visible = True Or .ButtonWinShadeTrueD.Visible = True Or DblClicked = True Then
            DblClicked = False
            If WinShade = False Then
                WinShade = True
                .WinShadeTime.Visible = True
                .WinShadeNext.Visible = True
                .WinShadeOpen.Visible = True
                .WinShadePause.Visible = True
                .WinShadePlay.Visible = True
                .WinShadePrev.Visible = True
                .WinShadeStop.Visible = True
                .SliderWSPos.Visible = True
                .ButtonWinShadeTrue.Visible = True
                .Bar(1).Visible = True
                If Playlist.Top = .Top + .Height Then MovePlaylist = True
                .Width = VBAmpMain.XTwips(.Bar(0).Width)
                .Height = VBAmpMain.YTwips(.Bar(0).Height)
                If MovePlaylist = True Then Playlist.Top = .Top + .Height
                MovePlaylist = False
            ElseIf WinShade = True Then
                WinShade = False
                .WinShadeTime.Visible = False
                .WinShadeNext.Visible = False
                .WinShadeOpen.Visible = False
                .WinShadePause.Visible = False
                .WinShadePlay.Visible = False
                .WinShadePrev.Visible = False
                .WinShadeStop.Visible = False
                .SliderWSPos.Visible = False
                .ButtonWinShadeTrue.Visible = False
                .WinShadeNext2.Visible = False
                .WinShadeOpen2.Visible = False
                .WinShadePause2.Visible = False
                .WinShadePlay2.Visible = False
                .WinShadePrev2.Visible = False
                .WinShadeStop2.Visible = False
                .ButtonWinShadeTrue2.Visible = False
                .Bar(1).Visible = False
                .Bar(3).Visible = False
                If Playlist.Top = .Top + .Height Then MovePlaylist = True
                .Width = VBAmpMain.XTwips(.Main.Width)
                .Height = VBAmpMain.YTwips(.Main.Height)
                If MovePlaylist = True Then Playlist.Top = .Top + .Height
                MovePlaylist = False
            End If
            .ButtonWinShadeD.Visible = False
            .ButtonWinShadeTrueD.Visible = False
            End If
        End If
    End If
    
    'Exit Button
    If ButtonDown = True And ButtonOver.Exit = True Then
        ExitD = True
        .ButtonExitD.Visible = True
    Else
        If ButtonDown = True Then
            ExitD = False
            .ButtonExitD.Visible = False
        End If
    End If
    If ButtonDown = False And ExitD = True Then
        ExitD = False
        ButtonNum = 1
    End If
    
    'Minimize Button
    If ButtonDown = True And ButtonOver.Minimize = True Then
        MinD = True
        .ButtonMinD.Visible = True
    Else
        If ButtonDown = True Then
            MinD = False
            .ButtonMinD.Visible = False
        End If
    End If
    If ButtonDown = False And MinD = True Then
        MinD = False
        ButtonNum = 1
    End If
    
    'Menu Button
    If ButtonDown = True And ButtonOver.Menu = True Then
        .ButtonMenuD.Visible = True
        ButtonDown = False
        'winamp style menu opening(winamp opens a menu as soon as the mouse is down and over the menu button)
        'make the extras form in the same position so you can get the popupmenu in the place you want it to be
        Extras.Left = .Left
        Extras.Top = .Top
        'don't ask me how I got these coordinates (guess and check)
        Extras.PopupMenu Extras.Options, 2, 3, -25
        .ButtonMenuD.Visible = False
    End If
    
    'WinShade Button
    If ButtonDown = True And ButtonOver.WinShade = True Then
        ShadeD = True
        If WinShade = False Then
            .ButtonWinShadeD.Visible = True
        Else
            .ButtonWinShadeTrueD.Visible = True
        End If
    Else
        If ButtonDown = True Then
        ShadeD = False
        If WinShade = False Then
            .ButtonWinShadeD.Visible = False
        Else
            .ButtonWinShadeTrueD.Visible = False
        End If
        End If
    End If
    If ButtonDown = False And ShadeD = True Then
        ShadeD = False
        ButtonNum = 1
    End If
    
    'Previous Button
    If ButtonDown = True And ButtonOver.Previous = True Then
        PrevD = True
        .ButtonPrevD.Visible = True
    Else
        If ButtonDown = True Then
        PrevD = False
        .ButtonPrevD.Visible = False
        End If
    End If
    If ButtonDown = False And PrevD = True Then
        PrevD = False
        ButtonNum = 1
    End If

    'Play Button
    If ButtonDown = True And ButtonOver.Play = True Then
        PlayD = True
        .ButtonPlayD.Visible = True
    Else
        If ButtonDown = True Then
            PlayD = False
            .ButtonPlayD.Visible = False
        End If
    End If
    If ButtonDown = False And PlayD = True Then
        PlayD = False
        ButtonNum = 1
    End If
    
    'Pause Button
    If ButtonDown = True And ButtonOver.Pause = True Then
        .ButtonPauseD.Visible = True
        PauseD = True
    Else
        If ButtonDown = True Then
        PauseD = False
        .ButtonPauseD.Visible = False
        End If
    End If
    If ButtonDown = False And PauseD = True Then
        PauseD = False
        ButtonNum = 1
    End If
    
    'Stop Button
    If ButtonDown = True And ButtonOver.Stop = True Then
        StopD = True
        .ButtonStopD.Visible = True
    Else
        If ButtonDown = True Then
            StopD = False
            .ButtonStopD.Visible = False
        End If
    End If
    If ButtonDown = False And StopD = True Then
        StopD = False
        ButtonNum = 1
    End If
    
    'Next Button
    If ButtonDown = True And ButtonOver.Next = True Then
        NextD = True
        .ButtonNextD.Visible = True
    Else
        If ButtonDown = True Then
        NextD = False
        .ButtonNextD.Visible = False
        End If
    End If
    If ButtonDown = False And NextD = True Then
        NextD = False
        ButtonNum = 1
    End If
    
    'Open Button
    If ButtonDown = True And ButtonOver.Open = True Then
        OpenD = True
        .ButtonOpenD.Visible = True
    Else
        If ButtonDown = True Then
        OpenD = False
        .ButtonOpenD.Visible = False
        End If
    End If
    If ButtonDown = False And OpenD = True Then
        OpenD = False
        ButtonNum = 1
    End If
    
    'Shuffle Button
    If ButtonDown = True And ButtonOver.Shuffle = True Then
        ShuffleD = True
        If ShuffleBln = False Then .ButtonShuffleD.Visible = True
        If ShuffleBln = True Then .ButtonShuffleTrueD.Visible = True
    Else
        If ButtonDown = True Then
            ShuffleD = False
            .ButtonShuffleD.Visible = False
            .ButtonShuffleTrueD.Visible = False
        End If
    End If
    If ButtonDown = False And ShuffleD = True Then
        ShuffleD = False
        If ShuffleBln = False Then
        ShuffleBln = True
        .ButtonShuffleTrue.Visible = True
        Else
        ShuffleBln = False
        .ButtonShuffleTrue.Visible = False
        End If
        ButtonNum = 1
    End If
    
    'Repeat Button
    If ButtonDown = True And ButtonOver.Repeat = True Then
        RepeatD = True
        If RepeatBln = False Then .ButtonRepeatD.Visible = True
        If RepeatBln = True Then .ButtonRepeatTrueD.Visible = True
    Else
        If ButtonDown = True Then
            RepeatD = False
            .ButtonRepeatD.Visible = False
            .ButtonRepeatTrueD.Visible = False
        End If
    End If
    If ButtonDown = False And RepeatD = True Then
        RepeatD = False
        If RepeatBln = False Then
        RepeatBln = True
        .ButtonRepeatTrue.Visible = True
        Else
        RepeatBln = False
        .ButtonRepeatTrue.Visible = False
        End If
        ButtonNum = 1
    End If
    
    'PL Button
    If ButtonDown = True And ButtonOver.PL = True Then
        PLD = True
        If PLBln = False Then .PLD.Visible = True
        If PLBln = True Then .PLTrueD.Visible = True
    Else
        If ButtonDown = True Then
            PLD = False
            .PLD.Visible = False
            .PLTrueD.Visible = False
        End If
    End If
    If ButtonDown = False And PLD = True Then
        PLD = False
        If PLBln = False Then
        PLBln = True
        Else
        PLBln = False
        End If
        ButtonNum = 1
    End If
    If PLBln = True Then .PLTrue.Visible = True Else .PLTrue.Visible = False

    If .WinMedia.PlayState = mpPlaying Or .WinMedia.PlayState = mpPaused Then
        NewMP3.FileName = .WinMedia.FileName
        'fill with the bitrate and frequency
        If NewMP3.BitRate < 100 Then StringToPicture " " & NewMP3.BitRate, .KBPS, 4, 4, 3
        If NewMP3.BitRate >= 100 Then StringToPicture NewMP3.BitRate, .KBPS, 4, 4, 3
        StringToPicture Round(NewMP3.Frequency / 1000), .KHZ, 4, 4, 2
        'make either mono or stereo lit up
        If NewMP3.Mode = "Mono" Then
            .Mono.PaintPicture Extras.MonoSter, 0, 0, , , 29, 0
            .Stereo.PaintPicture Extras.MonoSter, 0, 0, , , 0, 12
        Else
            .Mono.PaintPicture Extras.MonoSter, 0, 0, , , 29, 12
            .Stereo.PaintPicture Extras.MonoSter, 0, 0, , , 0, 0
        End If
        .KBPS.Visible = True
        .KHZ.Visible = True
        If WinShade = True Then .WinShadeTime.Visible = True
        .SliderPos.Visible = True
        If .WinMedia.PlayState = mpPlaying Then
            If ChoosingWSSongPos = False Then
            If Elapsed = False Then StringToPicture "-" & GetTime(.WinMedia.Duration - .WinMedia.CurrentPosition, True), .WinShadeTime, 3, 0, 6
            If Elapsed = True Then StringToPicture " " & GetTime(.WinMedia.CurrentPosition, True), .WinShadeTime, 3, 0, 6
            Else
            StringToPicture WSSongPosString, .WinShadeTime, 3, 0, 6
            End If
            'show the play symbol next to where it shows the current position in the song
            'i made it so it didn't show the left most column of pixels, because on one of my skins that part was discolored yet it didn't show up in winamp
            .Indicator.PaintPicture .Main, 0, 0, , , .Indicator.Left, .Indicator.Top, 1
            .Indicator.PaintPicture Extras.PlayPaus, 1, 0, , , 1, 0, 8, 9
            'fill the time remaining and the time elapsed in their picture boxes
            If Mins <> GetMinutes(Int(.WinMedia.CurrentPosition)) Or Update = True Then
                NumbertoPicture GetMinutes(Int(.WinMedia.CurrentPosition)), .ElapsedMin
                Mins = GetMinutes(Int(.WinMedia.CurrentPosition))
            End If
            If Secs <> GetSeconds(Int(.WinMedia.CurrentPosition)) Or Update = True Then
                NumbertoPicture GetSeconds(Int(.WinMedia.CurrentPosition)), .ElapsedSec
                Secs = GetSeconds(Int(.WinMedia.CurrentPosition))
            End If
            If RemMins <> GetMinutes(Int(.WinMedia.Duration - .WinMedia.CurrentPosition)) Or Update = True Then
                NumbertoPicture GetMinutes(Int(.WinMedia.Duration - .WinMedia.CurrentPosition)), .RemMin
                RemMins = GetMinutes(Int(.WinMedia.Duration - .WinMedia.CurrentPosition))
            End If
            If RemSecs <> GetSeconds(Int(.WinMedia.Duration - .WinMedia.CurrentPosition)) Or Update = True Then
                NumbertoPicture GetSeconds(Int(.WinMedia.Duration - .WinMedia.CurrentPosition)), .RemSec
                RemSecs = GetSeconds(Int(.WinMedia.Duration - .WinMedia.CurrentPosition))
            End If
            If Elapsed = True Then
            'show the elapsed time and no negative sign
            .ElapsedMin.Visible = True
            .ElapsedSec.Visible = True
            .RemMin.Visible = False
            .RemSec.Visible = False
            .Negative.Visible = False
            ElseIf Elapsed = False Then
            'show the remaining time and a negative sign
            .RemMin.Visible = True
            .RemSec.Visible = True
            .ElapsedMin.Visible = False
            .ElapsedSec.Visible = False
            .Negative.Visible = True
            End If
        ElseIf .WinMedia.PlayState = mpPaused Then
        'show pause symbol
            .Indicator.PaintPicture Extras.PlayPaus, 0, 0, , , 9, 0
        End If
    Else
        Secs = -1
        RemSecs = -1
        Mins = -1
        RemMins = -1
        'fill the bitrate and frequency pictures with blanks
        .KBPS.Visible = False
        .KHZ.Visible = False
        'show stop symbol
        .Indicator.PaintPicture Extras.PlayPaus.Picture, 0, 0, , , 18, 0
        'make mono and stereo pictures dimmed
        .Mono.PaintPicture Extras.MonoSter, 0, 0, , , 29, 12
        .Stereo.PaintPicture Extras.MonoSter, 0, 0, , , 0, 12
        'fill the winshade time with nothing
        StringToPicture "   :", .WinShadeTime, 3, 0, 6
        .ElapsedMin.Visible = False
        .ElapsedSec.Visible = False
        .RemMin.Visible = False
        .RemSec.Visible = False
        Minutes = 0
        Seconds = 0
        'don't show any time if it is stopped
        NumbertoPicture 0, .ElapsedMin, True
        NumbertoPicture 0, .ElapsedSec, True
        NumbertoPicture 0, .RemMin, True
        NumbertoPicture 0, .RemSec, True
        .Negative.Visible = False
        .SliderPos.Visible = False
        If WinShade = True Then .SliderWSPos.Visible = False
    End If
    
    If WinShade = True Then
        If .WinMedia.PlayState <> mpPaused Then .WinShadeTime.Visible = True
        If MainFocus = True Then
            .WinShadeNext2.Visible = False
            .WinShadeOpen2.Visible = False
            .WinShadePause2.Visible = False
            .WinShadePlay2.Visible = False
            .WinShadePrev2.Visible = False
            .WinShadeStop2.Visible = False
            .ButtonWinShadeTrue2.Visible = False
        Else
            .WinShadeNext2.Visible = True
            .WinShadeOpen2.Visible = True
            .WinShadePause2.Visible = True
            .WinShadePlay2.Visible = True
            .WinShadePrev2.Visible = True
            .WinShadeStop2.Visible = True
            .ButtonWinShadeTrue2.Visible = True
        End If
    End If

    With Playlist
    TempLong = GetActiveWindow()
    If TempLong = .hwnd And SkinBrowser.Visible = False And Playlist.Visible = False Then SetActiveWindow VBAmpMain.hwnd
    If TempLong = VBAmpMain.hwnd Then SetActiveWindow (VBAmpMain.hwnd)
    End With
    End With
    If ButtonNum < 5 Then ButtonNum = ButtonNum + 1
    
End Sub

Public Sub UpdateFormPlaylist()

    With Playlist
    If .hwnd <> GetActiveWindow() Then
        If PlaylistFocus = True Then
            PlaylistFocus = False
        End If
        PlayChoosing = False
        Menus False
        .Bar(1).Visible = True
        .Exit2.Visible = True
        .Shade2.Visible = True
    Else
        If PlaylistFocus = False Then
            PlaylistFocus = True
        End If
        .Bar(1).Visible = False
        .Exit2.Visible = False
        .Shade2.Visible = False
    End If
    
    If ButtonNum2 = 2 Then
        If PlayExitD = False And .ExitD.Visible = True Then
            .ExitD.Visible = False
            PLBln = False
        End If
        If PlayShadeD = False Then .ShadeD.Visible = False
    End If
    
    If PlayButtonDown = True And PlayOver.Exit = True Then
        PlayExitD = True
        .ExitD.Visible = True
    Else
        If PlayButtonDown = True Then
            PlayExitD = False
            .ExitD.Visible = False
        End If
    End If
    If PlayButtonDown = False And PlayExitD = True Then
        PlayExitD = False
        ButtonNum2 = 1
    End If
    
    If PlayButtonDown = True And PlayOver.Shade = True Then
        PlayShadeD = True
        .ShadeD.Visible = True
    Else
        If PlayButtonDown = True Then
            PlayShadeD = False
            .ShadeD.Visible = False
        End If
    End If
    If PlayButtonDown = False And PlayShadeD = True Then
        PlayShadeD = False
        ButtonNum2 = 1
    End If
            
    If AddMenu = False Then
        .AddMen.Visible = False
        .AddBar.Visible = False
    Else
        If PlayOver.AddURL = False Then .ButtonAdd(0).Visible = False Else .ButtonAdd(0).Visible = True
        If PlayOver.AddDir = False Then .ButtonAdd(1).Visible = False Else .ButtonAdd(1).Visible = True
        If PlayOver.AddFile = False Then .ButtonAdd(2).Visible = False Else .ButtonAdd(2).Visible = True
        .AddMen.Visible = True
        .AddBar.Visible = True
    End If
    
    If RemMenu = False Then
        .RemMen.Visible = False
        .RemBar.Visible = False
    Else
        If PlayOver.RemMisc = False Then .ButtonRem(0).Visible = False Else .ButtonRem(0).Visible = True
        If PlayOver.RemAll = False Then .ButtonRem(1).Visible = False Else .ButtonRem(1).Visible = True
        If PlayOver.RemCrop = False Then .ButtonRem(2).Visible = False Else .ButtonRem(2).Visible = True
        If PlayOver.RemFile = False Then .ButtonRem(3).Visible = False Else .ButtonRem(3).Visible = True
        .RemMen.Visible = True
        .RemBar.Visible = True
    End If
    
    If SelMenu = False Then
        .SelMen.Visible = False
        .SelBar.Visible = False
    Else
        If PlayOver.SelInverse = False Then .ButtonSel(0).Visible = False Else .ButtonSel(0).Visible = True
        If PlayOver.SelZero = False Then .ButtonSel(1).Visible = False Else .ButtonSel(1).Visible = True
        If PlayOver.SelAll = False Then .ButtonSel(2).Visible = False Else .ButtonSel(2).Visible = True
        .SelMen.Visible = True
        .SelBar.Visible = True
    End If
    
    If MisMenu = False Then
        .MisMen.Visible = False
        .MisBar.Visible = False
    Else
        If PopMenu = False Then
            If PlayOver.MisSort = False Then .ButtonMis(0).Visible = False Else .ButtonMis(0).Visible = True
            If PlayOver.MisInfo = False Then .ButtonMis(1).Visible = False Else .ButtonMis(1).Visible = True
            If PlayOver.MisOptions = False Then .ButtonMis(2).Visible = False Else .ButtonMis(2).Visible = True
        End If
        .MisMen.Visible = True
        .MisBar.Visible = True
    End If
    
    
    If LstMenu = False Then
        .ListMen.Visible = False
        .ListBar.Visible = False
    Else
        If PlayOver.LstNew = False Then .ButtonLst(0).Visible = False Else .ButtonLst(0).Visible = True
        If PlayOver.LstSave = False Then .ButtonLst(1).Visible = False Else .ButtonLst(1).Visible = True
        If PlayOver.LstLoad = False Then .ButtonLst(2).Visible = False Else .ButtonLst(2).Visible = True
        .ListMen.Visible = True
        .ListBar.Visible = True
    End If
    End With
    If ButtonNum2 < 5 Then ButtonNum2 = ButtonNum2 + 1
    
End Sub


Public Sub UpdateSliders()

    With VBAmpMain
    'get the volume in terms of hundred
    Volume = Int(VolNewX / 0.54)
    'set the volume (from -2500 to 0)
    .WinMedia.Volume = (Volume * 25) - 2500
    'get the volume so it is easily divisible by 27(there are 28 pictures for volume)
    VolVar = Volume * 1.35
    'get the Y of the picture of the correct volume
    VolVar = 15 * (Int(VolVar / 5))
    'set the volume picture
    .SliderVolume.PaintPicture Extras.VolumePic.Picture, 0, 0, , , 0, VolVar
    'set the button for the volume
    If ChoosingVolume = False Then .SliderVolume.PaintPicture Extras.VolumePic, VolNewX, 1, , , 15, 422, 14, 11
    If ChoosingVolume = True Then .SliderVolume.PaintPicture Extras.VolumePic, VolNewX, 1, , , 0, 422, 14, 11
    'get a volume string to show in the scrolltext picture
    VolumeString = "Volume: " & Volume & "%"
    
    'balance is the same as volume basically, refer to above if you don't understand
    Balance = Int(BalanceNewX / 0.12)
    Balance = Balance - 100
    .WinMedia.Balance = Balance * 25
    BalanceVar = Balance * 1.35
    BalanceVar = 15 * (Int((Abs(BalanceVar)) / 5))
    .SliderBalance.PaintPicture Extras.BalancePic, 0, 0, , , 9, BalanceVar
    If ChoosingBalance = False Then .SliderBalance.PaintPicture Extras.BalancePic, BalanceNewX, 1, , , 15, 422, 14, 11
    If ChoosingBalance = True Then .SliderBalance.PaintPicture Extras.BalancePic, BalanceNewX, 1, , , 0, 422, 14, 11
    If Balance = 0 Then BalanceString = "Balance: Center"
    If Balance > 0 Then BalanceString = "Balance: " & Balance & "% Right"
    If Balance < 0 Then BalanceString = "Balance: " & Abs(Balance) & "% Left"

    'same as volume and balance
    SongPosX = SongPosNewX
    Position = Int((SongPosNewX / 219) * .WinMedia.Duration)
    .SliderPos.PaintPicture Extras.PosBar, 0, 0, , , 0, 0
    If ChoosingSongPos = False Then .SliderPos.PaintPicture Extras.PosBar, SongPosNewX, 0, , , 248, 0, 29, 10
    If ChoosingSongPos = True Then .SliderPos.PaintPicture Extras.PosBar, SongPosNewX, 0, , , 278, 0, 29, 10
    SongPosString = "Seek to: " & GetTime(Position, True) & "/" & GetTime(.WinMedia.Duration, True) & " (" & Round(SongPosNewX / 2.19) & "%)"
    
    If WinShade = True Then
    WSSongPosX = WSSongPosNewX
    WSPosition = Int((WSSongPosNewX / 14) * .WinMedia.Duration)
    .SliderWSPos.PaintPicture Extras.Bars, 0, 0, , , 0, 36
    If .WinMedia.PlayState = mpPlaying Or .WinMedia.PlayState = mpPaused Then
        If WSSongPosNewX <= 4 Then .SliderWSPos.PaintPicture Extras.Bars, WSSongPosNewX, 0, , , 17, 36, 3, 7
        If WSSongPosNewX > 4 And WSSongPosNewX <= 9 Then .SliderWSPos.PaintPicture Extras.Bars, WSSongPosNewX, 0, , , 20, 36, 3, 7
        If WSSongPosNewX > 9 Then .SliderWSPos.PaintPicture Extras.Bars, WSSongPosNewX, 0, , , 23, 36, 3, 7
    End If
    If Elapsed = True Then WSSongPosString = " " & GetTime(WSPosition, True)
    If Elapsed = False Then WSSongPosString = "-" & GetTime(.WinMedia.Duration - WSPosition, True)
    End If

    'select which string to show in the scrolltext picture
    If ChoosingVolume = True Then StringToPicture VolumeString, .ScrollText, 4, 5, 30
    If ChoosingBalance = True Then StringToPicture BalanceString, .ScrollText, 4, 5, 30
    If ChoosingSongPos = True Then StringToPicture SongPosString, .ScrollText, 4, 5, 30
    If ChoosingScrollText = False And ChoosingVolume = False And ChoosingBalance = False And ChoosingSongPos = False Then StringToPicture SongString2, .ScrollText, 4, 5, 30
    End With

End Sub

Public Sub UpdatePlaylist()
On Error Resume Next
    Dim ItemString As String, SongString As String, CountNumber As Integer
    Dim AmpersandAdded As Boolean, ChangeCaption As Boolean
    With Playlist
        'Change the selected item number
        If SelectedItem <> .Playlist.SelectedItem.Index Then
            SelectedItem = .Playlist.SelectedItem.Index
        End If
        If SongNumber >= .Playlist.ListItems.Count Then SongNumber = 0
        If SelectedIndex >= .Playlist.ListItems.Count Then SelectedIndex = -1
        For I = 1 To .Item.UBound + 1
            ChangeCaption = False
            If I <= .Playlist.ListItems.Count - SongNumber Then
                CountNumber = Len(Str(.Playlist.ListItems.Count)) + 1
                SongString = .Playlist.ListItems.Item(I + SongNumber).Text
                'set the caption that has autosize so the string can be cut down
                .TestLabel2.Caption = GetTime(.SecondsList.ListItems.Item(I + SongNumber).Text)
                .TestLabel.Caption = SongString
                AmpersandAdded = False
                For B = 1 To Len(.TestLabel.Caption)
                    If Mid(.TestLabel.Caption, B, 1) = "&" Then
                        If B = 1 Then
                            .TestLabel.Caption = "&" & .TestLabel.Caption
                        ElseIf B = Len(.TestLabel.Caption) Then
                            .TestLabel.Caption = .TestLabel.Caption & "&"
                        Else
                            .TestLabel.Caption = Left(.TestLabel.Caption, B) & "&" & Right(.TestLabel.Caption, Len(.TestLabel.Caption) - B)
                        End If
                        B = B + 1
                        AmpersandAdded = True
                    End If
                Next B
                'put a number and period in front of song
                ItemString = SongNumber + I & ". "
                ItemString = Space(CountNumber - Len(ItemString)) & ItemString
                .TestLabel3.Caption = ItemString
                If .TestLabel.Width > 212 - .TestLabel3.Width Then
                    'cut down the string until it fits in the labels
                    Do Until .TestLabel.Width <= 203 - .TestLabel3.Width
                        .TestLabel.Caption = Left(.TestLabel.Caption, Len(.TestLabel.Caption) - 1)
                    Loop
                    'take out the extra ampersand
                    If AmpersandAdded = True Then
                    For B = 1 To Len(.TestLabel.Caption)
                        If Mid(.TestLabel.Caption, B, 2) = "&&" Then
                            If B = 1 Then
                                .TestLabel.Caption = Right(.TestLabel.Caption, Len(.TestLabel.Caption) - 1)
                            ElseIf B = Len(.TestLabel.Caption) - 1 Then
                                .TestLabel.Caption = Left(.TestLabel.Caption, Len(.TestLabel.Caption) - 1)
                            Else
                                .TestLabel.Caption = Left(.TestLabel.Caption, B) & Right(.TestLabel.Caption, Len(.TestLabel.Caption) - B - 1)
                            End If
                        End If
                    Next B
                    End If
                    SongString = .TestLabel.Caption & "..."
                End If
                'set the selected items to have either highlighted text or white text
                'the changecaption boolean is used later and it is used so that only the text that is changed is changed so the playlist doesn't flicker
                If I + SongNumber - 1 = SelectedIndex Then
                    If .Item(I - 1).BackColor = .ListView.BackColor Then
                    ChangeCaption = True
                    .Item(I - 1).BackColor = SelectedBackColor
                    .Time(I - 1).BackColor = SelectedBackColor
                    End If
                Else
                    If .Item(I - 1).BackColor = SelectedBackColor Then
                    ChangeCaption = True
                    .Item(I - 1).BackColor = .ListView.BackColor
                    .Time(I - 1).BackColor = .ListView.BackColor
                    End If
                End If
                If I + SongNumber = SelectedItem Then
                    ChangeCaption = True
                    .Item(I - 1).ForeColor = CurrentForeColor
                    .Time(I - 1).ForeColor = CurrentForeColor
                Else
                    ChangeCaption = True
                    .Item(I - 1).ForeColor = UnSelectedForeColor
                    .Time(I - 1).ForeColor = UnSelectedForeColor
                End If
                If .TestLabel2.Width < 27 Then .TestLabel2.Caption = "  " & .TestLabel2.Caption
                                
                If OldSongNumber <> SongNumber Then ChangeCaption = True
                'set the item caption to have the number of the song, a period, and the song
                If ChangeCaption = True Then
                .Item(I - 1).Cls
                .Time(I - 1).Cls
                .Item(I - 1).Print ItemString & SongString
                .Time(I - 1).Print .TestLabel2.Caption
                End If
                
            End If
            If I > .Playlist.ListItems.Count Then
                .Item(I - 1).Cls
                .Time(I - 1).Cls
            End If
        Next
        OldSongNumber = SongNumber
    End With

End Sub

Public Sub UpdatePlaylistSlider(NewSongNumber As Boolean)
On Error Resume Next
    With Playlist
    If NewSongNumber = True Then
    If .Playlist.ListItems.Count > VisibleItemCount Then
    If SongNumber <> Round((SongNewY / (.PlaylistSlider.Height - 18)) * (.Playlist.ListItems.Count - VisibleItemCount)) And Round((SongNewY / (.PlaylistSlider.Height - 18)) * (.Playlist.ListItems.Count - 4)) <= .Playlist.ListItems.Count - 4 Then
        SongNumber = Round((SongNewY / (.PlaylistSlider.Height - 18)) * (.Playlist.ListItems.Count - VisibleItemCount))
        UpdatePlaylist
    End If
    ElseIf .Playlist.ListItems.Count <= VisibleItemCount Then
    If SongNumber <> 0 Then
        SongNumber = 0
        UpdatePlaylist
    End If
    End If
    End If
    TemporaryNumber = .PlaylistSlider.Height / 29
    For A = 0 To TemporaryNumber
    .PlaylistSlider.PaintPicture .PLEdit, 0, A * 29, , , 26, 42, 25, 29
    Next A
    
    If .Playlist.ListItems.Count <> VisibleItemCount Then
    If ChoosingSongNumber = False Then .PlaylistSlider.PaintPicture .PLEdit, 10, Round((SongNumber / (.Playlist.ListItems.Count - VisibleItemCount)) * (.PlaylistSlider.Height - 18)), , , 52, 53, 8, 18
    If ChoosingSongNumber = True Then .PlaylistSlider.PaintPicture .PLEdit, 10, Round((SongNumber / (.Playlist.ListItems.Count - VisibleItemCount)) * (.PlaylistSlider.Height - 18)), , , 61, 53, 8, 18
    Else
    If ChoosingSongNumber = False Then .PlaylistSlider.PaintPicture .PLEdit, 10, 0, , , 52, 53, 8, 18
    If ChoosingSongNumber = True Then .PlaylistSlider.PaintPicture .PLEdit, 10, 0, , , 61, 53, 8, 18
    End If
    End With
    
End Sub

Public Sub DrawButtons()

    With VBAmpMain
    'paint buttons
    .ButtonPrev.PaintPicture Extras.Buttons, 0, 0, , , 0, 0
    .ButtonPrevD.PaintPicture Extras.Buttons, 0, 0, , , 0, 18
    .ButtonPlay.PaintPicture Extras.Buttons, 0, 0, , , 23, 0
    .ButtonPlayD.PaintPicture Extras.Buttons, 0, 0, , , 23, 18
    .ButtonPause.PaintPicture Extras.Buttons, 0, 0, , , 46, 0
    .ButtonPauseD.PaintPicture Extras.Buttons, 0, 0, , , 46, 18
    .ButtonStop.PaintPicture Extras.Buttons, 0, 0, , , 69, 0
    .ButtonStopD.PaintPicture Extras.Buttons, 0, 0, , , 69, 18
    .ButtonNext.PaintPicture Extras.Buttons, 0, 0, , , 92, 0
    .ButtonNextD.PaintPicture Extras.Buttons, 0, 0, , , 92, 18
    .ButtonOpen.PaintPicture Extras.Buttons, 0, 0, , , 114, 0
    .ButtonOpenD.PaintPicture Extras.Buttons, 0, 0, , , 114, 16
    .ButtonShuffle.PaintPicture Extras.Shuffle, 0, 0, , , 28, 0
    .ButtonShuffleTrue.PaintPicture Extras.Shuffle, 0, 0, , , 28, 30
    .ButtonShuffleD.PaintPicture Extras.Shuffle, 0, 0, , , 28, 15
    .ButtonShuffleTrueD.PaintPicture Extras.Shuffle, 0, 0, , , 28, 45
    .ButtonRepeat.PaintPicture Extras.Shuffle, 0, 0, , , 0, 0
    .ButtonRepeatTrue.PaintPicture Extras.Shuffle, 0, 0, , , 0, 30
    .ButtonRepeatD.PaintPicture Extras.Shuffle, 0, 0, , , 0, 15
    .ButtonRepeatTrueD.PaintPicture Extras.Shuffle, 0, 0, , , 0, 45
    .PL.PaintPicture Extras.Shuffle, 0, 0, , , 23, 61
    .PLTrue.PaintPicture Extras.Shuffle, 0, 0, , , 23, 73
    .PLD.PaintPicture Extras.Shuffle, 0, 0, , , 69, 61
    .PLTrueD.PaintPicture Extras.Shuffle, 0, 0, , , 69, 73
    .EQ.PaintPicture Extras.Shuffle, 0, 0, , , 0, 61
    .ScrollText.PaintPicture .Main, 0, 0, , , 107, 22
    .KBPS.PaintPicture .Main, 0, 0, , , 107, 39
    .KHZ.PaintPicture .Main, 0, 0, , , 152, 39
    .WinShadeTime.PaintPicture Extras.Bars, 0, 0, , , 152, 33
    .ButtonExit.PaintPicture Extras.Bars, 0, 0, , , 291, 3
    .ButtonExit2.PaintPicture Extras.Bars, 0, 0, , , 291, 18
    .ButtonExitD.PaintPicture Extras.Bars, 0, 0, , , 18, 9
    .ButtonMenu.PaintPicture Extras.Bars, 0, 0, , , 33, 3
    .ButtonMenu2.PaintPicture Extras.Bars, 0, 0, , , 33, 18
    .ButtonMenuD.PaintPicture Extras.Bars, 0, 0, , , 0, 9
    .ButtonMin.PaintPicture Extras.Bars, 0, 0, , , 271, 3
    .ButtonMin2.PaintPicture Extras.Bars, 0, 0, , , 271, 18
    .ButtonMinD.PaintPicture Extras.Bars, 0, 0, , , 9, 9
    .ButtonWinShade.PaintPicture Extras.Bars, 0, 0, , , 281, 3
    .ButtonWinShadeTrue.PaintPicture Extras.Bars, 0, 0, , , 281, 32
    .ButtonWinShadeD.PaintPicture Extras.Bars, 0, 0, , , 9, 18
    .ButtonWinShadeTrueD.PaintPicture Extras.Bars, 0, 0, , , 9, 27
    .ButtonWinShade2.PaintPicture Extras.Bars, 0, 0, , , 281, 18
    .ButtonWinShadeTrue2.PaintPicture Extras.Bars, 0, 0, , , 281, 45
    .WinShadeNext.PaintPicture Extras.Bars, 0, 0, , , .WinShadeNext.Left + 27, .WinShadeNext.Top + 29
    .WinShadePrev.PaintPicture Extras.Bars, 0, 0, , , .WinShadePrev.Left + 27, .WinShadePrev.Top + 29
    .WinShadeOpen.PaintPicture Extras.Bars, 0, 0, , , .WinShadeOpen.Left + 27, .WinShadeOpen.Top + 29
    .WinShadePlay.PaintPicture Extras.Bars, 0, 0, , , .WinShadePlay.Left + 27, .WinShadePlay.Top + 29
    .WinShadeStop.PaintPicture Extras.Bars, 0, 0, , , .WinShadeStop.Left + 27, .WinShadeStop.Top + 29
    .WinShadePause.PaintPicture Extras.Bars, 0, 0, , , .WinShadePause.Left + 27, .WinShadePause.Top + 29
    .WinShadeNext2.PaintPicture Extras.Bars, 0, 0, , , .WinShadeNext.Left + 27, .WinShadeNext.Top + 42
    .WinShadePrev2.PaintPicture Extras.Bars, 0, 0, , , .WinShadePrev.Left + 27, .WinShadePrev.Top + 42
    .WinShadeOpen2.PaintPicture Extras.Bars, 0, 0, , , .WinShadeOpen.Left + 27, .WinShadeOpen.Top + 42
    .WinShadePlay2.PaintPicture Extras.Bars, 0, 0, , , .WinShadePlay.Left + 27, .WinShadePlay.Top + 42
    .WinShadeStop2.PaintPicture Extras.Bars, 0, 0, , , .WinShadeStop.Left + 27, .WinShadeStop.Top + 42
    .WinShadePause2.PaintPicture Extras.Bars, 0, 0, , , .WinShadePause.Left + 27, .WinShadePause.Top + 42
    .Bar(0).PaintPicture Extras.Bars, 0, 0, , , 27, 0
    .Bar(1).PaintPicture Extras.Bars, 0, 0, , , 27, 29
    .Bar(2).PaintPicture Extras.Bars, 0, 0, , , 27, 15
    .Bar(3).PaintPicture Extras.Bars, 0, 0, , , 27, 42
    NumbertoPicture 0, .BlankMin, True
    NumbertoPicture 0, .BlankSec, True
    End With
    With Playlist
    .Bar(0).PaintPicture .PLEdit, 0, 0, , , 0, 0, 25, 21
    .Bar(0).PaintPicture .PLEdit, .Bar(0).Width - 25, 0, , , 153, 0, 25, 21
    .Bar(0).PaintPicture .PLEdit, 25, 0, .Bar(0).Width - 50, , 127, 0, 25, 21
    .Bar(0).PaintPicture .PLEdit, Int((.Bar(0).Width - 100) / 2), 0, , , 26, 0, 100, 21
    .Bar(1).PaintPicture .PLEdit, 0, 0, , , 0, 21, 25, 21
    .Bar(1).PaintPicture .PLEdit, .Bar(1).Width - 25, 0, , , 153, 21, 25, 21
    .Bar(1).PaintPicture .PLEdit, 25, 0, .Bar(1).Width - 50, , 127, 21, 25, 21
    .Bar(1).PaintPicture .PLEdit, Int((.Bar(1).Width - 100) / 2), 0, , , 26, 21, 100, 21
    .LeftBar.PaintPicture .PLEdit, 0, 20, 25, .LeftBar.Height, 0, 42, 25, 29
    .BotmBar.PaintPicture .PLEdit, 0, 0, , , 0, 72, 125, 38
    .BotmBar.PaintPicture .PLEdit, .BotmBar.Width - 150, 0, , , 126, 72, 150, 38
    .AddBar.PaintPicture .PLEdit, 0, 0, , , 48, 111
    .RemBar.PaintPicture .PLEdit, 0, 0, , , 100, 111
    .SelBar.PaintPicture .PLEdit, 0, 0, , , 150, 111
    .MisBar.PaintPicture .PLEdit, 0, 0, , , 200, 111
    .ListBar.PaintPicture .PLEdit, 0, 0, , , 250, 111
    .Add.PaintPicture .PLEdit, 0, 0, , , 14, 80
    .Remove.PaintPicture .PLEdit, 0, 0, , , 43, 80
    .Select.PaintPicture .PLEdit, 0, 0, , , 72, 80
    .Misc.PaintPicture .PLEdit, 0, 0, , , 101, 80
    .List.PaintPicture .PLEdit, 0, 0, , , 232, 80
    .AddMen.PaintPicture .PLEdit, 0, 0, , , 0, 111, 22, 18
    .AddMen.PaintPicture .PLEdit, 0, 18, , , 0, 130, 22, 18
    .AddMen.PaintPicture .PLEdit, 0, 36, , , 0, 149, 22, 18
    .RemMen.PaintPicture .PLEdit, 0, 0, , , 54, 168, 22, 18
    .RemMen.PaintPicture .PLEdit, 0, 18, , , 54, 111, 22, 18
    .RemMen.PaintPicture .PLEdit, 0, 36, , , 54, 130, 22, 18
    .RemMen.PaintPicture .PLEdit, 0, 54, , , 54, 149, 22, 18
    .SelMen.PaintPicture .PLEdit, 0, 0, , , 104, 111, 22, 18
    .SelMen.PaintPicture .PLEdit, 0, 18, , , 104, 130, 22, 18
    .SelMen.PaintPicture .PLEdit, 0, 36, , , 104, 149, 22, 18
    .MisMen.PaintPicture .PLEdit, 0, 0, , , 154, 111, 22, 18
    .MisMen.PaintPicture .PLEdit, 0, 18, , , 154, 130, 22, 18
    .MisMen.PaintPicture .PLEdit, 0, 36, , , 154, 149, 22, 18
    .ListMen.PaintPicture .PLEdit, 0, 0, , , 204, 111, 22, 18
    .ListMen.PaintPicture .PLEdit, 0, 18, , , 204, 130, 22, 18
    .ListMen.PaintPicture .PLEdit, 0, 36, , , 204, 149, 22, 18
    .ButtonAdd(0).PaintPicture .PLEdit, 0, 0, , , 23, 111, 22, 18
    .ButtonAdd(1).PaintPicture .PLEdit, 0, 0, , , 23, 130, 22, 18
    .ButtonAdd(2).PaintPicture .PLEdit, 0, 0, , , 23, 149, 22, 18
    .ButtonRem(0).PaintPicture .PLEdit, 0, 0, , , 77, 168, 22, 18
    .ButtonRem(1).PaintPicture .PLEdit, 0, 0, , , 77, 111, 22, 18
    .ButtonRem(2).PaintPicture .PLEdit, 0, 0, , , 77, 130, 22, 18
    .ButtonRem(3).PaintPicture .PLEdit, 0, 0, , , 77, 149, 22, 18
    .ButtonSel(0).PaintPicture .PLEdit, 0, 0, , , 127, 111, 22, 18
    .ButtonSel(1).PaintPicture .PLEdit, 0, 0, , , 127, 130, 22, 18
    .ButtonSel(2).PaintPicture .PLEdit, 0, 0, , , 127, 149, 22, 18
    .ButtonMis(0).PaintPicture .PLEdit, 0, 0, , , 177, 111, 22, 18
    .ButtonMis(1).PaintPicture .PLEdit, 0, 0, , , 177, 130, 22, 18
    .ButtonMis(2).PaintPicture .PLEdit, 0, 0, , , 177, 149, 22, 18
    .ButtonLst(0).PaintPicture .PLEdit, 0, 0, , , 227, 111, 22, 18
    .ButtonLst(1).PaintPicture .PLEdit, 0, 0, , , 227, 130, 22, 18
    .ButtonLst(2).PaintPicture .PLEdit, 0, 0, , , 227, 149, 22, 18
    .Exit.PaintPicture .PLEdit, 0, 0, , , 167, 3, 9, 9
    .Exit2.PaintPicture .PLEdit, 0, 0, , , 167, 24, 9, 9
    .ExitD.PaintPicture .PLEdit, 0, 0, , , 52, 42
    .Shade.PaintPicture .PLEdit, 0, 0, , , 158, 3, 9, 9
    .Shade2.PaintPicture .PLEdit, 0, 0, , , 158, 24, 9, 9
    .ShadeD.PaintPicture .PLEdit, 0, 0, , , 62, 42
    End With

End Sub

Public Sub MouseOver()
Dim UpdatePics As Boolean
    With VBAmpMain
    'vars used to save time
    X = .Left / Screen.TwipsPerPixelX
    Y = .Top / Screen.TwipsPerPixelY
    GetCursorPos MousePos
    'get the x and y of the current position
    UpdatePics = False
    If PosX <> MousePos.X Or PosY <> MousePos.Y Then UpdatePics = True
    PosX = MousePos.X
    PosY = MousePos.Y
    'check if the mouse is over any of the buttons
    If PosX >= X + .ButtonExit.Left And PosY >= Y + .ButtonExit.Top And PosX <= X + .ButtonExit.Left + .ButtonExit.Width And PosY <= Y + .ButtonExit.Top + .ButtonExit.Height Then ButtonOver.Exit = True Else ButtonOver.Exit = False
    If PosX >= X + .ButtonMenu.Left And PosY >= Y + .ButtonMenu.Top And PosX <= X + .ButtonMenu.Left + .ButtonMenu.Width And PosY <= Y + .ButtonMenu.Top + .ButtonMenu.Height Then ButtonOver.Menu = True Else ButtonOver.Menu = False
    If PosX >= X + .ButtonMin.Left And PosY >= Y + .ButtonMin.Top And PosX <= X + .ButtonMin.Left + .ButtonMin.Width And PosY <= Y + .ButtonMin.Top + .ButtonMin.Height Then ButtonOver.Minimize = True Else ButtonOver.Minimize = False
    If PosX >= X + .ButtonWinShade.Left And PosY >= Y + .ButtonWinShade.Top And PosX <= X + .ButtonWinShade.Left + .ButtonWinShade.Width And PosY <= Y + .ButtonWinShade.Top + .ButtonWinShade.Height Then ButtonOver.WinShade = True Else ButtonOver.WinShade = False
    If PosX > X + .ButtonPrev.Left And PosY >= Y + .ButtonPrev.Top And PosX < X + .ButtonPrev.Left + .ButtonPrev.Width And PosY <= Y + .ButtonPrev.Top + .ButtonPrev.Height Then ButtonOver.Previous = True Else ButtonOver.Previous = False
    If PosX > X + .ButtonPause.Left And PosY >= Y + .ButtonPause.Top And PosX < X + .ButtonPause.Left + .ButtonPause.Width And PosY <= Y + .ButtonPause.Top + .ButtonPause.Height Then ButtonOver.Pause = True Else ButtonOver.Pause = False
    If PosX > X + .ButtonStop.Left And PosY >= Y + .ButtonStop.Top And PosX < X + .ButtonStop.Left + .ButtonStop.Width And PosY <= Y + .ButtonStop.Top + .ButtonStop.Height Then ButtonOver.Stop = True Else ButtonOver.Stop = False
    If PosX > X + .ButtonNext.Left And PosY >= Y + .ButtonNext.Top And PosX < X + .ButtonNext.Left + .ButtonNext.Width And PosY <= Y + .ButtonNext.Top + .ButtonNext.Height Then ButtonOver.Next = True Else ButtonOver.Next = False
    If PosX > X + .ButtonPlay.Left And PosY >= Y + .ButtonPlay.Top And PosX < X + .ButtonPlay.Left + .ButtonPlay.Width And PosY <= Y + .ButtonPlay.Top + .ButtonPlay.Height Then ButtonOver.Play = True Else ButtonOver.Play = False
    If PosX >= X + .ButtonOpen.Left And PosY >= Y + .ButtonOpen.Top And PosX <= X + .ButtonOpen.Left + .ButtonOpen.Width And PosY <= Y + .ButtonOpen.Top + .ButtonOpen.Height Then ButtonOver.Open = True Else ButtonOver.Open = False
    If PosX >= X + .ButtonShuffle.Left And PosY >= Y + .ButtonShuffle.Top And PosX <= X + .ButtonShuffle.Left + .ButtonShuffle.Width And PosY <= Y + .ButtonShuffle.Top + .ButtonShuffle.Height Then ButtonOver.Shuffle = True Else ButtonOver.Shuffle = False
    If PosX >= X + .ButtonRepeat.Left And PosY >= Y + .ButtonRepeat.Top And PosX <= X + .ButtonRepeat.Left + .ButtonRepeat.Width And PosY <= Y + .ButtonRepeat.Top + .ButtonRepeat.Height Then ButtonOver.Repeat = True Else ButtonOver.Repeat = False
    If PosX > X + .PL.Left And PosY >= Y + .PL.Top And PosX <= X + .PL.Left + .PL.Width And PosY <= Y + .PL.Top + .PL.Height Then ButtonOver.PL = True Else ButtonOver.PL = False
    'check if mouse is over any of the sliders, etc.
    If PosX > X + .ScrollText.Left And PosY >= Y + .ScrollText.Top And PosX < X + .ScrollText.Left + .ScrollText.Width And PosY <= Y + .ScrollText.Top + .ScrollText.Height Then ButtonOver.SongScroll = True Else ButtonOver.SongScroll = False
    If PosX > X + .Volume.Left And PosY >= Y + .Volume.Top And PosX < X + .Volume.Left + .Volume.Width And PosY <= Y + .Volume.Top + .Volume.Height Then ButtonOver.Volume = True Else ButtonOver.Volume = False
    If PosX > X + .Balance.Left And PosY >= Y + .Balance.Top And PosX < X + .Balance.Left + .Balance.Width And PosY <= Y + .Balance.Top + .Balance.Height Then ButtonOver.Balance = True Else ButtonOver.Balance = False
    If PosX > X + .SliderPos.Left And PosY >= Y + .SliderPos.Top And PosX < X + .SliderPos.Left + .SliderPos.Width And PosY <= Y + .SliderPos.Top + .SliderPos.Height Then ButtonOver.SongPos = True Else ButtonOver.SongPos = False
    If PosX >= X And PosX <= X + .Width And PosY >= Y And PosY <= Y + .Height Then FormOver = True Else FormOver = False
    End With

    If UpdatePics = True And ButtonDown = True Then UpdatePictures

End Sub

Public Sub MouseOverPlaylist()
Dim X2, Y2
Dim PosX2, PosY2
    With Playlist
    GetCursorPos MousePos
    PosX2 = MousePos.X
    PosY2 = MousePos.Y
    X2 = .Left / Screen.TwipsPerPixelX
    Y2 = .Top / Screen.TwipsPerPixelY
    If PosX2 > X2 + .Exit.Left And PosX2 < X2 + .Exit.Left + .Exit.Width And PosY2 > Y2 + .Exit.Top And PosY2 < Y2 + .Exit.Top + .Exit.Height Then PlayOver.Exit = True Else PlayOver.Exit = False
    If PosX2 > X2 + .Shade.Left And PosX2 < X2 + .Shade.Left + .Shade.Width And PosY2 > Y2 + .Shade.Top And PosY2 < Y2 + .Shade.Top + .Shade.Height Then PlayOver.Shade = True Else PlayOver.Shade = False
    If PosX2 > X2 + .AddMen.Left And PosX2 < X2 + .AddMen.Left + .AddMen.Width And PosY2 > Y2 + .AddMen.Top And PosY2 < Y2 + .AddMen.Top + 18 Then PlayOver.AddURL = True Else PlayOver.AddURL = False
    If PosX2 > X2 + .AddMen.Left And PosX2 < X2 + .AddMen.Left + .AddMen.Width And PosY2 > Y2 + .AddMen.Top + 18 And PosY2 < Y2 + .AddMen.Top + 36 Then PlayOver.AddDir = True Else PlayOver.AddDir = False
    If PosX2 > X2 + .AddMen.Left And PosX2 < X2 + .AddMen.Left + .AddMen.Width And PosY2 > Y2 + .AddMen.Top + 36 And PosY2 < Y2 + .AddMen.Top + 54 Then PlayOver.AddFile = True Else PlayOver.AddFile = False
    If PosX2 > X2 + .RemMen.Left And PosX2 < X2 + .RemMen.Left + .RemMen.Width And PosY2 > Y2 + .RemMen.Top And PosY2 < Y2 + .RemMen.Top + 18 Then PlayOver.RemMisc = True Else PlayOver.RemMisc = False
    If PosX2 > X2 + .RemMen.Left And PosX2 < X2 + .RemMen.Left + .RemMen.Width And PosY2 > Y2 + .RemMen.Top + 18 And PosY2 < Y2 + .RemMen.Top + 36 Then PlayOver.RemAll = True Else PlayOver.RemAll = False
    If PosX2 > X2 + .RemMen.Left And PosX2 < X2 + .RemMen.Left + .RemMen.Width And PosY2 > Y2 + .RemMen.Top + 36 And PosY2 < Y2 + .RemMen.Top + 54 Then PlayOver.RemCrop = True Else PlayOver.RemCrop = False
    If PosX2 > X2 + .RemMen.Left And PosX2 < X2 + .RemMen.Left + .RemMen.Width And PosY2 > Y2 + .RemMen.Top + 54 And PosY2 < Y2 + .RemMen.Top + 72 Then PlayOver.RemFile = True Else PlayOver.RemFile = False
    If PosX2 > X2 + .SelMen.Left And PosX2 < X2 + .SelMen.Left + .SelMen.Width And PosY2 > Y2 + .SelMen.Top And PosY2 < Y2 + .SelMen.Top + 18 Then PlayOver.SelInverse = True Else PlayOver.SelInverse = False
    If PosX2 > X2 + .SelMen.Left And PosX2 < X2 + .SelMen.Left + .SelMen.Width And PosY2 > Y2 + .SelMen.Top + 18 And PosY2 < Y2 + .SelMen.Top + 36 Then PlayOver.SelZero = True Else PlayOver.SelZero = False
    If PosX2 > X2 + .SelMen.Left And PosX2 < X2 + .SelMen.Left + .SelMen.Width And PosY2 > Y2 + .SelMen.Top + 36 And PosY2 < Y2 + .SelMen.Top + 54 Then PlayOver.SelAll = True Else PlayOver.SelAll = False
    If PosX2 > X2 + .MisMen.Left And PosX2 < X2 + .MisMen.Left + .MisMen.Width And PosY2 > Y2 + .MisMen.Top And PosY2 < Y2 + .MisMen.Top + 18 Then PlayOver.MisSort = True Else PlayOver.MisSort = False
    If PosX2 > X2 + .MisMen.Left And PosX2 < X2 + .MisMen.Left + .MisMen.Width And PosY2 > Y2 + .MisMen.Top + 18 And PosY2 < Y2 + .MisMen.Top + 36 Then PlayOver.MisInfo = True Else PlayOver.MisInfo = False
    If PosX2 > X2 + .MisMen.Left And PosX2 < X2 + .MisMen.Left + .MisMen.Width And PosY2 > Y2 + .MisMen.Top + 36 And PosY2 < Y2 + .MisMen.Top + 54 Then PlayOver.MisOptions = True Else PlayOver.MisOptions = False
    If PosX2 > X2 + .ListMen.Left And PosX2 < X2 + .ListMen.Left + .ListMen.Width And PosY2 > Y2 + .ListMen.Top And PosY2 < Y2 + .ListMen.Top + 18 Then PlayOver.LstNew = True Else PlayOver.LstNew = False
    If PosX2 > X2 + .ListMen.Left And PosX2 < X2 + .ListMen.Left + .ListMen.Width And PosY2 > Y2 + .ListMen.Top + 18 And PosY2 < Y2 + .ListMen.Top + 36 Then PlayOver.LstSave = True Else PlayOver.LstSave = False
    If PosX2 > X2 + .ListMen.Left And PosX2 < X2 + .ListMen.Left + .ListMen.Width And PosY2 > Y2 + .ListMen.Top + 36 And PosY2 < Y2 + .ListMen.Top + 54 Then PlayOver.LstLoad = True Else PlayOver.LstLoad = False
    End With

End Sub

Public Sub StringToPicture(StringText As String, Picture As PictureBox, XStart As Integer, YStart As Integer, MaxLen As Integer)
Dim TextX, TextY
    'set textstrings to match with the letter that is being put into the picture
        'all of the symbols are not coded here, just the most used ones
    TextString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    TextString2 = "0123456789"
    TextString3 = ".:()-'!_+\/[]^&%,="
    
    'clear all text(the maxlen is the maximum number of letters that you can fit in the picture box
    For I = 1 To MaxLen
    Picture.PaintPicture Extras.Text, ((I - 1) * 5) + XStart, YStart, , , 143, 0, 5, 6
    Next I

    'go through all of the letters in the string to check for matches
    For I = 1 To Len(StringText)
        TextX = -1
        TextY = -1
        For A = 1 To Len(TextString)
        If UCase(Mid(StringText, I, 1)) = Mid(TextString, A, 1) Then
            'set the x and y of the position of the letter
            TextX = (A - 1) * 5
            TextY = 0
        End If
        Next A
        For B = 1 To Len(TextString2)
        If UCase(Mid(StringText, I, 1)) = Mid(TextString2, B, 1) Then
            'set the x and y of the position of the number
            TextX = (B - 1) * 5
            TextY = 6
        End If
        Next B
        For C = 1 To Len(TextString3)
        If Mid(StringText, I, 1) = Mid(TextString3, C, 1) Then
            'set the x and y of the position of the symbol
            TextX = (C * 5) + 50
            TextY = 6
        End If
        Next C
        'if there is a space or there was not a match then put in a space
        If Mid(StringText, I, 1) = " " Or TextX = -1 Then
            TextX = 143
            TextY = 0
        End If
        'symbol that is not coded for
        If Mid(StringText, I, 1) = "*" Then
            TextX = 20
            TextY = 12
        End If
        'put the letter, number or symbol into the picture
        If I <= MaxLen Then Picture.PaintPicture Extras.Text, ((I - 1) * 5) + XStart, YStart, , , TextX, TextY, 5, 6
    Next I
    
End Sub

Public Sub NumbertoPicture(Number As Integer, Picture As PictureBox, Optional ClearNumbers As Boolean)
Dim TempVar, TensDigit As Integer, OnesDigit As Integer

    TempVar = Str(Number)
    'get the ones and tens digits
    If Number >= 10 Then
    TensDigit = Int(Mid(TempVar, 2, 1))
    OnesDigit = Int(Mid(TempVar, 3, 1))
    End If
    If Number < 10 Then
    TensDigit = 0
    OnesDigit = Int(Mid(TempVar, 2, 1))
    End If
    'put the tens and ones digit into the picture
    Picture.PaintPicture VBAmpMain.Main, 0, 0, , , 48, 26, 21, 13
    Picture.PaintPicture Extras.Numbers, 0, 0, , , 9 * TensDigit, 0, 9, 13
    Picture.PaintPicture Extras.Numbers, 12, 0, , , 9 * OnesDigit, 0, 9, 13
    If ClearNumbers = True Then
    Picture.PaintPicture Extras.Numbers, 0, 0, , , 90, 0, 9, 13
    Picture.PaintPicture Extras.Numbers, 12, 0, , , 90, 0, 9, 13
    End If

End Sub

Public Function GetTime(SecondsInt As Integer, Optional TwoZeros As Boolean)
Dim Time
    
    'simple conversion of seconds to a nice looking string
    Minutes = GetMinutes(SecondsInt)
    Seconds = GetSeconds(SecondsInt)
    If Seconds < 10 Then Time = Minutes & ":0" & Seconds
    If Seconds >= 10 Then Time = Minutes & ":" & Seconds
    If TwoZeros = True Then
        If Minutes < 10 Then Time = "0" & Time
    End If
    GetTime = Time
    
End Function

Public Function GetMinutes(SecondsInt2 As Integer)
    
    'get the minutes out of a set number of seconds
    If SecondsInt2 > 0 Then GetMinutes = Int(SecondsInt2 / 60)
    If SecondsInt2 <= 0 Then GetMinutes = 0
    
End Function

Public Function GetSeconds(SecondsInt3 As Integer)

    'get the seconds after subtracting the minutes
    If SecondsInt3 > 0 Then GetSeconds = SecondsInt3 - ((GetMinutes(SecondsInt3)) * 60)
    If SecondsInt3 <= 0 Then GetSeconds = 0
    
End Function

Public Function FileExists(File As String) As Boolean
On Error GoTo FileDoesNotExist

    'open the file and if there is an error then the file does not exist
    Open File For Input As #1
    Close #1
    FileExists = True
    Exit Function
    
FileDoesNotExist:
    FileExists = False
End Function

Public Function GetFileName(File As String, Optional MinusExt As Boolean) As String

    TempNumber = 0
    TempNumber2 = 0
    For Number = 1 To Len(File)
        If Mid(File, Number, 1) = "\" Then TempNumber = Number
        If Mid(File, Number, 1) = "." Then TempNumber2 = Number
    Next Number
    GetFileName = Mid(File, TempNumber + 1, Len(File) - TempNumber)
    If MinusExt = True Then GetFileName = Mid(File, TempNumber + 1, TempNumber2 - 1)
    
End Function

Public Function Menus(Optional GetMenuState As Boolean) As Boolean

    If GetMenuState = True Then
        If AddMenu = False And RemMenu = False And SelMenu = False And MisMenu = False And LstMenu = False Then Menus = False Else Menus = True
    Else
        AddMenu = False
        RemMenu = False
        SelMenu = False
        MisMenu = False
        LstMenu = False
    End If

End Function

Public Function VisibleItemCount() As Integer
    Dim CountingNumber As Integer
    With Playlist
    CountingNumber = 0
    For I = .Item.LBound To .Item.UBound
        If .Item(I).Visible = True Then CountingNumber = CountingNumber + 1
    Next I
    End With
    VisibleItemCount = CountingNumber
End Function

Public Sub OpenM3U(FileName As String)
On Error Resume Next
    
    With Playlist
    .Playlist.ListItems.Clear
    .Playlist2.ListItems.Clear
    .Playlist3.ListItems.Clear
    .SecondsList.ListItems.Clear
    Dim FilePath, TmpString, FindComma, Data
    FilePath = Mid(FileName, 1, Len(FileName) - Len(GetFileName(FileName)))

    Open FileName For Input As #1
    Line Input #1, Data
    Do Until EOF(1)
    Line Input #1, Data
    .Playlist3.ListItems.Add , , Data
    Loop
    Close (1)
    
    For I = 1 To .Playlist3.ListItems.Count
    If Left(.Playlist3.ListItems.Item(I).Text, 7) <> "#EXTINF" Then
        If Left(.Playlist3.ListItems.Item(I).Text, 1) = "\" Then
        .Playlist2.ListItems.Add , , "C:" & .Playlist3.ListItems.Item(I).Text
        ElseIf Mid(.Playlist3.ListItems.Item(I).Text, 2, 2) = ":\" Then
        .Playlist2.ListItems.Add , , .Playlist3.ListItems.Item(I).Text
        Else
        .Playlist2.ListItems.Add , , FilePath & .Playlist3.ListItems.Item(I).Text
        End If
    
        If I - 1 > 0 And Left(.Playlist3.ListItems.Item(I - 1).Text, 7) = "#EXTINF" Then
            For FindComma = 1 To Len(.Playlist3.ListItems.Item(I - 1).Text)
            If Mid(.Playlist3.ListItems.Item(I - 1).Text, FindComma, 1) = "," Then
            If FindComma <> 9 Then .SecondsList.ListItems.Add , , Mid(.Playlist3.ListItems.Item(I - 1).Text, 9, FindComma - 9)
            .Playlist.ListItems.Add , , Mid(.Playlist3.ListItems.Item(I - 1).Text, FindComma + 1, Len(.Playlist3.ListItems.Item(I - 1).Text) - FindComma)
            Exit For
            End If
            Next FindComma
        Else
            .Playlist.ListItems.Add , , GetFileName(.Playlist3.ListItems.Item(I).Text, True)
        End If
    End If
    Next I
    End With
    UpdatePlaylist
    UpdatePlaylistSlider False
    VBAmpMain.OpenAFile Playlist.Playlist2.ListItems.Item(1).Text, "1. " & Playlist.Playlist.ListItems.Item(1).Text, True
    
End Sub

Public Sub SaveM3U(FileName As String)
On Error Resume Next

    With Playlist
    
    If .Playlist.ListItems.Count = 0 Then
        Open FileName For Output As #1
        Print #1, "#EXTM3U"
        Close (1)
        Exit Sub
    End If
    
    .Playlist3.ListItems.Clear
    .Playlist3.ListItems.Add , , "#EXTM3U"
    Dim FilePath
    FilePath = Mid(FileName, 1, Len(FileName) - Len(GetFileName(FileName)))
    For I = 1 To .Playlist.ListItems.Count
    'when it saves, it doesn't put the seconds, because that takes too long
    .Playlist3.ListItems.Add , , "#EXTINF:," & .Playlist.ListItems.Item(I).Text
    If Mid(.Playlist2.ListItems.Item(I).Text, 1, Len(.Playlist2.ListItems.Item(I).Text) - Len(GetFileName(.Playlist2.ListItems.Item(I).Text))) = FilePath Then
        .Playlist3.ListItems.Add , , GetFileName(.Playlist2.ListItems.Item(I).Text)
    Else
        .Playlist3.ListItems.Add , , Right(.Playlist2.ListItems.Item(I).Text, Len(.Playlist2.ListItems.Item(I).Text) - 2)
    End If
    Next
    
    Open FileName For Output As #1
    For I = 1 To .Playlist3.ListItems.Count
    Print #1, .Playlist3.ListItems.Item(I).Text
    Next
    Close (1)
    
    End With

End Sub
