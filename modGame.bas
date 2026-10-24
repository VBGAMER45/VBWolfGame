Attribute VB_Name = "modGame"
Option Explicit

' This file is part of VBWolf, a Visual Basic 6 port of the Wolf3D Browser
' Version GPL Source Code, Copyright (C) 2012 id Software LLC, a ZeniMax
' Media company.
'
' This program is free software: you can redistribute it and/or modify it
' under the terms of the GNU General Public License as published by the Free
' Software Foundation, either version 2 of the License, or (at your option)
' any later version. It is distributed WITHOUT ANY WARRANTY; see COPYING.txt.
'
' Ported from js/game.js: player controls, the 70 Hz game cycle, starting a
' level, dying, the intermission and level progression. The browser's timers
' and DOM are replaced by a single VB loop. The menus (js/menu.js used image
' art) are new, text-only screens: title, skill, controls and pause.

Public Const BUTTON_ATTACK As Long = 1
Public Const BUTTON_USE As Long = 2
Public Const BUTTON_ANY As Long = 128

Public Const BASEMOVE As Long = 35
Public Const RUNMOVE As Long = 70
Public Const MOVESCALE As Long = 150
Public Const BACKMOVESCALE As Long = 100
Public Const MAXMOUSETURN As Long = 10
Public Const TURNANGLESCALE As Long = 300

Public Const gd_baby As Long = 0
Public Const gd_easy As Long = 1
Public Const gd_medium As Long = 2
Public Const gd_hard As Long = 3

Public Const TICS_PER_SECOND As Long = 70

' screens
Public Const SCR_PLAYING As Long = 0
Public Const SCR_TITLE As Long = 1
Public Const SCR_SKILL As Long = 2
Public Const SCR_HELP As Long = 3
Public Const SCR_PAUSE As Long = 4
Public Const SCR_PACK As Long = 5

Public Running As Boolean
Public FormClosed As Boolean
Public Skill As Long
Public LevelNum As Long

' shown by the renderer
Public CurScreen As Long                   ' the menu that is up, or SCR_PLAYING
Public MenuIndex As Long
Public LevelActive As Boolean           ' a level is loaded (drawn behind the pause menu)
Public NotifyText As String
Public NotifyTics As Long
Public BonusFlashTics As Long
Public DamageFlashTics As Long
Public DeathFade As Double              ' 0..1 red fade after dying
Public GameOver As Boolean
Public InIntermission As Boolean
Public EpisodeDone As Boolean
Public IntermissionText(0 To 7) As String

Private LastTimeCount As Double         ' ms
Private QpcFreq As Currency
Private DeathTics As Long
Private PendingBonus As Long
Private KeyWas(0 To 255) As Boolean
Private HitEsc As Boolean
Private HitEnter As Boolean
Private HitUp As Boolean
Private HitDown As Boolean
Private HitSpace As Boolean
Private SumKill As Long, SumSecret As Long, SumTreasure As Long, LevelsDone As Long
Private EpisodeTime As Double
Private LevelsFromCommandLine As Boolean

Sub Main()
    Dim msg As String, tics As Long, path As String, packEpisode As String
    Math_BuildTables
    Art_Build
    Skill = gd_medium
    frmMain.Show
    Renderer_Init frmMain.hdc
    Sound_Init
    Pack_LoadSetting
    packEpisode = Pack_Apply()
    Renderer_RefreshBackground

    ' the command line may name an episode file (.txt) to use instead of the
    ' default, or a single map to play straight away, for testing levels
    path = Trim$(Command$)
    If Left$(path, 1) = """" Then path = Mid$(path, 2, Len(path) - 2)
    LevelsFromCommandLine = (path <> "")
    If path = "" Then path = DefaultEpisode(packEpisode)
    If LCase$(Right$(path, 4)) = ".txt" Then
        msg = Episodes_Load(path)
        CurScreen = SCR_TITLE
    Else
        Episodes_SetSingle path
        msg = StartGame()
    End If
    If msg <> "" Then
        MsgBox msg, vbCritical, "VBWolf"
        Sound_Shutdown
        Renderer_Shutdown
        Unload frmMain
        Exit Sub
    End If

    timeBeginPeriod 1
    QueryPerformanceFrequency QpcFreq
    LastTimeCount = NowMs()
    Running = True
    Do While Running
        tics = CalcTics()
        If tics > 0 Then
            GameCycle tics
            If Running Then Renderer_DrawFrame frmMain.hdc
        Else
            Sleep 1
        End If
        Sound_Update
        DoEvents
    Loop
    timeEndPeriod 1
    Sound_Shutdown
    Renderer_Shutdown
    If Not FormClosed Then Unload frmMain
End Sub

Private Function NowMs() As Double
    Dim c As Currency
    QueryPerformanceCounter c
    NowMs = c / QpcFreq * 1000#
End Function

' Whole 70 Hz tics since the last call; the remainder carries over.
Private Function CalcTics() As Long
    Dim delta As Double, tics As Long
    delta = (NowMs() - LastTimeCount) / 1000#
    tics = Int(TICS_PER_SECOND * delta)
    LastTimeCount = LastTimeCount + ToInt32(tics * 1000# / TICS_PER_SECOND)
    If tics > TICS_PER_SECOND Then tics = TICS_PER_SECOND   ' after a stall, don't jump more than a second
    CalcTics = tics
End Function

' True on the cycle a key goes down.
Private Function KeyHit(ByVal vk As Long, ByVal focused As Boolean) As Boolean
    Dim down As Boolean
    down = focused And IsKeyDown(vk)
    KeyHit = down And Not KeyWas(vk)
    KeyWas(vk) = down
End Function

Private Sub GameCycle(ByVal tics As Long)
    Dim focused As Boolean
    focused = (GetForegroundWindow() = frmMain.hWnd)
    HitEsc = KeyHit(VK_ESCAPE, focused)
    HitEnter = KeyHit(VK_RETURN, focused)
    HitUp = KeyHit(VK_UP, focused)
    HitDown = KeyHit(VK_DOWN, focused)
    HitSpace = KeyHit(VK_SPACE, focused)

    If NotifyTics > 0 Then NotifyTics = NotifyTics - tics
    If BonusFlashTics > 0 Then BonusFlashTics = BonusFlashTics - tics
    If DamageFlashTics > 0 Then DamageFlashTics = DamageFlashTics - tics

    If CurScreen <> SCR_PLAYING Then
        MenuInput
        Exit Sub
    End If
    If InIntermission Then
        If HitEnter Or HitSpace Then AdvanceLevel
        Exit Sub
    End If
    If GameOver Or EpisodeDone Then
        If HitEnter Then GoToTitle
        Exit Sub
    End If
    If HitEsc Then
        CurScreen = SCR_PAUSE
        MenuIndex = 0
        Exit Sub
    End If

    LevelTime = LevelTime + tics / TICS_PER_SECOND
    If Player.playstate <> ex_dead Then
        If focused Then UpdatePlayerControls tics Else ClearControls
        Player.angle = Math_NormalizeAngle(Player.angle)
        Player_Process tics
        Actors_Process tics
        PushWall_Process tics
        Doors_Process tics
    ElseIf Died(tics) Then
        ' turned to face the killer: fade to red, then respawn or end the game
        DeathTics = DeathTics + tics
        DeathFade = DeathTics / (TICS_PER_SECOND * 2)
        If DeathTics >= TICS_PER_SECOND * 2 Then
            DeathTics = 0
            DeathFade = 0
            If Player.lives > 0 Then Respawn Else GameOver = True
        End If
    End If
End Sub

'---------------------------------------------------------------- menus

' The entries of the current menu screen.
Public Function Game_MenuItems() As Variant
    Select Case CurScreen
        Case SCR_TITLE: Game_MenuItems = Array("NEW GAME", "CONTROLS", "CONTENT PACK", "QUIT")
        Case SCR_PACK: Game_MenuItems = Array("CHOOSE FOLDER...", "USE BUILT-IN CONTENT", "BACK")
        Case SCR_SKILL: Game_MenuItems = Array("RECRUIT", "SOLDIER", "VETERAN", "COMMANDO")
        Case SCR_PAUSE: Game_MenuItems = Array("RESUME", "QUIT TO TITLE")
        Case Else: Game_MenuItems = Array("BACK")
    End Select
End Function

Private Sub MenuInput()
    Dim count As Long
    count = UBound(Game_MenuItems()) + 1
    If HitUp Then MenuIndex = (MenuIndex + count - 1) Mod count
    If HitDown Then MenuIndex = (MenuIndex + 1) Mod count

    Select Case CurScreen
        Case SCR_TITLE
            If HitEsc Then Running = False
            If HitEnter Then
                Select Case MenuIndex
                    Case 0: CurScreen = SCR_SKILL: MenuIndex = gd_medium
                    Case 1: CurScreen = SCR_HELP: MenuIndex = 0
                    Case 2: CurScreen = SCR_PACK: MenuIndex = 0
                    Case 3: Running = False
                End Select
            End If
        Case SCR_SKILL
            If HitEsc Then CurScreen = SCR_TITLE: MenuIndex = 0
            If HitEnter Then
                Skill = MenuIndex
                Dim msg As String
                msg = StartGame()
                If msg <> "" Then
                    MsgBox msg, vbCritical, "VBWolf"
                    GoToTitle
                End If
            End If
        Case SCR_HELP
            If HitEsc Or HitEnter Then CurScreen = SCR_TITLE: MenuIndex = 1
        Case SCR_PACK
            If HitEsc Then CurScreen = SCR_TITLE: MenuIndex = 2
            If HitEnter Then
                Select Case MenuIndex
                    Case 0
                        Dim folder As String
                        folder = Pack_ChooseFolder(frmMain.hWnd)
                        If folder <> "" Then
                            PackFolder = folder
                            Pack_SaveSetting
                            ReloadContent
                        End If
                    Case 1
                        PackFolder = ""
                        Pack_SaveSetting
                        ReloadContent
                    Case 2
                        CurScreen = SCR_TITLE: MenuIndex = 2
                End Select
            End If
        Case SCR_PAUSE
            If HitEsc Then CurScreen = SCR_PLAYING
            If HitEnter Then
                If MenuIndex = 0 Then CurScreen = SCR_PLAYING Else GoToTitle
            End If
    End Select
End Sub

' The episode to play: a pack's own levels, or the built-in ones.
Private Function DefaultEpisode(ByVal packEpisode As String) As String
    If packEpisode <> "" Then DefaultEpisode = packEpisode Else DefaultEpisode = App.Path & "\maps\episode.txt"
End Function

' Rebuilds the built-in art and sounds and lays the current pack over them.
Private Sub ReloadContent()
    Dim packEpisode As String, msg As String
    Art_Build
    Sound_Rebuild
    packEpisode = Pack_Apply()
    Renderer_RefreshBackground
    If Not LevelsFromCommandLine Then
        msg = Episodes_Load(DefaultEpisode(packEpisode))
        If msg <> "" Then
            MsgBox msg & vbCrLf & "Using the built-in levels instead.", vbExclamation, "VBWolf"
            Episodes_Load App.Path & "\maps\episode.txt"
        End If
    End If
End Sub

Private Sub GoToTitle()
    LevelActive = False
    InIntermission = False
    GameOver = False
    EpisodeDone = False
    CurScreen = SCR_TITLE
    MenuIndex = 0
End Sub

' Starts the episode from its first level with a fresh player.
Private Function StartGame() As String
    SumKill = 0: SumSecret = 0: SumTreasure = 0: LevelsDone = 0: EpisodeTime = 0
    LevelNum = 0
    StartGame = Game_StartLevel(LevelFiles(0), False)
    If StartGame = "" Then CurScreen = SCR_PLAYING
End Function

'---------------------------------------------------------------- dying

' Turns the dead player toward the killer. Returns True once facing them.
Private Function Died(ByVal tics As Long) As Boolean
    Dim fangle As Double, dx As Double, dy As Double, iangle As Long, curangle As Long
    Dim clockwise As Long, counter As Long, change As Long
    If Player.lastAttacker < 0 Then
        Died = True
        Exit Function
    End If
    dx = Guards(Player.lastAttacker).x - Player.x
    dy = Player.y - Guards(Player.lastAttacker).y
    fangle = -Atan2(dy, dx)                         ' -pi to pi
    If fangle < 0 Then fangle = PI * 2 + fangle
    iangle = CLng(fangle / (PI * 2) * ANGLES)
    curangle = FINE2DEG(Player.angle)
    If curangle > iangle Then
        counter = curangle - iangle
        clockwise = ANGLES - curangle + iangle
    Else
        clockwise = iangle - curangle
        counter = curangle + ANGLES - iangle
    End If
    If clockwise < counter Then
        If curangle > iangle Then curangle = curangle - ANGLES
        If curangle = iangle Then
            Died = True
        Else
            change = tics * DEATHROTATE
            If curangle + change > iangle Then change = iangle - curangle
            curangle = curangle + change
            If curangle >= ANGLES Then curangle = curangle - ANGLES
            Player.angle = DEG2FINE(curangle)
        End If
    Else
        If curangle < iangle Then curangle = curangle + ANGLES
        If curangle = iangle Then
            Died = True
        Else
            change = -tics * DEATHROTATE
            If curangle + change < iangle Then change = iangle - curangle
            curangle = curangle + change
            If curangle < 0 Then curangle = curangle + ANGLES
            Player.angle = DEG2FINE(curangle)
        End If
    End If
End Function

' Restarts the level after losing a life, keeping the score from its start.
Private Sub Respawn()
    Dim lives As Long, score As Long, msg As String
    lives = Player.lives
    score = Player.startScore
    msg = Level_Reload()
    If msg <> "" Then
        MsgBox msg, vbCritical, "VBWolf"
        Running = False
        Exit Sub
    End If
    Level_ScanInfoPlane Skill
    Player_Spawn Skill, False
    Player.lives = lives - 1
    Player.score = score
    Player.startScore = score
End Sub

'---------------------------------------------------------------- controls

Private Sub ClearControls()
    Player.cmdButtons = 0
    Player.cmdForwardMove = 0
    Player.cmdSideMove = 0
End Sub

' Keys: arrows (or W/S) move and turn, Alt + arrows strafe, Shift runs,
' Ctrl fires, Space uses, 1-4 pick a weapon. A/D strafe as an addition.
Private Sub UpdatePlayerControls(ByVal tics As Long)
    Dim moveValue As Long, runKey As Boolean, strafing As Boolean
    Dim leftKey As Boolean, rightKey As Boolean, downKey As Boolean, upKey As Boolean
    Dim changeWeapon As Long

    ClearControls
    leftKey = IsKeyDown(VK_LEFT)
    rightKey = IsKeyDown(VK_RIGHT)
    downKey = IsKeyDown(VK_DOWN) Or IsKeyDown(VK_S)
    upKey = IsKeyDown(VK_UP) Or IsKeyDown(VK_W)
    runKey = IsKeyDown(VK_SHIFT)
    strafing = IsKeyDown(VK_MENU)

    If runKey Then moveValue = RUNMOVE Else moveValue = BASEMOVE

    If IsKeyDown(VK_CONTROL) Then Player.cmdButtons = Player.cmdButtons Or BUTTON_ATTACK

    If Not (upKey And downKey) Then
        If upKey Then Player.cmdForwardMove = Player.cmdForwardMove + moveValue * MOVESCALE
        If downKey Then Player.cmdForwardMove = Player.cmdForwardMove - moveValue * BACKMOVESCALE
    End If

    If leftKey Then
        If strafing Then
            Player.cmdSideMove = Player.cmdSideMove - moveValue * MOVESCALE
        Else
            Player.angle = Player.angle + TURNANGLESCALE * tics
        End If
    End If
    If rightKey Then
        If strafing Then
            Player.cmdSideMove = Player.cmdSideMove + moveValue * MOVESCALE
        Else
            Player.angle = Player.angle - TURNANGLESCALE * tics
        End If
    End If
    If IsKeyDown(VK_A) Then Player.cmdSideMove = Player.cmdSideMove - moveValue * MOVESCALE
    If IsKeyDown(VK_D) Then Player.cmdSideMove = Player.cmdSideMove + moveValue * MOVESCALE

    ' change weapon?
    changeWeapon = -1
    If IsKeyDown(VK_1) And (Player.items And ITEM_WEAPON_1) <> 0 Then
        changeWeapon = WEAPON_KNIFE
    ElseIf IsKeyDown(VK_2) And (Player.items And ITEM_WEAPON_2) <> 0 And Player.ammo(AMMO_BULLETS) > 0 Then
        changeWeapon = WEAPON_PISTOL
    ElseIf IsKeyDown(VK_3) And (Player.items And ITEM_WEAPON_3) <> 0 And Player.ammo(AMMO_BULLETS) > 0 Then
        changeWeapon = WEAPON_AUTO
    ElseIf IsKeyDown(VK_4) And (Player.items And ITEM_WEAPON_4) <> 0 And Player.ammo(AMMO_BULLETS) > 0 Then
        changeWeapon = WEAPON_CHAIN
    End If
    If changeWeapon > -1 Then
        Player.previousWeapon = WEAPON_KNIFE
        Player.weapon = changeWeapon
        Player.pendingWeapon = changeWeapon
    End If

    If IsKeyDown(VK_SPACE) Then Player.cmdButtons = Player.cmdButtons Or BUTTON_USE
End Sub

'---------------------------------------------------------------- levels

' Loads a level and places the player. Returns "" or an error message.
Public Function Game_StartLevel(ByVal path As String, ByVal keepStats As Boolean) As String
    Dim msg As String
    msg = Level_Load(path)
    If msg <> "" Then
        Game_StartLevel = msg
        Exit Function
    End If
    If SpawnX < 0 Then
        Game_StartLevel = "The level has no player start."
        Exit Function
    End If
    Level_ScanInfoPlane Skill           ' spawn items/guards
    Player_Spawn Skill, keepStats
    Player.startScore = Player.score
    LevelActive = True
    InIntermission = False
    GameOver = False
    EpisodeDone = False
    BonusFlashTics = 0
    DamageFlashTics = 0
    DeathFade = 0
    DeathTics = 0
    Game_Notify "Floor " & (LevelNum + 1) & ": " & LevelName
End Function

Public Sub Game_Notify(ByVal text As String)
    NotifyText = text
    NotifyTics = 3 * TICS_PER_SECOND
End Sub

Public Sub Game_StartBonusFlash()
    BonusFlashTics = 8
End Sub

Public Sub Game_StartDamageFlash(ByVal points As Long)
    DamageFlashTics = 21                ' about 300 ms, as the browser version's fade
End Sub

' Victory trigger tile: the browser version spins the camera; here the
' intermission is shown straight away.
Public Sub Game_Victory()
    Player.playstate = ex_victory
    Game_StartIntermission
End Sub

' End-of-floor tally, with the original's bonus rules: 500 points per second
' under par, 10000 for each ratio at 100%, a flat 15000 on the secret floor.
Public Sub Game_StartIntermission()
    Dim secs As Long, parSecs As Long, kill As Long, secret As Long, treasure As Long
    InIntermission = True
    secs = Int(LevelTime)
    parSecs = ParSeconds(LevelParTime)
    kill = Ratio(KilledMonsters, TotalMonsters)
    secret = Ratio(FoundSecrets, TotalSecrets)
    treasure = Ratio(FoundTreasure, TotalTreasure)

    If LevelNum = SecretLevel Then
        PendingBonus = 15000
    Else
        PendingBonus = 0
        If parSecs > 0 And parSecs > secs Then PendingBonus = PendingBonus + (parSecs - secs) * 500
        If kill = 100 Then PendingBonus = PendingBonus + 10000
        If secret = 100 Then PendingBonus = PendingBonus + 10000
        If treasure = 100 Then PendingBonus = PendingBonus + 10000
    End If
    SumKill = SumKill + kill: SumSecret = SumSecret + secret: SumTreasure = SumTreasure + treasure
    LevelsDone = LevelsDone + 1
    EpisodeTime = EpisodeTime + LevelTime

    If LevelNum = SecretLevel Then
        IntermissionText(0) = "SECRET FLOOR COMPLETED"
    Else
        IntermissionText(0) = "FLOOR " & (LevelNum + 1) & " COMPLETED"
    End If
    IntermissionText(1) = UCase$(LevelName)
    IntermissionText(2) = "BONUS  " & PendingBonus
    IntermissionText(3) = "TIME  " & Clock(secs)
    If parSecs > 0 Then IntermissionText(3) = IntermissionText(3) & "     PAR  " & Clock(parSecs)
    IntermissionText(4) = "KILL RATIO      " & kill & "%"
    IntermissionText(5) = "SECRET RATIO    " & secret & "%"
    IntermissionText(6) = "TREASURE RATIO  " & treasure & "%"
    IntermissionText(7) = "Press ENTER to continue"
End Sub

' Leaves the intermission for the next floor, or ends the episode.
Private Sub AdvanceLevel()
    Dim nextLevel As Long, msg As String
    Player_GivePoints PendingBonus
    PendingBonus = 0
    nextLevel = Episodes_NextLevel(LevelNum, Player.playstate = ex_secretlevel)
    If nextLevel >= NumLevels Then
        InIntermission = False
        EpisodeDone = True
        IntermissionText(0) = "EPISODE COMPLETE"
        IntermissionText(1) = UCase$(EpisodeName)
        IntermissionText(2) = "SCORE  " & Player.score
        IntermissionText(3) = "TOTAL TIME  " & Clock(Int(EpisodeTime))
        IntermissionText(4) = "AVERAGE KILLS      " & SumKill \ LevelsDone & "%"
        IntermissionText(5) = "AVERAGE SECRETS    " & SumSecret \ LevelsDone & "%"
        IntermissionText(6) = "AVERAGE TREASURE   " & SumTreasure \ LevelsDone & "%"
        IntermissionText(7) = "Press ENTER to return to the title"
        Exit Sub
    End If
    LevelNum = nextLevel
    msg = Game_StartLevel(LevelFiles(nextLevel), True)
    If msg <> "" Then
        MsgBox msg, vbCritical, "VBWolf"
        GoToTitle
    End If
End Sub

Private Function Ratio(ByVal found As Long, ByVal total As Long) As Long
    If total = 0 Then Ratio = 0 Else Ratio = (found * 100) \ total
End Function

' "MM:SS" to seconds; 0 if absent or malformed.
Private Function ParSeconds(ByVal par As String) As Long
    Dim parts() As String
    parts = Split(par, ":")
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function
    ParSeconds = CLng(parts(0)) * 60 + CLng(parts(1))
End Function

Private Function Clock(ByVal secs As Long) As String
    Clock = (secs \ 60) & ":" & Format$(secs Mod 60, "00")
End Function
