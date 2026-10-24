Attribute VB_Name = "modActors"
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
' Ported from js/actors.js: spawning and the state machine that drives each
' actor. Only the generic enemy types are included (no bosses, ghosts or the
' player-character actor). Removed actors are flagged instead of compacted,
' so actor indices stay valid.

Public Const SPDPATROL As Long = 512
Public Const SPDDOG As Long = 1500
Public Const FL_SHOOTABLE As Long = 1
Public Const FL_BONUS As Long = 2
Public Const FL_NEVERMARK As Long = 4
Public Const FL_VISABLE As Long = 8
Public Const FL_ATTACKMODE As Long = 16
Public Const FL_FIRSTATTACK As Long = 32
Public Const FL_AMBUSH As Long = 64
Public Const FL_NONMARK As Long = 128
Public Const MAX_GUARDS As Long = 255
Public Const NUMSTATES As Long = 34

Public Const en_guard As Long = 0
Public Const en_officer As Long = 1
Public Const en_ss As Long = 2
Public Const en_dog As Long = 3
Public Const en_fake As Long = 6
Public Const en_mutant As Long = 9

Public Const st_stand As Long = 0
Public Const st_path1 As Long = 1
Public Const st_path1s As Long = 2
Public Const st_path2 As Long = 3
Public Const st_path3 As Long = 4
Public Const st_path3s As Long = 5
Public Const st_path4 As Long = 6
Public Const st_pain As Long = 7
Public Const st_pain1 As Long = 8
Public Const st_shoot1 As Long = 9
Public Const st_shoot2 As Long = 10
Public Const st_shoot3 As Long = 11
Public Const st_shoot4 As Long = 12
Public Const st_shoot5 As Long = 13
Public Const st_shoot6 As Long = 14
Public Const st_shoot7 As Long = 15
Public Const st_shoot8 As Long = 16
Public Const st_shoot9 As Long = 17
Public Const st_chase1 As Long = 18
Public Const st_chase1s As Long = 19
Public Const st_chase2 As Long = 20
Public Const st_chase3 As Long = 21
Public Const st_chase3s As Long = 22
Public Const st_chase4 As Long = 23
Public Const st_die1 As Long = 24
Public Const st_die2 As Long = 25
Public Const st_die3 As Long = 26
Public Const st_die4 As Long = 27
Public Const st_die5 As Long = 28
Public Const st_die6 As Long = 29
Public Const st_die7 As Long = 30
Public Const st_die8 As Long = 31
Public Const st_die9 As Long = 32
Public Const st_dead As Long = 33
Public Const st_remove As Long = 34

Public Type TActor
    inUse As Boolean
    x As Long
    y As Long
    angle As Long
    actorType As Long
    health As Double            ' JS numbers: damage can be fractional
    speed As Long
    ticcount As Long
    temp2 As Double             ' reaction time
    distance As Long            ' left to move before reaching the next tile centre
    tileX As Long
    tileY As Long
    areanumber As Long
    waitfordoorx As Long        ' waiting on this door if non 0
    waitfordoory As Long
    flags As Long
    state As Long
    dir As Long
    sprite As Long
End Type

Public Guards(0 To MAX_GUARDS) As TActor
Public NumGuards As Long

Private Add8Dir(0 To 8) As Long
Private TablesReady As Boolean

Public Sub Actors_Reset()
    NumGuards = 0
    If Not TablesReady Then
        Add8Dir(0) = 4: Add8Dir(1) = 5: Add8Dir(2) = 6: Add8Dir(3) = 7
        Add8Dir(4) = 0: Add8Dir(5) = 1: Add8Dir(6) = 2: Add8Dir(7) = 3: Add8Dir(8) = 0
        ActStat_Init
        TablesReady = True
    End If
End Sub

Private Function GetNewActor() As Long
    GetNewActor = -1
    If NumGuards > MAX_GUARDS Then Exit Function
    With Guards(NumGuards)
        .inUse = True
        .x = 0: .y = 0: .angle = 0: .actorType = 0
        .health = 0: .speed = 0: .ticcount = 0: .temp2 = 0: .distance = 0
        .tileX = 0: .tileY = 0: .areanumber = 0
        .waitfordoorx = 0: .waitfordoory = 0
        .flags = 0: .state = 0: .dir = 0: .sprite = -1
    End With
    GetNewActor = NumGuards
    NumGuards = NumGuards + 1
End Function

' Runs one actor for the given tics. Returns False if it should be removed.
Private Function DoGuard(ByVal i As Long, ByVal tics As Long) As Boolean
    Dim think As Long, t As Long
    t = Guards(i).actorType
    If Guards(i).ticcount <> 0 Then
        Guards(i).ticcount = Guards(i).ticcount - tics
        Do While Guards(i).ticcount <= 0
            think = StAction(t, Guards(i).state)        ' end of state action
            If think <> TH_NONE Then
                AI_Run think, i, tics
                If Guards(i).state = st_remove Then Exit Function
            End If
            Guards(i).state = StNext(t, Guards(i).state)
            If Guards(i).state = st_remove Then Exit Function
            If StTimeout(t, Guards(i).state) = 0 Then
                Guards(i).ticcount = 0
                Exit Do
            End If
            Guards(i).ticcount = Guards(i).ticcount + StTimeout(t, Guards(i).state)
        Loop
    End If
    think = StThink(t, Guards(i).state)
    If think <> TH_NONE Then
        AI_Run think, i, tics
        If Guards(i).state = st_remove Then Exit Function
    End If
    DoGuard = True
End Function

Public Sub Actors_StateChange(ByVal i As Long, ByVal newState As Long)
    Guards(i).state = newState
    If newState = st_remove Then
        Guards(i).ticcount = 0
    Else
        Guards(i).ticcount = StTimeout(Guards(i).actorType, newState)
    End If
End Sub

Public Sub Actors_Process(ByVal tics As Long)
    Dim n As Long, tex As Long, t As Long
    For n = 0 To NumGuards - 1
        If Guards(n).inUse Then
            If Not DoGuard(n, tics) Then
                Sprites_Remove Guards(n).sprite
                Guards(n).inUse = False
            Else
                With Guards(n)
                    t = .actorType
                    Sprites_SetPos .sprite, .x, .y, .angle
                    tex = StTexture(t, .state)
                    If StRotate(t, .state) Then
                        tex = tex + Add8Dir(Math_Get8Dir(Angle_DistCW(FINE2RAD(Player.angle), FINE2RAD(.angle))))
                    End If
                    Sprites_SetTex .sprite, 0, tex
                End With
            End If
        End If
    Next
End Sub

Private Function Spawn(ByVal skill As Long, ByVal which As Long, ByVal x As Long, ByVal y As Long, ByVal dir As Long) As Long
    Dim i As Long
    i = GetNewActor()
    Spawn = i
    If i < 0 Then Exit Function
    With Guards(i)
        .x = TILE2POS(x)
        .y = TILE2POS(y)
        .tileX = x
        .tileY = y
        .angle = dir4angle(dir)
        .dir = dir4to8(dir)
        .areanumber = Areas(x, y)
        If .areanumber < 0 Then .areanumber = 0
        .actorType = which
        .health = StartHitPoints(skill, which)
        .sprite = Sprites_GetNew()
    End With
End Function

Private Function RandomTics(ByVal which As Long, ByVal st As Long) As Long
    If StTimeout(which, st) <> 0 Then RandomTics = Random_Rnd() Mod StTimeout(which, st) + 1
End Function

Public Sub Actors_SpawnDeadGuard(ByVal skill As Long, ByVal which As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long
    i = Spawn(skill, which, x, y, dir4_nodir)
    If i < 0 Then Exit Sub
    Guards(i).state = st_dead
    Guards(i).speed = 0
    Guards(i).health = 0
    Guards(i).ticcount = RandomTics(which, st_dead)
End Sub

Public Sub Actors_SpawnPatrol(ByVal skill As Long, ByVal which As Long, ByVal x As Long, ByVal y As Long, ByVal dir As Long)
    Dim i As Long
    i = Spawn(skill, which, x, y, dir)
    If i < 0 Then Exit Sub
    Guards(i).state = st_path1
    If which = en_dog Then Guards(i).speed = SPDDOG Else Guards(i).speed = SPDPATROL
    Guards(i).distance = TILEGLOBAL
    Guards(i).ticcount = RandomTics(which, st_path1)
    Guards(i).flags = Guards(i).flags Or FL_SHOOTABLE
    TotalMonsters = TotalMonsters + 1
End Sub

Public Sub Actors_SpawnStand(ByVal skill As Long, ByVal which As Long, ByVal x As Long, ByVal y As Long, ByVal dir As Long)
    Dim i As Long
    i = Spawn(skill, which, x, y, dir)
    If i < 0 Then Exit Sub
    Guards(i).state = st_stand
    Guards(i).speed = SPDPATROL
    Guards(i).ticcount = RandomTics(which, st_stand)
    Guards(i).flags = Guards(i).flags Or FL_SHOOTABLE
    If (TileMap(x, y) And AMBUSH_TILE) <> 0 Then Guards(i).flags = Guards(i).flags Or FL_AMBUSH
    TotalMonsters = TotalMonsters + 1
End Sub
