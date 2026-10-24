Attribute VB_Name = "modAI"
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
' Ported from js/ai.js and js/actorai.js for the generic enemy types. JS
' passes think routines as functions; here the state tables hold TH_*
' numbers and AI_Run dispatches them.

Public Const TH_NONE As Long = 0
Public Const TH_STAND As Long = 1
Public Const TH_PATH As Long = 2
Public Const TH_CHASE As Long = 3
Public Const TH_DOGCHASE As Long = 4
Public Const TH_SHOOT As Long = 5
Public Const TH_BITE As Long = 6
Public Const TH_DEATHSCREAM As Long = 7

Public Const RUNSPEED As Long = 6000
Public Const MINSIGHT As Long = &H18000

' the routines T_Advance calls each time an actor reaches a tile centre
Private Const ADV_PATH As Long = 1
Private Const ADV_CHASE As Long = 2
Private Const ADV_DODGE As Long = 3

Public Sub AI_Run(ByVal think As Long, ByVal i As Long, ByVal tics As Long)
    Select Case think
        Case TH_STAND: T_Stand i, tics
        Case TH_PATH: T_Path i, tics
        Case TH_CHASE: T_Chase i, tics
        Case TH_DOGCHASE: T_DogChase i, tics
        Case TH_SHOOT: T_Shoot i
        Case TH_BITE: T_Bite i
        Case TH_DEATHSCREAM: DeathScream i
    End Select
End Sub

'---------------------------------------------------------------- ai.js

Private Function CheckSight(ByVal i As Long) As Boolean
    Dim deltax As Long, deltay As Long
    With Guards(i)
        If (.flags And FL_AMBUSH) = 0 Then
            If Not AreaByPlayer(.areanumber) Then Exit Function
        End If
        deltax = Player.x - .x
        deltay = Player.y - .y
        If Abs(deltax) < MINSIGHT And Abs(deltay) < MINSIGHT Then
            CheckSight = True
            Exit Function
        End If
        Select Case .dir
            Case dir8_north: If deltay < 0 Then Exit Function
            Case dir8_east: If deltax < 0 Then Exit Function
            Case dir8_south: If deltay > 0 Then Exit Function
            Case dir8_west: If deltax > 0 Then Exit Function
        End Select
        CheckSight = Level_CheckLine(.x, .y, Player.x, Player.y)
    End With
End Function

' Tries to start moving one tile in newDir. Returns False if blocked.
Private Function ChangeDir(ByVal i As Long, ByVal newDir As Long) As Boolean
    Dim oldx As Long, oldy As Long, newx As Long, newy As Long, n As Long, moveok As Boolean

    oldx = POS2TILE(Guards(i).x)
    oldy = POS2TILE(Guards(i).y)
    newx = oldx + dx8dir(newDir)
    newy = oldy + dy8dir(newDir)

    If (newDir And 1) <> 0 Then                         ' diagonal
        If (TileMap(newx, oldy) And SOLID_TILE) <> 0 Then Exit Function
        If (TileMap(oldx, newy) And SOLID_TILE) <> 0 Then Exit Function
        If (TileMap(newx, newy) And SOLID_TILE) <> 0 Then Exit Function
        For n = 0 To NumGuards - 1
            If Guards(n).inUse And Guards(n).state < st_die1 Then
                If Guards(n).tileX = newx And Guards(n).tileY = newy Then Exit Function
                If Guards(n).tileX = oldx And Guards(n).tileY = newy Then Exit Function
                If Guards(n).tileX = newx And Guards(n).tileY = oldy Then Exit Function
            End If
        Next
    Else                                                ' E, N, W, S
        If (TileMap(newx, newy) And SOLID_TILE) <> 0 Then Exit Function
        If (TileMap(newx, newy) And DOOR_TILE) <> 0 Then
            If Guards(i).actorType = en_fake Or Guards(i).actorType = en_dog Then
                ' they can't open doors
                If Doors(DoorMap(newx, newy)).action <> dr_open Then Exit Function
            Else
                Guards(i).waitfordoorx = newx
                Guards(i).waitfordoory = newy
                moveok = True
            End If
        End If
        If Not moveok Then
            For n = 0 To NumGuards - 1
                If Guards(n).inUse And Guards(n).state < st_die1 Then
                    If Guards(n).tileX = newx And Guards(n).tileY = newy Then Exit Function
                End If
            Next
        End If
    End If

    With Guards(i)
        .tileX = newx
        .tileY = newy
        TileMap(oldx, oldy) = TileMap(oldx, oldy) And Not ACTOR_TILE
        TileMap(newx, newy) = TileMap(newx, newy) Or ACTOR_TILE
        If Areas(newx, newy) > 0 Then .areanumber = Areas(newx, newy)
        .distance = TILEGLOBAL
        .dir = newDir
    End With
    ChangeDir = True
End Function

Private Sub PathDir(ByVal i As Long)
    Dim tileinfo As Long
    tileinfo = TileMap(POS2TILE(Guards(i).x), POS2TILE(Guards(i).y))
    If (tileinfo And WAYPOINT_TILE) <> 0 Then
        If (tileinfo And TILE_IS_E_TURN) <> 0 Then
            Guards(i).dir = dir8_east
        ElseIf (tileinfo And TILE_IS_NE_TURN) <> 0 Then
            Guards(i).dir = dir8_northeast
        ElseIf (tileinfo And TILE_IS_N_TURN) <> 0 Then
            Guards(i).dir = dir8_north
        ElseIf (tileinfo And TILE_IS_NW_TURN) <> 0 Then
            Guards(i).dir = dir8_northwest
        ElseIf (tileinfo And TILE_IS_W_TURN) <> 0 Then
            Guards(i).dir = dir8_west
        ElseIf (tileinfo And TILE_IS_SW_TURN) <> 0 Then
            Guards(i).dir = dir8_southwest
        ElseIf (tileinfo And TILE_IS_S_TURN) <> 0 Then
            Guards(i).dir = dir8_south
        ElseIf (tileinfo And TILE_IS_SE_TURN) <> 0 Then
            Guards(i).dir = dir8_southeast
        End If
    End If
    If Not ChangeDir(i, Guards(i).dir) Then Guards(i).dir = dir8_nodir
End Sub

' Returns True once the actor has noticed the player and reacted.
Private Function FindTarget(ByVal i As Long, ByVal tics As Long) As Boolean
    With Guards(i)
        If .temp2 <> 0 Then                         ' count down reaction time
            .temp2 = .temp2 - tics
            If .temp2 > 0 Then Exit Function
            .temp2 = 0                              ' time to react
        Else
            If (Player.flags And FL_NOTARGET) <> 0 Then Exit Function
            If (.flags And FL_AMBUSH) = 0 And Not AreaByPlayer(.areanumber) Then Exit Function
            If Not CheckSight(i) Then
                If (.flags And FL_AMBUSH) <> 0 Or Not Player.madenoise Then Exit Function
            End If
            .flags = .flags And Not FL_AMBUSH
            Select Case .actorType
                Case en_guard: .temp2 = 1 + Random_Rnd() / 4
                Case en_officer: .temp2 = 2
                Case en_mutant: .temp2 = 1 + Random_Rnd() / 6
                Case en_ss: .temp2 = 1 + Random_Rnd() / 6
                Case en_dog: .temp2 = 1 + Random_Rnd() / 8
                Case Else: .temp2 = 1
            End Select
            Exit Function                           ' amazed, waiting to understand what to do
        End If
    End With
    FirstSighting i
    FindTarget = True
End Function

Private Sub Chase(ByVal i As Long)
    Dim deltax As Long, deltay As Long, d0 As Long, d1 As Long, tdir As Long
    Dim olddir As Long, turnaround As Long

    If Player.playstate = ex_victory Then Exit Sub
    olddir = Guards(i).dir
    turnaround = opposite8(olddir)
    d0 = dir8_nodir: d1 = dir8_nodir

    deltax = POS2TILE(Player.x) - POS2TILE(Guards(i).x)
    deltay = POS2TILE(Player.y) - POS2TILE(Guards(i).y)
    If deltax > 0 Then
        d0 = dir8_east
    ElseIf deltax < 0 Then
        d0 = dir8_west
    End If
    If deltay > 0 Then
        d1 = dir8_north
    ElseIf deltay < 0 Then
        d1 = dir8_south
    End If
    If Abs(deltay) > Abs(deltax) Then
        tdir = d0: d0 = d1: d1 = tdir
    End If
    If d0 = turnaround Then d0 = dir8_nodir
    If d1 = turnaround Then d1 = dir8_nodir

    If d0 <> dir8_nodir Then
        If ChangeDir(i, d0) Then Exit Sub
    End If
    If d1 <> dir8_nodir Then
        If ChangeDir(i, d1) Then Exit Sub
    End If
    If olddir <> dir8_nodir Then
        If ChangeDir(i, olddir) Then Exit Sub
    End If
    If Random_Rnd() > 128 Then                      ' randomly determine direction of search
        For tdir = dir8_east To dir8_south Step 2
            If tdir <> turnaround Then
                If ChangeDir(i, tdir) Then Exit Sub
            End If
        Next
    Else
        For tdir = dir8_south To dir8_east Step -2
            If tdir <> turnaround Then
                If ChangeDir(i, tdir) Then Exit Sub
            End If
        Next
    End If
    If turnaround <> dir8_nodir Then
        If ChangeDir(i, turnaround) Then Exit Sub
    End If
    Guards(i).dir = dir8_nodir                      ' can't move
End Sub

Private Sub Dodge(ByVal i As Long)
    Dim deltax As Long, deltay As Long, n As Long, turnaround As Long, tdir As Long
    Dim dirtry(0 To 4) As Long

    If Player.playstate = ex_victory Then Exit Sub
    If (Guards(i).flags And FL_FIRSTATTACK) <> 0 Then
        turnaround = dir8_nodir
        Guards(i).flags = Guards(i).flags And Not FL_FIRSTATTACK
    Else
        turnaround = opposite8(Guards(i).dir)
    End If

    deltax = POS2TILE(Player.x) - POS2TILE(Guards(i).x)
    deltay = POS2TILE(Player.y) - POS2TILE(Guards(i).y)
    If deltax > 0 Then
        dirtry(1) = dir8_east: dirtry(3) = dir8_west
    Else
        dirtry(1) = dir8_west: dirtry(3) = dir8_east
    End If
    If deltay > 0 Then
        dirtry(2) = dir8_north: dirtry(4) = dir8_south
    Else
        dirtry(2) = dir8_south: dirtry(4) = dir8_north
    End If
    If Abs(deltax) > Abs(deltay) Then
        tdir = dirtry(1): dirtry(1) = dirtry(2): dirtry(2) = tdir
        tdir = dirtry(3): dirtry(3) = dirtry(4): dirtry(4) = tdir
    End If
    If Random_Rnd() < 128 Then
        tdir = dirtry(1): dirtry(1) = dirtry(2): dirtry(2) = tdir
        tdir = dirtry(3): dirtry(3) = dirtry(4): dirtry(4) = tdir
    End If
    dirtry(0) = diagonal(dirtry(1), dirtry(2))

    For n = 0 To 4
        If dirtry(n) <> dir8_nodir And dirtry(n) <> turnaround Then
            If ChangeDir(i, dirtry(n)) Then Exit Sub
        End If
    Next
    If turnaround <> dir8_nodir Then
        If ChangeDir(i, turnaround) Then Exit Sub
    End If
    Guards(i).dir = dir8_nodir
End Sub

Private Sub RunAdvance(ByVal which As Long, ByVal i As Long)
    Select Case which
        Case ADV_PATH: PathDir i
        Case ADV_CHASE: Chase i
        Case ADV_DODGE: Dodge i
    End Select
End Sub

Private Sub T_Stand(ByVal i As Long, ByVal tics As Long)
    FindTarget i, tics
End Sub

Private Sub T_Path(ByVal i As Long, ByVal tics As Long)
    If FindTarget(i, tics) Then Exit Sub
    If Guards(i).speed = 0 Then Exit Sub            ' patrolling with a speed of 0
    If Guards(i).dir = dir8_nodir Then
        PathDir i
        If Guards(i).dir = dir8_nodir Then Exit Sub ' all movement is blocked
    End If
    T_Advance i, ADV_PATH, tics
End Sub

Private Sub T_Shoot(ByVal i As Long)
    Dim dx As Long, dy As Long, dist As Double, hitchance As Double, damage As Long, trans As Double

    With Guards(i)
        If Not AreaByPlayer(.areanumber) Then Exit Sub
        If Not Level_CheckLine(.x, .y, Player.x, Player.y) Then Exit Sub   ' player is behind a wall

        dx = Abs(POS2TILE(.x) - POS2TILE(Player.x))
        dy = Abs(POS2TILE(.y) - POS2TILE(Player.y))
        If dx > dy Then dist = dx Else dist = dy
        If .actorType = en_ss Then dist = dist * 2 / 3      ' ss are better shots

        If Player.speed >= RUNSPEED Then hitchance = 160 Else hitchance = 256

        ' (the original compares radians with degrees here; kept as is)
        trans = Math_TransformPoint(.x, .y, Player.x, Player.y)
        If Angle_Diff(trans, FINE2DEG(Player.angle)) < PI / 3 Then
            hitchance = hitchance - dist * 16
        Else
            hitchance = hitchance - dist * 8
        End If

        If Random_Rnd() < hitchance Then
            If dist < 2 Then
                damage = Random_Rnd() \ 4
            ElseIf dist < 4 Then
                damage = Random_Rnd() \ 8
            Else
                damage = Random_Rnd() \ 16
            End If
            Player_Damage i, damage
        End If

        If .actorType = en_ss Then
            Sound_Start "sfx/024.wav", .x, .y
        Else
            Sound_Start "sfx/049.wav", .x, .y
        End If
    End With
End Sub

Private Sub T_Chase(ByVal i As Long, ByVal tics As Long)
    Dim dx As Long, dy As Long, dist As Long, chance As Double, shouldDodge As Boolean

    With Guards(i)
        If Level_CheckLine(.x, .y, Player.x, Player.y) Then     ' got a shot at the player?
            dx = Abs(POS2TILE(.x) - POS2TILE(Player.x))
            dy = Abs(POS2TILE(.y) - POS2TILE(Player.y))
            If dx > dy Then dist = dx Else dist = dy
            If dist = 0 Or (dist = 1 And .distance < 16) Then
                chance = 300
            Else
                chance = (tics * 16) / dist
            End If
            If Random_Rnd() < chance Then
                Actors_StateChange i, st_shoot1
                Exit Sub
            End If
            shouldDodge = True
        End If
    End With

    If Guards(i).dir = dir8_nodir Then
        If shouldDodge Then Dodge i Else Chase i
        If Guards(i).dir = dir8_nodir Then Exit Sub         ' blocked in
        Guards(i).angle = dir8angle(Guards(i).dir)
    End If
    If shouldDodge Then T_Advance i, ADV_DODGE, tics Else T_Advance i, ADV_CHASE, tics
End Sub

Private Sub T_DogChase(ByVal i As Long, ByVal tics As Long)
    Dim dx As Long, dy As Long
    If Guards(i).dir = dir8_nodir Then
        Dodge i
        Guards(i).angle = dir8angle(Guards(i).dir)
        If Guards(i).dir = dir8_nodir Then Exit Sub         ' blocked in
    End If
    dx = Abs(Player.x - Guards(i).x) - TILEGLOBAL \ 2
    If dx <= MINACTORDIST Then
        dy = Abs(Player.y - Guards(i).y) - TILEGLOBAL \ 2
        If dy <= MINACTORDIST Then
            Actors_StateChange i, st_shoot1
            Exit Sub                                        ' bite the player
        End If
    End If
    T_Advance i, ADV_DODGE, tics
End Sub

' Moves an actor along its path; at each tile centre the given routine
' picks the next direction.
Private Sub T_Advance(ByVal i As Long, ByVal which As Long, ByVal tics As Long)
    Dim move As Long, d As Long
    move = Guards(i).speed * tics
    Do While move > 0
        If Guards(i).waitfordoorx <> 0 Then
            d = DoorMap(Guards(i).waitfordoorx, Guards(i).waitfordoory)
            Doors_Open d
            If Doors(d).action <> dr_open Then Exit Sub     ' not opened yet
            Guards(i).waitfordoorx = 0                      ' go ahead, the door is now open
            Guards(i).waitfordoory = 0
        End If
        If move < Guards(i).distance Then
            T_Move i, move
            Exit Do
        End If
        Guards(i).x = TILE2POS(Guards(i).tileX)
        Guards(i).y = TILE2POS(Guards(i).tileY)
        move = move - Guards(i).distance
        RunAdvance which, i
        Guards(i).angle = dir8angle(Guards(i).dir)
        If Guards(i).dir = dir8_nodir Then Exit Sub         ' all movement is blocked
    Loop
End Sub

Private Sub T_Move(ByVal i As Long, ByVal dist As Long)
    With Guards(i)
        If .dir = dir8_nodir Or dist = 0 Then Exit Sub
        .x = .x + dist * dx8dir(.dir)
        .y = .y + dist * dy8dir(.dir)
        If Abs(.x - Player.x) <= MINACTORDIST Then
            If Abs(.y - Player.y) <= MINACTORDIST Then
                .x = .x - dist * dx8dir(.dir)
                .y = .y - dist * dy8dir(.dir)
                Exit Sub
            End If
        End If
        .distance = .distance - dist
        If .distance < 0 Then .distance = 0
    End With
End Sub

Private Sub T_Bite(ByVal i As Long)
    Dim dx As Long, dy As Long
    Sound_Start "sfx/002.wav", Guards(i).x, Guards(i).y
    dx = Abs(Player.x - Guards(i).x) - TILEGLOBAL
    If dx <= MINACTORDIST Then
        dy = Abs(Player.y - Guards(i).y) - TILEGLOBAL
        If dy <= MINACTORDIST Then
            If Random_Rnd() < 180 Then Player_Damage i, Random_Rnd() \ 16
        End If
    End If
End Sub

'---------------------------------------------------------------- actorai.js

Private Sub DeathScream(ByVal i As Long)
    Select Case Guards(i).actorType
        Case en_mutant: Sound_Start "sfx/037.wav", Guards(i).x, Guards(i).y
        Case en_officer: Sound_Start "sfx/074.wav", Guards(i).x, Guards(i).y
        Case en_ss: Sound_Start "sfx/046.wav", Guards(i).x, Guards(i).y
        Case en_dog: Sound_Start "sfx/035.wav", Guards(i).x, Guards(i).y
        Case en_guard
            Sound_Start "sfx/" & Choose(Random_Rnd() Mod 6 + 1, "025", "026", "086", "088", "105", "107") & ".wav", Guards(i).x, Guards(i).y
    End Select
End Sub

Private Sub FirstSighting(ByVal i As Long)
    With Guards(i)
        Select Case .actorType
            Case en_guard: .speed = .speed * 3      ' go faster when chasing the player
            Case en_officer: .speed = .speed * 5
            Case en_mutant: .speed = .speed * 3
            Case en_ss: .speed = .speed * 4
            Case en_dog: .speed = .speed * 2
            Case Else: Exit Sub
        End Select
    End With
    Sound_Start "robot/alert", Guards(i).x, Guards(i).y     ' VB port: a robot chirp in place of the voice clips
    Actors_StateChange i, st_chase1
    Guards(i).waitfordoorx = 0                      ' ignore the door opening command
    Guards(i).waitfordoory = 0
    Guards(i).dir = dir8_nodir
    Guards(i).flags = Guards(i).flags Or FL_ATTACKMODE Or FL_FIRSTATTACK
End Sub

Public Sub AI_DamageActor(ByVal i As Long, ByVal damage As Double)
    Player.madenoise = True
    If (Guards(i).flags And FL_ATTACKMODE) = 0 Then damage = ToInt32(damage) * 2
    Guards(i).health = Guards(i).health - damage
    If Guards(i).health <= 0 Then
        KillActor i
    Else
        If (Guards(i).flags And FL_ATTACKMODE) = 0 Then FirstSighting i   ' put into combat mode
        Select Case Guards(i).actorType                 ' dogs only have one hit point
            Case en_guard, en_officer, en_mutant, en_ss
                If (ToInt32(Guards(i).health) And 1) <> 0 Then
                    Actors_StateChange i, st_pain
                Else
                    Actors_StateChange i, st_pain1
                End If
        End Select
    End If
End Sub

Private Sub KillActor(ByVal i As Long)
    Dim tilex As Long, tiley As Long
    tilex = POS2TILE(Guards(i).x): Guards(i).tileX = tilex     ' drop items on the centre
    tiley = POS2TILE(Guards(i).y): Guards(i).tileY = tiley
    Select Case Guards(i).actorType
        Case en_guard
            Player_GivePoints 100
            Powerups_Spawn tilex, tiley, pow_clip2
        Case en_officer
            Player_GivePoints 400
            Powerups_Spawn tilex, tiley, pow_clip2
        Case en_mutant
            Player_GivePoints 700
            Powerups_Spawn tilex, tiley, pow_clip2
        Case en_ss
            Player_GivePoints 500
            If (Player.items And ITEM_WEAPON_3) <> 0 Then
                Powerups_Spawn tilex, tiley, pow_clip2
            Else
                Powerups_Spawn tilex, tiley, pow_machinegun
            End If
        Case en_dog
            Player_GivePoints 200
    End Select
    Actors_StateChange i, st_die1
    KilledMonsters = KilledMonsters + 1
    If KilledMonsters = TotalMonsters Then Game_Notify "You killed the last enemy!"
    Guards(i).flags = Guards(i).flags And Not FL_SHOOTABLE
    Guards(i).flags = Guards(i).flags Or FL_NONMARK
End Sub
