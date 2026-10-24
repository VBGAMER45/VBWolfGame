Attribute VB_Name = "modPlayer"
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
' Ported from js/player.js: spawning, movement and clipping, using doors,
' push walls and elevators, attacking, taking damage and the give* helpers.

Public Const PLAYERSIZE As Long = MINDIST   ' player radius
Public Const STOPSPEED As Long = &HD00&
Public Const FRICTION As Double = 0.25
Public Const MAXMOVE As Long = MINDIST * 2 - 1
Public Const EXTRAPOINTS As Long = 40000    ' points for an extra life
Public Const MINACTORDIST As Long = &H10000 ' from actors.js

Public Const ITEM_KEY_1 As Long = 1
Public Const ITEM_KEY_2 As Long = 2
Public Const ITEM_KEY_3 As Long = 4
Public Const ITEM_KEY_4 As Long = 8
Public Const ITEM_WEAPON_1 As Long = 16
Public Const ITEM_WEAPON_2 As Long = 32
Public Const ITEM_WEAPON_3 As Long = 64
Public Const ITEM_WEAPON_4 As Long = 128
Public Const ITEM_WEAPON_5 As Long = 256
Public Const ITEM_WEAPON_6 As Long = 512
Public Const ITEM_WEAPON_7 As Long = 1024
Public Const ITEM_WEAPON_8 As Long = 2048
Public Const ITEM_BACKPACK As Long = &H1000&    ' doubles carrying capacity
Public Const ITEM_AUGMENT As Long = &H2000&     ' adds 50 to maximum health
Public Const ITEM_UNIFORM As Long = &H4000&     ' allows you to pass guards
Public Const ITEM_AUTOMAP As Long = &H8000&
Public Const ITEM_FREE As Long = &H10000

Public Const PL_FLAG_REUSE As Long = 1      ' use button pressed
Public Const PL_FLAG_ATTCK As Long = 2      ' attacking
Public Const FL_GODMODE As Long = 16
Public Const FL_NOTARGET As Long = 64

Public Const WEAPON_KNIFE As Long = 0
Public Const WEAPON_PISTOL As Long = 1
Public Const WEAPON_AUTO As Long = 2
Public Const WEAPON_CHAIN As Long = 3
Public Const WEAPON_TYPES As Long = 4

Public Const KEY_GOLD As Long = 0
Public Const KEY_SILVER As Long = 1

Public Const AMMO_BULLETS As Long = 0
Public Const AMMO_TYPES As Long = 1

Public Const ex_notingame As Long = 0
Public Const ex_playing As Long = 1
Public Const ex_dead As Long = 2
Public Const ex_secretlevel As Long = 3
Public Const ex_victory As Long = 4
Public Const ex_complete As Long = 5

Public Type TPlayer
    x As Long                   ' position, 16.16 fixed point
    y As Long
    angle As Long               ' fines
    tileX As Long
    tileY As Long
    movX As Long
    movY As Long
    speed As Long
    health As Long
    score As Long
    lives As Long
    startScore As Long
    nextExtra As Long
    items As Long               ' keys and weapons
    weapon As Long
    pendingWeapon As Long
    previousWeapon As Long
    ammo(0 To AMMO_TYPES - 1) As Long
    armor As Long
    cmdForwardMove As Long      ' movement / action command for this cycle
    cmdSideMove As Long
    cmdButtons As Long
    cmdImpulse As Long
    attackFrame As Long
    attackCount As Long
    weaponFrame As Long
    madenoise As Boolean
    faceFrame As Long
    faceCount As Long
    faceGotGun As Boolean
    faceOuch As Boolean
    flags As Long
    areanumber As Long
    playstate As Long
    skill As Long
    lastAttacker As Long        ' actor index, or -1
End Type

Public Player As TPlayer

' attackinfo: per weapon, frames of (tics, attack, weapon view frame).
' attack: -1 = done, 0 = nothing, 1 = fire a bullet, 2 = knife,
'         3 = loop back while firing (auto), 4 = fire and loop (rotary)
Private AttackTics(0 To WEAPON_TYPES - 1, 0 To 3) As Long
Private AttackKind(0 To WEAPON_TYPES - 1, 0 To 3) As Long
Private AttackFrameNum(0 To WEAPON_TYPES - 1, 0 To 3) As Long
Private AttackInfoReady As Boolean

Private Sub InitAttackInfo()
    Dim w As Long, f As Long
    For w = 0 To WEAPON_TYPES - 1
        For f = 0 To 3
            AttackTics(w, f) = 6
            AttackFrameNum(w, f) = (f + 1) Mod 4
        Next
        AttackKind(w, 0) = 0: AttackKind(w, 3) = -1
    Next
    AttackKind(WEAPON_KNIFE, 1) = 2: AttackKind(WEAPON_KNIFE, 2) = 0
    AttackKind(WEAPON_PISTOL, 1) = 1: AttackKind(WEAPON_PISTOL, 2) = 0
    AttackKind(WEAPON_AUTO, 1) = 1: AttackKind(WEAPON_AUTO, 2) = 3
    AttackKind(WEAPON_CHAIN, 1) = 1: AttackKind(WEAPON_CHAIN, 2) = 4
    AttackInfoReady = True
End Sub

' Places the player at the level's start. Keeps score, lives and inventory
' when keepStats is True (moving on to another level).
Public Sub Player_Spawn(ByVal skill As Long, ByVal keepStats As Boolean)
    Dim old As TPlayer, i As Long
    old = Player
    With Player
        .x = SpawnX
        .y = SpawnY
        .angle = SpawnAngle
        .tileX = POS2TILE(SpawnX)
        .tileY = POS2TILE(SpawnY)
        .movX = 0: .movY = 0: .speed = 0
        .health = 100
        .score = 0: .lives = 0: .startScore = 0: .nextExtra = 0
        .items = 0: .weapon = 0: .pendingWeapon = -1: .previousWeapon = -1
        For i = 0 To AMMO_TYPES - 1
            .ammo(i) = 0
        Next
        .armor = 0
        .cmdForwardMove = 0: .cmdSideMove = 0: .cmdButtons = 0: .cmdImpulse = 0
        .attackFrame = 0: .attackCount = 0: .weaponFrame = 0
        .madenoise = False
        .faceFrame = 0: .faceCount = 0: .faceGotGun = False: .faceOuch = False
        .flags = 0
        .areanumber = Areas(.tileX, .tileY)
        .playstate = ex_playing
        .skill = skill
        .lastAttacker = -1
        If .areanumber < 0 Then .areanumber = 36
    End With
    Areas_Init Player.areanumber
    Areas_Connect Player.areanumber
    If keepStats Then
        CopyPlayer old
    Else
        Player_NewGame
    End If
End Sub

Private Sub CopyPlayer(ByRef src As TPlayer)
    Dim i As Long
    With Player
        .health = src.health
        .score = src.score
        .startScore = src.startScore
        .lives = src.lives
        .nextExtra = src.nextExtra
        .items = src.items And (ITEM_WEAPON_1 Or ITEM_WEAPON_2 Or ITEM_WEAPON_3 Or ITEM_WEAPON_4)
        .weapon = src.weapon
        .pendingWeapon = src.pendingWeapon
        .previousWeapon = src.previousWeapon
        For i = 0 To AMMO_TYPES - 1
            .ammo(i) = src.ammo(i)
        Next
    End With
End Sub

Public Sub Player_NewGame()
    With Player
        .health = 100
        .ammo(AMMO_BULLETS) = 8
        .score = 0
        .startScore = 0
        .lives = 3
        .previousWeapon = WEAPON_KNIFE
        .weapon = WEAPON_PISTOL
        .pendingWeapon = WEAPON_PISTOL
        .items = ITEM_WEAPON_1 Or ITEM_WEAPON_2
        .nextExtra = EXTRAPOINTS
    End With
End Sub

Private Function TryMove() As Boolean
    Dim xl As Long, yl As Long, xh As Long, yh As Long, x As Long, y As Long, n As Long, d As Long
    xl = POS2TILE(Player.x - PLAYERSIZE)
    yl = POS2TILE(Player.y - PLAYERSIZE)
    xh = POS2TILE(Player.x + PLAYERSIZE)
    yh = POS2TILE(Player.y + PLAYERSIZE)

    ' checking for solid walls
    For y = yl To yh
        For x = xl To xh
            If (TileMap(x, y) And SOLID_TILE) <> 0 Then Exit Function
            If (TileMap(x, y) And DOOR_TILE) <> 0 Then
                If Doors_Opened(DoorMap(x, y)) <> DOOR_FULLOPEN Then
                    ' iphone hack to allow the player to move halfway into door tiles:
                    ' if the player bounds don't cross the middle of the tile, allow it
                    If Abs(Player.x - TILE2POS(x)) <= &H9000& And Abs(Player.y - TILE2POS(y)) <= &H9000& Then Exit Function
                End If
            End If
        Next
    Next
    ' checking for actors
    For n = 0 To NumGuards - 1
        If Guards(n).inUse And Guards(n).state < st_die1 And (Guards(n).flags And FL_SHOOTABLE) <> 0 Then
            d = Player.x - Guards(n).x
            If d >= -MINACTORDIST And d <= MINACTORDIST Then
                d = Player.y - Guards(n).y
                If d >= -MINACTORDIST And d <= MINACTORDIST Then Exit Function
            End If
        End If
    Next
    TryMove = True
End Function

Private Sub ClipMove(ByVal xmove As Long, ByVal ymove As Long)
    Dim basex As Long, basey As Long
    basex = Player.x
    basey = Player.y

    Player.x = Player.x + xmove
    Player.y = Player.y + ymove
    If TryMove() Then Exit Sub              ' we moved as we wanted

    If xmove <> 0 Then                      ' maybe we can move along x only
        Player.x = basex + xmove
        Player.y = basey
        If TryMove() Then Exit Sub
    End If
    If ymove <> 0 Then                      ' maybe along y only
        Player.x = basex
        Player.y = basey + ymove
        If TryMove() Then Exit Sub
    End If
    ' movement blocked; stay in one place
    Player.x = basex
    Player.y = basey
End Sub

Public Sub Player_ControlMovement(ByVal tics As Long)
    Dim angle As Long, speed As Double, x As Long, y As Long, tx As Long, ty As Long

    angle = Player.angle
    Player.movX = 0: Player.movY = 0        ' clear accumulated movement

    If Player.cmdForwardMove <> 0 Then
        speed = tics * Player.cmdForwardMove
        Player.movX = Player.movX + ToInt32(speed * CosTable(angle))
        Player.movY = Player.movY + ToInt32(speed * SinTable(angle))
    End If
    If Player.cmdSideMove <> 0 Then
        speed = tics * Player.cmdSideMove
        Player.movX = Player.movX + ToInt32(speed * SinTable(angle))
        Player.movY = Player.movY - ToInt32(speed * CosTable(angle))
    End If
    If Player.movX = 0 And Player.movY = 0 Then Exit Sub

    Player.speed = Player.movX + Player.movY

    ' bound movement
    If Player.movX > MAXMOVE Then
        Player.movX = MAXMOVE
    ElseIf Player.movX < -MAXMOVE Then
        Player.movX = -MAXMOVE
    End If
    If Player.movY > MAXMOVE Then
        Player.movY = MAXMOVE
    ElseIf Player.movY < -MAXMOVE Then
        Player.movY = -MAXMOVE
    End If

    ' move the player and clip the movement to walls
    ClipMove Player.movX, Player.movY
    Player.tileX = POS2TILE(Player.x)
    Player.tileY = POS2TILE(Player.y)

    ' pick up items easier: any tile you touch, not just the midpoint tile
    For x = -1 To 1 Step 2
        tx = POS2TILE(Player.x + x * PLAYERSIZE)
        For y = -1 To 1 Step 2
            ty = POS2TILE(Player.y + y * PLAYERSIZE)
            Powerups_PickUp tx, ty
        Next
    Next

    ' area change (ambush tiles and doors have negative values)
    If Areas(Player.tileX, Player.tileY) >= 0 And Areas(Player.tileX, Player.tileY) <> Player.areanumber Then
        Player.areanumber = Areas(Player.tileX, Player.tileY)
        Areas_Connect Player.areanumber
    End If

    If (TileMap(Player.tileX, Player.tileY) And EXIT_TILE) <> 0 Then Game_Victory
End Sub

Private Function Use() As Boolean
    Dim x As Long, y As Long, dir As Long

    dir = Math_Get4Dir(FINE2RAD(Player.angle))
    x = Player.tileX + dx4dir(dir)
    y = Player.tileY + dy4dir(dir)

    If (TileMap(x, y) And DOOR_TILE) <> 0 Then
        Use = Doors_TryUse(DoorMap(x, y))
        Exit Function
    End If
    If (TileMap(x, y) And SECRET_TILE) <> 0 Then
        Use = PushWall_Push(x, y, dir)
        Exit Function
    End If
    If (TileMap(x, y) And ELEVATOR_TILE) <> 0 Then
        Select Case dir
            Case dir4_east, dir4_west
                WallTexX(x, y) = WallTexX(x, y) + 2     ' flip the switch
            Case dir4_north, dir4_south
                Exit Function                           ' don't allow pressing the elevator rails
        End Select
        If (TileMap(Player.tileX, Player.tileY) And SECRETLEVEL_TILE) <> 0 Then
            Player.playstate = ex_secretlevel
        Else
            Player.playstate = ex_complete
        End If
        Sound_Start "lsfx/040.wav", -1, -1
        Game_StartIntermission
        Use = True
    End If
End Function

Public Sub Player_Process(ByVal tics As Long)
    If Not AttackInfoReady Then InitAttackInfo
    Player.madenoise = False
    Player_ControlMovement tics

    If (Player.flags And PL_FLAG_ATTCK) <> 0 Then
        Attack (Player.cmdButtons And BUTTON_ATTACK) <> 0, tics
    Else
        If (Player.cmdButtons And BUTTON_USE) <> 0 Then
            If (Player.flags And PL_FLAG_REUSE) = 0 Then
                If Use() Then Player.flags = Player.flags Or PL_FLAG_REUSE
            End If
        Else
            Player.flags = Player.flags And Not PL_FLAG_REUSE
        End If
        If (Player.cmdButtons And BUTTON_ATTACK) <> 0 Then
            Player.flags = Player.flags Or PL_FLAG_ATTCK
            Player.attackFrame = 0
            Player.attackCount = AttackTics(Player.weapon, 0)
            Player.weaponFrame = AttackFrameNum(Player.weapon, 0)
        End If
    End If
End Sub

Private Sub Attack(ByVal reAttack As Boolean, ByVal tics As Long)
    Dim w As Long, f As Long, kind As Long, fire As Boolean
    Player.attackCount = Player.attackCount - tics
    Do While Player.attackCount <= 0
        w = Player.weapon: f = Player.attackFrame
        kind = AttackKind(w, f)
        fire = False
        Select Case kind
            Case -1                                         ' done
                Player.flags = Player.flags And Not PL_FLAG_ATTCK
                If Player.ammo(AMMO_BULLETS) = 0 Then
                    Player.weapon = WEAPON_KNIFE
                ElseIf Player.weapon <> Player.pendingWeapon Then
                    Player.weapon = Player.pendingWeapon
                End If
                Player.attackFrame = 0
                Player.weaponFrame = 0
                Exit Sub
            Case 4                                          ' rotary: fire, and keep going
                If Player.ammo(AMMO_BULLETS) <> 0 Then
                    If reAttack Then Player.attackFrame = Player.attackFrame - 2
                    fire = True
                End If
            Case 1
                fire = True
            Case 2
                Weapon_FireHit
            Case 3                                          ' auto: keep going while held
                If Player.ammo(AMMO_BULLETS) <> 0 And reAttack Then Player.attackFrame = Player.attackFrame - 2
        End Select
        If fire Then
            If Player.ammo(AMMO_BULLETS) = 0 Then           ' can only happen with the rotary gun
                Player.attackFrame = Player.attackFrame + 1
            Else
                Weapon_FireLead
                Player.ammo(AMMO_BULLETS) = Player.ammo(AMMO_BULLETS) - 1
            End If
        End If
        Player.attackCount = Player.attackCount + AttackTics(w, f)
        Player.attackFrame = Player.attackFrame + 1
        If Player.attackFrame > 3 Then Player.attackFrame = 3
        Player.weaponFrame = AttackFrameNum(Player.weapon, Player.attackFrame)
    Loop
End Sub

' attacker: index of the actor that did the damage.
Public Sub Player_Damage(ByVal attacker As Long, ByVal points As Long)
    If Player.playstate = ex_dead Or Player.playstate = ex_complete Or Player.playstate = ex_victory Then Exit Sub
    Player.lastAttacker = attacker
    If Player.skill = gd_baby Then points = points \ 4
    If (Player.flags And FL_GODMODE) = 0 Then Player.health = Player.health - points
    If Player.health <= 0 Then
        Game_Notify "You have died"
        Player.health = 0
        Player.playstate = ex_dead
        Sound_Start "lsfx/009.wav", -1, -1
    End If
    Game_StartDamageFlash points
    Player.faceGotGun = False
    If points > 30 And Player.health <> 0 Then
        Player.faceOuch = True
        Player.faceCount = 0
    End If
End Sub

'---------------------------------------------------------------- give* helpers

Public Sub Player_GivePoints(ByVal points As Long)
    Player.score = Player.score + points
    Do While Player.score >= Player.nextExtra
        Player.nextExtra = Player.nextExtra + EXTRAPOINTS
        Player_GiveLife
    Loop
End Sub

' Returns True if the player needed the health. max = 0 means the natural limit.
Public Function Player_GiveHealth(ByVal points As Long, ByVal max As Long) As Boolean
    If max = 0 Then
        If (Player.items And ITEM_AUGMENT) <> 0 Then max = 150 Else max = 100
    End If
    If Player.health >= max Then Exit Function     ' doesn't need this health
    Player.health = Player.health + points
    If Player.health > max Then Player.health = max
    Player.faceGotGun = False
    Player_GiveHealth = True
End Function

Public Sub Player_GiveLife()
    If Player.lives < 9 Then Player.lives = Player.lives + 1
End Sub

Public Sub Player_GiveKey(ByVal key As Long)
    Player.items = Player.items Or (ITEM_KEY_1 * 2 ^ key)
End Sub

Public Sub Player_GiveWeapon(ByVal weapon As Long)
    Dim itemflag As Long
    Player_GiveAmmo AMMO_BULLETS, 6         ' give some ammo with a weapon
    itemflag = ITEM_WEAPON_1 * 2 ^ weapon
    If (Player.items And itemflag) <> 0 Then Exit Sub   ' player owns this weapon
    Player.items = Player.items Or itemflag
    ' don't switch if already using a better weapon
    If Player.weapon < weapon Then
        Player.weapon = weapon
        Player.pendingWeapon = weapon
    End If
End Sub

Public Function Player_GiveAmmo(ByVal ammoType As Long, ByVal ammo As Long) As Boolean
    Dim maxAmmo As Long
    maxAmmo = 99
    If (Player.items And ITEM_BACKPACK) <> 0 Then maxAmmo = maxAmmo * 2
    If Player.ammo(ammoType) >= maxAmmo Then Exit Function  ' don't need
    If Player.ammo(ammoType) = 0 And Player.attackFrame = 0 Then
        Player.weapon = Player.pendingWeapon                ' knife was out
    End If
    Player.ammo(ammoType) = Player.ammo(ammoType) + ammo
    If Player.ammo(ammoType) > maxAmmo Then Player.ammo(ammoType) = maxAmmo
    Player_GiveAmmo = True
End Function
