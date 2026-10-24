Attribute VB_Name = "modPushWall"
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
' Ported from js/pushwall.js. Only one push wall moves at a time.

Public Type TPushWall
    active As Boolean
    tilesMoved As Long
    pointsMoved As Long         ' 0..127 within the current tile
    dir As Long
    x As Long
    y As Long
    dx As Long
    dy As Long
    texX As Long
    texY As Long
End Type

Public PWall As TPushWall

Public Sub PushWall_Reset()
    With PWall
        .active = False
        .tilesMoved = 0
        .pointsMoved = 0
        .dir = 0
        .x = 0: .y = 0
        .dx = 0: .dy = 0
        .texX = 0: .texY = 0
    End With
End Sub

Public Function PushWall_Push(ByVal x As Long, ByVal y As Long, ByVal dir As Long) As Boolean
    Dim dx As Long, dy As Long
    If PWall.active Then Exit Function          ' another push wall is moving

    dx = dx4dir(dir)
    dy = dy4dir(dir)
    If (TileMap(x + dx, y + dy) And (SOLID_TILE Or DOOR_TILE)) <> 0 Then
        PushWall_Push = True
        Exit Function
    End If

    TileMap(x, y) = TileMap(x, y) And Not SECRET_TILE
    TileMap(x, y) = TileMap(x, y) And Not WALL_TILE
    TileMap(x, y) = TileMap(x, y) Or PUSHWALL_TILE

    FoundSecrets = FoundSecrets + 1
    If FoundSecrets = TotalSecrets Then
        Game_Notify "You found the last secret!"
    Else
        Game_Notify "You found a secret!"
    End If
    Sound_Start "sfx/034.wav", -1, -1

    TileMap(x + dx, y + dy) = TileMap(x + dx, y + dy) Or PUSHWALL_TILE
    WallTexX(x + dx, y + dy) = WallTexX(x, y)
    WallTexY(x + dx, y + dy) = WallTexY(x, y)

    With PWall
        .active = True
        .tilesMoved = 0
        .pointsMoved = 0
        .dir = dir
        .x = x: .y = y
        .dx = dx: .dy = dy
        .texX = WallTexX(x, y)
        .texY = WallTexY(x, y)
    End With
    PushWall_Push = True
End Function

Public Sub PushWall_Process(ByVal tics As Long)
    If Not PWall.active Then Exit Sub

    PWall.pointsMoved = PWall.pointsMoved + tics
    If PWall.pointsMoved < 128 Then Exit Sub

    PWall.pointsMoved = PWall.pointsMoved - 128
    PWall.tilesMoved = PWall.tilesMoved + 1

    TileMap(PWall.x, PWall.y) = TileMap(PWall.x, PWall.y) And Not PUSHWALL_TILE
    PWall.x = PWall.x + PWall.dx
    PWall.y = PWall.y + PWall.dy

    If (TileMap(PWall.x + PWall.dx, PWall.y + PWall.dy) And (SOLID_TILE Or DOOR_TILE Or ACTOR_TILE Or POWERUP_TILE)) <> 0 Or PWall.tilesMoved = 3 Then
        TileMap(PWall.x, PWall.y) = TileMap(PWall.x, PWall.y) And Not PUSHWALL_TILE
        TileMap(PWall.x, PWall.y) = TileMap(PWall.x, PWall.y) Or WALL_TILE     ' wall now
        WallTexX(PWall.x, PWall.y) = PWall.texX
        WallTexY(PWall.x, PWall.y) = PWall.texY
        PWall.active = False                                                   ' free the push wall
    Else
        TileMap(PWall.x + PWall.dx, PWall.y + PWall.dy) = TileMap(PWall.x + PWall.dx, PWall.y + PWall.dy) Or PUSHWALL_TILE
        WallTexX(PWall.x + PWall.dx, PWall.y + PWall.dy) = PWall.texX
        WallTexY(PWall.x + PWall.dx, PWall.y + PWall.dy) = PWall.texY
    End If
End Sub
