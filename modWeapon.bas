Attribute VB_Name = "modWeapon"
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
' Ported from js/weapon.js.

' The nearest shootable actor near the line of fire with a clear line of
' sight, or -1. dist receives its distance along the line of fire.
Private Function ClosestTarget(ByRef dist As Long) As Long
    Dim n As Long, shotDist As Long, d1 As Long
    ClosestTarget = -1
    dist = &H7FFFFFFF
    For n = 0 To NumGuards - 1
        If Guards(n).inUse And (Guards(n).flags And FL_SHOOTABLE) <> 0 Then
            shotDist = Math_Point2LineDist(Guards(n).x - Player.x, Guards(n).y - Player.y, Player.angle)
            If shotDist <= (2 * TILEGLOBAL / 3) Then
                d1 = Math_LineLen2Point(Guards(n).x - Player.x, Guards(n).y - Player.y, Player.angle)
                If d1 >= 0 And d1 <= dist Then
                    If Level_CheckLine(Guards(n).x, Guards(n).y, Player.x, Player.y) Then
                        dist = d1
                        ClosestTarget = n
                    End If
                End If
            End If
        End If
    Next
End Function

' Knife.
Public Sub Weapon_FireHit()
    Dim closest As Long, dist As Long
    Sound_Start "lsfx/023.wav", -1, -1
    closest = ClosestTarget(dist)
    If closest < 0 Or dist > TILE2POS(1) Then Exit Sub     ' missed if further than 1.5 tiles
    AI_DamageActor closest, Random_Rnd() \ 16
End Sub

' Pistol, machine gun and rotary gun.
Public Sub Weapon_FireLead()
    Dim closest As Long, dist As Long, dx As Long, dy As Long, damage As Double
    Dim tp As TTracePoint

    Select Case Player.weapon
        Case WEAPON_PISTOL: Sound_Start "sfx/012.wav", -1, -1
        Case WEAPON_AUTO: Sound_Start "sfx/011.wav", -1, -1
        Case WEAPON_CHAIN: Sound_Start "sfx/013.wav", -1, -1
    End Select
    Player.madenoise = True

    closest = ClosestTarget(dist)
    If closest < 0 Then                             ' missed
        tp.angle = Math_NormalizeAngle(Player.angle - DEG2FINE(2) + JsMod(Rnd * 65536#, DEG2FINE(4)))
        tp.x = Player.x
        tp.y = Player.y
        tp.flags = TRACE_BULLET
        Raycaster_Trace tp, False
        If (tp.flags And TRACE_HIT_DOOR) <> 0 Then Sound_Start "lsfx/028.wav", -1, -1
        Exit Sub
    End If

    dx = Abs(Guards(closest).tileX - Player.tileX)
    dy = Abs(Guards(closest).tileY - Player.tileY)
    If dx > dy Then dist = dx Else dist = dy
    If dist < 2 Then
        damage = Random_Rnd() / 4
    ElseIf dist < 4 Then
        damage = Random_Rnd() / 6
    Else
        If Random_Rnd() / 12 < dist Then Exit Sub   ' missed
        damage = Random_Rnd() / 6
    End If
    AI_DamageActor closest, damage
End Sub
