Attribute VB_Name = "modArt"
Option Explicit

' This file is part of VBWolf, a Visual Basic 6 port of the Wolf3D Browser
' Version GPL Source Code, Copyright (C) 2012 id Software LLC, a ZeniMax
' Media company.
'
' This program is free software: you can redistribute it and/or modify it
' under the terms of the GNU General Public License as published bodyY the Free
' Software Foundation, either version 2 of the License, or (at your option)
' any later version. It is distributed WITHOUT ANY WARRANTY; see COPYING.txt.
'
' New for the VB6 port. The GPL release ships no game data, so this module
' draws original replacement art at startup, filling the texture and sprite
' slots the engine refers to:
'   wall texture 2k (light) and 2k+1 (dark) for map wall code k+1,
'   98..105 for doors, sprite 2+n for static object n, and the actor sprite
'   slots used bodyY the state tables in modActStat (robot enemies).
' Slots without art get a magenta checker so they are easy to spot.
' Pixels are &H00RRGGBB, stored column-major (index = x * 64 + y).

Public Const TEX_SIZE As Long = 64
Public Const TEX_PIXELS As Long = 4096
Public Const CLEAR_PIXEL As Long = &H1000000
Public Const NUM_TEX_SLOTS As Long = 106
Public Const NUM_SPR_SLOTS As Long = ACTOR_SPRITE_MAX + 1
Public Const NUM_WEAPON_VIEWS As Long = 16       ' weapon * 4 + frame

Public TexPix() As Long
Public SprPix() As Long
Public SprTop() As Long         ' first opaque row per sprite column (TEX_SIZE if empty)
Public SprBot() As Long         ' last opaque row per sprite column (-1 if empty)
Public WeapPix() As Long        ' first-person weapon views, same layout as sprites
Public WeapTop() As Long
Public WeapBot() As Long

Private SeedD As Double
Private CurBuf As Long          ' 0 = sprite, 1 = weapon view (target of the S* helpers)
Private Drawn() As Boolean      ' actor sprite slots already drawn (states can share slots)

Public Sub Art_Build()
    Dim i As Long, t As Long, w As Long, f As Long
    ReDim TexPix(0 To NUM_TEX_SLOTS * TEX_PIXELS - 1)
    ReDim SprPix(0 To NUM_SPR_SLOTS * TEX_PIXELS - 1)
    ReDim SprTop(0 To NUM_SPR_SLOTS * TEX_SIZE - 1)
    ReDim SprBot(0 To NUM_SPR_SLOTS * TEX_SIZE - 1)
    ReDim WeapPix(0 To NUM_WEAPON_VIEWS * TEX_PIXELS - 1)
    ReDim WeapTop(0 To NUM_WEAPON_VIEWS * TEX_SIZE - 1)
    ReDim WeapBot(0 To NUM_WEAPON_VIEWS * TEX_SIZE - 1)
    SeedD = 12345#

    For t = 0 To NUM_TEX_SLOTS - 1
        MakeMissingTex t
    Next
    ' wall codes 1..8, light faces; the dark faces are derived below
    MakeStone 0
    MakeStone 2: AddVent 2
    MakeBrick 4
    MakeBrick 6: AddPlaque 6
    MakeWood 8
    MakeWood 10: AddPainting 10
    MakeTiles 12
    MakeMetal 14
    For t = 0 To 14 Step 2
        DarkCopy t, t + 1
    Next
    ' wall code 21: elevator rails on north/south faces, the switch on east/west faces
    MakeRails 40
    MakeSwitch 41, False
    MakeRails 42
    MakeSwitch 43, True
    ' doors
    MakeDoor 98, 0: DarkCopy 98, 99
    MakePlate 100: DarkCopy 100, 101
    MakeElevatorDoor 102: DarkCopy 102, 103
    MakeDoor 104, 1
    MakeDoor 105, 2

    CurBuf = 0
    For i = 0 To NUM_SPR_SLOTS * TEX_PIXELS - 1
        SprPix(i) = CLEAR_PIXEL
    Next
    MakePuddle SPR_STAT_0 + 0
    MakeDrum SPR_STAT_0 + 1, 86, 110, 70
    MakeTable SPR_STAT_0 + 2
    MakeFloorLamp SPR_STAT_0 + 3
    MakeChandelier SPR_STAT_0 + 4
    MakePillar SPR_STAT_0 + 7
    MakePottedTree SPR_STAT_0 + 8
    MakePlant SPR_STAT_0 + 11
    MakeUrn SPR_STAT_0 + 12
    MakeCeilingLight SPR_STAT_0 + 14
    MakeKey SPR_STAT_0 + 20, 230, 185, 50
    MakeKey SPR_STAT_0 + 21, 200, 205, 215
    MakeFood SPR_STAT_0 + 24
    MakeFirstAid SPR_STAT_0 + 25
    MakeClip SPR_STAT_0 + 26
    MakeGunPickup SPR_STAT_0 + 27, False
    MakeGunPickup SPR_STAT_0 + 28, True
    MakeCoins SPR_STAT_0 + 29
    MakeGoblet SPR_STAT_0 + 30
    MakeBook SPR_STAT_0 + 31
    MakeCrown SPR_STAT_0 + 32
    MakeOrb SPR_STAT_0 + 33
    MakeDrum SPR_STAT_0 + 35, 125, 85, 50
    MakeAmmoBox SPR_STAT_0 + 49
    For i = SPR_STAT_0 To SPR_STAT_0 + 51
        If SpriteIsEmpty(i) Then MakeMissingSprite i
    Next
    Actors_Reset                        ' loads the state tables that say which slots each actor uses
    ReDim Drawn(0 To NUM_SPR_SLOTS - 1)
    MakeActorArt en_guard
    MakeActorArt en_officer
    MakeActorArt en_ss
    MakeActorArt en_mutant
    MakeActorArt en_dog
    Bounds SprPix, SprTop, SprBot, NUM_SPR_SLOTS

    CurBuf = 1
    For i = 0 To NUM_WEAPON_VIEWS * TEX_PIXELS - 1
        WeapPix(i) = CLEAR_PIXEL
    Next
    For w = 0 To 3
        For f = 0 To 3
            MakeWeaponView w * 4 + f, w, f
        Next
    Next
    Bounds WeapPix, WeapTop, WeapBot, NUM_WEAPON_VIEWS
    CurBuf = 0
End Sub

'---------------------------------------------------------------- helpers

Public Function RGB32(ByVal r As Double, ByVal g As Double, ByVal b As Double) As Long
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
    RGB32 = CLng(Int(r)) * 65536 + CLng(Int(g)) * 256 + CLng(Int(b))
End Function

Public Function Shade(ByVal c As Long, ByVal f As Double) As Long
    Shade = RGB32(((c \ 65536) And 255) * f, ((c \ 256) And 255) * f, (c And 255) * f)
End Function

' deterministic 0..1 noise, so the art is identical on every run
Private Function Frand() As Double
    SeedD = SeedD * 16807#
    SeedD = SeedD - Int(SeedD / 2147483647#) * 2147483647#
    Frand = SeedD / 2147483647#
End Function

Private Sub TPut(ByVal t As Long, ByVal x As Long, ByVal y As Long, ByVal c As Long)
    If x < 0 Or y < 0 Or x > 63 Or y > 63 Then Exit Sub
    TexPix(t * TEX_PIXELS + x * TEX_SIZE + y) = c
End Sub

Private Function TGet(ByVal t As Long, ByVal x As Long, ByVal y As Long) As Long
    TGet = TexPix(t * TEX_PIXELS + x * TEX_SIZE + y)
End Function

Private Sub TRect(ByVal t As Long, ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal c As Long)
    Dim x As Long, y As Long
    For x = x0 To x1
        For y = y0 To y1
            TPut t, x, y, c
        Next
    Next
End Sub

Private Sub DarkCopy(ByVal src As Long, ByVal dst As Long)
    Dim i As Long
    For i = 0 To TEX_PIXELS - 1
        TexPix(dst * TEX_PIXELS + i) = Shade(TexPix(src * TEX_PIXELS + i), 0.72)
    Next
End Sub

Private Sub SPut(ByVal s As Long, ByVal x As Long, ByVal y As Long, ByVal c As Long)
    If x < 0 Or y < 0 Or x > 63 Or y > 63 Then Exit Sub
    If CurBuf = 0 Then
        SprPix(s * TEX_PIXELS + x * TEX_SIZE + y) = c
    Else
        WeapPix(s * TEX_PIXELS + x * TEX_SIZE + y) = c
    End If
End Sub

Private Function SGet(ByVal s As Long, ByVal x As Long, ByVal y As Long) As Long
    SGet = CLEAR_PIXEL
    If x < 0 Or y < 0 Or x > 63 Or y > 63 Then Exit Function
    If CurBuf = 0 Then
        SGet = SprPix(s * TEX_PIXELS + x * TEX_SIZE + y)
    Else
        SGet = WeapPix(s * TEX_PIXELS + x * TEX_SIZE + y)
    End If
End Function

Private Sub SRect(ByVal s As Long, ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal c As Long)
    Dim x As Long, y As Long
    For x = x0 To x1
        For y = y0 To y1
            SPut s, x, y, c
        Next
    Next
End Sub

Private Sub SDisc(ByVal s As Long, ByVal cx As Double, ByVal cy As Double, ByVal r As Double, ByVal c As Long)
    Dim x As Long, y As Long, dx As Double, dy As Double
    For y = Int(cy - r) To Int(cy + r)
        For x = Int(cx - r) To Int(cx + r)
            dx = x + 0.5 - cx: dy = y + 0.5 - cy
            If dx * dx + dy * dy <= r * r Then SPut s, x, y, c
        Next
    Next
End Sub

' Ellipse filled with a flat colour.
Private Sub SEllipse(ByVal s As Long, ByVal cx As Double, ByVal cy As Double, ByVal rx As Double, ByVal ry As Double, ByVal c As Long)
    Dim x As Long, y As Long, dx As Double, dy As Double
    For y = Int(cy - ry) To Int(cy + ry)
        For x = Int(cx - rx) To Int(cx + rx)
            dx = (x + 0.5 - cx) / rx: dy = (y + 0.5 - cy) / ry
            If dx * dx + dy * dy <= 1 Then SPut s, x, y, c
        Next
    Next
End Sub

' Lit sphere, light from the upper left.
Private Sub SSphere(ByVal s As Long, ByVal cx As Double, ByVal cy As Double, ByVal r As Double, ByVal cr As Double, ByVal cg As Double, ByVal cb As Double)
    Dim x As Long, y As Long, nx As Double, ny As Double, q As Double, dif As Double, spec As Double, f As Double
    For y = Int(cy - r) To Int(cy + r)
        For x = Int(cx - r) To Int(cx + r)
            nx = (x + 0.5 - cx) / r: ny = (y + 0.5 - cy) / r
            q = nx * nx + ny * ny
            If q <= 1 Then
                dif = -0.45 * nx - 0.55 * ny + 0.7 * Sqr(1 - q)
                If dif < 0 Then dif = 0
                spec = (dif ^ 20) * 140
                f = 0.25 + 0.85 * dif
                SPut s, x, y, RGB32(cr * f + spec, cg * f + spec, cb * f + spec)
            End If
        Next
    Next
End Sub

' Vertical cylinder from x0 to x1, rows y0..y1, shaded across its width.
Private Sub SCylinder(ByVal s As Long, ByVal x0 As Long, ByVal x1 As Long, ByVal y0 As Long, ByVal y1 As Long, ByVal cr As Double, ByVal cg As Double, ByVal cb As Double)
    Dim x As Long, y As Long, u As Double, f As Double, half As Double, mid As Double
    half = (x1 - x0 + 1) / 2: mid = x0 + half
    For x = x0 To x1
        u = (x + 0.5 - mid) / half
        If u > 0.999 Then u = 0.999
        If u < -0.999 Then u = -0.999
        f = 0.3 + 0.8 * Sqr(1 - u * u) * (1 - 0.25 * u)
        For y = y0 To y1
            SPut s, x, y, RGB32(cr * f, cg * f, cb * f)
        Next
    Next
End Sub

Private Sub Bounds(ByRef pix() As Long, ByRef topArr() As Long, ByRef botArr() As Long, ByVal count As Long)
    Dim s As Long, x As Long, y As Long, t As Long, b As Long
    For s = 0 To count - 1
        For x = 0 To TEX_SIZE - 1
            t = TEX_SIZE: b = -1
            For y = 0 To TEX_SIZE - 1
                If pix(s * TEX_PIXELS + x * TEX_SIZE + y) <> CLEAR_PIXEL Then
                    If t = TEX_SIZE Then t = y
                    b = y
                End If
            Next
            topArr(s * TEX_SIZE + x) = t
            botArr(s * TEX_SIZE + x) = b
        Next
    Next
End Sub

'---------------------------------------------------------------- wall textures

Private Sub MakeMissingTex(ByVal t As Long)
    Dim x As Long, y As Long
    For x = 0 To 63
        For y = 0 To 63
            If ((x \ 8) + (y \ 8)) Mod 2 = 0 Then TPut t, x, y, RGB32(200, 0, 200) Else TPut t, x, y, RGB32(30, 0, 30)
        Next
    Next
End Sub

Private Sub MakeStone(ByVal t As Long)
    Dim x As Long, y As Long, xo As Long, bRow As Long, bCol As Long, lx As Long, ly As Long
    Dim tone(0 To 7) As Double, v As Double
    For x = 0 To 7
        tone(x) = 92 + Frand() * 46
    Next
    For y = 0 To 63
        bRow = y \ 16
        For x = 0 To 63
            xo = x
            If (bRow Mod 2) = 1 Then xo = (x + 16) Mod 64
            bCol = xo \ 32
            lx = xo Mod 32: ly = y Mod 16
            v = tone(bRow * 2 + bCol) + (Frand() - 0.5) * 24
            If lx = 0 Or ly = 0 Then
                v = 46 + Frand() * 10
            ElseIf lx = 1 Or ly = 1 Then
                v = v + 22
            ElseIf lx = 31 Or ly = 15 Then
                v = v - 26
            End If
            TPut t, x, y, RGB32(v, v * 0.98, v * 1.06)
        Next
    Next
End Sub

Private Sub AddVent(ByVal t As Long)
    Dim x As Long, y As Long
    TRect t, 18, 18, 45, 45, RGB32(40, 40, 44)
    TRect t, 20, 20, 43, 43, RGB32(12, 12, 14)
    For x = 22 To 42 Step 4
        For y = 20 To 43
            TPut t, x, y, RGB32(90, 92, 98)
            TPut t, x + 1, y, RGB32(60, 62, 66)
        Next
    Next
    For x = 20 To 43
        TPut t, x, 31, RGB32(80, 82, 88)
    Next
End Sub

Private Sub MakeBrick(ByVal t As Long)
    Dim x As Long, y As Long, xo As Long, bRow As Long, bCol As Long, lx As Long, ly As Long
    Dim tone(0 To 31) As Double, f As Double, n As Double
    For x = 0 To 31
        tone(x) = 0.8 + Frand() * 0.35
    Next
    For y = 0 To 63
        bRow = y \ 8
        For x = 0 To 63
            xo = x
            If (bRow Mod 2) = 1 Then xo = (x + 8) Mod 64
            bCol = xo \ 16
            lx = xo Mod 16: ly = y Mod 8
            n = (Frand() - 0.5) * 0.14
            If lx = 0 Or ly = 0 Then
                f = 0.95 + n
                TPut t, x, y, RGB32(150 * f, 142 * f, 130 * f)
            Else
                f = tone(bRow * 4 + bCol) + n
                If ly = 1 Then f = f * 1.15
                If ly = 7 Or lx = 15 Then f = f * 0.78
                TPut t, x, y, RGB32(152 * f, 64 * f, 48 * f)
            End If
        Next
    Next
End Sub

' A blank brass plate with corner rivets.
Private Sub AddPlaque(ByVal t As Long)
    Dim x As Long, y As Long, f As Double
    For x = 16 To 47
        For y = 22 To 41
            f = 1 - (y - 22) / 60
            If x = 16 Or y = 22 Then f = f * 1.25
            If x = 47 Or y = 41 Then f = f * 0.6
            TPut t, x, y, RGB32(205 * f, 160 * f, 70 * f)
        Next
    Next
    TRect t, 20, 26, 43, 37, RGB32(170, 128, 52)
    TPut t, 18, 24, RGB32(90, 70, 30): TPut t, 45, 24, RGB32(90, 70, 30)
    TPut t, 18, 39, RGB32(90, 70, 30): TPut t, 45, 39, RGB32(90, 70, 30)
End Sub

Private Sub MakeWood(ByVal t As Long)
    Dim x As Long, y As Long, plank As Long, lx As Long, tone(0 To 3) As Double, f As Double
    For x = 0 To 3
        tone(x) = 0.85 + Frand() * 0.25
    Next
    For x = 0 To 63
        plank = x \ 16: lx = x Mod 16
        For y = 0 To 63
            f = tone(plank) * (1 + 0.12 * Sin(y * 0.31 + plank * 1.7 + Sin(x * 0.7 + plank) * 1.3) + (Frand() - 0.5) * 0.08)
            If lx = 0 Then f = 0.3
            If lx = 1 Then f = f * 0.8
            If lx = 15 Then f = f * 0.85
            If y >= 30 And y <= 33 Then f = f * 0.7
            If y = 30 Then f = f * 1.3
            TPut t, x, y, RGB32(130 * f, 84 * f, 48 * f)
        Next
    Next
End Sub

' A gilt frame around a simple landscape: sky, sun, hills.
Private Sub AddPainting(ByVal t As Long)
    Dim x As Long, y As Long, h1 As Double, h2 As Double, c As Long
    TRect t, 12, 12, 51, 45, RGB32(200, 160, 60)
    TRect t, 13, 13, 50, 44, RGB32(150, 110, 40)
    For x = 15 To 48
        h1 = 34 - 5 * Sin((x - 15) * 0.18)
        h2 = 38 - 3 * Sin((x - 10) * 0.3 + 1)
        For y = 15 To 42
            If y > h2 Then
                c = RGB32(60, 110, 50)
            ElseIf y > h1 Then
                c = RGB32(90, 120, 80)
            Else
                c = RGB32(110 + (y - 15) * 3, 150 + (y - 15) * 2, 200)
            End If
            TPut t, x, y, c
        Next
    Next
    For x = 38 To 44
        For y = 18 To 24
            If (x - 41) ^ 2 + (y - 21) ^ 2 <= 9 Then TPut t, x, y, RGB32(250, 220, 120)
        Next
    Next
End Sub

Private Sub MakeTiles(ByVal t As Long)
    Dim x As Long, y As Long, lx As Long, ly As Long, f As Double, tone(0 To 15) As Double
    For x = 0 To 15
        tone(x) = 0.88 + Frand() * 0.18
    Next
    For x = 0 To 63
        For y = 0 To 63
            lx = x Mod 16: ly = y Mod 16
            If lx = 0 Or ly = 0 Then
                TPut t, x, y, RGB32(175, 180, 185)
            Else
                f = tone((y \ 16) * 4 + (x \ 16)) + (Frand() - 0.5) * 0.04
                If lx = 1 Or ly = 1 Then f = f * 1.2
                If lx = 15 Or ly = 15 Then f = f * 0.8
                TPut t, x, y, RGB32(60 * f, 100 * f, 160 * f)
            End If
        Next
    Next
End Sub

Private Sub MakeMetal(ByVal t As Long)
    Dim x As Long, y As Long, lx As Long, ly As Long, rx As Long, ry As Long, f As Double
    For x = 0 To 63
        For y = 0 To 63
            lx = x Mod 32: ly = y Mod 32
            f = 0.95 + (Frand() - 0.5) * 0.06 + (lx + ly) / 400
            If lx = 0 Or ly = 0 Then
                f = 0.45
            ElseIf lx = 1 Or ly = 1 Then
                f = f * 1.3
            ElseIf lx = 31 Or ly = 31 Then
                f = f * 0.7
            End If
            If lx < 16 Then rx = 5 Else rx = 26
            If ly < 16 Then ry = 5 Else ry = 26
            If (lx - rx) * (lx - rx) + (ly - ry) * (ly - ry) <= 2 Then f = 1.6
            TPut t, x, y, RGB32(92 * f, 104 * f, 122 * f)
        Next
    Next
End Sub

Private Sub MakeRails(ByVal t As Long)
    Dim x As Long, y As Long, f As Double
    For x = 0 To 63
        f = 0.75 + 0.25 * Sin(x * PI / 8)
        If (x Mod 16) = 0 Then f = 0.4
        For y = 0 To 63
            TPut t, x, y, RGB32(150 * f, 150 * f, 160 * f)
        Next
    Next
End Sub

Private Sub MakeSwitch(ByVal t As Long, ByVal down As Boolean)
    Dim y As Long
    MakeRails t
    TRect t, 20, 14, 43, 49, RGB32(50, 52, 58)
    TRect t, 22, 16, 41, 47, RGB32(90, 92, 100)
    TRect t, 30, 22, 33, 41, RGB32(25, 25, 28)        ' lever slot
    If down Then
        TRect t, 28, 34, 35, 41, RGB32(200, 200, 205)
        TRect t, 26, 50, 37, 53, RGB32(220, 40, 40)   ' red: used
    Else
        TRect t, 28, 22, 35, 29, RGB32(200, 200, 205)
        TRect t, 26, 50, 37, 53, RGB32(40, 220, 60)   ' green: ready
    End If
    For y = 16 To 47
        TPut t, 22, y, RGB32(130, 132, 140)
    Next
End Sub

' kind 0 = plain, 1 = gold lock, 2 = silver lock
Private Sub MakeDoor(ByVal t As Long, ByVal kind As Long)
    Dim x As Long, y As Long, f As Double, c As Long, lockColor As Long
    For x = 0 To 63
        For y = 0 To 63
            f = 0.92 + (Frand() - 0.5) * 0.06 - y / 500
            If (y Mod 8) = 0 Then f = f * 1.25
            If (y Mod 8) = 7 Then f = f * 0.72
            c = RGB32(78 * f, 96 * f, 112 * f)
            If x >= 20 And x <= 43 And y >= 10 And y <= 21 Then
                If x = 20 Or x = 43 Or y = 10 Or y = 21 Then c = RGB32(130, 140, 150) Else c = RGB32(16, 22, 28)
            End If
            If x >= 4 And x <= 8 And y >= 28 And y <= 39 Then c = RGB32(200, 175, 60)
            If x = 0 Or x = 63 Then c = RGB32(30, 34, 40)
            TPut t, x, y, c
        Next
    Next
    If kind > 0 Then
        If kind = 1 Then lockColor = RGB32(225, 180, 50) Else lockColor = RGB32(200, 205, 215)
        TRect t, 24, 30, 39, 47, Shade(lockColor, 0.6)
        TRect t, 25, 31, 38, 46, lockColor
        TRect t, 30, 35, 33, 38, RGB32(20, 20, 20)    ' keyhole
        TRect t, 31, 39, 32, 42, RGB32(20, 20, 20)
    End If
End Sub

' The wall faces beside a door: a recessed steel channel.
Private Sub MakePlate(ByVal t As Long)
    Dim x As Long, y As Long, f As Double
    For x = 0 To 63
        For y = 0 To 63
            f = 0.9 + (Frand() - 0.5) * 0.1
            If (x Mod 16) = 0 Then f = 0.55
            If (x Mod 16) = 1 Then f = f * 1.3
            TPut t, x, y, RGB32(60 * f, 62 * f, 70 * f)
        Next
    Next
End Sub

Private Sub MakeElevatorDoor(ByVal t As Long)
    Dim x As Long, y As Long, f As Double
    For x = 0 To 63
        For y = 0 To 63
            f = 0.85 + 0.1 * Sin(y * 0.8) + (Frand() - 0.5) * 0.05
            If (y Mod 12) = 0 Then f = 0.55
            If x = 31 Or x = 32 Then f = 0.4                ' centre seam
            TPut t, x, y, RGB32(170 * f, 172 * f, 180 * f)
        Next
    Next
End Sub

'---------------------------------------------------------------- object sprites
' Sprites span floor to ceiling: row 0 touches the ceiling, row 63 the floor.

Private Function SpriteIsEmpty(ByVal s As Long) As Boolean
    Dim i As Long
    For i = 0 To TEX_PIXELS - 1
        If SprPix(s * TEX_PIXELS + i) <> CLEAR_PIXEL Then Exit Function
    Next
    SpriteIsEmpty = True
End Function

Private Sub MakeMissingSprite(ByVal s As Long)
    SRect s, 20, 40, 43, 63, RGB32(200, 0, 200)
    SRect s, 23, 43, 40, 60, RGB32(30, 0, 30)
End Sub

Private Sub MakePuddle(ByVal s As Long)
    SEllipse s, 32, 59, 18, 4, RGB32(60, 90, 140)
    SEllipse s, 30, 58, 10, 2, RGB32(110, 150, 200)
End Sub

Private Sub MakeDrum(ByVal s As Long, ByVal r As Double, ByVal g As Double, ByVal b As Double)
    Dim x As Long, y As Long, c As Long
    SCylinder s, 20, 43, 30, 63, r, g, b
    For x = 20 To 43
        For y = 30 To 63
            c = SGet(s, x, y)
            If y = 34 Or y = 35 Or y = 48 Or y = 49 Or y = 60 Or y = 61 Then SPut s, x, y, Shade(c, 0.6)
            If y <= 31 Then SPut s, x, y, Shade(c, 1.3)
        Next
    Next
End Sub

Private Sub MakeTable(ByVal s As Long)
    SRect s, 10, 40, 53, 43, RGB32(150, 100, 60)
    SRect s, 10, 44, 53, 45, RGB32(100, 66, 40)
    SRect s, 13, 46, 16, 63, RGB32(110, 72, 44)
    SRect s, 47, 46, 50, 63, RGB32(110, 72, 44)
    SRect s, 22, 46, 24, 58, RGB32(90, 60, 36)
    SRect s, 39, 46, 41, 58, RGB32(90, 60, 36)
End Sub

Private Sub MakeFloorLamp(ByVal s As Long)
    SRect s, 31, 22, 32, 61, RGB32(70, 70, 76)
    SEllipse s, 32, 62, 8, 2, RGB32(60, 60, 66)
    Dim y As Long, hw As Long
    For y = 8 To 21
        hw = 5 + (y - 8) \ 2
        SRect s, 32 - hw, y, 31 + hw, y, RGB32(230 - (y - 8) * 4, 205 - (y - 8) * 4, 140)
    Next
    SRect s, 26, 21, 37, 21, RGB32(255, 240, 180)
End Sub

Private Sub MakeChandelier(ByVal s As Long)
    Dim i As Long, x As Long
    SRect s, 31, 0, 32, 8, RGB32(80, 70, 50)
    SEllipse s, 32, 12, 14, 3, RGB32(170, 135, 60)
    SEllipse s, 32, 12, 11, 1.5, RGB32(110, 85, 40)
    For i = 0 To 4
        x = 20 + i * 6
        SRect s, x, 6, x + 1, 10, RGB32(235, 230, 210)
        SDisc s, x + 1, 4.5, 1.6, RGB32(255, 210, 90)
    Next
End Sub

Private Sub MakePillar(ByVal s As Long)
    SCylinder s, 24, 39, 6, 57, 150, 146, 138
    SCylinder s, 20, 43, 0, 5, 165, 160, 150
    SCylinder s, 20, 43, 58, 63, 165, 160, 150
End Sub

Private Sub MakePottedTree(ByVal s As Long)
    SCylinder s, 24, 39, 50, 63, 160, 90, 60
    SRect s, 31, 22, 32, 50, RGB32(90, 60, 30)
    SSphere s, 32, 18, 14, 50, 120, 50
    SSphere s, 24, 26, 8, 45, 110, 45
    SSphere s, 40, 25, 8, 55, 125, 55
End Sub

Private Sub MakePlant(ByVal s As Long)
    Dim i As Long, x As Long, y As Long
    SCylinder s, 26, 37, 52, 63, 170, 100, 70
    For i = 0 To 6
        For y = 30 To 52
            x = 32 + (i - 3) * (52 - y) \ 4
            SPut s, x, y, RGB32(40, 120 + i * 10, 50)
            SPut s, x + 1, y, RGB32(30, 100 + i * 10, 40)
        Next
    Next
End Sub

Private Sub MakeUrn(ByVal s As Long)
    Dim y As Long, hw As Double
    For y = 34 To 63
        hw = 7 + 6 * Sin((y - 34) / 29 * PI)
        SCylinder s, Int(32 - hw), Int(31 + hw), y, y, 175, 110, 70
    Next
    SRect s, 25, 32, 38, 34, RGB32(140, 85, 55)
    SRect s, 22, 46, 41, 47, RGB32(90, 50, 30)
End Sub

Private Sub MakeCeilingLight(ByVal s As Long)
    SRect s, 31, 0, 32, 3, RGB32(60, 60, 64)
    SEllipse s, 32, 6, 9, 3, RGB32(90, 90, 96)
    SEllipse s, 32, 8, 6, 2, RGB32(255, 245, 200)
End Sub

Private Sub MakeKey(ByVal s As Long, ByVal r As Double, ByVal g As Double, ByVal b As Double)
    SDisc s, 25, 55, 5, RGB32(r, g, b)
    SDisc s, 25, 55, 2, CLEAR_PIXEL
    SRect s, 29, 54, 42, 56, RGB32(r * 0.9, g * 0.9, b * 0.9)
    SRect s, 38, 57, 39, 60, RGB32(r * 0.8, g * 0.8, b * 0.8)
    SRect s, 41, 57, 42, 59, RGB32(r * 0.8, g * 0.8, b * 0.8)
End Sub

Private Sub MakeFood(ByVal s As Long)
    SEllipse s, 32, 60, 15, 3, RGB32(230, 230, 225)
    SEllipse s, 27, 56, 6, 4, RGB32(200, 140, 70)
    SEllipse s, 37, 57, 5, 3, RGB32(150, 60, 50)
    SDisc s, 33, 54, 2.5, RGB32(90, 160, 60)
End Sub

Private Sub MakeFirstAid(ByVal s As Long)
    SRect s, 20, 48, 43, 63, RGB32(40, 150, 70)
    SRect s, 20, 48, 43, 50, RGB32(55, 185, 90)
    SRect s, 30, 51, 33, 60, RGB32(245, 245, 245)
    SRect s, 26, 54, 37, 57, RGB32(245, 245, 245)
End Sub

Private Sub MakeClip(ByVal s As Long)
    SRect s, 27, 46, 36, 63, RGB32(60, 62, 66)
    SRect s, 28, 47, 35, 62, RGB32(95, 98, 104)
    SRect s, 29, 43, 34, 46, RGB32(200, 160, 60)
End Sub

Private Sub MakeGunPickup(ByVal s As Long, ByVal rotary As Boolean)
    If rotary Then
        SRect s, 12, 50, 52, 57, RGB32(70, 72, 78)
        SRect s, 8, 51, 12, 52, RGB32(40, 40, 44)
        SRect s, 8, 54, 12, 55, RGB32(40, 40, 44)
        SRect s, 36, 57, 41, 62, RGB32(55, 40, 30)
    Else
        SRect s, 14, 52, 50, 56, RGB32(65, 66, 72)
        SRect s, 8, 53, 14, 54, RGB32(40, 40, 44)
        SRect s, 30, 56, 34, 63, RGB32(50, 50, 54)
        SRect s, 40, 56, 46, 61, RGB32(90, 60, 40)
    End If
End Sub

Private Sub MakeCoins(ByVal s As Long)
    Dim i As Long
    For i = 0 To 5
        SEllipse s, 32, 61 - i * 2.2, 8, 2, RGB32(230 - i * 5, 185, 60)
        SEllipse s, 32, 60.5 - i * 2.2, 7, 1.2, RGB32(250, 215, 100)
    Next
End Sub

Private Sub MakeGoblet(ByVal s As Long)
    Dim y As Long, hw As Double
    For y = 40 To 50
        hw = 7 - (y - 40) * 0.45
        SCylinder s, Int(32 - hw), Int(31 + hw), y, y, 225, 185, 70
    Next
    SRect s, 31, 51, 32, 59, RGB32(200, 160, 60)
    SEllipse s, 32, 61, 6, 2, RGB32(210, 170, 65)
    SDisc s, 32, 45, 1.5, RGB32(200, 30, 60)
End Sub

Private Sub MakeBook(ByVal s As Long)
    SRect s, 20, 52, 43, 62, RGB32(110, 50, 35)
    SRect s, 21, 53, 42, 55, RGB32(235, 225, 200)
    SRect s, 20, 52, 22, 54, RGB32(220, 180, 70)
    SRect s, 41, 52, 43, 54, RGB32(220, 180, 70)
    SRect s, 20, 60, 22, 62, RGB32(220, 180, 70)
    SRect s, 41, 60, 43, 62, RGB32(220, 180, 70)
End Sub

Private Sub MakeCrown(ByVal s As Long)
    Dim i As Long
    SRect s, 20, 52, 43, 61, RGB32(225, 180, 50)
    For i = 0 To 4
        SRect s, 20 + i * 5, 45, 23 + i * 5, 52, RGB32(225, 180, 50)
        SDisc s, 21.5 + i * 5, 44, 1.7, RGB32(250, 230, 140)
    Next
    SDisc s, 26, 56, 1.8, RGB32(40, 90, 220)
    SDisc s, 32, 56, 1.8, RGB32(210, 30, 50)
    SDisc s, 38, 56, 1.8, RGB32(40, 180, 90)
End Sub

Private Sub MakeOrb(ByVal s As Long)
    SDisc s, 32, 50, 11, RGB32(40, 80, 160)
    SSphere s, 32, 50, 9, 90, 170, 255
End Sub

Private Sub MakeAmmoBox(ByVal s As Long)
    SRect s, 18, 50, 45, 63, RGB32(95, 105, 60)
    SRect s, 18, 50, 45, 52, RGB32(120, 132, 78)
    SRect s, 18, 56, 45, 57, RGB32(220, 190, 70)
End Sub

'---------------------------------------------------------------- first-person weapon views
' Drawn at double size from the bottom of the view; everything sits in rows
' 26..63 so the weapon stays below the centre of the screen. Each weapon has
' four frames: 0 idle, 1 raised, 2 firing (or stabbing), 3 recoil.

Private Sub MakeWeaponView(ByVal s As Long, ByVal weapon As Long, ByVal frame As Long)
    Dim dx As Long, dy As Long, mx As Double, my As Double
    Select Case weapon
        Case 0: MakeKnifeView s
        Case 1: MakePistolView s: mx = 31.5: my = 32
        Case 2: MakeSmgView s: mx = 31.5: my = 27
        Case Else: MakeRotaryView s: mx = 31: my = 24
    End Select
    If weapon = 0 Then
        Select Case frame
            Case 1: dx = 4: dy = 3
            Case 2: dx = -7: dy = -9
            Case 3: dx = 2: dy = 1
        End Select
    Else
        Select Case frame
            Case 1: dy = -3
            Case 2: dy = -5
            Case 3: dy = 1
        End Select
    End If
    ShiftView s, dx, dy
    If weapon > 0 And frame = 2 Then
        SDisc s, mx + dx, my + dy - 3, 6, RGB32(255, 200, 80)
        SDisc s, mx + dx, my + dy - 3, 3, RGB32(255, 250, 220)
    End If
End Sub

Private Sub ShiftView(ByVal s As Long, ByVal dx As Long, ByVal dy As Long)
    Dim tmp(0 To 4095) As Long, x As Long, y As Long
    If dx = 0 And dy = 0 Then Exit Sub
    For x = 0 To 4095
        tmp(x) = WeapPix(s * TEX_PIXELS + x)
        WeapPix(s * TEX_PIXELS + x) = CLEAR_PIXEL
    Next
    For x = 0 To 63
        For y = 0 To 63
            If tmp(x * 64 + y) <> CLEAR_PIXEL Then SPut s, x + dx, y + dy, tmp(x * 64 + y)
        Next
    Next
End Sub

Private Sub MakeKnifeView(ByVal s As Long)
    Dim y As Long, hw As Double
    For y = 28 To 50
        hw = 0.5 + (y - 28) * 0.09
        SRect s, Int(40 - hw), y, Int(40 + hw), y, RGB32(205, 210, 220)
        SPut s, Int(40 - hw), y, RGB32(245, 248, 252)
    Next
    SRect s, 34, 51, 46, 53, RGB32(90, 92, 98)
    SRect s, 37, 54, 43, 63, RGB32(70, 45, 30)
    SRect s, 38, 54, 38, 63, RGB32(100, 70, 48)
End Sub

Private Sub MakePistolView(ByVal s As Long)
    SRect s, 26, 52, 37, 63, RGB32(70, 50, 38)          ' grip
    SRect s, 27, 53, 28, 63, RGB32(95, 70, 52)
    SRect s, 25, 46, 38, 52, RGB32(52, 54, 60)          ' frame
    SRect s, 23, 36, 40, 46, RGB32(64, 66, 74)          ' slide, seen from behind
    SRect s, 23, 36, 40, 37, RGB32(110, 114, 124)
    SRect s, 23, 38, 24, 46, RGB32(90, 94, 102)
    SRect s, 30, 33, 33, 35, RGB32(40, 40, 44)          ' front sight
    SRect s, 24, 34, 27, 35, RGB32(40, 40, 44)          ' rear sight
    SRect s, 36, 34, 39, 35, RGB32(40, 40, 44)
End Sub

Private Sub MakeSmgView(ByVal s As Long)
    SRect s, 24, 38, 39, 52, RGB32(56, 58, 64)          ' receiver
    SRect s, 24, 38, 39, 39, RGB32(100, 104, 112)
    SRect s, 29, 30, 34, 38, RGB32(44, 45, 50)          ' barrel shroud
    SRect s, 30, 28, 33, 30, RGB32(30, 30, 33)
    SRect s, 18, 44, 23, 60, RGB32(45, 46, 50)          ' magazine
    SRect s, 27, 52, 36, 63, RGB32(72, 52, 40)          ' grip
    SRect s, 28, 53, 29, 63, RGB32(96, 72, 54)
End Sub

Private Sub MakeRotaryView(ByVal s As Long)
    Dim i As Long
    For i = 0 To 3
        SRect s, 24 + i * 4, 26, 25 + i * 4, 44, RGB32(80 + i * 12, 82 + i * 12, 90 + i * 12)
    Next
    SRect s, 22, 44, 41, 50, RGB32(60, 62, 68)
    SRect s, 20, 50, 43, 63, RGB32(52, 54, 60)
    SRect s, 21, 50, 23, 63, RGB32(85, 88, 96)
End Sub

'---------------------------------------------------------------- robot enemies
' Original designs for the generic enemy slots. Rotation r follows the engine:
' 0 faces the viewer, 2 faces screen-left, 4 faces away, 6 faces screen-right.

Private Sub SLine(ByVal s As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal w As Double, ByVal c As Long)
    Dim n As Long, i As Long, t As Double
    n = Int(Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)) + 1
    For i = 0 To n
        t = i / n
        SDisc s, x0 + (x1 - x0) * t, y0 + (y1 - y0) * t, w / 2, c
    Next
End Sub

' A box lit from the left.
Private Sub SBox(ByVal s As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal c As Long)
    Dim x As Long, y As Long, w As Double
    w = x1 - x0
    If w < 1 Then w = 1
    For x = Int(x0) To Int(x1)
        For y = Int(y0) To Int(y1)
            SPut s, x, y, Shade(c, 1.15 - 0.4 * (x - x0) / w)
        Next
    Next
End Sub

' Bleaches every opaque pixel toward white-red and adds sparks: the hit flash.
Private Sub HitFlash(ByVal s As Long)
    Dim i As Long, c As Long
    For i = 0 To TEX_PIXELS - 1
        c = SprPix(s * TEX_PIXELS + i)
        If c <> CLEAR_PIXEL Then
            SprPix(s * TEX_PIXELS + i) = RGB32(((c \ 65536) And 255) * 0.5 + 128, ((c \ 256) And 255) * 0.5 + 100, (c And 255) * 0.5 + 100)
        End If
    Next
    For i = 0 To 6
        SPut s, 16 + Int(Frand() * 32), 10 + Int(Frand() * 40), RGB32(255, 230, 120)
    Next
End Sub

' Walks an actor type's state table and draws every sprite slot it uses.
Private Sub MakeActorArt(ByVal t As Long)
    Dim st As Long, tex As Long, r As Long, walkBase As Long, dieCount As Long, shootCount As Long
    walkBase = StTexture(t, st_chase1)
    For st = st_die1 To st_die9
        If StTexture(t, st) > 0 Then dieCount = dieCount + 1
    Next
    For st = st_shoot1 To st_shoot9
        If StTexture(t, st) > 0 Then shootCount = shootCount + 1
    Next
    For st = 0 To NUMSTATES - 1
        tex = StTexture(t, st)
        If tex > 0 Then
            If StRotate(t, st) Then
                For r = 0 To 7
                    If st = st_stand Then
                        DrawActor tex + r, t, r, 0, 0, 1
                    Else
                        DrawActor tex + r, t, r, 1, ((tex - walkBase) \ 8) And 3, 4
                    End If
                Next
            ElseIf st = st_pain Or st = st_pain1 Then
                If DrawActor(tex, t, 0, 0, 0, 1) Then HitFlash tex
            ElseIf st >= st_shoot1 And st <= st_shoot9 Then
                DrawActor tex, t, 0, 3, st - st_shoot1, shootCount
            ElseIf st >= st_die1 And st <= st_die9 Then
                DrawActor tex, t, 0, 4, st - st_die1, dieCount
            ElseIf st = st_dead Then
                DrawActor tex, t, 0, 5, 0, 1
            End If
        End If
    Next
End Sub

' pose: 0 stand, 1 walk, 3 shoot (a lunge for the hound), 4 die, 5 dead.
' Returns False if another state already drew this slot (the first one wins).
Private Function DrawActor(ByVal s As Long, ByVal t As Long, ByVal r As Long, ByVal pose As Long, ByVal phase As Long, ByVal phases As Long) As Boolean
    Dim i As Long
    If Drawn(s) Then Exit Function
    Drawn(s) = True
    For i = 0 To TEX_PIXELS - 1
        SprPix(s * TEX_PIXELS + i) = CLEAR_PIXEL
    Next
    Select Case t
        Case en_dog
            DrawHound s, r, pose, phase, phases
        Case en_officer     ' scout: slim, white, blue visor
            DrawBot s, RGB32(205, 205, 195), RGB32(90, 95, 105), RGB32(60, 170, 255), RGB32(60, 120, 200), 0.85, False, r, pose, phase, phases
        Case en_ss          ' enforcer: heavy, dark blue, yellow visor
            DrawBot s, RGB32(72, 82, 112), RGB32(35, 38, 52), RGB32(255, 200, 40), RGB32(230, 190, 40), 1.25, False, r, pose, phase, phases
        Case en_mutant      ' brute: wide, green, two eyes
            DrawBot s, RGB32(96, 122, 82), RGB32(48, 60, 40), RGB32(255, 130, 40), RGB32(150, 90, 50), 1.3, True, r, pose, phase, phases
        Case Else           ' sentry: olive grey, red visor
            DrawBot s, RGB32(125, 130, 112), RGB32(60, 62, 58), RGB32(230, 40, 40), RGB32(200, 170, 60), 1, False, r, pose, phase, phases
    End Select
    DrawActor = True
End Function

' Two-legged robot. twoEyes draws a pair of eyes instead of a visor bar.
Private Sub DrawBot(ByVal s As Long, ByVal body As Long, ByVal dark As Long, ByVal eye As Long, ByVal accent As Long, ByVal bulk As Double, ByVal twoEyes As Boolean, ByVal r As Long, ByVal pose As Long, ByVal phase As Long, ByVal phases As Long)
    Dim f As Double, sx As Double, rightX As Double, wf As Double, hw As Double, hh As Double
    Dim cx As Double, hx As Double, spread As Double, sw As Double, L As Long
    Dim hipX As Double, footX As Double, footY As Double, lift As Double, shX As Double
    Dim vw As Double, vx As Double, c As Double, drop As Double, i As Long, fPos As Double

    cx = 32
    f = Cos(r * PI / 4)                     ' 1 = facing the viewer, -1 = facing away
    sx = -Sin(r * PI / 4)                   ' -1 = facing screen-left, 1 = screen-right
    If pose = 3 Then f = 1: sx = 0          ' shooting frames always face the viewer
    rightX = -f                             ' screen direction of the robot's right side
    wf = 0.55 + 0.45 * Abs(f)
    hw = 9 * bulk * wf
    hh = 6 * Sqr(bulk) * (0.7 + 0.3 * Abs(f))

    If pose = 5 Then                        ' wreck on the floor
        SBox s, 16, 53, 44, 61, Shade(body, 0.7)
        SBox s, 40, 49, 51, 58, Shade(body, 0.8)
        SLine s, 12, 62, 22, 58, 4, dark
        SLine s, 46, 63, 56, 60, 4, dark
        SDisc s, 45, 53, 1.5, Shade(eye, 0.4)
        For i = 0 To 30
            SPut s, 16 + Int(Frand() * 30), 52 + Int(Frand() * 10), RGB32(28, 26, 24)
        Next
        Exit Sub
    End If

    If pose = 4 Then                        ' collapsing: the body drops, the head comes loose
        c = (phase + 1) / (phases + 1)
        drop = c * 16
        SLine s, cx - 5, 45 + drop, cx - 9 - c * 6, 63, 4 * bulk, dark
        SLine s, cx + 5, 45 + drop, cx + 8 + c * 6, 63, 4 * bulk, dark
        SBox s, cx - hw, 25 + drop, cx + hw, 46 + drop * 0.7, body
        SBox s, cx - hh + c * 12, 12 + c * 34, cx + hh + c * 12, 23 + c * 34, body
        For i = 0 To 10
            SPut s, 18 + Int(Frand() * 28), 18 + Int(Frand() * 30 + drop), RGB32(255, 200 - i * 10, 80)
        Next
        If phase > 0 Then
            SDisc s, cx - 6, 18 + drop, 4, RGB32(90, 90, 95)
            SDisc s, cx - 2, 13 + drop, 3, RGB32(110, 110, 115)
        End If
        Exit Sub
    End If

    ' legs
    spread = 4.5 * bulk * wf
    If pose = 1 Then sw = Sin(phase * PI / 2)
    For L = -1 To 1 Step 2
        hipX = cx + L * spread
        footX = hipX + L * sw * 7 * sx
        lift = L * sw * 3 * Abs(f)
        If lift < 0 Then lift = 0
        footY = 61 - lift
        SLine s, hipX, 45, footX, footY, 4 * bulk, dark
        SRect s, Int(footX - 3), Int(footY), Int(footX + 3 + 2 * sx), Int(footY + 2), Shade(dark, 0.8)
    Next

    ' the arm on the far side, and the gun arm when it is behind the body
    shX = cx - rightX * (hw + 1)
    SLine s, shX, 28, shX - sx * 2, 41, 3.5 * bulk, dark
    If f < -0.3 Then SLine s, cx + rightX * (hw + 1), 28, cx + rightX * (hw + 1) + sx * 5, 40, 3.5 * bulk, dark

    ' torso: chest light from the front, backpack from behind, stripe in profile
    SBox s, cx - hw, 25, cx + hw, 46, body
    SRect s, Int(cx - hw + 1), 44, Int(cx + hw - 1), 46, dark
    If f > 0.3 Then
        SRect s, Int(cx - 3), 30, Int(cx + 3), 35, dark
        SRect s, Int(cx - 2), 31, Int(cx + 2), 34, accent
    ElseIf f < -0.3 Then
        SBox s, cx - hw * 0.7, 27, cx + hw * 0.7, 42, Shade(body, 0.75)
        For i = 29 To 40 Step 3
            SRect s, Int(cx - hw * 0.6), i, Int(cx + hw * 0.6), i, dark
        Next
    Else
        SRect s, Int(cx - 1), 27, Int(cx + 1), 44, accent
    End If

    ' head
    hx = cx + sx * 2.5
    SBox s, hx - hh, 12, hx + hh, 23, body
    SRect s, Int(hx - hh), 22, Int(hx + hh), 23, dark
    If Not twoEyes Then
        SLine s, hx - sx * 2, 11, hx - sx * 2, 6, 1.5, dark
        SPut s, Int(hx - sx * 2), 5, accent
    End If
    If f > -0.3 Then
        vx = hx + sx * hh * 0.5
        If twoEyes Then
            SDisc s, vx - 3 * (0.4 + 0.6 * f), 17, 1.8, eye
            SDisc s, vx + 3 * (0.4 + 0.6 * f), 17, 1.8, eye
        Else
            fPos = f
            If fPos < 0 Then fPos = 0
            vw = hh * (0.35 + 0.65 * fPos) + 1.5
            SRect s, Int(vx - vw), 16, Int(vx + vw), 18, eye
            SRect s, Int(vx - vw + 1), 16, Int(vx - vw + 2), 16, RGB32(255, 220, 220)
        End If
    End If

    ' gun arm
    shX = cx + rightX * (hw + 1)
    If pose = 3 Then                        ' aiming at the viewer, flash on odd frames
        SLine s, shX, 28, cx - 4, 33, 4 * bulk, dark
        SDisc s, cx - 4, 33, 3.5, RGB32(30, 30, 34)
        SDisc s, cx - 4, 33, 1.5, RGB32(10, 10, 12)
        If (phase Mod 2) = 1 Then
            SDisc s, cx - 4, 33, 7, RGB32(255, 200, 80)
            SDisc s, cx - 4, 33, 3.5, RGB32(255, 250, 220)
        End If
    ElseIf f >= -0.3 Then
        SLine s, shX, 28, shX + sx * 5, 38, 3.5 * bulk, dark
        If Abs(sx) > 0.3 Then
            SLine s, shX + sx * 5, 38, shX + sx * 15, 37, 3, RGB32(40, 40, 45)
        Else
            SDisc s, shX, 38, 3, RGB32(40, 40, 45)
            SDisc s, shX, 38, 1.2, RGB32(10, 10, 12)
        End If
    End If
End Sub

' Four-legged robot hound (the dog slots).
Private Sub DrawHound(ByVal s As Long, ByVal r As Long, ByVal pose As Long, ByVal phase As Long, ByVal phases As Long)
    Dim body As Long, dark As Long, eye As Long, f As Double, sx As Double, side As Double
    Dim cx As Double, bodyY As Double, bodyLen As Double, sw As Double, k As Long, legX As Double
    Dim headX As Double, headY As Double, c As Double, i As Long, gap As Double

    body = RGB32(100, 100, 112): dark = RGB32(50, 50, 58): eye = RGB32(255, 60, 40)
    cx = 32
    f = Cos(r * PI / 4): sx = -Sin(r * PI / 4): side = Abs(sx)

    If pose = 5 Then                        ' wreck
        SBox s, 14, 55, 46, 62, Shade(body, 0.7)
        SBox s, 44, 52, 54, 60, Shade(body, 0.8)
        SLine s, 10, 63, 18, 58, 3, dark
        SDisc s, 49, 56, 1.3, Shade(eye, 0.4)
        Exit Sub
    End If
    If pose = 4 Then                        ' collapsing
        c = (phase + 1) / (phases + 1)
        SBox s, 14, 44 + c * 10, 48, 54 + c * 7, body
        SBox s, 44, 40 + c * 14, 54, 49 + c * 12, body
        For i = 0 To 8
            SPut s, 16 + Int(Frand() * 34), 38 + Int(Frand() * 20), RGB32(255, 190, 70)
        Next
        Exit Sub
    End If
    If pose = 3 Then                        ' lunge at the viewer: crouch, leap, land
        bodyY = 46 - Choose((phase Mod 3) + 1, 0, 8, 2)
        SLine s, 24, bodyY + 6, 20, 63, 3.5, dark
        SLine s, 40, bodyY + 6, 44, 63, 3.5, dark
        SBox s, 20, bodyY - 6, 44, bodyY + 7, body
        gap = Choose((phase Mod 3) + 1, 1, 4, 2)
        SBox s, 24, bodyY - 14, 40, bodyY - 6 - gap, body             ' upper jaw
        SBox s, 25, bodyY - 6 + gap, 39, bodyY + 1, Shade(body, 0.8)  ' lower jaw
        SRect s, 26, Int(bodyY - 6 - gap + 1), 38, Int(bodyY - 6 + gap - 1), RGB32(20, 20, 24)
        SDisc s, 28, bodyY - 11, 1.6, eye
        SDisc s, 36, bodyY - 11, 1.6, eye
        Exit Sub
    End If

    ' walking, any rotation
    bodyY = 47
    bodyLen = 6 + 11 * side
    If pose = 1 Then sw = Sin(phase * PI / 2)
    For k = -1 To 1 Step 2                  ' k = 1: front legs, -1: back legs
        legX = cx + k * sx * bodyLen * 0.75
        SLine s, legX - 3 * Abs(f), bodyY + 4, legX - 3 * Abs(f) + k * sw * 4 * sx, 63, 3, dark
        SLine s, legX + 3 * Abs(f), bodyY + 4, legX + 3 * Abs(f) - k * sw * 4 * sx, 63, 3, dark
    Next
    If f < -0.3 Then SLine s, cx, bodyY - 4, cx, bodyY - 13, 1.5, dark       ' antenna tail, seen from behind
    SBox s, cx - bodyLen, bodyY - 5, cx + bodyLen, bodyY + 5, body
    SRect s, Int(cx - bodyLen + 1), Int(bodyY - 1), Int(cx + bodyLen - 1), Int(bodyY), Shade(body, 0.7)
    headX = cx + sx * (bodyLen + 2)
    headY = bodyY - 7
    If f >= -0.3 Then
        SBox s, headX - 5 - side * 2, headY - 5, headX + 5 + side * 2, headY + 5, body
        SRect s, Int(headX - 4 + sx * 3), Int(headY + 2), Int(headX + 4 + sx * 3), Int(headY + 3), RGB32(25, 25, 30)
        If side > 0.7 Then
            SDisc s, headX + sx * 2, headY - 1, 1.5, eye
        Else
            SDisc s, headX - 2.5, headY - 1, 1.5, eye
            SDisc s, headX + 2.5, headY - 1, 1.5, eye
        End If
    End If
End Sub

'---------------------------------------------------------------- content pack hooks

Public Sub Art_SetTexture(ByVal slot As Long, ByRef pix() As Long)
    Dim i As Long
    For i = 0 To TEX_PIXELS - 1
        TexPix(slot * TEX_PIXELS + i) = pix(i)
    Next
End Sub

Public Sub Art_SetSprite(ByVal slot As Long, ByRef pix() As Long)
    Dim i As Long
    For i = 0 To TEX_PIXELS - 1
        SprPix(slot * TEX_PIXELS + i) = pix(i)
    Next
End Sub

Public Sub Art_SetWeaponView(ByVal slot As Long, ByRef pix() As Long)
    Dim i As Long
    For i = 0 To TEX_PIXELS - 1
        WeapPix(slot * TEX_PIXELS + i) = pix(i)
    Next
End Sub

Public Sub Art_DarkCopy(ByVal src As Long, ByVal dst As Long)
    DarkCopy src, dst
End Sub

' Recomputes the opaque row range of every sprite and weapon column.
Public Sub Art_RefreshBounds()
    Bounds SprPix, SprTop, SprBot, NUM_SPR_SLOTS
    Bounds WeapPix, WeapTop, WeapBot, NUM_WEAPON_VIEWS
End Sub
