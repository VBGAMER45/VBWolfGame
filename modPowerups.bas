Attribute VB_Name = "modPowerups"
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
' Ported from js/powerups.js.

Public Const pow_gibs As Long = 0           '  1% if <=10%
Public Const pow_gibs2 As Long = 1          '  1% if <=10%
Public Const pow_alpo As Long = 2           '  4% if <100%
Public Const pow_firstaid As Long = 3       ' 25% if <100%
Public Const pow_key1 As Long = 4           ' gold key
Public Const pow_key2 As Long = 5           ' silver key
Public Const pow_key3 As Long = 6           ' not used
Public Const pow_key4 As Long = 7           ' not used
Public Const pow_cross As Long = 8          '  100pts
Public Const pow_chalice As Long = 9        '  500pts
Public Const pow_bible As Long = 10         ' 1000pts
Public Const pow_crown As Long = 11         ' 5000pts
Public Const pow_clip As Long = 12          ' 8 bullets if <99
Public Const pow_clip2 As Long = 13         ' 4 bullets if <99
Public Const pow_machinegun As Long = 14
Public Const pow_chaingun As Long = 15
Public Const pow_food As Long = 16          ' 10% if <100%
Public Const pow_fullheal As Long = 17      ' 99%, 25 bullets
Public Const pow_25clip As Long = 18        ' 25 bullets if <99
Public Const pow_spear As Long = 19
Public Const pow_last As Long = 20

Public Type TPowerup
    x As Long               ' tile, or -1 once picked up
    y As Long
    powType As Long
    sprite As Long
End Type

Public Powerups(0 To MAX_POWERUPS - 1) As TPowerup
Public NumPowerups As Long

Private PowTexture(0 To pow_last - 1) As Long
Private TexturesReady As Boolean

Private Sub InitTextures()
    PowTexture(pow_gibs) = SPR_STAT_0 + 34
    PowTexture(pow_gibs2) = SPR_STAT_0 + 38
    PowTexture(pow_alpo) = SPR_STAT_0 + 6
    PowTexture(pow_firstaid) = SPR_STAT_0 + 25
    PowTexture(pow_key1) = SPR_STAT_0 + 20
    PowTexture(pow_key2) = SPR_STAT_0 + 21
    PowTexture(pow_key3) = SPR_STAT_0 + 20
    PowTexture(pow_key4) = SPR_STAT_0 + 20
    PowTexture(pow_cross) = SPR_STAT_0 + 29
    PowTexture(pow_chalice) = SPR_STAT_0 + 30
    PowTexture(pow_bible) = SPR_STAT_0 + 31
    PowTexture(pow_crown) = SPR_STAT_0 + 32
    PowTexture(pow_clip) = SPR_STAT_0 + 26
    PowTexture(pow_clip2) = SPR_STAT_0 + 26
    PowTexture(pow_machinegun) = SPR_STAT_0 + 27
    PowTexture(pow_chaingun) = SPR_STAT_0 + 28
    PowTexture(pow_food) = SPR_STAT_0 + 24
    PowTexture(pow_fullheal) = SPR_STAT_0 + 33
    PowTexture(pow_25clip) = SPR_STAT_0 + 49
    PowTexture(pow_spear) = SPR_STAT_0 + 51
    TexturesReady = True
End Sub

Private Sub Remove(ByVal i As Long)
    Powerups(i).x = -1
    Powerups(i).y = -1
End Sub

Private Function AddNew() As Long
    Dim i As Long
    For i = 0 To NumPowerups - 1
        If Powerups(i).x = -1 Then
            AddNew = i
            Exit Function
        End If
    Next
    If NumPowerups = MAX_POWERUPS Then
        AddNew = 0
        Exit Function
    End If
    NumPowerups = NumPowerups + 1
    With Powerups(NumPowerups - 1)
        .x = -1: .y = -1: .powType = 0: .sprite = -1
    End With
    AddNew = NumPowerups - 1
End Function

Public Sub Powerups_Reset()
    NumPowerups = 0
End Sub

Public Sub Powerups_Spawn(ByVal x As Long, ByVal y As Long, ByVal powType As Long)
    Dim p As Long
    If Not TexturesReady Then InitTextures
    p = AddNew()
    With Powerups(p)
        .sprite = Sprites_GetNew()
        .powType = powType
        .x = x
        .y = y
        Sprites_SetPos .sprite, TILE2POS(x), TILE2POS(y), 0
        Sprites_SetTex .sprite, -1, PowTexture(powType)
    End With
    TileMap(x, y) = TileMap(x, y) Or POWERUP_TILE
End Sub

Private Sub TreasureFound()
    FoundTreasure = FoundTreasure + 1
    If FoundTreasure = TotalTreasure Then Game_Notify "You found the last treasure!"
End Sub

' Returns True if the player took the item.
Private Function Give(ByVal powType As Long) As Boolean
    Select Case powType
        Case pow_key1, pow_key2, pow_key3, pow_key4
            Player_GiveKey powType - pow_key1
            Sound_Start "lsfx/012.wav", -1, -1
            Select Case powType - pow_key1
                Case 0: Game_Notify "Gold key"
                Case 1: Game_Notify "Silver key"
                Case Else: Game_Notify "? key"
            End Select
        Case pow_cross
            Player_GivePoints 100
            Sound_Start "lsfx/035.wav", -1, -1
            TreasureFound
        Case pow_chalice
            Player_GivePoints 500
            Sound_Start "lsfx/036.wav", -1, -1
            TreasureFound
        Case pow_bible
            Player_GivePoints 1000
            Sound_Start "lsfx/037.wav", -1, -1
            TreasureFound
        Case pow_crown
            Player_GivePoints 5000
            Sound_Start "lsfx/045.wav", -1, -1
            TreasureFound
        Case pow_gibs
            If Not Player_GiveHealth(1, 11) Then Exit Function
            Sound_Start "lsfx/061.wav", -1, -1
        Case pow_alpo
            If Not Player_GiveHealth(4, 0) Then Exit Function
            Sound_Start "lsfx/033.wav", -1, -1
        Case pow_food
            If Not Player_GiveHealth(10, 0) Then Exit Function
            Sound_Start "lsfx/033.wav", -1, -1
        Case pow_firstaid
            If Not Player_GiveHealth(25, 0) Then Exit Function
            Sound_Start "lsfx/034.wav", -1, -1
        Case pow_clip
            If Not Player_GiveAmmo(AMMO_BULLETS, 8) Then Exit Function
            Sound_Start "lsfx/031.wav", -1, -1
        Case pow_clip2
            If Not Player_GiveAmmo(AMMO_BULLETS, 4) Then Exit Function
            Sound_Start "lsfx/031.wav", -1, -1
        Case pow_25clip
            If Not Player_GiveAmmo(AMMO_BULLETS, 25) Then Exit Function
            Sound_Start "lsfx/031.wav", -1, -1
        Case pow_machinegun
            Player_GiveWeapon WEAPON_AUTO
            Sound_Start "lsfx/030.wav", -1, -1
            Game_Notify "Machinegun"
        Case pow_chaingun
            Player_GiveWeapon WEAPON_CHAIN
            Sound_Start "lsfx/038.wav", -1, -1
            Game_Notify "Chaingun"
            Player.faceCount = -100
            Player.faceGotGun = True
        Case pow_fullheal
            Player_GiveHealth 99, 99
            Player_GiveAmmo AMMO_BULLETS, 25
            Player_GiveLife
            FoundTreasure = FoundTreasure + 1
            If FoundTreasure = TotalTreasure Then
                Game_Notify "You found the last treasure!"
            Else
                Game_Notify "Full Heal"
            End If
            Sound_Start "lsfx/034.wav", -1, -1
        Case Else
            ' unknown item type
    End Select
    Game_StartBonusFlash
    Give = True
End Function

Public Sub Powerups_PickUp(ByVal x As Long, ByVal y As Long)
    Dim i As Long, pLeft As Boolean
    For i = 0 To NumPowerups - 1
        If Powerups(i).x = x And Powerups(i).y = y Then
            If Give(Powerups(i).powType) Then
                Sprites_Remove Powerups(i).sprite
                Remove i
            Else
                pLeft = True
            End If
        End If
    Next
    If pLeft Then
        TileMap(x, y) = TileMap(x, y) Or POWERUP_TILE
    Else
        TileMap(x, y) = TileMap(x, y) And Not POWERUP_TILE
    End If
End Sub
