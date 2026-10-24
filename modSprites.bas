Attribute VB_Name = "modSprites"
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
' Ported from the sprite list code in js/sprites.js. Removed sprites keep
' their slot (flagged SPRT_REMOVE) and are recycled, as in the C original,
' so that other modules can hold sprite indices.

Public Const SPRT_ONE_TEX As Long = 1       ' use one texture for each rotation
Public Const SPRT_NO_ROT As Long = 2        ' do not rotate sprite (fence)
Public Const SPRT_CHG_POS As Long = 4
Public Const SPRT_CHG_TEX As Long = 8
Public Const SPRT_REMOVE As Long = 16
Public Const MAX_SPRITES As Long = 1024
Public Const MAX_VIS_SPRITES As Long = 128

Public Const SPR_DEMO As Long = 0
Public Const SPR_DEATHCAM As Long = 1
Public Const SPR_STAT_0 As Long = 2         ' static sprite n is SPR_STAT_0 + n (n = 0..51)

Public Type TSprite
    x As Long
    y As Long
    angle As Long
    tileX As Long           ' tile at or before the sprite, so only 4 tiles need checking
    tileY As Long
    flags As Long
    tex(0 To 7) As Long     ' one texture per rotation, or tex(0) with SPRT_ONE_TEX
End Type

Public Type TVisSprite
    dist As Long
    x As Long
    y As Long
    angle As Long
    tex As Long
    sprite As Long
End Type

Public Sprites(0 To MAX_SPRITES - 1) As TSprite
Public NumSprites As Long
Public VisList(0 To MAX_VIS_SPRITES) As TVisSprite
Public NumVisible As Long

Public Sub Sprites_Reset()
    NumSprites = 0
    NumVisible = 0
End Sub

' Returns a free sprite index, or -1 when the list is full.
Public Function Sprites_GetNew() As Long
    Dim n As Long, i As Long
    n = -1
    For i = 0 To NumSprites - 1
        If (Sprites(i).flags And SPRT_REMOVE) <> 0 Then
            n = i
            Exit For
        End If
    Next
    If n < 0 Then
        If NumSprites >= MAX_SPRITES Then
            Sprites_GetNew = -1
            Exit Function
        End If
        n = NumSprites
        NumSprites = NumSprites + 1
    End If
    With Sprites(n)
        .x = 0: .y = 0: .angle = 0
        .tileX = 0: .tileY = 0
        .flags = 0
        For i = 0 To 7
            .tex(i) = 0
        Next
    End With
    Sprites_GetNew = n
End Function

Public Sub Sprites_SetPos(ByVal s As Long, ByVal x As Long, ByVal y As Long, ByVal angle As Long)
    If s < 0 Then Exit Sub
    With Sprites(s)
        .x = x
        .y = y
        .angle = angle
        .tileX = POS2TILE(x)
        .tileY = POS2TILE(y)
        .flags = .flags Or SPRT_CHG_POS
        If (x And HALFTILE) = 0 Then .tileX = .tileX - 1    ' x within the low half of its tile
        If (y And HALFTILE) = 0 Then .tileY = .tileY - 1
    End With
End Sub

Public Sub Sprites_SetTex(ByVal s As Long, ByVal index As Long, ByVal tex As Long)
    If s < 0 Then Exit Sub
    With Sprites(s)
        If index = -1 Then          ' one texture for each phase
            .tex(0) = tex
            .flags = .flags Or SPRT_ONE_TEX
        Else
            .tex(index) = tex
        End If
        .flags = .flags Or SPRT_CHG_TEX
    End With
End Sub

Public Sub Sprites_Remove(ByVal s As Long)
    If s < 0 Then Exit Sub
    Sprites(s).flags = Sprites(s).flags Or SPRT_REMOVE
End Sub

' Builds VisList from the sprites on tiles the last ray cast marked visible.
Public Sub Sprites_CreateVisList(ByVal viewX As Long, ByVal viewY As Long, ByVal viewAngle As Long)
    Dim n As Long, tx As Long, ty As Long, seen As Boolean
    NumVisible = 0
    For n = 0 To NumSprites - 1
        If (Sprites(n).flags And SPRT_REMOVE) = 0 Then
            tx = Sprites(n).tileX
            ty = Sprites(n).tileY
            If tx > 63 Then tx = 63
            If ty > 63 Then ty = 63
            ' can be in any of 4 surrounding tiles; not 9 - see Sprites_SetPos
            seen = False
            If tx >= 0 And ty >= 0 Then
                seen = VisibleTiles(tx, ty) Or VisibleTiles(tx + 1, ty) Or VisibleTiles(tx, ty + 1) Or VisibleTiles(tx + 1, ty + 1)
            End If
            If seen Then
                With VisList(NumVisible)
                    .dist = Math_LineLen2Point(Sprites(n).x - viewX, Sprites(n).y - viewY, viewAngle)
                    .x = Sprites(n).x
                    .y = Sprites(n).y
                    .angle = Sprites(n).angle
                    .tex = Sprites(n).tex(0)
                    .sprite = n
                End With
                NumVisible = NumVisible + 1
                If NumVisible > MAX_VIS_SPRITES Then Exit For   ' vislist full
            End If
        End If
    Next
End Sub
