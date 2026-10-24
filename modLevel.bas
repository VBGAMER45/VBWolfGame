Attribute VB_Name = "modLevel"
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
' Ported from js/level.js. The JS level object becomes module-level state
' here, since only one level is loaded at a time.

Public Const WALL_TILE As Long = 1
Public Const DOOR_TILE As Long = 2
Public Const SECRET_TILE As Long = 4
Public Const DRESS_TILE As Long = 8
Public Const BLOCK_TILE As Long = 16
Public Const ACTOR_TILE As Long = 32
Public Const DEADACTOR_TILE As Long = 64
Public Const POWERUP_TILE As Long = 128
Public Const AMBUSH_TILE As Long = 256
Public Const EXIT_TILE As Long = 512
Public Const SECRETLEVEL_TILE As Long = 1024
Public Const ELEVATOR_TILE As Long = &H800&
Public Const TILE_IS_E_TURN As Long = &H1000&
Public Const TILE_IS_NE_TURN As Long = &H2000&
Public Const TILE_IS_N_TURN As Long = &H4000&
Public Const TILE_IS_NW_TURN As Long = &H8000&
Public Const TILE_IS_W_TURN As Long = &H10000
Public Const TILE_IS_SW_TURN As Long = &H20000
Public Const TILE_IS_S_TURN As Long = &H40000
Public Const TILE_IS_SE_TURN As Long = &H80000
Public Const PUSHWALL_TILE As Long = &H100000
Public Const SOLID_TILE As Long = WALL_TILE Or BLOCK_TILE Or PUSHWALL_TILE
Public Const BLOCKS_MOVE_TILE As Long = SOLID_TILE Or ACTOR_TILE
Public Const WAYPOINT_TILE As Long = TILE_IS_E_TURN Or TILE_IS_NE_TURN Or TILE_IS_N_TURN Or TILE_IS_NW_TURN Or TILE_IS_W_TURN Or TILE_IS_SW_TURN Or TILE_IS_S_TURN Or TILE_IS_SE_TURN

Public Const MAPHEADER_SIZE As Long = 49
Public Const MAP_SIGNATURE As Long = &H21444921     ' "!ID!"
Public Const MAX_POWERUPS As Long = 1000
Public Const NUM_STATICS As Long = 48

' read-only after level load
Public TileMap(0 To 63, 0 To 63) As Long
Public Areas(0 To 63, 0 To 63) As Long
Public WallTexX(0 To 63, 0 To 63) As Long
Public WallTexY(0 To 63, 0 To 63) As Long
Public Plane1(0 To 4095) As Long
Public Plane2(0 To 4095) As Long
Public Plane3(0 To 4095) As Long
Public LevelWidth As Long
Public LevelHeight As Long
Public LevelCeiling(0 To 3) As Long     ' RGBA
Public LevelFloor(0 To 3) As Long
Public LevelName As String
Public LevelMusic As String
Public LevelParTime As String
Public LevelFile As String
Public SpawnX As Long
Public SpawnY As Long
Public SpawnAngle As Long

' level.state
Public TotalSecrets As Long
Public FoundSecrets As Long
Public TotalTreasure As Long
Public FoundTreasure As Long
Public TotalMonsters As Long
Public KilledMonsters As Long
Public LevelTime As Double              ' seconds played on this level

Private StatBlock(0 To NUM_STATICS - 1) As Boolean
Private StatPowerup(0 To NUM_STATICS - 1) As Long
Private StatInfoReady As Boolean

' statinfo: which statics block movement and which are really powerups.
Private Sub InitStatInfo()
    Const BLOCKS As String = "011101011011110011100010000000000001110110000110"
    Dim i As Long
    For i = 0 To NUM_STATICS - 1
        StatBlock(i) = (Mid$(BLOCKS, i + 1, 1) = "1")
        StatPowerup(i) = -1
    Next
    StatPowerup(6) = pow_alpo           ' bad food
    StatPowerup(20) = pow_key1          ' key 1
    StatPowerup(21) = pow_key2          ' key 2
    StatPowerup(24) = pow_food          ' good food
    StatPowerup(25) = pow_firstaid      ' first aid
    StatPowerup(26) = pow_clip          ' clip
    StatPowerup(27) = pow_machinegun    ' machine gun
    StatPowerup(28) = pow_chaingun      ' gatling gun
    StatPowerup(29) = pow_cross         ' treasure 1
    StatPowerup(30) = pow_chalice       ' treasure 2
    StatPowerup(31) = pow_bible         ' treasure 3
    StatPowerup(32) = pow_crown         ' treasure 4
    StatPowerup(33) = pow_fullheal      ' one up
    StatPowerup(34) = pow_gibs          ' gibs
    StatPowerup(38) = pow_gibs          ' gibs 2
    StatInfoReady = True
End Sub

Private Sub NewLevel()
    Dim x As Long, y As Long
    SpawnX = -1: SpawnY = -1: SpawnAngle = 0
    LevelParTime = "": LevelName = "": LevelMusic = ""
    TotalSecrets = 0: FoundSecrets = 0
    TotalTreasure = 0: FoundTreasure = 0
    TotalMonsters = 0: KilledMonsters = 0
    LevelTime = 0
    Sprites_Reset
    Powerups_Reset
    PushWall_Reset
    For x = 0 To 63
        For y = 0 To 63
            Areas(x, y) = 0
            TileMap(x, y) = 0
            WallTexX(x, y) = 0
            WallTexY(x, y) = 0
        Next
    Next
End Sub

' Loads and parses a .map file. Returns "" on success or an error message.
Public Function Level_Load(ByVal path As String) As String
    Dim data() As Byte
    If Not StatInfoReady Then InitStatInfo
    If Not File_Load(path, data) Then
        Level_Load = "Could not read " & path
        Exit Function
    End If
    LevelFile = path
    Level_Load = ParseMapData(data)
End Function

Public Function Level_Reload() As String
    Level_Reload = Level_Load(LevelFile)
End Function

Private Function ParseMapData(ByRef data() As Byte) As String
    Dim pos As Long, rle As Long, i As Long, size As Long
    Dim length(0 To 2) As Long, offset(0 To 2) As Long
    Dim mapNameLength As Long, musicNameLength As Long
    Dim x As Long, y As Long, y0 As Long, layer1 As Long, layer2 As Long

    NewLevel
    size = UBound(data) + 1
    If size < MAPHEADER_SIZE Then
        ParseMapData = "Map file size is smaller than mapheader size"
        Exit Function
    End If
    If File_ReadUInt32(data, pos) <> MAP_SIGNATURE Then
        ParseMapData = "File signature does not match MAP_SIGNATURE"
        Exit Function
    End If

    rle = File_ReadUInt16(data, pos)
    LevelWidth = File_ReadUInt16(data, pos)
    LevelHeight = File_ReadUInt16(data, pos)
    For i = 0 To 3
        LevelCeiling(i) = File_ReadUInt8(data, pos)
    Next
    For i = 0 To 3
        LevelFloor(i) = File_ReadUInt8(data, pos)
    Next
    For i = 0 To 2
        length(i) = File_ReadUInt16(data, pos)
    Next
    For i = 0 To 2
        offset(i) = File_ReadUInt32(data, pos)
    Next
    mapNameLength = File_ReadUInt16(data, pos)
    musicNameLength = File_ReadUInt16(data, pos)
    pos = pos + 4                               ' (single) fpartime
    LevelParTime = File_ReadString(data, pos, 5)

    If size < MAPHEADER_SIZE + mapNameLength + musicNameLength + length(0) + length(1) + length(2) Then
        ParseMapData = "filesize is less than MAPHEADER_SIZE + mapNameLength + musicNameLength + etc"
        Exit Function
    End If
    LevelName = File_ReadString(data, pos, mapNameLength)
    LevelMusic = File_ReadString(data, pos, musicNameLength)

    ReadPlaneData data, offset(0), length(0), rle, Plane1
    ReadPlaneData data, offset(1), length(1), rle, Plane2
    ReadPlaneData data, offset(2), length(2), rle, Plane3

    Doors_Reset

    ' map rows are stored north first; the engine's y axis points north
    For y0 = 0 To 63
        For x = 0 To 63
            y = 63 - y0
            layer1 = Plane1(y0 * 64 + x)
            layer2 = Plane2(y0 * 64 + x)

            If layer2 <> 0 Then SpawnObj layer2, x, y

            If layer1 = 0 Then
                Areas(x, y) = -3                    ' unknown area
            ElseIf layer1 < &H6A& Then              ' solid map object
                If (layer1 >= &H5A& And layer1 <= &H5F&) Or layer1 = &H64& Or layer1 = &H65& Then
                    TileMap(x, y) = TileMap(x, y) Or DOOR_TILE
                    Doors_Spawn x, y, layer1
                    Areas(x, y) = -2                ' door area
                Else
                    TileMap(x, y) = TileMap(x, y) Or WALL_TILE
                    WallTexX(x, y) = (layer1 - 1) * 2 + 1
                    WallTexY(x, y) = (layer1 - 1) * 2
                    Areas(x, y) = -1                ' wall area
                    If layer1 = &H15& Then TileMap(x, y) = TileMap(x, y) Or ELEVATOR_TILE
                End If
            ElseIf layer1 = &H6A& Then              ' ambush floor tile
                TileMap(x, y) = TileMap(x, y) Or AMBUSH_TILE
                Areas(x, y) = -3
            ElseIf layer1 >= FIRSTAREA And layer1 < FIRSTAREA + NUMAREAS Then
                If layer1 = FIRSTAREA Then TileMap(x, y) = TileMap(x, y) Or SECRETLEVEL_TILE
                Areas(x, y) = layer1 - FIRSTAREA    ' spawn area
            Else
                Areas(x, y) = -3
            End If
        Next
    Next

    ' JDC: replace unknown areas with an adjacent area, to avoid the silent
    ' attack / no damage problem when an ambush guard is stuck on its tile.
    ' (The last test reads [x+1][y+1] but copies [x][y+1], as in the original.)
    For x = 1 To 62
        For y = 1 To 62
            If Areas(x, y) = -3 Then
                If Areas(x - 1, y) >= 0 Then
                    Areas(x, y) = Areas(x - 1, y)
                ElseIf Areas(x + 1, y) >= 0 Then
                    Areas(x, y) = Areas(x + 1, y)
                ElseIf Areas(x, y - 1) >= 0 Then
                    Areas(x, y) = Areas(x, y - 1)
                ElseIf Areas(x + 1, y + 1) >= 0 Then
                    Areas(x, y) = Areas(x, y + 1)
                End If
            End If
        Next
    Next

    Doors_SetAreas
    ParseMapData = ""
End Function

Private Sub ReadPlaneData(ByRef data() As Byte, ByVal offset As Long, ByVal length As Long, ByVal rle As Long, ByRef plane() As Long)
    Dim pos As Long, expandedLength As Long, words() As Long
    pos = offset
    expandedLength = File_ReadUInt16(data, pos)
    CarmackExpand data, pos, expandedLength, words
    ' the expanded data starts with the RLEW length word, which is skipped
    RlewExpand words, 1, 64 * 64 * 2, rle, plane
End Sub

Private Sub RlewExpand(ByRef source() As Long, ByVal inptr As Long, ByVal length As Long, ByVal rlewtag As Long, ByRef dest() As Long)
    Dim value As Long, count As Long, i As Long, outptr As Long, endptr As Long
    endptr = length \ 2
    Do
        value = source(inptr): inptr = inptr + 1
        If value <> rlewtag Then
            dest(outptr) = value                ' uncompressed
            outptr = outptr + 1
        Else
            count = source(inptr)               ' compressed string
            value = source(inptr + 1)
            inptr = inptr + 2
            For i = 1 To count
                If outptr < endptr Then dest(outptr) = value
                outptr = outptr + 1
            Next
        End If
    Loop While outptr < endptr
End Sub

Private Sub CarmackExpand(ByRef source() As Byte, ByVal inptr As Long, ByVal length As Long, ByRef dest() As Long)
    Const NEARTAG As Long = &HA7&
    Const FARTAG As Long = &HA8&
    Dim ch As Long, chhigh As Long, count As Long, offset As Long, copyptr As Long, outptr As Long

    length = length \ 2
    ReDim dest(0 To length + 256)
    Do While length > 0
        ch = source(inptr) + source(inptr + 1) * 256&
        inptr = inptr + 2
        chhigh = ch \ 256
        If chhigh = NEARTAG Or chhigh = FARTAG Then
            count = ch And &HFF&
            If count = 0 Then
                ' a word containing the tag byte
                ch = ch Or source(inptr)
                inptr = inptr + 1
                dest(outptr) = ch
                outptr = outptr + 1
                length = length - 1
            Else
                If chhigh = NEARTAG Then
                    offset = source(inptr)
                    inptr = inptr + 1
                    copyptr = outptr - offset
                Else
                    offset = source(inptr) + source(inptr + 1) * 256&
                    inptr = inptr + 2
                    copyptr = offset
                End If
                length = length - count
                Do While count > 0
                    dest(outptr) = dest(copyptr)
                    outptr = outptr + 1
                    copyptr = copyptr + 1
                    count = count - 1
                Loop
            End If
        Else
            dest(outptr) = ch
            outptr = outptr + 1
            length = length - 1
        End If
    Loop
End Sub

Private Sub SpawnObj(ByVal objType As Long, ByVal x As Long, ByVal y As Long)
    If objType >= 23 And objType < 23 + NUM_STATICS Then
        SpawnStatic objType - 23, x, y
        Exit Sub
    End If
    Select Case objType
        Case &H13                               ' start N
            SpawnX = TILE2POS(x): SpawnY = TILE2POS(y): SpawnAngle = ANG_90
        Case &H14                               ' start E
            SpawnX = TILE2POS(x): SpawnY = TILE2POS(y): SpawnAngle = ANG_0
        Case &H15                               ' start S
            SpawnX = TILE2POS(x): SpawnY = TILE2POS(y): SpawnAngle = ANG_270
        Case &H16                               ' start W
            SpawnX = TILE2POS(x): SpawnY = TILE2POS(y): SpawnAngle = ANG_180
        Case &H5A: TileMap(x, y) = TileMap(x, y) Or TILE_IS_E_TURN
        Case &H5B: TileMap(x, y) = TileMap(x, y) Or TILE_IS_NE_TURN
        Case &H5C: TileMap(x, y) = TileMap(x, y) Or TILE_IS_N_TURN
        Case &H5D: TileMap(x, y) = TileMap(x, y) Or TILE_IS_NW_TURN
        Case &H5E: TileMap(x, y) = TileMap(x, y) Or TILE_IS_W_TURN
        Case &H5F: TileMap(x, y) = TileMap(x, y) Or TILE_IS_SW_TURN
        Case &H60: TileMap(x, y) = TileMap(x, y) Or TILE_IS_S_TURN
        Case &H61: TileMap(x, y) = TileMap(x, y) Or TILE_IS_SE_TURN
        Case &H62                               ' pushwall modifier
            TileMap(x, y) = TileMap(x, y) Or SECRET_TILE
            TotalSecrets = TotalSecrets + 1
        Case &H63                               ' victory trigger
            TileMap(x, y) = TileMap(x, y) Or EXIT_TILE
    End Select
End Sub

Private Sub SpawnStatic(ByVal statType As Long, ByVal x As Long, ByVal y As Long)
    Dim s As Long, pu As Long
    If StatPowerup(statType) = -1 Then
        If StatBlock(statType) Then
            TileMap(x, y) = TileMap(x, y) Or BLOCK_TILE     ' blocking static
        Else
            TileMap(x, y) = TileMap(x, y) Or DRESS_TILE     ' dressing static
        End If
        s = Sprites_GetNew()
        If s < 0 Then Exit Sub
        Sprites_SetPos s, TILE2POS(x), TILE2POS(y), 0
        Sprites_SetTex s, 0, SPR_STAT_0 + statType
    Else
        pu = StatPowerup(statType)
        Powerups_Spawn x, y, pu
        If pu = pow_cross Or pu = pow_chalice Or pu = pow_bible Or pu = pow_crown Or pu = pow_fullheal Then
            TotalTreasure = TotalTreasure + 1
        End If
    End If
End Sub

' Spawns the actors on the info plane for the given skill. Each enemy has a
' block of codes: stand facing E/N/W/S, then patrol facing E/N/W/S; the same
' block 36 codes higher (18 for mutants) appears from medium skill, and 72
' (36) higher only on hard. Bosses, ghosts and other special actors are not
' part of this port.
Public Sub Level_ScanInfoPlane(ByVal skill As Long)
    Dim x As Long, y As Long, tile As Long
    Actors_Reset
    For y = 0 To 63
        For x = 0 To 63
            tile = Plane2((63 - y) * 64 + x)
            If tile <> 0 Then
                If tile = 124 Then
                    Actors_SpawnDeadGuard skill, en_guard, x, y
                Else
                    SpawnEnemy skill, tile, x, y
                End If
            End If
        Next
    Next
End Sub

Private Sub SpawnEnemy(ByVal skill As Long, ByVal tile As Long, ByVal x As Long, ByVal y As Long)
    Dim bases As Variant, kinds As Variant, steps As Variant, k As Long, tier As Long, code As Long
    bases = Array(108, 116, 126, 134, 216)          ' stand codes; patrol codes are base + 4
    kinds = Array(en_guard, en_officer, en_ss, en_dog, en_mutant)
    steps = Array(36, 36, 36, 36, 18)
    For k = 0 To 4
        For tier = 0 To 2
            code = tile - bases(k) - tier * steps(k)
            If code >= 0 And code <= 7 Then
                If (tier = 1 And skill < gd_medium) Or (tier = 2 And skill < gd_hard) Then Exit Sub
                If kinds(k) = en_dog And code < 4 Then Exit Sub         ' dogs only patrol
                If code < 4 Then
                    Actors_SpawnStand skill, kinds(k), x, y, code
                Else
                    Actors_SpawnPatrol skill, kinds(k), x, y, code - 4
                End If
                Exit Sub
            End If
        Next
    Next
End Sub

' True if nothing blocks the line between two points (walls and the closed
' part of doors). Fixed point with 8 bits of fraction, as in the original.
Public Function Level_CheckLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Const FRACBITS As Long = 8
    Dim xt1 As Long, yt1 As Long, xt2 As Long, yt2 As Long, x As Long, y As Long
    Dim xdist As Long, ydist As Long, xstep As Double, ystep As Double
    Dim deltafrac As Long, frac As Double, partial As Long, intercept As Long, d As Long

    xt1 = Sar(x1, TILESHIFT): yt1 = Sar(y1, TILESHIFT)
    xt2 = Sar(x2, TILESHIFT): yt2 = Sar(y2, TILESHIFT)
    xdist = Abs(xt2 - xt1)
    ydist = Abs(yt2 - yt1)
    x1 = Sar(x1, FRACBITS): y1 = Sar(y1, FRACBITS)
    x2 = Sar(x2, FRACBITS): y2 = Sar(y2, FRACBITS)

    If xdist <> 0 Then
        If xt2 > xt1 Then
            partial = 256 - (x1 And &HFF&)
            xstep = 1
        Else
            partial = x1 And &HFF&
            xstep = -1
        End If
        deltafrac = Abs(x2 - x1)
        ystep = ToInt32((y2 - y1) * 256#) / deltafrac
        frac = y1 + Sar(ToInt32(ystep * partial), FRACBITS)
        x = xt1 + xstep
        xt2 = xt2 + xstep
        Do
            y = Sar(ToInt32(frac), FRACBITS)
            frac = frac + ystep
            If (TileMap(x, y) And WALL_TILE) <> 0 Then Exit Function
            If (TileMap(x, y) And DOOR_TILE) <> 0 Then
                d = DoorMap(x, y)
                If Doors(d).action <> dr_open Then
                    If Doors(d).action = dr_closed Then Exit Function
                    intercept = Sar(ToInt32(frac - ystep / 2) And &HFF&, 4)     ' 1/64 of a tile
                    If intercept < (63 - Doors(d).ticcount) Then Exit Function
                End If
            End If
            x = x + xstep
        Loop While x <> xt2
    End If

    If ydist <> 0 Then
        If yt2 > yt1 Then
            partial = 256 - (y1 And &HFF&)
            ystep = 1
        Else
            partial = y1 And &HFF&
            ystep = -1
        End If
        deltafrac = Abs(y2 - y1)
        xstep = ToInt32((x2 - x1) * 256#) / deltafrac
        frac = x1 + Sar(ToInt32(xstep * partial), FRACBITS)
        y = yt1 + ystep
        yt2 = yt2 + ystep
        Do
            x = Sar(ToInt32(frac), FRACBITS)
            frac = frac + xstep
            If (TileMap(x, y) And WALL_TILE) <> 0 Then Exit Function
            If (TileMap(x, y) And DOOR_TILE) <> 0 Then
                d = DoorMap(x, y)
                If Doors(d).action <> dr_open Then
                    If Doors(d).action = dr_closed Then Exit Function
                    intercept = Sar(ToInt32(frac - xstep / 2) And &HFF&, 4)
                    If intercept < Doors(d).ticcount Then Exit Function
                End If
            End If
            y = y + ystep
        Loop While y <> yt2
    End If
    Level_CheckLine = True
End Function
