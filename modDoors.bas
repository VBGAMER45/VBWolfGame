Attribute VB_Name = "modDoors"
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
' Ported from js/doors.js. Doors are kept in an array; DoorMap holds the
' index of the door on each tile, or -1.

Public Const CLOSEWALL As Long = MINDIST    ' space between wall & player
Public Const MAXDOORS As Long = 64
Public Const DOOR_TIMEOUT As Long = 300
Public Const DOOR_MINOPEN As Long = 50
Public Const DOOR_FULLOPEN As Long = 63
Public Const DOOR_VERT As Long = 255
Public Const DOOR_HORIZ As Long = 254
Public Const DOOR_E_VERT As Long = 253
Public Const DOOR_E_HORIZ As Long = 252
Public Const DOOR_G_VERT As Long = 251
Public Const DOOR_G_HORIZ As Long = 250
Public Const DOOR_S_VERT As Long = 249
Public Const DOOR_S_HORIZ As Long = 248
Public Const FIRST_DOOR As Long = 248
Public Const LAST_LOCK As Long = 251
Public Const TEX_DOOR As Long = 98
Public Const TEX_DDOOR As Long = TEX_DOOR           ' simple door
Public Const TEX_PLATE As Long = TEX_DOOR + 2       ' door plate (the wall beside a door)
Public Const TEX_DELEV As Long = TEX_DOOR + 4       ' elevator door
Public Const TEX_DLOCK As Long = TEX_DOOR + 6       ' locked door

Public Const dr_closing As Long = -1
Public Const dr_closed As Long = 0
Public Const dr_opening As Long = 1
Public Const dr_open As Long = 2

Public Type TDoor
    doorType As Long
    vertical As Boolean
    texture As Long
    ticcount As Long
    tileX As Long
    tileY As Long
    action As Long
    area1 As Long
    area2 As Long
End Type

Public Doors(0 To MAXDOORS - 1) As TDoor
Public NumDoors As Long
Public DoorMap(0 To 63, 0 To 63) As Long

Public Sub Doors_Reset()
    Dim x As Long, y As Long
    NumDoors = 0
    For x = 0 To 63
        For y = 0 To 63
            DoorMap(x, y) = -1
        Next
    Next
End Sub

Public Function Doors_Spawn(ByVal x As Long, ByVal y As Long, ByVal doorType As Long) As Long
    Dim d As Long
    If NumDoors >= MAXDOORS Then Err.Raise vbObjectError + 1, "Doors_Spawn", "Too many Doors on level!"
    d = NumDoors
    DoorMap(x, y) = d
    With Doors(d)
        .doorType = -1: .vertical = False: .texture = -1: .ticcount = 0
        Select Case doorType
            Case &H5A: .doorType = DOOR_VERT: .vertical = True: .texture = TEX_DDOOR + 1
            Case &H5B: .doorType = DOOR_HORIZ: .vertical = False: .texture = TEX_DDOOR
            Case &H5C: .doorType = DOOR_G_VERT: .vertical = True: .texture = TEX_DLOCK
            Case &H5D: .doorType = DOOR_G_HORIZ: .vertical = False: .texture = TEX_DLOCK
            Case &H5E: .doorType = DOOR_S_VERT: .vertical = True: .texture = TEX_DLOCK + 1
            Case &H5F: .doorType = DOOR_S_HORIZ: .vertical = False: .texture = TEX_DLOCK + 1
            Case &H64: .doorType = DOOR_E_VERT: .vertical = True: .texture = TEX_DELEV + 1
            Case &H65: .doorType = DOOR_E_HORIZ: .vertical = False: .texture = TEX_DELEV
            Case Else: Err.Raise vbObjectError + 2, "Doors_Spawn", "Unknown door type: " & doorType
        End Select
        .tileX = x: .tileY = y
        .action = dr_closed
        .area1 = 0: .area2 = 0
    End With
    NumDoors = NumDoors + 1
    Doors_Spawn = d
End Function

' 0 if the door is closed, DOOR_FULLOPEN if open, in between while moving.
Public Function Doors_Opened(ByVal d As Long) As Long
    If Doors(d).action = dr_open Then Doors_Opened = DOOR_FULLOPEN Else Doors_Opened = Doors(d).ticcount
End Function

Public Sub Doors_Process(ByVal tics As Long)
    Dim n As Long, doorPosX As Long, doorPosY As Long
    If Player.playstate = ex_victory Then Exit Sub

    For n = 0 To NumDoors - 1
        With Doors(n)
            doorPosX = TILE2POS(.tileX): doorPosY = TILE2POS(.tileY)
            Select Case .action
                Case dr_closed
                    ' this door is closed
                Case dr_opening
                    If .ticcount >= DOOR_FULLOPEN Then      ' door fully opened
                        .action = dr_open
                        .ticcount = 0
                    Else
                        If .ticcount = 0 Then
                            Areas_Join .area1, .area2
                            Areas_Connect Player.areanumber
                            If AreaByPlayer(.area1) Then Sound_Start "sfx/010.wav", doorPosX, doorPosY
                        End If
                        .ticcount = .ticcount + tics
                        If .ticcount > DOOR_FULLOPEN Then .ticcount = DOOR_FULLOPEN
                    End If
                Case dr_closing
                    If .ticcount <= 0 Then                  ' fully closed: disconnect areas
                        Areas_Disconnect .area1, .area2
                        Areas_Connect Player.areanumber
                        .ticcount = 0
                        .action = dr_closed
                    Else
                        If .ticcount = DOOR_FULLOPEN Then
                            If AreaByPlayer(.area1) Then Sound_Start "sfx/007.wav", doorPosX, doorPosY
                        End If
                        .ticcount = .ticcount - tics
                        If .ticcount < 0 Then .ticcount = 0
                    End If
                Case dr_open
                    If .ticcount > DOOR_MINOPEN Then
                        ' do not close the door immediately if something is in the way
                        If Not CanCloseDoor(.tileX, .tileY, .vertical) Then .ticcount = DOOR_MINOPEN
                    End If
                    If .ticcount >= DOOR_TIMEOUT Then
                        .action = dr_closing
                        .ticcount = DOOR_FULLOPEN
                    Else
                        .ticcount = .ticcount + tics
                    End If
            End Select
        End With
    Next
End Sub

Public Sub Doors_SetAreas()
    Dim n As Long, x As Long, y As Long
    For n = 0 To NumDoors - 1
        x = Doors(n).tileX: y = Doors(n).tileY
        If Doors(n).vertical Then
            Doors(n).area1 = AreaOrZero(x + 1, y)
            Doors(n).area2 = AreaOrZero(x - 1, y)
        Else
            Doors(n).area1 = AreaOrZero(x, y + 1)
            Doors(n).area2 = AreaOrZero(x, y - 1)
        End If
    Next
End Sub

Private Function AreaOrZero(ByVal x As Long, ByVal y As Long) As Long
    If x < 0 Or y < 0 Or x > 63 Or y > 63 Then Exit Function
    If Areas(x, y) >= 0 Then AreaOrZero = Areas(x, y)
End Function

Public Sub Doors_Open(ByVal d As Long)
    If Doors(d).action = dr_open Then
        Doors(d).ticcount = 0           ' reset opened time
    Else
        Doors(d).action = dr_opening    ' start opening it
    End If
End Sub

Private Sub ChangeDoorState(ByVal d As Long)
    If Doors(d).action < dr_opening Then
        Doors_Open d
    End If
    ' (the browser version does not let the player close an open door)
End Sub

Private Function CanCloseDoor(ByVal x As Long, ByVal y As Long, ByVal vert As Boolean) As Boolean
    Dim tileX As Long, tileY As Long, n As Long
    tileX = POS2TILE(Player.x): tileY = POS2TILE(Player.y)
    If tileX = x And tileY = y Then Exit Function

    If vert Then
        If tileY = y Then
            If POS2TILE(Player.x + CLOSEWALL) = x Then Exit Function
            If POS2TILE(Player.x - CLOSEWALL) = x Then Exit Function
        End If
        For n = 0 To NumGuards - 1
            If Guards(n).inUse Then
                With Guards(n)
                    If .tileX = x And .tileY = y Then Exit Function                                 ' guard in door
                    If .tileX = x - 1 And .tileY = y And POS2TILE(.x + CLOSEWALL) = x Then Exit Function
                    If .tileX = x + 1 And .tileY = y And POS2TILE(.x - CLOSEWALL) = x Then Exit Function
                End With
            End If
        Next
    Else
        If tileX = x Then
            If POS2TILE(Player.y + CLOSEWALL) = y Then Exit Function
            If POS2TILE(Player.y - CLOSEWALL) = y Then Exit Function
        End If
        For n = 0 To NumGuards - 1
            If Guards(n).inUse Then
                With Guards(n)
                    If .tileX = x And .tileY = y Then Exit Function
                    If .tileX = x And .tileY = y - 1 And POS2TILE(.y + CLOSEWALL) = y Then Exit Function
                    If .tileX = x And .tileY = y + 1 And POS2TILE(.y - CLOSEWALL) = y Then Exit Function
                End With
            End If
        Next
    End If
    CanCloseDoor = True
End Function

Public Function Doors_TryUse(ByVal d As Long) As Boolean
    Select Case Doors(d).doorType
        Case DOOR_VERT, DOOR_HORIZ, DOOR_E_VERT, DOOR_E_HORIZ
            ChangeDoorState d               ' does not require a key
        Case DOOR_G_VERT, DOOR_G_HORIZ
            If (Player.items And ITEM_KEY_1) <> 0 Then
                ChangeDoorState d
            Else
                Game_Notify "You need a gold key"
            End If
        Case DOOR_S_VERT, DOOR_S_HORIZ
            If (Player.items And ITEM_KEY_2) <> 0 Then
                ChangeDoorState d
            Else
                Game_Notify "You need a silver key"
            End If
    End Select
    Doors_TryUse = True
End Function
