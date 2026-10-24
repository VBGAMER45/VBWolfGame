Attribute VB_Name = "modWolf"
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
' Ported from js/wolf.js (core constants), js/angle.js and js/random.js,
' plus helpers that reproduce JavaScript integer semantics.

Public Const XRES As Long = 320             ' 3D view size (the browser version used 608x304)
Public Const YRES As Long = 160
Public Const NUMAREAS As Long = 37
Public Const FIRSTAREA As Long = &H6B&      ' first area code in map data
Public Const AMBUSHTILE As Long = &H6A&
Public Const AMBUSH As Long = -2
Public Const TILEGLOBAL As Long = &H10000
Public Const HALFTILE As Long = &H8000&
Public Const TILESHIFT As Long = 16
Public Const MINDIST As Long = &H5800&
Public Const FLOATTILE As Double = 65536#

Public Const ANG_0 As Long = 0              ' angles are in "fines": 128 per degree
Public Const ANG_1 As Long = 128
Public Const ANG_6 As Long = 768
Public Const ANG_15 As Long = 1920
Public Const ANG_22_5 As Long = 2880
Public Const ANG_30 As Long = 3840
Public Const ANG_45 As Long = 5760
Public Const ANG_67_5 As Long = 8640
Public Const ANG_90 As Long = 11520
Public Const ANG_112_5 As Long = 14400
Public Const ANG_135 As Long = 17280
Public Const ANG_157_5 As Long = 20160
Public Const ANG_180 As Long = 23040
Public Const ANG_202_5 As Long = 25920
Public Const ANG_225 As Long = 28800
Public Const ANG_247_5 As Long = 31680
Public Const ANG_270 As Long = 34560
Public Const ANG_292_5 As Long = 37440
Public Const ANG_315 As Long = 40320
Public Const ANG_337_5 As Long = 43200
Public Const ANG_360 As Long = 46080
Public Const ANGLES As Long = 360
Public Const DEATHROTATE As Long = 2

Public Const PI As Double = 3.14159265358979

Private RndTable(0 To 255) As Long
Private RndIndex As Long
Private RndReady As Boolean

'---------------------------------------------------------------- JS integer semantics

' ToInt32: what JavaScript's "x >> 0" and "x | 0" do (truncate, then wrap to 32 bits).
Public Function ToInt32(ByVal d As Double) As Long
    d = Fix(d)
    If d >= -2147483648# And d <= 2147483647# Then
        ToInt32 = d
        Exit Function
    End If
    d = d - Int(d / 4294967296#) * 4294967296#
    If d >= 2147483648# Then d = d - 4294967296#
    ToInt32 = d
End Function

' Arithmetic right shift ("a >> n") of a 32-bit value.
Public Function Sar(ByVal a As Long, ByVal n As Long) As Long
    Sar = Int(a / (2# ^ n))
End Function

' JavaScript "%" on numbers: the result takes the sign of the dividend.
Public Function JsMod(ByVal a As Double, ByVal m As Double) As Double
    JsMod = a - Fix(a / m) * m
End Function

'---------------------------------------------------------------- wolf.js

Public Function TILE2POS(ByVal a As Long) As Long
    TILE2POS = a * TILEGLOBAL + HALFTILE
End Function

Public Function POS2TILE(ByVal a As Double) As Long
    POS2TILE = Int(ToInt32(a) / 65536#)
End Function

Public Function POS2TILEf(ByVal a As Double) As Double
    POS2TILEf = a / FLOATTILE
End Function

Public Function FINE2RAD(ByVal a As Double) As Double
    FINE2RAD = a * PI / ANG_180
End Function

Public Function RAD2FINE(ByVal a As Double) As Double
    RAD2FINE = a * ANG_180 / PI
End Function

Public Function FINE2DEG(ByVal a As Double) As Long
    FINE2DEG = ToInt32(a / ANG_1)
End Function

Public Function FINE2DEGf(ByVal a As Double) As Double
    FINE2DEGf = a / ANG_1
End Function

Public Function DEG2FINE(ByVal a As Double) As Double
    DEG2FINE = a * ANG_1
End Function

'---------------------------------------------------------------- angle.js

Public Function DEG2RAD(ByVal a As Double) As Double
    DEG2RAD = a * 1.74532925199433E-02
End Function

Public Function RAD2DEG(ByVal a As Double) As Double
    RAD2DEG = a / 1.74532925199433E-02
End Function

' Absolute difference between two angles in radians, always 0..PI.
Public Function Angle_Diff(ByVal angle1 As Double, ByVal angle2 As Double) As Double
    Dim d As Double
    If angle1 > angle2 Then d = angle1 - angle2 Else d = angle2 - angle1
    If d > PI Then Angle_Diff = 2 * PI - d Else Angle_Diff = d
End Function

' Clockwise distance from angle2 to angle1; may exceed PI.
Public Function Angle_DistCW(ByVal angle1 As Double, ByVal angle2 As Double) As Double
    If angle1 > angle2 Then
        Angle_DistCW = angle1 - angle2
    Else
        Angle_DistCW = angle1 + 2 * PI - angle2
    End If
End Function

Public Function Angle_Normalize(ByVal angle As Double) As Double
    Do While angle < 0
        angle = angle + 2 * PI
    Loop
    Do While angle >= 2 * PI
        angle = angle - 2 * PI
    Loop
    Angle_Normalize = angle
End Function

' Linear interpolation that allows for the modulo-360 wrap (degrees).
Public Function Angle_Lerp(ByVal fromA As Double, ByVal toA As Double, ByVal frac As Double) As Double
    If toA - fromA > 180 Then toA = toA - 360
    If toA - fromA < -180 Then toA = toA + 360
    Angle_Lerp = fromA + frac * (toA - fromA)
End Function

'---------------------------------------------------------------- random.js
' John Carmack's table driven pseudo-random number generator.

Public Sub Random_Init(ByVal randomizeSeed As Boolean)
    If Not RndReady Then LoadRndTable
    If randomizeSeed Then
        RndIndex = CLng(Timer * 1000) And &HFF&
    Else
        RndIndex = 0
    End If
End Sub

Public Function Random_Rnd() As Long
    If Not RndReady Then LoadRndTable
    RndIndex = (RndIndex + 1) And &HFF&
    Random_Rnd = RndTable(RndIndex)
End Function

Private Sub LoadRndTable()
    Dim s As String, parts() As String, i As Long
    s = "0,8,109,220,222,241,149,107,75,248,254,140,16,66,74,21,211,47,80,242,154,27,205,128,161,89,77,36,"
    s = s & "95,110,85,48,212,140,211,249,22,79,200,50,28,188,52,140,202,120,68,145,62,70,184,190,91,197,152,224,"
    s = s & "149,104,25,178,252,182,202,182,141,197,4,81,181,242,145,42,39,227,156,198,225,193,219,93,122,175,249,0,"
    s = s & "175,143,70,239,46,246,163,53,163,109,168,135,2,235,25,92,20,145,138,77,69,166,78,176,173,212,166,113,"
    s = s & "94,161,41,50,239,49,111,164,70,60,2,37,171,75,136,156,11,56,42,146,138,229,73,146,77,61,98,196,"
    s = s & "135,106,63,197,195,86,96,203,113,101,170,247,181,113,80,250,108,7,255,237,129,226,79,107,112,166,103,241,"
    s = s & "24,223,239,120,198,58,60,82,128,3,184,66,143,224,145,224,81,206,163,45,63,90,168,114,59,33,159,95,"
    s = s & "28,139,123,98,125,196,15,70,194,253,54,14,109,226,71,17,161,93,186,87,244,138,20,52,123,251,26,36,"
    s = s & "17,46,52,231,232,76,31,221,84,37,216,165,212,106,197,242,98,43,39,175,254,145,190,84,118,222,187,136,"
    s = s & "120,163,236,249"
    parts = Split(s, ",")
    For i = 0 To 255
        RndTable(i) = CLng(parts(i))
    Next
    RndReady = True
End Sub
