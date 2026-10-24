Attribute VB_Name = "modMath"
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
' Ported from js/math.js: lookup tables, quadrants and direction helpers.

Public Const q_first As Long = 0
Public Const q_second As Long = 1
Public Const q_third As Long = 2
Public Const q_fourth As Long = 3

Public Const dir4_east As Long = 0
Public Const dir4_north As Long = 1
Public Const dir4_west As Long = 2
Public Const dir4_south As Long = 3
Public Const dir4_nodir As Long = 4

Public Const dir8_east As Long = 0
Public Const dir8_northeast As Long = 1
Public Const dir8_north As Long = 2
Public Const dir8_northwest As Long = 3
Public Const dir8_west As Long = 4
Public Const dir8_southwest As Long = 5
Public Const dir8_south As Long = 6
Public Const dir8_southeast As Long = 7
Public Const dir8_nodir As Long = 8

Public SinTable(0 To ANG_360 + ANG_90) As Double
Public CosTable(0 To ANG_360) As Double
Public TanTable(0 To ANG_360) As Double
Public XnextTable(0 To ANG_360) As Double
Public YnextTable(0 To ANG_360) As Double
Public ColumnAngle(0 To XRES - 1) As Long     ' view angle offset of each screen column, in fines

Public dx4dir(0 To 4) As Long
Public dy4dir(0 To 4) As Long
Public dx8dir(0 To 8) As Long
Public dy8dir(0 To 8) As Long
Public opposite4(0 To 4) As Long
Public opposite8(0 To 8) As Long
Public dir4to8(0 To 4) As Long
Public dir8angle(0 To 8) As Long
Public dir4angle(0 To 4) As Long
Public diagonal(0 To 8, 0 To 8) As Long

Public Sub Math_BuildTables()
    Dim n As Long, angle As Double, value As Double, tanfov2 As Double, tanval As Double

    For n = 0 To ANG_90
        angle = FINE2RAD(n)
        value = Sin(angle)
        SinTable(n) = value
        SinTable(ANG_180 - n) = value
        SinTable(n + ANG_360) = value
        SinTable(ANG_180 + n) = -value
        SinTable(ANG_360 - n) = -value
    Next
    For n = 0 To ANG_360
        CosTable(n) = SinTable(n + ANG_90)
    Next

    For n = 0 To ANG_360
        angle = FINE2RAD(n)
        If n = ANG_90 Or n = ANG_270 Then
            TanTable(n) = Tan(FINE2RAD(n - 0.5))            ' infinity
            YnextTable(n) = ToInt32(FLOATTILE * Tan(FINE2RAD(n - 0.5)))
        Else
            TanTable(n) = Tan(angle)
            YnextTable(n) = ToInt32(FLOATTILE * Tan(angle))
        End If
        If n = ANG_0 Or n = ANG_360 Then
            XnextTable(n) = ToInt32(FLOATTILE / Tan(FINE2RAD(n + 0.5)))     ' infinity
        ElseIf n = ANG_180 Then
            XnextTable(n) = ToInt32(FLOATTILE / Tan(FINE2RAD(n - 0.5)))     ' -infinity
        ElseIf n = ANG_90 Or n = ANG_270 Then
            XnextTable(n) = 0
        Else
            XnextTable(n) = ToInt32(FLOATTILE / Tan(angle))
        End If
    Next

    tanfov2 = Tan(DEG2RAD(Math_CalcFov(75, XRES, YRES) / 2#)) * (XRES / YRES)
    For n = 0 To XRES - 1
        tanval = tanfov2 * (-1# + 2# * n / (XRES - 1))
        ColumnAngle(n) = ToInt32(RAD2FINE(Atn(tanval)))
    Next

    SetDirectionTables
    Random_Init True
End Sub

Private Sub SetDirectionTables()
    Dim i As Long, j As Long
    dx4dir(0) = 1: dx4dir(1) = 0: dx4dir(2) = -1: dx4dir(3) = 0: dx4dir(4) = 0
    dy4dir(0) = 0: dy4dir(1) = 1: dy4dir(2) = 0: dy4dir(3) = -1: dy4dir(4) = 0
    dx8dir(0) = 1: dx8dir(1) = 1: dx8dir(2) = 0: dx8dir(3) = -1: dx8dir(4) = -1
    dx8dir(5) = -1: dx8dir(6) = 0: dx8dir(7) = 1: dx8dir(8) = 0
    dy8dir(0) = 0: dy8dir(1) = 1: dy8dir(2) = 1: dy8dir(3) = 1: dy8dir(4) = 0
    dy8dir(5) = -1: dy8dir(6) = -1: dy8dir(7) = -1: dy8dir(8) = 0
    opposite4(0) = 2: opposite4(1) = 3: opposite4(2) = 0: opposite4(3) = 1: opposite4(4) = 4
    For i = 0 To 7
        opposite8(i) = (i + 4) Mod 8
    Next
    opposite8(8) = 8
    dir4to8(0) = 0: dir4to8(1) = 2: dir4to8(2) = 4: dir4to8(3) = 6: dir4to8(4) = 8
    For i = 0 To 7
        dir8angle(i) = i * ANG_45
    Next
    dir8angle(8) = ANG_0
    dir4angle(0) = ANG_0: dir4angle(1) = ANG_90: dir4angle(2) = ANG_180: dir4angle(3) = ANG_270: dir4angle(4) = ANG_0

    For i = 0 To 8
        For j = 0 To 8
            diagonal(i, j) = dir8_nodir
        Next
    Next
    diagonal(dir8_east, dir8_north) = dir8_northeast
    diagonal(dir8_east, dir8_south) = dir8_southeast
    diagonal(dir8_north, dir8_east) = dir8_northeast
    diagonal(dir8_north, dir8_west) = dir8_northwest
    diagonal(dir8_west, dir8_north) = dir8_northwest
    diagonal(dir8_west, dir8_south) = dir8_southwest
    diagonal(dir8_south, dir8_east) = dir8_southeast
    diagonal(dir8_south, dir8_west) = dir8_southwest
End Sub

' Vertical field of view (degrees) for a horizontal fov on a width x height view.
Public Function Math_CalcFov(ByVal fovX As Double, ByVal width As Double, ByVal height As Double) As Double
    Math_CalcFov = RAD2DEG(Atn(height / (width / Tan(fovX / 360 * PI)))) * 2
End Function

Public Function Math_NormalizeAngle(ByVal alpha As Long) As Long
    If alpha > ANG_360 Then alpha = alpha Mod ANG_360
    If alpha < ANG_0 Then alpha = ANG_360 - ((-alpha) Mod ANG_360)
    Math_NormalizeAngle = alpha
End Function

Public Function Math_GetQuadrant(ByVal angle As Double) As Long
    angle = Angle_Normalize(angle)
    If angle < PI / 2 Then
        Math_GetQuadrant = q_first
    ElseIf angle < PI Then
        Math_GetQuadrant = q_second
    ElseIf angle < 3 * PI / 2 Then
        Math_GetQuadrant = q_third
    Else
        Math_GetQuadrant = q_fourth
    End If
End Function

Public Function Math_Get4Dir(ByVal angle As Double) As Long
    angle = Angle_Normalize(angle + PI / 4)
    If angle < PI / 2 Then
        Math_Get4Dir = dir4_east
    ElseIf angle < PI Then
        Math_Get4Dir = dir4_north
    ElseIf angle < 3 * PI / 2 Then
        Math_Get4Dir = dir4_west
    Else
        Math_Get4Dir = dir4_south
    End If
End Function

Public Function Math_Get8Dir(ByVal angle As Double) As Long
    angle = Angle_Normalize(angle + PI / 12)
    If angle <= PI / 4 Then
        Math_Get8Dir = dir8_east
    ElseIf angle < PI / 2 Then
        Math_Get8Dir = dir8_northeast
    ElseIf angle <= 3 * PI / 4 Then
        Math_Get8Dir = dir8_north
    ElseIf angle < PI Then
        Math_Get8Dir = dir8_northwest
    ElseIf angle <= 5 * PI / 4 Then
        Math_Get8Dir = dir8_west
    ElseIf angle < 3 * PI / 2 Then
        Math_Get8Dir = dir8_southwest
    ElseIf angle <= 7 * PI / 4 Then
        Math_Get8Dir = dir8_south
    Else
        Math_Get8Dir = dir8_southeast
    End If
End Function

Public Function Math_Point2LineDist(ByVal x As Double, ByVal y As Double, ByVal a As Long) As Long
    Math_Point2LineDist = Abs(ToInt32(x * SinTable(a) - y * CosTable(a)))
End Function

' Distance from the origin to (x, y) measured along direction a: depth in front of a viewer.
Public Function Math_LineLen2Point(ByVal x As Double, ByVal y As Double, ByVal a As Long) As Long
    Math_LineLen2Point = ToInt32(x * CosTable(a) + y * SinTable(a))
End Function

' Angle of the vector from point2 to point1, normalized to 0..2*PI.
Public Function Math_TransformPoint(ByVal p1x As Double, ByVal p1y As Double, ByVal p2x As Double, ByVal p2y As Double) As Double
    Math_TransformPoint = Angle_Normalize(Atan2(p1y - p2y, p1x - p2x))
End Function

Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atan2 = Atn(y / x) + PI Else Atan2 = Atn(y / x) - PI
    ElseIf y > 0 Then
        Atan2 = PI / 2
    ElseIf y < 0 Then
        Atan2 = -PI / 2
    Else
        Atan2 = 0
    End If
End Function
