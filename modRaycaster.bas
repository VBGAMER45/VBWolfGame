Attribute VB_Name = "modRaycaster"
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
' Ported from js/raycaster.js. Intercepts are kept as Doubles (JS numbers)
' and converted with ToInt32 wherever the JS code applies ">>".

Public Const UPPERZCOORD As Double = 0.6
Public Const LOWERZCOORD As Double = -0.6

Public Const TRACE_MARK_MAP As Long = 1     ' marks traced area in the automap
Public Const TRACE_SIGHT As Long = 2        ' player sight
Public Const TRACE_SIGHT_AI As Long = 4     ' enemy sight
Public Const TRACE_BULLET As Long = 8       ' bullet
Public Const TRACE_OBJECT As Long = 16      ' object
Public Const TRACE_HIT_VERT As Long = 32    ' vertical wall was hit
Public Const TRACE_HIT_DOOR As Long = 64    ' door was hit
Public Const TRACE_HIT_PWALL As Long = 128  ' pushwall was hit

Public Type TTracePoint
    x As Double
    y As Double
    angle As Long
    flags As Long
    oob As Boolean
    tileX As Long
    tileY As Long
    frac As Double
End Type

Public VisibleTiles(0 To 64, 0 To 64) As Boolean    ' one extra row/column for the 4-tile sprite test
Public Tracers(0 To XRES - 1) As TTracePoint

Private MarkVisible As Boolean

Private Function TraceCheck(ByVal x As Long, ByVal y As Long, ByVal frac As Double, ByVal dfrac As Double, ByVal vert As Boolean, ByVal flip As Boolean, ByRef tp As TTracePoint) As Boolean
    Dim d As Long, offset As Double, flipOfs As Double

    If (TileMap(x, y) And WALL_TILE) <> 0 Then
        If flip Then flipOfs = TILEGLOBAL
        If vert Then
            tp.x = x * 65536# + flipOfs
            tp.y = y * 65536# + frac
            tp.flags = tp.flags Or TRACE_HIT_VERT
        Else
            tp.x = x * 65536# + frac
            tp.y = y * 65536# + flipOfs
            tp.flags = tp.flags And Not TRACE_HIT_VERT
        End If
        tp.tileX = x
        tp.tileY = y
        tp.frac = frac / TILEGLOBAL
        TraceCheck = True                   ' wall, stop tracing
        Exit Function
    End If

    If MarkVisible Then VisibleTiles(x, y) = True   ' this tile is visible

    If (TileMap(x, y) And DOOR_TILE) <> 0 Then
        d = DoorMap(x, y)
        If Doors(d).action <> dr_open Then
            frac = frac + Sar(ToInt32(dfrac), 1)
            If POS2TILE(frac) <> 0 Then Exit Function
            If vert Then
                If Doors(d).action <> dr_closed And Sar(ToInt32(frac), 10) > DOOR_FULLOPEN - Doors_Opened(d) Then
                    Exit Function           ' opened enough
                End If
                tp.x = TILE2POS(x)
                tp.y = y * 65536# + frac
                tp.flags = tp.flags Or TRACE_HIT_VERT
                tp.frac = frac / TILEGLOBAL
            Else
                If Doors(d).action <> dr_closed And Sar(ToInt32(frac), 10) < Doors_Opened(d) Then
                    Exit Function           ' opened enough
                End If
                tp.y = TILE2POS(y)
                tp.x = x * 65536# + frac
                tp.flags = tp.flags And Not TRACE_HIT_VERT
                tp.frac = 1 - frac / TILEGLOBAL
            End If
            tp.flags = tp.flags Or TRACE_HIT_DOOR
            tp.tileX = x
            tp.tileY = y
            tp.frac = tp.frac + Doors_Opened(d) / DOOR_FULLOPEN
            TraceCheck = True               ' closed door, stop tracing
            Exit Function
        End If
    End If

    If (TileMap(x, y) And PUSHWALL_TILE) <> 0 Then
        offset = PWall.pointsMoved / 128
        frac = frac + dfrac * offset
        If POS2TILE(frac) <> 0 Then Exit Function
        If flip Then flipOfs = TILEGLOBAL
        If vert Then
            tp.x = x * 65536# + flipOfs + offset * TILEGLOBAL * IIf(flip, -1, 1)
            tp.y = y * 65536# + frac
            tp.flags = tp.flags Or TRACE_HIT_VERT
        Else
            tp.x = x * 65536# + frac
            tp.y = y * 65536# + flipOfs + offset * TILEGLOBAL * IIf(flip, -1, 1)
            tp.flags = tp.flags And Not TRACE_HIT_VERT
        End If
        tp.flags = tp.flags Or TRACE_HIT_PWALL
        tp.tileX = x
        tp.tileY = y
        tp.frac = frac / TILEGLOBAL
        TraceCheck = True
        Exit Function
    End If
    ' no intersection, go on
End Function

Public Sub Raycaster_Trace(ByRef tp As TTracePoint, ByVal markTiles As Boolean)
    Dim xtilestep As Long, ytilestep As Long, xtile As Long, ytile As Long
    Dim xstep As Double, ystep As Double, xintercept As Double, yintercept As Double
    Dim XmapPos As Long, YmapPos As Long, q As Long, traceCount As Long, edge As Long
    Dim tanv As Double

    MarkVisible = markTiles

    ' setup for ray casting
    q = Math_GetQuadrant(FINE2RAD(tp.angle))
    Select Case q
        Case q_first: xtilestep = 1: ytilestep = 1
        Case q_second: xtilestep = -1: ytilestep = 1
        Case q_third: xtilestep = -1: ytilestep = -1
        Case Else: xtilestep = 1: ytilestep = -1
    End Select

    xtile = POS2TILE(tp.x) + xtilestep
    ytile = POS2TILE(tp.y) + ytilestep

    xstep = ytilestep * XnextTable(tp.angle)
    ystep = xtilestep * YnextTable(tp.angle)

    tanv = TanTable(tp.angle)
    If ytilestep = -1 Then edge = ytile + 1 Else edge = ytile
    If tanv = 0 Then
        xintercept = tp.x                   ' JS: (n / 0) >> 0 is 0
    Else
        xintercept = ToInt32((edge * 65536# - tp.y) / tanv) + tp.x
    End If
    If xtilestep = -1 Then edge = xtile + 1 Else edge = xtile
    yintercept = ToInt32((edge * 65536# - tp.x) * tanv) + tp.y

    YmapPos = Sar(ToInt32(yintercept), TILESHIFT)   ' toXray
    XmapPos = Sar(ToInt32(xintercept), TILESHIFT)   ' toYray

    If markTiles Then VisibleTiles(POS2TILE(tp.x), POS2TILE(tp.y)) = True

    Do
        traceCount = traceCount + 1

        ' vertical loop: an analogue for X-Ray
        Do While Not (ytilestep = -1 And YmapPos <= ytile) And Not (ytilestep = 1 And YmapPos >= ytile)
            If xtile < 0 Or xtile >= 64 Or YmapPos < 0 Or YmapPos >= 64 Then
                tp.oob = True
                Exit Sub
            End If
            If TraceCheck(xtile, YmapPos, JsMod(yintercept, TILEGLOBAL), ystep, True, (xtilestep = -1), tp) Then
                If xstep < 0 Then tp.frac = 1 - tp.frac
                Exit Sub
            End If
            ' prepare for next step
            xtile = xtile + xtilestep
            yintercept = yintercept + ystep
            YmapPos = Sar(ToInt32(yintercept), TILESHIFT)
        Loop

        ' horizontal loop: an analogue for Y-Ray
        Do While Not (xtilestep = -1 And XmapPos <= xtile) And Not (xtilestep = 1 And XmapPos >= xtile)
            If ytile < 0 Or ytile >= 64 Or XmapPos < 0 Or XmapPos >= 64 Then
                tp.oob = True
                Exit Sub
            End If
            If TraceCheck(XmapPos, ytile, JsMod(xintercept, TILEGLOBAL), xstep, False, (ytilestep = -1), tp) Then
                If ystep > 0 Then tp.frac = 1 - tp.frac
                Exit Sub
            End If
            ' prepare for next step
            ytile = ytile + ytilestep
            xintercept = xintercept + xstep
            XmapPos = Sar(ToInt32(xintercept), TILESHIFT)
        Loop

        If traceCount > 1000 Then Exit Sub
    Loop
End Sub

' Casts one ray per screen column and records which tiles were seen.
Public Sub Raycaster_TraceRays(ByVal viewX As Long, ByVal viewY As Long, ByVal viewAngle As Long)
    Dim n As Long, i As Long, j As Long
    For i = 0 To 64
        For j = 0 To 64
            VisibleTiles(i, j) = False
        Next
    Next
    For n = 0 To XRES - 1
        With Tracers(n)
            .x = viewX
            .y = viewY
            .angle = Math_NormalizeAngle(viewAngle - ColumnAngle(n))
            .flags = TRACE_SIGHT Or TRACE_MARK_MAP
            .oob = False
        End With
        Raycaster_Trace Tracers(n), True
        ' hack against the "blank slice" glitch from out-of-bounds rays: reuse the previous slice
        If Tracers(n).oob Then
            If n > 0 Then
                If Not Tracers(n - 1).oob Then Tracers(n) = Tracers(n - 1)
            End If
        End If
    Next
End Sub
