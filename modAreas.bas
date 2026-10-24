Attribute VB_Name = "modAreas"
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
' Ported from js/areas.js.
'
' Open doors connect two areas, so sounds travel between them and sight is
' checked when the player is in a connected area. AreaConnect is incremented
' and decremented by each door; above zero means connected. AreaByPlayer is
' recalculated every time a door opens or closes and is true for every area
' that connects with the player's current one.

Public AreaConnect(0 To NUMAREAS - 1, 0 To NUMAREAS - 1) As Long
Public AreaByPlayer(0 To NUMAREAS - 1) As Boolean

Public Sub Areas_Init(ByVal areanumber As Long)
    Dim i As Long, j As Long
    For i = 0 To NUMAREAS - 1
        For j = 0 To NUMAREAS - 1
            AreaConnect(i, j) = 0
        Next
        AreaByPlayer(i) = False
    Next
    AreaByPlayer(areanumber) = True
End Sub

Private Sub RecursiveConnect(ByVal areanumber As Long)
    Dim i As Long
    For i = 0 To NUMAREAS - 1
        If AreaConnect(areanumber, i) <> 0 And Not AreaByPlayer(i) Then
            AreaByPlayer(i) = True
            RecursiveConnect i
        End If
    Next
End Sub

Public Sub Areas_Connect(ByVal areanumber As Long)
    Dim i As Long
    If areanumber >= NUMAREAS Then Err.Raise vbObjectError + 10, "Areas_Connect", "areanumber >= NUMAREAS"
    For i = 0 To NUMAREAS - 1
        AreaByPlayer(i) = False
    Next
    AreaByPlayer(areanumber) = True
    RecursiveConnect areanumber
End Sub

Public Sub Areas_Join(ByVal area1 As Long, ByVal area2 As Long)
    If area1 < 0 Or area1 >= NUMAREAS Then Err.Raise vbObjectError + 11, "Areas_Join", "area1 < 0 || area1 >= NUMAREAS"
    If area2 < 0 Or area2 >= NUMAREAS Then Err.Raise vbObjectError + 12, "Areas_Join", "area2 < 0 || area2 >= NUMAREAS"
    AreaConnect(area1, area2) = AreaConnect(area1, area2) + 1
    AreaConnect(area2, area1) = AreaConnect(area2, area1) + 1
End Sub

Public Sub Areas_Disconnect(ByVal area1 As Long, ByVal area2 As Long)
    If area1 < 0 Or area1 >= NUMAREAS Then Err.Raise vbObjectError + 13, "Areas_Disconnect", "area1 < 0 || area1 >= NUMAREAS"
    If area2 < 0 Or area2 >= NUMAREAS Then Err.Raise vbObjectError + 14, "Areas_Disconnect", "area2 < 0 || area2 >= NUMAREAS"
    AreaConnect(area1, area2) = AreaConnect(area1, area2) - 1
    AreaConnect(area2, area1) = AreaConnect(area2, area1) - 1
End Sub
