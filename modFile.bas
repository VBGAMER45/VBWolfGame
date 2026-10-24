Attribute VB_Name = "modFile"
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
' Replaces js/file.js. The browser version decoded base64 strings; here files
' are read from disk into a byte array and read little-endian from a position.

Public Function File_Load(ByVal path As String, ByRef data() As Byte) As Boolean
    Dim f As Integer, size As Long
    On Error GoTo Failed
    f = FreeFile
    Open path For Binary Access Read As #f
    size = LOF(f)
    If size > 0 Then
        ReDim data(0 To size - 1)
        Get #f, , data
    End If
    Close #f
    File_Load = (size > 0)
    Exit Function
Failed:
    File_Load = False
End Function

Public Function File_ReadUInt8(ByRef data() As Byte, ByRef pos As Long) As Long
    File_ReadUInt8 = data(pos)
    pos = pos + 1
End Function

Public Function File_ReadUInt16(ByRef data() As Byte, ByRef pos As Long) As Long
    File_ReadUInt16 = data(pos) + data(pos + 1) * 256&
    pos = pos + 2
End Function

Public Function File_ReadUInt32(ByRef data() As Byte, ByRef pos As Long) As Long
    Dim v As Double
    v = data(pos) + data(pos + 1) * 256# + data(pos + 2) * 65536# + data(pos + 3) * 16777216#
    File_ReadUInt32 = ToInt32(v)
    pos = pos + 4
End Function

' Reads a fixed-length string, stopping at the first NUL.
Public Function File_ReadString(ByRef data() As Byte, ByRef pos As Long, ByVal length As Long) As String
    Dim i As Long, s As String
    For i = 0 To length - 1
        If data(pos + i) = 0 Then Exit For
        s = s & Chr$(data(pos + i))
    Next
    pos = pos + length
    File_ReadString = s
End Function
