Attribute VB_Name = "modEpisodes"
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
' Replaces js/episodes.js, which listed the original episodes. The level
' order is read from a text file (maps\episode.txt) instead:
'   @episode <name>
'   <map file> [secret]          one per line, in play order
' A secret level is reached only through a secret elevator; afterwards play
' continues with the level after the one it was entered from.

Public EpisodeName As String
Public LevelFiles() As String
Public NumLevels As Long
Public SecretLevel As Long          ' index of the secret level, or -1
Public SecretFrom As Long           ' level the secret elevator was used on

Public Function Episodes_Load(ByVal path As String) As String
    Dim f As Integer, content As String, lines() As String, i As Long, ln As String, folder As String
    On Error GoTo Failed
    f = FreeFile
    Open path For Binary Access Read As #f
    content = Space$(LOF(f))
    Get #f, , content
    Close #f
    On Error GoTo 0

    folder = Left$(path, InStrRev(path, "\"))
    EpisodeName = ""
    NumLevels = 0
    SecretLevel = -1
    ReDim LevelFiles(0 To 63)
    lines = Split(Replace(content, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If ln = "" Or Left$(ln, 1) = ";" Then
            ' blank or comment
        ElseIf LCase$(Left$(ln, 9)) = "@episode " Then
            EpisodeName = Trim$(Mid$(ln, 10))
        ElseIf NumLevels <= UBound(LevelFiles) Then
            If LCase$(Right$(ln, 7)) = " secret" Then
                SecretLevel = NumLevels
                ln = Trim$(Left$(ln, Len(ln) - 7))
            End If
            LevelFiles(NumLevels) = folder & ln
            NumLevels = NumLevels + 1
        End If
    Next
    If NumLevels = 0 Then Episodes_Load = "No levels are listed in " & path
    Exit Function
Failed:
    Episodes_Load = "Could not read " & path
End Function

' A one-level episode, for playing a map given on the command line.
Public Sub Episodes_SetSingle(ByVal mapPath As String)
    EpisodeName = "Test map"
    ReDim LevelFiles(0 To 0)
    LevelFiles(0) = mapPath
    NumLevels = 1
    SecretLevel = -1
End Sub

' Level to play after finishing level num, or NumLevels when the episode is over.
Public Function Episodes_NextLevel(ByVal num As Long, ByVal tookSecretExit As Boolean) As Long
    If tookSecretExit And SecretLevel >= 0 Then
        SecretFrom = num
        Episodes_NextLevel = SecretLevel
    ElseIf num = SecretLevel Then
        Episodes_NextLevel = SecretFrom + 1
    Else
        Episodes_NextLevel = num + 1
        If Episodes_NextLevel = SecretLevel Then Episodes_NextLevel = Episodes_NextLevel + 1
    End If
End Function
