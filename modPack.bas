Attribute VB_Name = "modPack"
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
' New for the VB6 port: content packs. A pack is a folder of replacement
' content laid over the built-in art, sound and levels, one slot at a time
' (see PACKS.md):
'   walls\<slot>.png|bmp        wall and door textures
'   sprites\<slot>.png|bmp      objects, pickups and enemies
'   weapons\<weapon>_<frame>.png|bmp
'   sounds\<effect>.wav
'   maps\episode.txt            a level set, replacing the built-in episode
' Images may be any size (they are scaled to 64 x 64); transparency comes
' from the alpha channel or pure magenta (255, 0, 255). The chosen folder is
' remembered in VBWolf.ini.

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, height As Long) As Long
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, color As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const SOUND_NAMES As String = "pistol auto rotary knife enemy_shot heavy_shot bite door_open door_close pushwall ammo health treasure key weapon robot_death hound_death player_death elevator alert ricochet gibs"

Public PackFolder As String         ' "" = built-in content only
Public PackName As String
Public PackSummary As String

'---------------------------------------------------------------- settings

Private Function IniPath() As String
    IniPath = App.Path & "\VBWolf.ini"
End Function

Public Sub Pack_LoadSetting()
    Dim buf As String, n As Long
    buf = Space$(1024)
    n = GetPrivateProfileString("pack", "folder", "", buf, Len(buf), IniPath())
    PackFolder = Left$(buf, n)
    If PackFolder <> "" Then
        If Not FolderExists(PackFolder) Then PackFolder = ""
    End If
End Sub

Public Sub Pack_SaveSetting()
    WritePrivateProfileString "pack", "folder", PackFolder, IniPath()
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(path) And vbDirectory) = 0
    If Err.Number <> 0 Then FileExists = False
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(path) And vbDirectory) <> 0
    If Err.Number <> 0 Then FolderExists = False
End Function

' Shows the folder picker. Returns the chosen folder, or "" if cancelled.
Public Function Pack_ChooseFolder(ByVal owner As Long) As String
    Dim bi As BROWSEINFO, pidl As Long, path As String
    bi.hOwner = owner
    bi.pszDisplayName = Space$(260)
    bi.lpszTitle = "Choose a VBWolf content pack folder"
    bi.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    pidl = SHBrowseForFolder(bi)
    If pidl = 0 Then Exit Function
    path = Space$(260)
    If SHGetPathFromIDList(pidl, path) <> 0 Then Pack_ChooseFolder = Left$(path, InStr(path, vbNullChar) - 1)
    CoTaskMemFree pidl
End Function

'---------------------------------------------------------------- applying a pack

' Lays the current pack over freshly built art and sounds. Returns the pack's
' episode file, or "" when the built-in levels should be used.
Public Function Pack_Apply() As String
    Dim slot As Long, w As Long, f As Long, i As Long, path As String, pix() As Long
    Dim loaded(0 To NUM_TEX_SLOTS - 1) As Boolean, token As Long, si As GdiplusStartupInput
    Dim nTex As Long, nSpr As Long, nWeap As Long, nSnd As Long, names() As String
    Dim samples() As Integer, count As Long

    PackName = "Built-in content"
    PackSummary = ""
    If PackFolder = "" Then Exit Function
    PackName = ReadPackName()
    ReDim pix(0 To TEX_PIXELS - 1)

    si.GdiplusVersion = 1
    If GdiplusStartup(token, si, 0) = 0 Then
        For slot = 0 To NUM_TEX_SLOTS - 1
            path = FindImage(PackFolder & "\walls\", slot)
            If path <> "" Then
                If LoadImage(path, pix, False) Then
                    Art_SetTexture slot, pix
                    loaded(slot) = True
                    nTex = nTex + 1
                End If
            End If
        Next
        ' a light wall face without its dark partner gets a darkened copy
        For slot = 0 To 102 Step 2
            If slot < 40 Or slot >= 98 Then
                If loaded(slot) And Not loaded(slot + 1) Then Art_DarkCopy slot, slot + 1
            End If
        Next
        For slot = SPR_STAT_0 To NUM_SPR_SLOTS - 1
            path = FindImage(PackFolder & "\sprites\", slot)
            If path <> "" Then
                If LoadImage(path, pix, True) Then
                    Art_SetSprite slot, pix
                    nSpr = nSpr + 1
                End If
            End If
        Next
        For w = 0 To 3
            For f = 0 To 3
                path = PackFolder & "\weapons\" & w & "_" & f
                If FileExists(path & ".png") Then
                    path = path & ".png"
                ElseIf FileExists(path & ".bmp") Then
                    path = path & ".bmp"
                Else
                    path = ""
                End If
                If path <> "" Then
                    If LoadImage(path, pix, True) Then
                        Art_SetWeaponView w * 4 + f, pix
                        nWeap = nWeap + 1
                    End If
                End If
            Next
        Next
        GdiplusShutdown token
        Art_RefreshBounds
    End If

    names = Split(SOUND_NAMES, " ")
    For i = 0 To UBound(names)
        path = PackFolder & "\sounds\" & names(i) & ".wav"
        If FileExists(path) Then
            If LoadWav(path, samples, count) Then
                Sound_ReplaceEffect i, samples, count
                nSnd = nSnd + 1
            End If
        End If
    Next

    PackSummary = Plural(nTex, "texture") & ", " & Plural(nSpr, "sprite") & ", " & Plural(nWeap, "weapon frame") & ", " & Plural(nSnd, "sound")
    If FileExists(PackFolder & "\maps\episode.txt") Then
        Pack_Apply = PackFolder & "\maps\episode.txt"
        PackSummary = PackSummary & ", and its own levels"
    End If
End Function

Private Function Plural(ByVal n As Long, ByVal word As String) As String
    If n = 1 Then Plural = n & " " & word Else Plural = n & " " & word & "s"
End Function

' "@name" from pack.txt, or else the folder's name.
Private Function ReadPackName() As String
    Dim f As Integer, ln As String
    ReadPackName = Mid$(PackFolder, InStrRev(PackFolder, "\") + 1)
    If Not FileExists(PackFolder & "\pack.txt") Then Exit Function
    f = FreeFile
    Open PackFolder & "\pack.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If LCase$(Left$(Trim$(ln), 6)) = "@name " Then ReadPackName = Trim$(Mid$(Trim$(ln), 7))
    Loop
    Close #f
End Function

' <folder><slot>.png or .bmp, also accepting a three-digit slot ("014.png").
Private Function FindImage(ByVal folder As String, ByVal slot As Long) As String
    Dim base As Variant, ext As Variant
    For Each base In Array(CStr(slot), Format$(slot, "000"))
        For Each ext In Array(".png", ".bmp")
            If FileExists(folder & base & ext) Then
                FindImage = folder & base & ext
                Exit Function
            End If
        Next
    Next
End Function

' Reads an image into a 64 x 64 column-major buffer, scaling as needed.
Private Function LoadImage(ByVal path As String, ByRef pix() As Long, ByVal transparent As Boolean) As Boolean
    Dim img As Long, w As Long, h As Long, x As Long, y As Long, argb As Long, c As Long, a As Long
    If GdipLoadImageFromFile(StrPtr(path), img) <> 0 Then Exit Function
    GdipGetImageWidth img, w
    GdipGetImageHeight img, h
    If w > 0 And h > 0 Then
        For x = 0 To TEX_SIZE - 1
            For y = 0 To TEX_SIZE - 1
                GdipBitmapGetPixel img, Int((x + 0.5) * w / TEX_SIZE), Int((y + 0.5) * h / TEX_SIZE), argb
                c = argb And &HFFFFFF
                a = ((argb And &HFF000000) \ &H1000000) And &HFF
                If transparent And (a < 128 Or c = &HFF00FF) Then
                    pix(x * TEX_SIZE + y) = CLEAR_PIXEL
                Else
                    pix(x * TEX_SIZE + y) = c
                End If
            Next
        Next
        LoadImage = True
    End If
    GdipDisposeImage img
End Function

' Reads a PCM .wav (8 or 16 bit, any rate, mono or stereo) and converts it to
' the mixer's 16-bit mono at SOUND_RATE. Up to 5 seconds are kept.
Private Function LoadWav(ByVal path As String, ByRef out() As Integer, ByRef count As Long) As Boolean
    Dim d() As Byte, pos As Long, id As String, size As Long, nextPos As Long
    Dim fmtTag As Long, channels As Long, rate As Long, bits As Long, dataPos As Long, dataLen As Long
    Dim frames As Long, frameBytes As Long, i As Long, srcPos As Double, i0 As Long, t As Double

    If Not File_Load(path, d) Then Exit Function
    If UBound(d) < 44 Then Exit Function
    If Chr$(d(0)) & Chr$(d(1)) & Chr$(d(2)) & Chr$(d(3)) <> "RIFF" Then Exit Function
    If Chr$(d(8)) & Chr$(d(9)) & Chr$(d(10)) & Chr$(d(11)) <> "WAVE" Then Exit Function
    pos = 12
    Do While pos + 8 <= UBound(d) + 1
        id = Chr$(d(pos)) & Chr$(d(pos + 1)) & Chr$(d(pos + 2)) & Chr$(d(pos + 3))
        nextPos = pos + 8
        size = File_ReadUInt32(d, nextPos - 4)
        If id = "fmt " Then
            i = pos + 8
            fmtTag = File_ReadUInt16(d, i)
            channels = File_ReadUInt16(d, i)
            rate = File_ReadUInt32(d, i)
            i = pos + 22
            bits = File_ReadUInt16(d, i)
        ElseIf id = "data" Then
            dataPos = pos + 8
            dataLen = size
            If dataPos + dataLen > UBound(d) + 1 Then dataLen = UBound(d) + 1 - dataPos
        End If
        pos = pos + 8 + size + (size And 1)
    Loop
    If fmtTag <> 1 Or channels < 1 Or rate <= 0 Or (bits <> 8 And bits <> 16) Or dataLen <= 0 Then Exit Function

    frameBytes = channels * bits \ 8
    frames = dataLen \ frameBytes
    count = Int(frames * CDbl(SOUND_RATE) / rate)
    If count > SOUND_RATE * 5 Then count = SOUND_RATE * 5
    If count < 1 Then Exit Function
    ReDim out(0 To count - 1)
    For i = 0 To count - 1
        srcPos = i * CDbl(rate) / SOUND_RATE
        i0 = Int(srcPos)
        t = srcPos - i0
        If i0 >= frames - 1 Then
            out(i) = FrameValue(d, dataPos, frames - 1, channels, bits)
        Else
            out(i) = FrameValue(d, dataPos, i0, channels, bits) * (1 - t) + FrameValue(d, dataPos, i0 + 1, channels, bits) * t
        End If
    Next
    LoadWav = True
End Function

' One frame as a 16-bit value, channels averaged.
Private Function FrameValue(ByRef d() As Byte, ByVal dataPos As Long, ByVal frame As Long, ByVal channels As Long, ByVal bits As Long) As Double
    Dim c As Long, p As Long, v As Long, sum As Double
    For c = 0 To channels - 1
        If bits = 8 Then
            p = dataPos + frame * channels + c
            sum = sum + (CLng(d(p)) - 128) * 256
        Else
            p = dataPos + (frame * channels + c) * 2
            v = d(p) + d(p + 1) * 256&
            If v >= 32768 Then v = v - 65536
            sum = sum + v
        End If
    Next
    FrameValue = sum / channels
End Function
