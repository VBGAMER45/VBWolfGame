Attribute VB_Name = "modRenderer"
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
' Replaces js/renderer.js, which positioned HTML image strips and divs. The
' projection (processTrace, drawWall, drawDoor, drawSprites) follows the JS
' code; drawing goes into a 32-bit pixel buffer that GDI scales to the window.
' The status bar, HUD text and the text menus are new.

Public Const SCR_W As Long = XRES
Public Const SCR_H As Long = 200
Public Const STATUS_Y As Long = YRES        ' status bar occupies rows YRES..SCR_H-1
Public Const WIN_SCALE As Long = 3
Public Const FOV_RAD As Double = 75 * PI / 180

Public Buf() As Long
Private ViewDist As Double                  ' (XRES / 2) / tan(FOV / 2)
Private ZBuf(0 To XRES - 1) As Double
Private SortIdx(0 To MAX_VIS_SPRITES) As Long
Private BMI As BITMAPINFO
Private MemDC As Long
Private MemBmp As Long
Private OldBmp As Long
Private FontSmall As Long
Private FontMid As Long
Private FontBig As Long
Private OldFont As Long
Private FontTitle As Long
Private MenuBg() As Long                    ' backdrop for the title screens

' results of ProcessTrace
Private pDist As Double
Private pH As Long
Private pFrac As Double

Public Sub Renderer_Init(ByVal hdc As Long)
    ViewDist = (XRES / 2) / Tan(FOV_RAD / 2)
    ReDim Buf(0 To SCR_W * SCR_H - 1)
    With BMI.bmiHeader
        .biSize = Len(BMI.bmiHeader)
        .biWidth = SCR_W
        .biHeight = -SCR_H                  ' top-down rows
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
    End With
    MemDC = CreateCompatibleDC(hdc)
    MemBmp = CreateCompatibleBitmap(hdc, SCR_W * WIN_SCALE, SCR_H * WIN_SCALE)
    OldBmp = SelectObject(MemDC, MemBmp)
    SetStretchBltMode MemDC, COLORONCOLOR
    SetBkMode MemDC, TRANSPARENT
    FontSmall = CreateFont(15, 0, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 0, "Consolas")
    FontMid = CreateFont(26, 0, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 0, "Consolas")
    FontBig = CreateFont(48, 0, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 0, "Consolas")
    FontTitle = CreateFont(96, 0, 0, 0, FW_BOLD, 0, 0, 0, 0, 0, 0, 0, 0, "Consolas")
    BuildMenuBackground
    OldFont = SelectObject(MemDC, FontMid)
End Sub

Public Sub Renderer_Shutdown()
    If MemDC = 0 Then Exit Sub
    SelectObject MemDC, OldFont
    SelectObject MemDC, OldBmp
    DeleteObject FontSmall
    DeleteObject FontMid
    DeleteObject FontBig
    DeleteObject FontTitle
    DeleteObject MemBmp
    DeleteDC MemDC
    MemDC = 0
End Sub

Public Sub Renderer_DrawFrame(ByVal hdc As Long)
    Dim n As Long
    If LevelActive Then
        ClearView
        Raycaster_TraceRays Player.x, Player.y, Player.angle
        For n = 0 To XRES - 1
            ZBuf(n) = 1E+30
            If Not Tracers(n).oob Then
                If (Tracers(n).flags And TRACE_HIT_DOOR) <> 0 Then
                    DrawDoor n
                Else
                    DrawWall n
                End If
            End If
        Next
        DrawSprites
        If Player.playstate <> ex_dead Then DrawWeapon
        If BonusFlashTics > 0 Then TintView RGB32(255, 220, 80), 0.25
        If DamageFlashTics > 0 Then TintView RGB32(255, 0, 0), 0.4 * DamageFlashTics / 21
        If DeathFade > 0 Then TintView RGB32(160, 0, 0), DeathFade
        If InIntermission Or GameOver Or EpisodeDone Then TintView 0, 0.7
        If CurScreen <> SCR_PLAYING Then TintView 0, 0.6
        DrawStatusPanel
    Else
        For n = 0 To SCR_W * SCR_H - 1
            Buf(n) = MenuBg(n)
        Next
    End If
    StretchDIBits MemDC, 0, 0, SCR_W * WIN_SCALE, SCR_H * WIN_SCALE, 0, 0, SCR_W, SCR_H, Buf(0), BMI, DIB_RGB_COLORS, SRCCOPY
    If LevelActive Then DrawHudText
    If CurScreen <> SCR_PLAYING Then DrawMenu
    BitBlt hdc, 0, 0, SCR_W * WIN_SCALE, SCR_H * WIN_SCALE, MemDC, 0, 0, SRCCOPY
End Sub

' Dark metal panels with a vignette and a red rule under the title.
Private Sub BuildMenuBackground()
    Dim x As Long, y As Long, c As Long, f As Double, dx As Double, dy As Double
    ReDim MenuBg(0 To SCR_W * SCR_H - 1)
    For y = 0 To SCR_H - 1
        For x = 0 To SCR_W - 1
            c = TexPix(14 * TEX_PIXELS + (x And 63) * TEX_SIZE + (y And 63))
            dx = (x - SCR_W / 2) / (SCR_W / 2): dy = (y - SCR_H / 2) / (SCR_H / 2)
            f = 0.45 - 0.25 * (dx * dx + dy * dy)
            If f < 0.08 Then f = 0.08
            MenuBg(y * SCR_W + x) = Shade(c, f)
        Next
    Next
    For x = 60 To SCR_W - 61
        MenuBg(50 * SCR_W + x) = RGB32(170, 30, 30)
        MenuBg(51 * SCR_W + x) = RGB32(90, 15, 15)
    Next
End Sub

Private Sub DrawMenu()
    Dim items As Variant, i As Long, y0 As Long, keys As Variant, acts As Variant
    items = Game_MenuItems()
    Select Case CurScreen
        Case SCR_TITLE
            HudText XRES \ 2, 8, "VBWOLF", 3, True, RGB(230, 60, 50)
            HudText XRES \ 2, 58, "EPISODE:  " & UCase$(EpisodeName), 0, True
            If PackFolder <> "" Then HudText XRES \ 2, 67, "CONTENT PACK:  " & UCase$(PackName), 0, True
            y0 = 84
        Case SCR_PACK
            HudText XRES \ 2, 10, "CONTENT PACK", 2, True
            HudText XRES \ 2, 58, PackName, 1, True, RGB(255, 210, 60)
            If PackFolder <> "" Then
                HudText XRES \ 2, 76, PackSummary, 0, True
                HudText XRES \ 2, 86, PackFolder, 0, True
            Else
                HudText XRES \ 2, 76, "Choose a folder of your own art, sounds or levels.", 0, True
                HudText XRES \ 2, 86, "PACKS.md describes the folder layout.", 0, True
            End If
            y0 = 112
        Case SCR_SKILL
            HudText XRES \ 2, 30, "CHOOSE YOUR SKILL", 2, True
            y0 = 74
        Case SCR_PAUSE
            HudText XRES \ 2, 40, "PAUSED", 2, True
            y0 = 84
        Case Else
            HudText XRES \ 2, 10, "CONTROLS", 2, True
            keys = Array("Up/Down  W/S", "Left/Right", "Alt+arrows  A/D", "Shift", "Ctrl", "Space", "1  2  3  4", "Esc")
            acts = Array("move", "turn", "strafe", "run", "fire", "open, push, use", "choose weapon", "pause")
            For i = 0 To UBound(keys)
                HudText 62, 42 + i * 15, keys(i), 1, False, RGB(255, 210, 60)
                HudText 182, 42 + i * 15, acts(i), 1, False
            Next
            y0 = 170
    End Select
    For i = 0 To UBound(items)
        If i = MenuIndex Then
            HudText XRES \ 2, y0 + i * 18, "> " & items(i) & " <", 1, True, RGB(255, 210, 60)
        Else
            HudText XRES \ 2, y0 + i * 18, items(i), 1, True
        End If
    Next
    If CurScreen = SCR_TITLE Then
        HudText XRES \ 2, 180, "A Visual Basic 6 port of the Wolf3D Browser GPL source", 0, True
        HudText XRES \ 2, 189, "Original art, sound and levels", 0, True
    End If
End Sub

Private Sub ClearView()
    Dim c As Long, f As Long, i As Long, half As Long
    c = RGB32(LevelCeiling(0), LevelCeiling(1), LevelCeiling(2))
    f = RGB32(LevelFloor(0), LevelFloor(1), LevelFloor(2))
    half = (YRES \ 2) * SCR_W
    For i = 0 To half - 1
        Buf(i) = c
    Next
    For i = half To YRES * SCR_W - 1
        Buf(i) = f
    Next
End Sub

'---------------------------------------------------------------- walls (renderer.js)

Private Sub ProcessTrace(ByRef tp As TTracePoint)
    Dim dx As Double, dy As Double
    dx = Player.x - tp.x
    dy = Player.y - tp.y
    pDist = Sqr(dx * dx + dy * dy)
    pDist = pDist * Cos(FINE2RAD(tp.angle - Player.angle))     ' correct for fisheye
    If pDist < 1 Then pDist = 1
    pH = ToInt32(ViewDist / pDist * TILEGLOBAL)

    If (tp.flags And TRACE_HIT_DOOR) <> 0 Then
        If (tp.flags And TRACE_HIT_VERT) <> 0 Then
            If tp.x < Player.x Then pFrac = tp.frac Else pFrac = 1 - tp.frac
        Else
            If tp.y < Player.y Then pFrac = 1 - tp.frac Else pFrac = tp.frac
        End If
    Else
        pFrac = 1 - tp.frac
    End If
End Sub

Private Sub DrawWall(ByVal n As Long)
    Dim x As Long, y As Long, vx As Long, vy As Long, texture As Long
    ProcessTrace Tracers(n)
    x = Tracers(n).tileX: y = Tracers(n).tileY
    vx = POS2TILE(Player.x): vy = POS2TILE(Player.y)
    If (Tracers(n).flags And TRACE_HIT_VERT) <> 0 Then
        texture = WallTexX(x, y)
    Else
        texture = WallTexY(x, y)
    End If

    ' door sides
    If (Tracers(n).flags And TRACE_HIT_VERT) <> 0 Then
        If x >= vx And x > 0 Then
            If (TileMap(x - 1, y) And DOOR_TILE) <> 0 Then texture = TEX_PLATE
        End If
        If x < vx And x < 63 Then
            If (TileMap(x + 1, y) And DOOR_TILE) <> 0 Then texture = TEX_PLATE
        End If
    Else
        If y >= vy And y > 0 Then
            If (TileMap(x, y - 1) And DOOR_TILE) <> 0 Then texture = TEX_PLATE
        End If
        If y < vy And y < 63 Then
            If (TileMap(x, y + 1) And DOOR_TILE) <> 0 Then texture = TEX_PLATE
        End If
    End If
    DrawColumn n, texture
End Sub

Private Sub DrawDoor(ByVal n As Long)
    ProcessTrace Tracers(n)
    DrawColumn n, Doors(DoorMap(Tracers(n).tileX, Tracers(n).tileY)).texture
End Sub

Private Sub DrawColumn(ByVal n As Long, ByVal texture As Long)
    Dim texX As Long, h As Long, top As Double, y0 As Long, y1 As Long, y As Long
    Dim ts As Long, tpos As Long, colBase As Long, p As Long

    If texture < 0 Or texture >= NUM_TEX_SLOTS Then texture = NUM_TEX_SLOTS - 1
    texX = Int(pFrac * TEX_SIZE)
    If texX < 0 Then texX = 0
    If texX > TEX_SIZE - 1 Then texX = TEX_SIZE - 1
    ZBuf(n) = pDist

    h = pH
    If h < 1 Then h = 1
    top = (YRES - h) / 2
    y0 = Int(top + 0.5)
    y1 = Int(top + h - 0.5)
    If y0 < 0 Then y0 = 0
    If y1 > YRES - 1 Then y1 = YRES - 1
    If y0 > y1 Then Exit Sub

    ts = ToInt32(4194304# / h)                          ' 64 texels per wall height, 16.16 fixed point
    tpos = ToInt32((y0 + 0.5 - top) * 4194304# / h)
    colBase = texture * TEX_PIXELS + texX * TEX_SIZE
    p = y0 * SCR_W + n
    For y = y0 To y1
        Buf(p) = TexPix(colBase + ((tpos \ 65536) And 63))
        tpos = tpos + ts
        p = p + SCR_W
    Next
End Sub

'---------------------------------------------------------------- sprites (renderer.js drawSprites)

Private Sub DrawSprites()
    Dim i As Long, j As Long, k As Long, angle As Double, size As Long
    Dim dx As Double, dy As Double, leftX As Double, topY As Double

    Sprites_CreateVisList Player.x, Player.y, Player.angle
    For i = 0 To NumVisible - 1
        SortIdx(i) = i
    Next
    ' far to near, so nearer sprites paint over farther ones
    For i = 1 To NumVisible - 1
        k = SortIdx(i): j = i - 1
        Do While j >= 0
            If VisList(SortIdx(j)).dist >= VisList(k).dist Then Exit Do
            SortIdx(j + 1) = SortIdx(j)
            j = j - 1
        Loop
        SortIdx(j + 1) = k
    Next

    For i = 0 To NumVisible - 1
        With VisList(SortIdx(i))
            If .dist > 0 Then
                dx = .x - Player.x
                dy = .y - Player.y
                angle = Atan2(dy, dx) - FINE2RAD(Player.angle)
                Do While angle > PI
                    angle = angle - 2 * PI
                Loop
                Do While angle <= -PI
                    angle = angle + 2 * PI
                Loop
                If Abs(angle) < PI / 2 Then
                    size = ToInt32(ViewDist / .dist * TILEGLOBAL)
                    If size > 4000 Then size = 4000
                    leftX = XRES / 2 - size / 2 - Tan(angle) * ViewDist
                    topY = YRES / 2 - size / 2
                    If size > 0 Then DrawSpriteImage .tex, leftX, topY, size, .dist
                End If
            End If
        End With
    Next
End Sub

Private Sub DrawSpriteImage(ByVal tex As Long, ByVal leftX As Double, ByVal topY As Double, ByVal size As Long, ByVal dist As Double)
    Dim x0 As Long, x1 As Long, col As Long, texX As Long, colIdx As Long
    Dim ys As Long, ye As Long, y As Long, p As Long, c As Long, colBase As Long
    Dim texScale As Double, ty As Long

    If tex < 0 Or tex >= NUM_SPR_SLOTS Then Exit Sub   ' actor sprites arrive with the actor port
    texScale = TEX_SIZE / size
    x0 = Int(leftX + 0.5)
    x1 = Int(leftX + size - 0.5)
    If x0 < 0 Then x0 = 0
    If x1 > XRES - 1 Then x1 = XRES - 1
    For col = x0 To x1
        If dist < ZBuf(col) Then
            texX = Int((col + 0.5 - leftX) * texScale)
            If texX < 0 Then texX = 0
            If texX > TEX_SIZE - 1 Then texX = TEX_SIZE - 1
            colIdx = tex * TEX_SIZE + texX
            If SprTop(colIdx) <= SprBot(colIdx) Then
                ys = Int(topY + SprTop(colIdx) / texScale)
                ye = Int(topY + (SprBot(colIdx) + 1) / texScale)
                If ys < 0 Then ys = 0
                If ye > YRES - 1 Then ye = YRES - 1
                colBase = tex * TEX_PIXELS + texX * TEX_SIZE
                p = ys * SCR_W + col
                For y = ys To ye
                    ty = Int((y + 0.5 - topY) * texScale)
                    If ty >= 0 And ty < TEX_SIZE Then
                        c = SprPix(colBase + ty)
                        If c <> CLEAR_PIXEL Then Buf(p) = c
                    End If
                    p = p + SCR_W
                Next
            End If
        End If
    Next
End Sub

'---------------------------------------------------------------- weapon, flashes, status bar

Private Sub DrawWeapon()
    Dim s As Long, tx As Long, ty As Long, c As Long, sx As Long, sy As Long
    Dim left0 As Long, top0 As Long, colIdx As Long, dx As Long, dy As Long
    s = Player.weapon * 4 + Player.weaponFrame
    If s < 0 Or s >= NUM_WEAPON_VIEWS Then Exit Sub
    left0 = XRES \ 2 - TEX_SIZE
    top0 = YRES - TEX_SIZE * 2 + 8
    For tx = 0 To TEX_SIZE - 1
        colIdx = s * TEX_SIZE + tx
        For ty = WeapTop(colIdx) To WeapBot(colIdx)
            c = WeapPix(s * TEX_PIXELS + tx * TEX_SIZE + ty)
            If c <> CLEAR_PIXEL Then
                For dx = 0 To 1
                    For dy = 0 To 1
                        sx = left0 + tx * 2 + dx: sy = top0 + ty * 2 + dy
                        If sx >= 0 And sx < XRES And sy >= 0 And sy < YRES Then Buf(sy * SCR_W + sx) = c
                    Next
                Next
            End If
        Next
    Next
End Sub

' Blends the 3D view toward a colour by the given amount (0..1).
Private Sub TintView(ByVal color As Long, ByVal amount As Double)
    Dim i As Long, c As Long, r As Long, g As Long, b As Long, keep As Double
    Dim tr As Double, tg As Double, tb As Double
    keep = 1 - amount
    tr = ((color \ 65536) And 255) * amount
    tg = ((color \ 256) And 255) * amount
    tb = (color And 255) * amount
    For i = 0 To YRES * SCR_W - 1
        c = Buf(i)
        r = ((c \ 65536) And 255) * keep + tr
        g = ((c \ 256) And 255) * keep + tg
        b = (c And 255) * keep + tb
        Buf(i) = r * 65536 + g * 256 + b
    Next
End Sub

' Status bar columns, in 320x200 units: FLOOR, SCORE, LIVES, HEALTH, AMMO, KEYS, WEAPON.
Private Sub DrawStatusPanel()
    Dim x As Long, y As Long, c As Long, i As Long, edges As Variant
    For y = STATUS_Y To SCR_H - 1
        If y = STATUS_Y Then
            c = RGB32(110, 120, 150)
        ElseIf y = STATUS_Y + 1 Then
            c = RGB32(20, 24, 40)
        Else
            c = RGB32(36, 44, 76)
        End If
        For x = 0 To SCR_W - 1
            Buf(y * SCR_W + x) = c
        Next
    Next
    edges = Array(36, 106, 140, 196, 238, 270)
    For i = 0 To UBound(edges)
        For y = STATUS_Y + 4 To SCR_H - 4
            Buf(y * SCR_W + edges(i)) = RGB32(20, 24, 40)
            Buf(y * SCR_W + edges(i) + 1) = RGB32(70, 80, 115)
        Next
    Next
    If (Player.items And ITEM_KEY_1) <> 0 Then KeyIcon 246, RGB32(230, 185, 50)
    If (Player.items And ITEM_KEY_2) <> 0 Then KeyIcon 257, RGB32(200, 205, 215)
End Sub

Private Sub KeyIcon(ByVal x0 As Long, ByVal c As Long)
    Dim x As Long, y As Long
    For x = x0 To x0 + 7
        For y = 176 To 192
            If y < 182 Or (x >= x0 + 3 And x <= x0 + 4) Or (y >= 187 And y <= 188 And x >= x0 + 5) Then
                Buf(y * SCR_W + x) = c
            End If
        Next
    Next
    For x = x0 + 2 To x0 + 5
        For y = 178 To 179
            Buf(y * SCR_W + x) = RGB32(36, 44, 76)
        Next
    Next
End Sub

Private Sub DrawHudText()
    Dim weaponName As String, i As Long
    HudText 18, 164, "FLOOR", 0, True
    HudText 71, 164, "SCORE", 0, True
    HudText 123, 164, "LIVES", 0, True
    HudText 168, 164, "HEALTH", 0, True
    HudText 217, 164, "AMMO", 0, True
    HudText 254, 164, "KEYS", 0, True
    HudText 295, 164, "WEAPON", 0, True
    HudText 18, 176, CStr(LevelNum + 1), 1, True
    HudText 71, 176, CStr(Player.score), 1, True
    HudText 123, 176, CStr(Player.lives), 1, True
    HudText 168, 176, Player.health & "%", 1, True
    HudText 217, 176, CStr(Player.ammo(AMMO_BULLETS)), 1, True
    Select Case Player.weapon
        Case WEAPON_KNIFE: weaponName = "KNIFE"
        Case WEAPON_PISTOL: weaponName = "PISTOL"
        Case WEAPON_AUTO: weaponName = "AUTO"
        Case Else: weaponName = "ROTARY"
    End Select
    HudText 295, 178, weaponName, 0, True

    If NotifyTics > 0 And Not InIntermission And Not GameOver And Not EpisodeDone Then HudText XRES \ 2, 4, NotifyText, 1, True
    If GameOver Then
        HudText XRES \ 2, 50, "GAME OVER", 2, True
        HudText XRES \ 2, 100, "Press ENTER to return to the title", 1, True
    End If
    If InIntermission Or EpisodeDone Then
        HudText XRES \ 2, 6, IntermissionText(0), 2, True
        HudText XRES \ 2, 30, IntermissionText(1), 1, True, RGB(255, 210, 60)
        For i = 2 To 6
            HudText XRES \ 2, 50 + (i - 2) * 16, IntermissionText(i), 1, True
        Next
        HudText XRES \ 2, 140, IntermissionText(7), 1, True
    End If
End Sub

' Text at 320x200 coordinates with a drop shadow. size: 0 small, 1 medium,
' 2 big, 3 title. color: a VB RGB() value, or -1 for the default.
Private Sub HudText(ByVal x As Long, ByVal y As Long, ByVal s As String, ByVal size As Long, ByVal centered As Boolean, Optional ByVal color As Long = -1)
    Select Case size
        Case 0: SelectObject MemDC, FontSmall
        Case 1: SelectObject MemDC, FontMid
        Case 2: SelectObject MemDC, FontBig
        Case Else: SelectObject MemDC, FontTitle
    End Select
    If centered Then SetTextAlign MemDC, TA_CENTER Else SetTextAlign MemDC, TA_LEFT
    SetTextColor MemDC, 0
    TextOut MemDC, x * WIN_SCALE + 2, y * WIN_SCALE + 2, s, Len(s)
    If color <> -1 Then
        SetTextColor MemDC, color
    ElseIf size = 0 Then
        SetTextColor MemDC, RGB(170, 185, 225)
    Else
        SetTextColor MemDC, RGB(240, 240, 240)
    End If
    TextOut MemDC, x * WIN_SCALE, y * WIN_SCALE, s, Len(s)
End Sub

' Rebuilds the title backdrop from the current wall textures (after a pack change).
Public Sub Renderer_RefreshBackground()
    BuildMenuBackground
End Sub
