Attribute VB_Name = "modSound"
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
' Replaces js/sound.js. The browser version played sampled sounds that are
' not part of the GPL release; these effects are original, synthesized at
' startup. Calls keep the browser version's sample names, which map to the
' effects below. A small software mixer feeds the waveOut API, so several
' sounds can play at once. Without an audio device the game stays silent.

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Private Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long

Public Const SOUND_RATE As Long = 22050
Private Const RATE As Long = SOUND_RATE
Private Const BUF_SAMPLES As Long = 512         ' about 23 ms per buffer
Private Const NUM_BUFS As Long = 4
Private Const MAX_VOICES As Long = 8
Private Const WAVE_MAPPER As Long = -1
Private Const WHDR_DONE As Long = 1
Private Const FALLOFF_TILES As Double = 20      ' sounds fade out over this distance

' effects
Private Const SND_PISTOL As Long = 0
Private Const SND_AUTO As Long = 1
Private Const SND_ROTARY As Long = 2
Private Const SND_KNIFE As Long = 3
Private Const SND_ZAP As Long = 4               ' robot shot
Private Const SND_ZAP_HEAVY As Long = 5         ' enforcer shot
Private Const SND_BITE As Long = 6
Private Const SND_DOOR_OPEN As Long = 7
Private Const SND_DOOR_CLOSE As Long = 8
Private Const SND_PUSHWALL As Long = 9
Private Const SND_AMMO As Long = 10
Private Const SND_HEALTH As Long = 11
Private Const SND_TREASURE As Long = 12
Private Const SND_KEY As Long = 13
Private Const SND_WEAPON As Long = 14
Private Const SND_ROBOT_DEATH As Long = 15
Private Const SND_HOUND_DEATH As Long = 16
Private Const SND_PLAYER_DEATH As Long = 17
Private Const SND_ELEVATOR As Long = 18
Private Const SND_ALERT As Long = 19
Private Const SND_RICOCHET As Long = 20
Private Const SND_GIBS As Long = 21
Private Const NUM_SND As Long = 22

Private hWave As Long
Private Hdr(0 To NUM_BUFS - 1) As WAVEHDR
Private Mem(0 To NUM_BUFS * BUF_SAMPLES - 1) As Integer     ' fixed array: stays put while queued
Private Queued(0 To NUM_BUFS - 1) As Boolean
Private MixAcc(0 To BUF_SAMPLES - 1) As Double
Private Ready As Boolean

Private SndData() As Integer
Private SndStart(0 To NUM_SND - 1) As Long
Private SndLen(0 To NUM_SND - 1) As Long
Private DataLen As Long

Private VoiceSnd(0 To MAX_VOICES - 1) As Long   ' effect playing, -1 = free
Private VoicePos(0 To MAX_VOICES - 1) As Long
Private VoiceVol(0 To MAX_VOICES - 1) As Double
Private VoiceAge(0 To MAX_VOICES - 1) As Long
Private AgeCounter As Long

Private Gen() As Double                         ' effect being synthesized
Private GenLen As Long
Private SeedS As Double

Public Sub Sound_Init()
    Dim fmt As WAVEFORMATEX, i As Long
    BuildEffects
    For i = 0 To MAX_VOICES - 1
        VoiceSnd(i) = -1
    Next
    With fmt
        .wFormatTag = 1                         ' PCM
        .nChannels = 1
        .nSamplesPerSec = RATE
        .wBitsPerSample = 16
        .nBlockAlign = 2
        .nAvgBytesPerSec = RATE * 2
    End With
    If waveOutOpen(hWave, WAVE_MAPPER, fmt, 0, 0, 0) <> 0 Then
        hWave = 0
        Exit Sub
    End If
    For i = 0 To NUM_BUFS - 1
        Hdr(i).lpData = VarPtr(Mem(i * BUF_SAMPLES))
        Hdr(i).dwBufferLength = BUF_SAMPLES * 2
        Hdr(i).dwFlags = 0
        waveOutPrepareHeader hWave, Hdr(i), Len(Hdr(i))
        Queued(i) = False
    Next
    Ready = True
    Sound_Update
End Sub

Public Sub Sound_Shutdown()
    Dim i As Long
    If Not Ready Then Exit Sub
    waveOutReset hWave
    For i = 0 To NUM_BUFS - 1
        waveOutUnprepareHeader hWave, Hdr(i), Len(Hdr(i))
    Next
    waveOutClose hWave
    Ready = False
End Sub

' Refills every buffer the device has finished with. Call this often.
Public Sub Sound_Update()
    Dim i As Long
    If Not Ready Then Exit Sub
    For i = 0 To NUM_BUFS - 1
        If Not Queued(i) Or (Hdr(i).dwFlags And WHDR_DONE) <> 0 Then
            MixInto i
            waveOutWrite hWave, Hdr(i), Len(Hdr(i))
            Queued(i) = True
        End If
    Next
End Sub

' sample: the browser version's sample name. x, y: source position, or -1
' for sounds that play at full volume (the player's own).
Public Sub Sound_Start(ByVal sample As String, ByVal x As Long, ByVal y As Long)
    Dim id As Long, vol As Double, dist As Double, v As Long, best As Long
    If Not Ready Then Exit Sub
    id = EffectFor(sample)
    If id < 0 Then Exit Sub
    vol = 1
    If x >= 0 Then
        dist = Sqr((CDbl(x) - Player.x) ^ 2 + (CDbl(y) - Player.y) ^ 2) / TILEGLOBAL
        vol = 1 - dist / FALLOFF_TILES
        If vol <= 0.03 Then Exit Sub
    End If
    ' a free voice, or else the oldest one
    best = 0
    For v = 0 To MAX_VOICES - 1
        If VoiceSnd(v) < 0 Then
            best = v
            Exit For
        End If
        If VoiceAge(v) < VoiceAge(best) Then best = v
    Next
    AgeCounter = AgeCounter + 1
    VoiceSnd(best) = id
    VoicePos(best) = 0
    VoiceVol(best) = vol * 0.8
    VoiceAge(best) = AgeCounter
End Sub

Private Function EffectFor(ByVal sample As String) As Long
    Select Case sample
        Case "sfx/012.wav": EffectFor = SND_PISTOL
        Case "sfx/011.wav": EffectFor = SND_AUTO
        Case "sfx/013.wav": EffectFor = SND_ROTARY
        Case "lsfx/023.wav": EffectFor = SND_KNIFE
        Case "sfx/049.wav": EffectFor = SND_ZAP
        Case "sfx/024.wav", "sfx/022.wav": EffectFor = SND_ZAP_HEAVY
        Case "sfx/002.wav": EffectFor = SND_BITE
        Case "sfx/010.wav": EffectFor = SND_DOOR_OPEN
        Case "sfx/007.wav": EffectFor = SND_DOOR_CLOSE
        Case "sfx/034.wav": EffectFor = SND_PUSHWALL
        Case "lsfx/031.wav": EffectFor = SND_AMMO
        Case "lsfx/033.wav", "lsfx/034.wav": EffectFor = SND_HEALTH
        Case "lsfx/035.wav", "lsfx/036.wav", "lsfx/037.wav", "lsfx/045.wav": EffectFor = SND_TREASURE
        Case "lsfx/012.wav": EffectFor = SND_KEY
        Case "lsfx/030.wav", "lsfx/038.wav": EffectFor = SND_WEAPON
        Case "sfx/025.wav", "sfx/026.wav", "sfx/086.wav", "sfx/088.wav", "sfx/105.wav", "sfx/107.wav", _
             "sfx/037.wav", "sfx/074.wav", "sfx/046.wav": EffectFor = SND_ROBOT_DEATH
        Case "sfx/035.wav": EffectFor = SND_HOUND_DEATH
        Case "lsfx/009.wav": EffectFor = SND_PLAYER_DEATH
        Case "lsfx/040.wav": EffectFor = SND_ELEVATOR
        Case "lsfx/028.wav": EffectFor = SND_RICOCHET
        Case "lsfx/061.wav": EffectFor = SND_GIBS
        Case "robot/alert": EffectFor = SND_ALERT
        Case Else: EffectFor = -1
    End Select
End Function

Private Sub MixInto(ByVal b As Long)
    Dim n As Long, v As Long, s As Long, st As Long, pos As Long, cnt As Long, vol As Double, a As Double
    For n = 0 To BUF_SAMPLES - 1
        MixAcc(n) = 0
    Next
    For v = 0 To MAX_VOICES - 1
        s = VoiceSnd(v)
        If s >= 0 Then
            st = SndStart(s): pos = VoicePos(v): vol = VoiceVol(v)
            cnt = SndLen(s) - pos
            If cnt > BUF_SAMPLES Then cnt = BUF_SAMPLES
            For n = 0 To cnt - 1
                MixAcc(n) = MixAcc(n) + SndData(st + pos + n) * vol
            Next
            VoicePos(v) = pos + cnt
            If VoicePos(v) >= SndLen(s) Then VoiceSnd(v) = -1
        End If
    Next
    For n = 0 To BUF_SAMPLES - 1
        a = MixAcc(n)
        If a > 32767 Then a = 32767
        If a < -32768 Then a = -32768
        Mem(b * BUF_SAMPLES + n) = a
    Next
End Sub

'---------------------------------------------------------------- synthesis

Private Function Frand() As Double
    SeedS = SeedS * 16807#
    SeedS = SeedS - Int(SeedS / 2147483647#) * 2147483647#
    Frand = SeedS / 2147483647#
End Function

Private Sub BeginEffect(ByVal secs As Double)
    GenLen = Int(secs * RATE)
    ReDim Gen(0 To GenLen - 1)
End Sub

' Scales the effect to a fixed peak and appends it to SndData.
Private Sub FinishEffect(ByVal id As Long)
    Dim i As Long, peak As Double, k As Double
    For i = 0 To GenLen - 1
        If Abs(Gen(i)) > peak Then peak = Abs(Gen(i))
    Next
    If peak < 1 Then peak = 1
    k = 26000 / peak
    SndStart(id) = DataLen
    SndLen(id) = GenLen
    For i = 0 To GenLen - 1
        SndData(DataLen + i) = Gen(i) * k
    Next
    DataLen = DataLen + GenLen
End Sub

' A tone sliding from f0 to f1 Hz between t0 and t1 seconds.
' wave: 0 sine, 1 square, 2 saw, 3 triangle. Envelope: linear attack, exponential decay.
Private Sub Tone(ByVal t0 As Double, ByVal t1 As Double, ByVal f0 As Double, ByVal f1 As Double, ByVal amp As Double, ByVal wave As Long, ByVal attack As Double, ByVal decay As Double)
    Dim i As Long, i0 As Long, i1 As Long, t As Double, phase As Double, ph As Double, w As Double, env As Double
    i0 = Int(t0 * RATE): i1 = Int(t1 * RATE) - 1
    If i1 > GenLen - 1 Then i1 = GenLen - 1
    For i = i0 To i1
        t = (i - i0) / RATE
        phase = phase + (f0 + (f1 - f0) * t / (t1 - t0)) / RATE
        ph = phase - Int(phase)
        Select Case wave
            Case 0: w = Sin(2 * PI * ph)
            Case 1: If ph < 0.5 Then w = 1 Else w = -1
            Case 2: w = 2 * ph - 1
            Case Else: w = 1 - 4 * Abs(ph - 0.5)
        End Select
        env = amp * Exp(-decay * t)
        If t < attack Then env = env * t / attack
        Gen(i) = Gen(i) + w * env
    Next
End Sub

' Filtered noise between t0 and t1 seconds. lowpass 0..1: lower is darker.
Private Sub Noise(ByVal t0 As Double, ByVal t1 As Double, ByVal amp As Double, ByVal lowpass As Double, ByVal attack As Double, ByVal decay As Double)
    Dim i As Long, i0 As Long, i1 As Long, t As Double, lp As Double, env As Double, gain As Double
    i0 = Int(t0 * RATE): i1 = Int(t1 * RATE) - 1
    If i1 > GenLen - 1 Then i1 = GenLen - 1
    gain = 1 / Sqr(lowpass)
    For i = i0 To i1
        t = (i - i0) / RATE
        lp = lp + (Frand() * 2 - 1 - lp) * lowpass
        env = amp * Exp(-decay * t)
        If t < attack Then env = env * t / attack
        Gen(i) = Gen(i) + lp * gain * env
    Next
End Sub

Private Sub BuildEffects()
    ReDim SndData(0 To RATE * 12)
    DataLen = 0
    SeedS = 4321#

    BeginEffect 0.22: Noise 0, 0.2, 0.9, 0.5, 0.001, 22: Tone 0, 0.1, 140, 55, 0.7, 0, 0.001, 30: FinishEffect SND_PISTOL
    BeginEffect 0.15: Noise 0, 0.14, 0.8, 0.6, 0.001, 32: Tone 0, 0.07, 170, 70, 0.5, 0, 0.001, 40: FinishEffect SND_AUTO
    BeginEffect 0.13: Noise 0, 0.12, 0.7, 0.35, 0.001, 38: Tone 0, 0.08, 95, 60, 0.45, 1, 0.001, 35: FinishEffect SND_ROTARY
    BeginEffect 0.2: Noise 0, 0.18, 0.45, 0.12, 0.06, 18: FinishEffect SND_KNIFE
    BeginEffect 0.2: Tone 0, 0.18, 1500, 300, 0.5, 1, 0.002, 12: Noise 0, 0.05, 0.15, 0.7, 0.001, 50: FinishEffect SND_ZAP
    BeginEffect 0.28: Tone 0, 0.26, 750, 110, 0.6, 2, 0.002, 9: Tone 0, 0.12, 1500, 900, 0.2, 1, 0.002, 25: FinishEffect SND_ZAP_HEAVY
    BeginEffect 0.15: Tone 0, 0.05, 320, 200, 0.5, 1, 0.001, 30: Noise 0.04, 0.12, 0.6, 0.5, 0.001, 30: Tone 0.04, 0.14, 1900, 1700, 0.25, 0, 0.001, 25: FinishEffect SND_BITE
    BeginEffect 0.65: Noise 0, 0.65, 0.6, 0.05, 0.05, 3: Tone 0, 0.6, 55, 85, 0.3, 2, 0.05, 3: FinishEffect SND_DOOR_OPEN
    BeginEffect 0.8: Noise 0, 0.55, 0.6, 0.05, 0.05, 3: Tone 0, 0.55, 85, 55, 0.3, 2, 0.05, 3: Tone 0.55, 0.8, 75, 40, 0.8, 0, 0.002, 14: FinishEffect SND_DOOR_CLOSE
    BeginEffect 1.5: Noise 0, 1.5, 0.7, 0.04, 0.1, 1.5: Tone 0, 1.5, 38, 34, 0.25, 2, 0.1, 1.5: FinishEffect SND_PUSHWALL
    BeginEffect 0.12: Noise 0, 0.02, 0.7, 0.8, 0.001, 80: Noise 0.07, 0.09, 0.7, 0.8, 0.001, 80: Tone 0.07, 0.1, 2200, 2200, 0.2, 0, 0.001, 60: FinishEffect SND_AMMO
    BeginEffect 0.26: Tone 0, 0.09, 523, 523, 0.4, 3, 0.005, 6: Tone 0.09, 0.26, 784, 784, 0.4, 3, 0.005, 8: FinishEffect SND_HEALTH
    BeginEffect 0.6: Tone 0, 0.45, 880, 880, 0.35, 0, 0.002, 7: Tone 0.07, 0.52, 1109, 1109, 0.35, 0, 0.002, 7: Tone 0.14, 0.6, 1319, 1319, 0.35, 0, 0.002, 7: FinishEffect SND_TREASURE
    BeginEffect 0.3: Tone 0, 0.1, 1500, 1500, 0.3, 0, 0.001, 30: Tone 0.05, 0.15, 2250, 2250, 0.3, 0, 0.001, 30: Tone 0.1, 0.3, 1800, 1800, 0.3, 0, 0.001, 20: FinishEffect SND_KEY
    BeginEffect 0.42: Noise 0, 0.06, 0.6, 0.3, 0.001, 50: Tone 0.08, 0.3, 440, 440, 0.35, 3, 0.005, 8: Tone 0.18, 0.42, 660, 660, 0.35, 3, 0.005, 8: FinishEffect SND_WEAPON
    BeginEffect 0.7: Noise 0, 0.5, 0.6, 0.6, 0.001, 6: Tone 0, 0.4, 620, 80, 0.35, 1, 0.002, 4: Tone 0.35, 0.65, 60, 40, 0.7, 0, 0.002, 12: FinishEffect SND_ROBOT_DEATH
    BeginEffect 0.6: Noise 0, 0.45, 0.5, 0.7, 0.001, 7: Tone 0, 0.4, 950, 150, 0.35, 1, 0.002, 5: FinishEffect SND_HOUND_DEATH
    BeginEffect 1.2: Tone 0, 1.15, 220, 55, 0.5, 2, 0.01, 2: Noise 0, 1, 0.2, 0.05, 0.01, 2: FinishEffect SND_PLAYER_DEATH
    BeginEffect 0.9: Noise 0, 0.08, 0.7, 0.3, 0.001, 40: Tone 0.1, 0.8, 110, 110, 0.3, 3, 0.05, 2: Tone 0.3, 0.9, 660, 660, 0.25, 0, 0.005, 5: Tone 0.45, 0.9, 880, 880, 0.25, 0, 0.005, 5: FinishEffect SND_ELEVATOR
    BeginEffect 0.2: Tone 0, 0.07, 880, 880, 0.3, 1, 0.002, 5: Tone 0.09, 0.17, 1320, 1320, 0.3, 1, 0.002, 5: FinishEffect SND_ALERT
    BeginEffect 0.15: Tone 0, 0.14, 2600, 1700, 0.3, 0, 0.001, 22: FinishEffect SND_RICOCHET
    BeginEffect 0.2: Noise 0, 0.18, 0.5, 0.15, 0.01, 15: FinishEffect SND_GIBS
End Sub

'---------------------------------------------------------------- content pack hooks

' Rebuilds the synthesized effects, dropping any replaced by a pack.
Public Sub Sound_Rebuild()
    Dim v As Long
    For v = 0 To MAX_VOICES - 1
        VoiceSnd(v) = -1
    Next
    BuildEffects
End Sub

' Replaces effect id with 16-bit mono samples at SOUND_RATE.
Public Sub Sound_ReplaceEffect(ByVal id As Long, ByRef samples() As Integer, ByVal count As Long)
    Dim i As Long
    If id < 0 Or id >= NUM_SND Or count < 1 Then Exit Sub
    ReDim Preserve SndData(0 To DataLen + count)
    For i = 0 To count - 1
        SndData(DataLen + i) = samples(i)
    Next
    SndStart(id) = DataLen
    SndLen(id) = count
    DataLen = DataLen + count
End Sub
