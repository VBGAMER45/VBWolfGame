VERSION 5.00
Begin VB.Form frmMain
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBWolf"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   960
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim w As Single, h As Single
    ' size the window so the client area is exactly the scaled frame
    w = Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX + SCR_W * WIN_SCALE * Screen.TwipsPerPixelX
    h = Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY + SCR_H * WIN_SCALE * Screen.TwipsPerPixelY
    Me.Move (Screen.Width - w) / 2, (Screen.Height - h) / 2, w, h
    HookWindow Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindow Me.hWnd
    Running = False
    FormClosed = True
End Sub
