VERSION 5.00
Begin VB.Form FrmClock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   Icon            =   "FrmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrTime 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FrmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
FrmSplash.TmrSplash.Enabled = False
Unload FrmSplash
SetWinLng Me
MakePNGS
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 5 Then Unload Me
If KeyAscii = 19 Then FrmCPL.Show
If KeyAscii = 20 Then
    uTopMost = Not uTopMost
    FormTop
End If
End Sub

Public Sub FormTop()
On Error Resume Next
    If uTopMost = True Then
        MakeFormTop Me.hwnd, True
        If CPLLoaded = True Then FrmCPL.CHKTOPMOST.Value = vbChecked
    Else
        MakeFormTop Me.hwnd, False
        If CPLLoaded = True Then FrmCPL.CHKTOPMOST.Value = vbUnchecked
    End If
End Sub

Private Sub Form_Load()
FormTop
Me.Left = uLeft
Me.Top = uTop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button = 1 Then
        If Shift = 2 Then
            FormDrag Me
            uLeft = Me.Left
            uTop = Me.Top
        End If
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting
TotalEnd
End Sub

Private Sub TmrTime_Timer()
'SetWinLng Me
MakePNGS
End Sub



