VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrSplash 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
If App.PrevInstance Then
    MsgBox App.EXEName & " is already runing!", vbInformation + vbOKOnly, App.EXEName & ". Loading System."
   End
End If
If GDIAvailable = False Then MsgBox "ERROR LOADING GDI+", vbCritical, "U11D Checking System": End
MakeFormTop Me.hwnd, True
SetWinLng Me
MakePNG App.Path & "\THEMES\Splash.png", Me, 240, False

CM.m(0, 0) = 1
CM.m(1, 1) = 1
CM.m(2, 2) = 1
CM.m(3, 3) = 1
End Sub

Private Sub TmrSplash_Timer()
    FrmClock.Show
End Sub





