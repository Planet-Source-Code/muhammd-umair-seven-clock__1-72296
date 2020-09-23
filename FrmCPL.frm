VERSION 5.00
Begin VB.Form FrmCPL 
   BorderStyle     =   0  'None
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraTH 
      Caption         =   "THEMES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2625
      Begin VB.ComboBox CmbTH 
         Height          =   315
         ItemData        =   "FrmCPL.frx":0000
         Left            =   120
         List            =   "FrmCPL.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2400
      End
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   495
      Left            =   1470
      TabIndex        =   5
      Top             =   2955
      Width           =   1215
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   2955
      Width           =   1215
   End
   Begin VB.CheckBox CHKSTARTUP 
      Caption         =   "Start When windows startup."
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   2625
      Width           =   2385
   End
   Begin VB.CheckBox CHKTOPMOST 
      Caption         =   "Top Most Other Windows."
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   2400
      Width           =   2355
   End
   Begin VB.Frame FraCLOpa 
      Caption         =   "Clock Transparency - "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2625
      Begin VB.HScrollBar HSOpacity 
         Height          =   255
         Left            =   60
         Max             =   255
         TabIndex        =   1
         Top             =   300
         Width           =   2475
      End
   End
   Begin VB.Label LblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BY:Muhammad Umair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15
      TabIndex        =   9
      Top             =   3480
      Width           =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   345
      TabIndex        =   8
      Top             =   -105
      Width           =   2175
   End
End
Attribute VB_Name = "FrmCPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHKSTARTUP_Click()
uStartup = CHKSTARTUP.Value
SetStartUp CBool(uStartup)
End Sub

Private Sub CHKTOPMOST_Click()
uTopMost = CBool(CHKTOPMOST.Value)
FrmClock.FormTop
End Sub

Private Sub CmbTH_Click()
CTN = CmbTH.List(CmbTH.ListIndex)
End Sub

Private Sub CmdAbout_Click()
FrmAbout.Show
End Sub

Private Sub CmdOK_Click()
uCTN = CmbTH.List(CmbTH.ListIndex)
CTN = uCTN
uClockMASTrans = HSOpacity.Value
SaveSetting
CPLLoaded = False
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
CPLLoaded = True
CmbTH.Text = CTN
HSOpacity.Value = uClockMASTrans
If uTopMost = True Then
    CHKTOPMOST.Value = vbChecked
Else
    CHKTOPMOST.Value = vbUnchecked
End If

CHKSTARTUP.Value = uStartup
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub HSOpacity_Change()
FraCLOpa.Caption = "Clock Transparency - " & CInt(HSOpacity.Value * 100 / 255) & "%"
uClockMASTrans = HSOpacity.Value
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
