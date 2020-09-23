VERSION 5.00
Begin VB.Form frmRes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mr. Eat Gum Guy 2 Options"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "frmRes.frx":0000
   LinkTopic       =   "frmRes"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMusic 
      Caption         =   "Play Game with Music"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkSoundFX 
      Caption         =   "Allow Sound Effects"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Mr. Eat Gum Guy"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start Game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test Resolution"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pick A Resolution:"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   2055
      Begin VB.ComboBox cmbBPP 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton opt1024x768 
         Caption         =   "1024x768"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton opt800x600 
         Caption         =   "800x600"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opt640x480 
         Caption         =   "640x480"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "BPP:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1110
         Width           =   375
      End
   End
   Begin VB.Line Line2 
      X1              =   8
      X2              =   272
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The testing process may make your monitor flicker!"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label lblTest 
      Alignment       =   2  'Center
      Caption         =   "You must test a resolution before you can start!"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Line Line1 
      X1              =   152
      X2              =   272
      Y1              =   40
      Y2              =   40
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All of the code below is really self explanitory if you read it line by line, which
'you really should do if you want to understand the concepts I used.

Public tFormTop As Integer
Public tForp1left As Integer
Public tFormWidth As Integer
Public tFormHeight As Integer

Private Sub cmbBPP_Click()
    With lblTest
        .FontBold = False
        .ForeColor = vbBlack
        .Caption = "You must test the resolution before you can start!"
    End With
    cmdStart.Enabled = False
End Sub

Private Sub Form_Load()
    With cmbBPP
        .AddItem "8"
        .AddItem "16"
        .AddItem "24"
        .AddItem "32"
        .ListIndex = 1
    End With

    With frmRes
        tFormTop = (Screen.Height \ 2) - (.Height \ 2)
        tForp1left = (Screen.Width \ 2) - (.Width \ 2)
        tFormWidth = .Width
        tFormHeight = .Height

        Form_Resize
    End With
End Sub

Private Sub Form_Resize()
    With frmRes
        .Top = tFormTop
        .Left = tForp1left
        .Width = tFormWidth
        .Height = tFormHeight
    End With
End Sub

Private Sub cmdExit_Click()
    Unload frmMain
    Unload frmRes
    End
End Sub

Private Sub cmdStart_Click()
    If opt640x480.Value = True Then
        ChosenRes = "640x480"
    ElseIf opt800x600.Value = True Then
        ChosenRes = "800x600"
    ElseIf opt1024x768.Value = True Then
        ChosenRes = "1024x768"
    End If

    If chkMusic.Value = 1 Then
        AllowMusic = True
    Else
        AllowMusic = False
    End If

    If chkSoundFX.Value = 1 Then
        AllowSound = True
    Else
        AllowSound = False
    End If

    Unload frmRes
    frmMain.Show
End Sub

Private Sub cmdTest_Click()
Dim dxTest As New DirectX7
Dim ddTest As DirectDraw7

    If opt640x480.Value = True Then
        ScreenCurW = 640
        ScreenCurH = 480
    ElseIf opt800x600.Value = True Then
        ScreenCurW = 800
        ScreenCurH = 600
    ElseIf opt1024x768.Value = True Then
        ScreenCurW = 1024
        ScreenCurH = 768
    End If

    Select Case cmbBPP.ListIndex
        Case 0
            ScreenCurBPP = 8
        Case 1
            ScreenCurBPP = 16
        Case 2
            ScreenCurBPP = 24
        Case 3
            ScreenCurBPP = 32
    End Select

On Error GoTo errResTest
    Set ddTest = dxTest.DirectDrawCreate("")
    ddTest.SetCooperativeLevel frmRes.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX
    ddTest.SetDisplayMode ScreenCurW, ScreenCurH, ScreenCurBPP, 0, DDSDM_DEFAULT
    ddTest.RestoreDisplayMode
    ddTest.SetCooperativeLevel frmRes.hWnd, DDSCL_NORMAL
    Set ddTest = Nothing
    Set dxTest = Nothing
    Form_Resize
    With lblTest
        .FontBold = True
        .ForeColor = vbBlue
        .Caption = "The Selected Resolution is supported!"
    End With
    cmdStart.Enabled = True
    Exit Sub

errResTest:
    Form_Resize
    With lblTest
        .FontBold = True
        .ForeColor = vbRed
        .Caption = "The Selected Resolution is not supported!"
    End With
    cmdStart.Enabled = False
End Sub

Private Sub opt640x480_Click()
    With lblTest
        .FontBold = False
        .ForeColor = vbBlack
        .Caption = "You must test the resolution before you can start!"
    End With
    cmdStart.Enabled = False
End Sub

Private Sub opt800x600_Click()
    With lblTest
        .FontBold = False
        .ForeColor = vbBlack
        .Caption = "You must test the resolution before you can start!"
    End With
    cmdStart.Enabled = False
End Sub

Private Sub opt1024x768_Click()
    With lblTest
        .FontBold = False
        .ForeColor = vbBlack
        .Caption = "You must test the resolution before you can start!"
    End With
    cmdStart.Enabled = False
End Sub
