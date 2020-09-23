VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMusic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   Begin VB.CheckBox chkLoop 
      Caption         =   "Loop"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox picTarget 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   1680
      ScaleHeight     =   179
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   3000
   End
   Begin MSComctlLib.Slider Volume 
      Height          =   2415
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   4260
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   2000
      SmallChange     =   50
      Max             =   10000
      TickStyle       =   2
      TickFrequency   =   1000
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open"
      Height          =   615
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMusic.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMusic.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMusic.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMusic.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3360
   End
   Begin MSComDlg.CommonDialog cdOpenFile 
      Left            =   480
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider Balance 
      Height          =   2415
      Left            =   6300
      TabIndex        =   7
      Top             =   360
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   4260
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   2000
      SmallChange     =   1000
      Min             =   -10000
      Max             =   10000
      TickStyle       =   2
      TickFrequency   =   2000
   End
   Begin MSComctlLib.Slider Speed 
      Height          =   2415
      Left            =   7560
      TabIndex        =   8
      Top             =   360
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   4260
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   2
      Max             =   100
      SelStart        =   50
      TickStyle       =   2
      TickFrequency   =   10
      Value           =   50
   End
   Begin MSComctlLib.Slider Position 
      Height          =   630
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1111
      _Version        =   393216
      Max             =   100
      TickStyle       =   2
      TickFrequency   =   10
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Position"
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Speed"
      Height          =   255
      Left            =   7440
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Balance"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Volume"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblState 
      Caption         =   "State:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2235
      Width           =   2175
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1950
      Width           =   2175
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1665
      Width           =   2055
   End
   Begin VB.Label lblDuration 
      Caption         =   "Duration:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblPosition 
      Caption         =   "Position:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   525
      Width           =   2055
   End
   Begin VB.Label lblRate 
      Caption         =   "Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   810
      Width           =   2055
   End
   Begin VB.Label lblBalance 
      Caption         =   "Balance:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1095
      Width           =   2055
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume: "
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1380
      Width           =   2175
   End
End
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' To play Ogg Vorbis format, download the Ogg DirectShow filters from
' http://tobias.everwicked.com/

Private Music As clsMusic


Private Sub Form_Load()
    ' Create an instance of the music class
    Show
    Set Music = New clsMusic
    Music.Window = picTarget.hWnd
End Sub


Private Sub cmdOpenFile_Click()
    ' Play a music file
    On Error Resume Next
    cdOpenFile.CancelError = True
    cdOpenFile.ShowOpen
    If Err.Number <> 0 Then
        Exit Sub
    End If
    Music.FileName = cdOpenFile.FileName
    tmrTimer.Enabled = True
    Position.Min = 0
    Position.Max = Music.Duration
    Position.TickFrequency = Position.Max / 10
    Position.SmallChange = Position.TickFrequency / 2
    Position.LargeChange = Position.Max / 10
    Position = 0
    Music.Volume = 0 - Volume.value
    Music.Balance = Balance.value
    Music.Speed = Speed.value / 50
    Music.Position = Position.value
    If Music.HasVideo = False Then
        picTarget.Cls
        picTarget.Print LastPart(cdOpenFile.FileName)
    End If
    Music.Play
End Sub


Private Function LastPart(ByVal s As String)
    Dim s2() As String
    s2 = Split(s, "\")
    LastPart = s2(UBound(s2, 1))
End Function


Private Sub tmrTimer_Timer()
    'Update all label controls
    Dim State As enumState
    Dim s As String
    Position.value = Music.Position
    lblPosition.Caption = "Position: " & Music.Position
    lblDuration.Caption = "Duration: " & Music.Duration
    lblRate.Caption = "Rate: " & Music.Speed
    lblVolume.Caption = "Volume: " & Music.Volume
    lblBalance.Caption = "Balance: " & Music.Balance
    lblWidth.Caption = "Width: " & Music.Width
    lblHeight.Caption = "Height: " & Music.Height
    State = Music.State
    If State = stStopped Then
        s = "Stopped"
    ElseIf State = stPaused Then
        s = "Paused"
    ElseIf State = stPlaying Then
        s = "Playing"
    End If
    lblState.Caption = "State: " & s
    If chkLoop.value = 1 Then
        If Music.Position = Music.Duration Then
            Music.Position = 0
            Music.Play
        End If
    End If
End Sub


Private Sub cmdPlay_Click()
    ' Play
    Music.Play
End Sub


Private Sub cmdPause_Click()
    ' Pause
    Music.Pause
End Sub


Private Sub cmdStop_Click()
    ' Stop
    Music.StopPlaying
End Sub


Private Sub Volume_Scroll()
    Music.Volume = 0 - Volume.value
End Sub


Private Sub Balance_Scroll()
    Music.Balance = Balance.value
End Sub


Private Sub Speed_Scroll()
    Music.Speed = Speed.value / 50
End Sub


Private Sub Position_Scroll()
    Music.Position = Position.value
End Sub


