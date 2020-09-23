VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10455
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Player"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuW 
      Caption         =   "&Window"
      Begin VB.Menu mnuWCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWTile 
         Caption         =   "&Tile"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
    OpenNewWindow
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileNew_Click()
    OpenNewWindow
End Sub


Private Sub OpenNewWindow()
    Dim frm As frmMusic
    Set frm = New frmMusic
    frm.Show
End Sub


Private Sub mnuWCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuWTile_Click()
    Me.Arrange vbTileHorizontal
End Sub
