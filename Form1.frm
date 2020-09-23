VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   60
      ScaleHeight     =   1275
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   420
      Width           =   7515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1740
      Top             =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Black"
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mike toye
Option Explicit
Dim lNumParticles As Long
Dim lParticles(6, 1000) As Long
Const Pi As Single = 3.14159265359
Function GimmeX(ByVal aIn As Single, lIn As Long) As Long
    GimmeX = Sin(aIn * (Pi / 180)) * lIn
End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Long
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function

Private Sub Command1_Click()
    If Command1.Caption = "Start" Then
        Command1.Caption = "Stop"
        Timer1.Enabled = True
    Else
        Command1.Caption = "Start"
        Timer1.Enabled = False
    End If
        
End Sub

Sub LoadParticles()
Dim x As Long
    Randomize Timer ^ Format(Now, "ss")
    For x = 0 To lNumParticles - 1
        lParticles(0, x) = Rnd * picM.Width 'x pos
        lParticles(1, x) = Rnd * picM.Height 'y pos
        lParticles(2, x) = Int(Rnd * 360) 'angle
        lParticles(3, x) = Rnd * 30 'speed
        lParticles(4, x) = Int(Rnd * 255) 'color
        lParticles(5, x) = 1 'for collision detection
    Next x
End Sub

Private Sub Command2_Click()
Dim x As Integer
Dim iCol As Integer
    If Command2.Caption = "Black" Then
        Command2.Caption = "White"
        For x = 0 To lNumParticles - 1: lParticles(4, x) = 0: Next x
    Else
        Command2.Caption = "Black"
        For x = 0 To lNumParticles - 1: lParticles(4, x) = Int(Rnd * 255): Next x
    End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
    lNumParticles = 1000
    LoadParticles

End Sub

Private Sub Form_Resize()
    picM.Width = Me.Width - picM.Left - 120
    picM.Height = Me.Height - picM.Top - 400
End Sub

Private Sub picM_Click()
    Unload Me
    End
End Sub

Private Sub Timer1_Timer()
Dim x As Long
Dim y As Long
Dim iX As Integer
Dim iY As Integer
Const iArc As Integer = 10
Dim bDirPositive As Boolean
    If Me.Width = Screen.Width Then Me.BorderStyle = 0
    'picM.Refresh
    For x = 0 To lNumParticles - 1
        picM.PSet (lParticles(0, x), lParticles(1, x)), RGB(lParticles(4, x), lParticles(4, x), lParticles(4, x))
    Next x
    DoEvents
    For x = 0 To lNumParticles - 1
        'new angle
        bDirPositive = IIf(Rnd * 1000 > 500, True, False)
        lParticles(2, x) = lParticles(2, x) + (Int(Rnd * iArc) * IIf(bDirPositive, 1, -1))
        If lParticles(2, x) > 359 Then lParticles(2, x) = lParticles(2, x) - 359
        If lParticles(2, x) < 0 Then lParticles(2, x) = lParticles(2, x) + 359
        'get the new x,y coordinates based on the new angle
        iX = GimmeX(lParticles(2, x), lParticles(3, x))
        iY = GimmeY(lParticles(2, x), lParticles(3, x))
        lParticles(0, x) = lParticles(0, x) + iX
        If lParticles(0, x) < 0 Then
            lParticles(0, x) = picM.Width
        End If
        If lParticles(0, x) > picM.Width Then
            lParticles(0, x) = 0
        End If
        lParticles(1, x) = lParticles(1, x) + iY
        If lParticles(1, x) < 0 Then
            lParticles(1, x) = picM.Height
        End If
        If lParticles(1, x) > picM.Height Then
            lParticles(1, x) = 0
        End If
    Next x
End Sub
