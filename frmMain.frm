VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "AI - TEST"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10065
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer timShowIA 
      Interval        =   1
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer timMoveEnSp 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer timMoveEnLin 
      Interval        =   250
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox picSp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   750
      Left            =   240
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox picEn 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   750
      Left            =   8400
      Picture         =   "frmMain.frx":1DF2
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************'
'***    AI - Test   ***'
'**********************'

Private Const MAX_DISTANCE = 250
 
Private bEnDirL As Boolean
Private bFollow As Boolean

Private Sub Form_Load()
    bEnDirL = True
    bFollow = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        picSp.Top = y
        picSp.Left = X
    End If
    If (GetDistance(picSp.Left, picSp.Top, picEn.Left, picEn.Top) < MAX_DISTANCE) Then
        timMoveEnLin.Enabled = False
        timMoveEnSp.Enabled = True
    End If
End Sub

Private Sub timMoveEnLin_Timer()
    
    If (GetDistance(picSp.Left, picSp.Top, picEn.Left, picEn.Top) < MAX_DISTANCE) Then
        timMoveEnLin.Enabled = False
        timMoveEnSp.Enabled = True
        Exit Sub
    End If
    
    If bEnDirL Then
        If (picEn.Left > 1) Then
            picEn.Left = picEn.Left - 10
        Else
            bEnDirL = False
        End If
    Else
        If (picEn.Left + 50 < frmMain.ScaleWidth) Then
            picEn.Left = picEn.Left + 10
        Else
            bEnDirL = True
        End If
    End If
End Sub

Function GetDistance(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Long
    Dim X As Long, y As Long
    Dim Length As Long
    
    X1 = X1 + 25
    Y1 = Y1 + 25
    X2 = X2 + 25
    Y2 = Y2 + 25
    
    X = X2 - X1
    y = Y2 - Y1
    
    Length = Sqr(X * X + y * y)
    
    GetDistance = Length
End Function

Private Sub timMoveEnSp_Timer()
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    Dim bXOk As Boolean, bYOk As Boolean
    
    X1 = picEn.Left
    Y1 = picEn.Top
    X2 = picSp.Left
    Y2 = picSp.Top
    
    If (GetDistance(X2, Y2, X1, Y1) > MAX_DISTANCE) Then
        If picEn.Top > 5 Then
            Call MoveSpriteToLine
            Exit Sub
        Else
            timMoveEnSp.Enabled = False
            timMoveEnLin.Enabled = True
        End If
    End If
    If Collision Then
        MsgBox "*** |~`CRASH ´~| ***"
        timMoveEnSp.Enabled = False
        picSp.Move 20, 380
        picEn.Move 570, 0
        timMoveEnLin.Enabled = True
        Exit Sub
    End If
    If (X1 = X2) Then
        bXOk = True
    End If
    If (Y1 = Y2) Then
        bYOk = True
    End If
    If bXOk Or (X1 < X2) Then
        picEn.Left = picEn.Left + 5
    Else
        picEn.Left = picEn.Left - 5
    End If
    If bYOk Or (Y1 < Y2) Then
        picEn.Top = picEn.Top + 5
    Else
        picEn.Top = picEn.Top - 5
    End If
End Sub

Private Sub MoveSpriteToLine()
    picEn.Top = picEn.Top - 10
End Sub

Private Function Collision() As Boolean
    If Not ((picSp.Left > picEn.Left + 50) Or (picSp.Left + 50 < picEn.Left)) Then
        Collision = Not ((picSp.Top > picEn.Top + 50) Or (picSp.Top + 50 < picEn.Top))
    End If
End Function

Private Sub timShowIA_Timer()
    Dim X As Integer, y As Integer
    
    Refresh
    X = picEn.Left + 25
    y = picEn.Top + 25

    frmMain.Circle (X, y), MAX_DISTANCE
End Sub
