VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7845
   LinkTopic       =   "Form2"
   ScaleHeight     =   6705
   ScaleWidth      =   7845
   Begin VB.CommandButton p2back 
      Caption         =   "Back"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton p2go 
      Caption         =   "Go"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox p1name 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   " "
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox p2name 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   " "
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7200
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   2760
   End
   Begin VB.Shape b2 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   6360
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape b1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   6480
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Player 1 Name:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Player 2 Name:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label twop 
      BackColor       =   &H80000007&
      Caption         =   "2 Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label onep 
      BackColor       =   &H80000007&
      Caption         =   "1 Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   480
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Image paddle 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1560
   End
   Begin VB.Image paddle2 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label twoname 
      BackColor       =   &H80000008&
      Caption         =   " "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label score1 
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label onename 
      BackColor       =   &H80000008&
      Caption         =   " "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label score2 
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   3720
      X2              =   3960
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape ball 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   720
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu new 
         Caption         =   "New Game"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu fasterenabled 
         Caption         =   "Speed Up"
      End
      Begin VB.Menu fasternotenabled 
         Caption         =   "Slow Down"
      End
      Begin VB.Menu changeon 
         Caption         =   "Color Change on"
      End
      Begin VB.Menu changeoff 
         Caption         =   "Color Change off"
      End
   End
   Begin VB.Menu comp 
      Caption         =   "Comp Lv."
      Begin VB.Menu comphardest 
         Caption         =   "Comp. Hardest"
      End
      Begin VB.Menu comphard 
         Caption         =   "Comp. Hard"
      End
      Begin VB.Menu compnormal 
         Caption         =   "Comp. Normal"
      End
      Begin VB.Menu compeasy 
         Caption         =   "Comp. Easy"
      End
      Begin VB.Menu compeasiest 
         Caption         =   "Comp. Easiest"
      End
   End
   Begin VB.Menu mnustars 
      Caption         =   "Stars"
      Begin VB.Menu starson 
         Caption         =   "Stars on"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Program is made by Vincent Chao Gr.8 and is programmed in simple code'
'You may change or modify this program as long as the original concept excluding the name "Pong" belongs to Vincent'
'Thank you to unknown user for 3d stars'

'***********'
'REQUIREMENTS'
  'Pentium 3 or above'
  '500+'
  '96 Mb Ram'
  '**********'


Dim leftup As Boolean
Dim rightup As Boolean
Dim leftdown As Boolean
Dim rightdown As Boolean
Dim hitpaddleleft As Integer
Dim hitpaddletop As Integer
Dim hitpaddleleft2 As Long
Dim hitpaddletop2 As Long
Dim speedup As Boolean
Dim ballspeed As Long
Dim oneplayer As Boolean
Dim twoplayers As Boolean
Dim noupdown As Boolean
Dim colours As Integer
Dim colourchange As Boolean
Dim start As Boolean
Dim startcount As Integer
Dim compspeed As Integer
Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
Dim ran As Integer
Dim qran As Integer


Private Sub Form_Activate()
If starson.Checked = True Then
    Speed = -1
    K = 2038
    Zoom = 256
    Timer1.Interval = 1


    For i = 0 To 100
        X(i) = Int(Rnd * 1024) - 512
        Y(i) = Int(Rnd * 1024) - 512
        Z(i) = Int(Rnd * 512) - 256
    Next i
    End If
End Sub


Private Sub starson_Click()
If starson.Checked = True Then
starson.Checked = False
Else
starson.Checked = True
End If
End Sub

Private Sub Timer1_Timer()

If b2.Visible = False And oneplayer = True And start = False And ran = 1 Then
b2.Top = paddle2.Top
b2.Left = paddle2.Left + paddle2.Width / 2
b2.Visible = True
ran = (Rnd * qran)
End If





If colourchange = True Then
colours = colours + 1
End If
If colours = 1 Then
Form2.BackColor = &HFFFFFF
End If
If colours = 2 Then
Form2.BackColor = &HFFFF&
End If
If colours = 3 Then
Form2.BackColor = &HFFFFFF
End If
If colours = 4 Then
Form2.BackColor = &HFF
End If
If colours = 5 Then
Form2.BackColor = &HFF00&
End If
If colours = 6 Then
Form2.BackColor = &H80FF&
End If
If colours = 7 Then
Form2.BackColor = &HFFFF00
End If
If colours = 8 Then
Form2.BackColor = &HFF0000
End If
If colours = 9 Then
Form2.BackColor = &HFF00FF
End If
If colours = 10 Then
Form2.BackColor = &H8080FF
End If
If colours = 11 Then
colours = 0
End If

onep.BackColor = Form2.BackColor
twop.BackColor = Form2.BackColor
Label1.BackColor = Form2.BackColor
Label2.BackColor = Form2.BackColor
Line1.BorderColor = Form2.BackColor
score1.BackColor = Form2.BackColor
score2.BackColor = Form2.BackColor



If ballspeed < 31 Then
fasternotenabled.Enabled = False
End If
If ballspeed > 32 Then
fasternotenabled.Enabled = True
End If
If ballspeed < 238 Then
fasterenabled.Enabled = True
End If
If ballspeed > 239 Then
fasterenabled.Enabled = False
End If


If compnormal.Checked = True Then
compspeed = ballspeed / 1.1
qran = 13
End If
If comphard.Checked = True Then
compspeed = ballspeed / 1.07
qran = 11
End If
If comphardest.Checked = True Then
compspeed = ballspeed / 1.04
qran = 7
End If
If compeasy.Checked = True Then
compspeed = ballspeed / 1.15
qran = 14
End If
If compeasiest.Checked = True Then
compspeed = ballspeed / 1.2
qran = 15
End If
If ball.Left > (paddle2.Left + paddle.Width / 2) And oneplayer = True Then
paddle2.Left = paddle2.Left + compspeed
End If
If ball.Left < (paddle2.Left + paddle.Width / 2) And oneplayer = True Then
paddle2.Left = paddle2.Left - compspeed
End If

hitpaddleleft = ball.Left - paddle.Left
hitpaddleleft2 = ball.Left - paddle2.Left



'Normal Ball Movement'

If leftup = True And leftdown = False And rightup = False And rightdown = False Then
ball.Left = ball.Left - ballspeed
ball.Top = ball.Top - ballspeed
End If
If rightup = True And leftdown = False And leftup = False And rightdown = False Then
ball.Left = ball.Left + ballspeed
ball.Top = ball.Top - ballspeed
End If
If leftdown = True And leftup = False And rightup = False And rightdown = False Then
ball.Left = ball.Left - ballspeed
ball.Top = ball.Top + ballspeed
End If
If rightdown = True And leftdown = False And rightup = False And leftup = False Then
ball.Left = ball.Left + ballspeed
ball.Top = ball.Top + ballspeed
End If

'End Normal Ball Movement'



If ball.Top > 6480 And leftdown = True And rightdown = False Then
leftdown = False
leftup = True
MsgBox "Player 2 + 1"
score2.Caption = score2.Caption + 1
End If
If ball.Top > 6480 And rightdown = True And leftdown = False Then
rightdown = False
rightup = True
MsgBox "Player 2 + 1"
score2.Caption = score2.Caption + 1
End If
If ball.Left > 7560 And rightup = True And rightdown = False Then
rightup = False
leftup = True
End If
If ball.Left > 7560 And rightdown = True And rightup = False Then
rightdown = False
leftdown = True
End If
If ball.Top < 0 And leftup = True And rightup = False Then
leftup = False
leftdown = True
MsgBox "Player 1 + 1"
score1.Caption = score1.Caption + 1
End If
If ball.Top < 0 And rightup = True And leftup = False Then
rightup = False
rightdown = True
MsgBox "Player 1 + 1"
score1.Caption = score1.Caption + 1
End If
If ball.Left < 0 And leftup = True And leftdown = False Then
leftup = False
rightup = True
End If
If ball.Left < 0 And leftdown = True And leftup = False Then
leftdown = False
rightdown = True
End If

If hitpaddleleft < paddle.Width And hitpaddleleft > -(paddle.Width / -6.5) And ball.Top > paddle.Top And leftdown = True Then
leftdown = False
leftup = True
End If
If hitpaddleleft < paddle.Width And hitpaddleleft > -(paddle.Width / -6.5) And ball.Top > paddle.Top And rightdown = True Then
rightdown = False
rightup = True
End If
If hitpaddleleft2 < paddle2.Width And hitpaddleleft2 > -(paddle.Width / -6.5) And ball.Top < paddle2.Top And leftup = True Then
leftup = False
leftdown = True
End If
If hitpaddleleft2 < paddle2.Width And hitpaddleleft2 > -(paddle.Width / -6.5) And ball.Top < paddle2.Top And rightup = True Then
rightup = False
rightdown = True
End If
If starson.Checked Then
    For i = 0 To 50
        Circle (tmpX(i), tmpY(i)), 5, BackColor
        Z(i) = Z(i) + Speed
        If Z(i) > 255 Then Z(i) = -255
        If Z(i) < -255 Then Z(i) = 255
        tmpZ(i) = Z(i) + Zoom
        tmpX(i) = (X(i) * K / tmpZ(i)) + (Form2.Width / 2)
        tmpY(i) = (Y(i) * K / tmpZ(i)) + (Form2.Height / 2)
        Radius = 1
        StarColor = 256 - Z(i)
        Circle (tmpX(i), tmpY(i)), 5, RGB(StarColor, StarColor, StarColor)
    Next i
    End If
    If b1.Visible = True Then
    b1.Top = b1.Top - 75
    End If
        If b2.Visible = True Then
    b2.Top = b2.Top + 75
    End If
        If b1.Top < 0 Then
        b1.Visible = False
    End If
            If b2.Top > 6480 Then
        b2.Visible = False
    End If
    If paddle2.Width > 719 And b1.Left < paddle2.Left + 1 + paddle2.Width And b1.Left > paddle2.Left - 1 And b1.Top < paddle2.Top And b1.Visible = True Then
    paddle2.Width = paddle2.Width - 15
    End If
        If paddle.Width > 719 And b2.Left < paddle.Left + 1 + paddle.Width And b2.Left > paddle.Left - 1 And b2.Top > paddle.Top And b2.Visible = True Then
    paddle.Width = paddle.Width - 15
    End If
End Sub

































Private Sub changeoff_Click()
colourchange = False
changeon.Enabled = True
changeoff.Enabled = False
End Sub

Private Sub changeon_Click()
colourchange = True
changeon.Enabled = False
changeoff.Enabled = True
End Sub

Private Sub compeasiest_Click()
compeasiest.Checked = True
compeasy.Checked = False
compnormal.Checked = False
comphard.Checked = False
comphardest.Checked = False

End Sub



Private Sub compeasy_Click()
compeasiest.Checked = False
compeasy.Checked = True
compnormal.Checked = False
comphard.Checked = False
comphardest.Checked = False
End Sub

Private Sub comphard_Click()
compeasiest.Checked = False
compeasy.Checked = False
compnormal.Checked = False
comphard.Checked = True
comphardest.Checked = False

End Sub

Private Sub comphardest_Click()
compeasiest.Checked = False
compeasy.Checked = False
compnormal.Checked = False
comphard.Checked = False
comphardest.Checked = True

End Sub

Private Sub compnormal_Click()
compeasiest.Checked = False
compeasy.Checked = False
compnormal.Checked = True
comphard.Checked = False
comphardest.Checked = False

End Sub

Private Sub exit_Click()
Unload Form2
End
End Sub

Private Sub fasterenabled_Click()
ballspeed = ballspeed + 30
End Sub

Private Sub fasternotenabled_Click()
ballspeed = ballspeed - 30
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyW Then
speedup = False
End If
If KeyCode = vbKeyQ Then
speedup = False
End If
If KeyCode = vbKeyRight Then
paddle.Left = paddle.Left + 480
End If
If KeyCode = vbKeyLeft Then
paddle.Left = paddle.Left - 480
End If
If KeyCode = vbKeyC And twoplayers = True Then
paddle2.Left = paddle2.Left + 480
End If
If KeyCode = vbKeyZ And twoplayers = True Then
paddle2.Left = paddle2.Left - 480
End If
If KeyCode = vbKeyUp And noupdown = False Then
paddle.Top = paddle.Top - 480
End If
If KeyCode = vbKeyDown And noupdown = False Then
paddle.Top = paddle.Top + 480
End If
If KeyCode = vbKeyS And noupdown = False And twoplayers = True Then
paddle2.Top = paddle2.Top - 480
End If
If KeyCode = vbKeyX And noupdown = False And twoplayers = True Then
paddle2.Top = paddle2.Top + 480
End If
If KeyCode = vbKeyUp Then
If b1.Visible = False And (oneplayer = True Or twoplayers = True) And start = False Then
b1.Top = paddle.Top
b1.Left = paddle.Left + paddle.Width / 2
b1.Visible = True
End If
End If
If KeyCode = vbKeyX And twoplayers = True Then
If b2.Visible = False And (oneplayer = True Or twoplayers = True) And start = False Then
b2.Top = paddle2.Top
b2.Left = paddle2.Left + paddle2.Width / 2
b2.Visible = True
End If
End If
End Sub

Private Sub Form_Load()
Randomize
noupdown = True
speedup = True
colourchange = False
End Sub

Private Sub new_Click()
start = True
End Sub

Private Sub onep_Click()
start = True
onep.Visible = False
twop.Visible = False
oneplayer = True
End Sub

Private Sub p2back_Click()
onep.Visible = True
twoplayers = False
p2go.Visible = False
p2name.Visible = False
p1name.Visible = False
Label1.Visible = False
Label2.Visible = False
p2back.Visible = False
End Sub

Private Sub p2go_Click()
p2go.Visible = False
p2name.Visible = False
p1name.Visible = False
Label1.Visible = False
Label2.Visible = False
p2back.Visible = False
twop.Visible = False
start = True
End Sub


Private Sub Timer2_Timer()
ran = (Rnd * qran)
If start = True Then
Line1.Visible = True
score1.Visible = True
score2.Visible = True
rightdown = False
leftdown = False
rightup = False
leftup = False
onename.Caption = "Let's get ready"
twoname.Caption = "to RUMBLE!!!"
startcount = startcount + 1
End If
If startcount = 5 Then
score1.Caption = "5"
score2.Caption = "5"
End If
If start = True And startcount = 10 Then
score1.Caption = "4"
score2.Caption = "4"
End If
If start = True And startcount = 15 Then
score1.Caption = "3"
score2.Caption = "3"
End If
If start = True And startcount = 20 Then
score1.Caption = "2"
score2.Caption = "2"
End If
If start = True And startcount = 25 Then
score1.Caption = "1"
score2.Caption = "1"
End If
If start = True And startcount = 30 Then
twoname.Caption = p1name.Text
onename.Caption = p2name.Text
score1.Caption = "0"
score2.Caption = "0"
rightdown = True
ballspeed = 60
start = False
startcount = 0
comphard.Checked = True
End If
End Sub



Private Sub twop_Click()
oneplayer = False
onep.Visible = False
twoplayers = True
p2go.Visible = True
p2name.Visible = True
p1name.Visible = True
Label1.Visible = True
Label2.Visible = True
p2back.Visible = True
End Sub

