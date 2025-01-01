VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "Form1.frx":6988A
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调试模式"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   360
      Left            =   2340
      TabIndex        =   0
      Top             =   1900
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Dim qdtm As Integer
Dim qdsd As Integer
Dim tmd As String
Dim xt As String

Private Sub Form_Load()
Open "data\aud.txt" For Input As #1
Line Input #1, aud
Close #1
Open "data\tmdu.txt" For Input As #1
Line Input #1, tmd
Close #1
tmdu = Val(tmd)
Open "themes\themes.txt" For Input As #1
Line Input #1, theme
Close #1
Open "themes\" & theme & "\textc.txt" For Input As #1
Line Input #1, textc
Close #1
Open "data\dt.txt" For Input As #1
Line Input #1, dtrs
Close #1
Open "data\xtbb.txt" For Input As #1
Line Input #1, xt
Close #1
xtbb = Val(xt)
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
qdtm = 0
qdsd = 1
Form2.Show
Open "data\dt.txt" For Input As #1
Line Input #1, dt
Close #1
If exitproc("qnbf.exe") Then
 Label1.Visible = False
 If dt = "1" Then
  If Not exitproc("dual.exe") Then
   Call Shell("cmd /c start dual.exe")
  End If
 End If
Else
 Label1.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
If qdtm + 10 <= 240 Then
 qdtm = qdtm + 10
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
Else
 qdtm = 240
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
 Timer2.Enabled = True
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
If qdtm - 15 >= 0 Then
 qdtm = qdtm - 15
 qdsd = qdsd + 2
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
 Me.Top = Me.Top - qdsd
Else
 qdtm = 0
 qdsd = 1
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
 Unload Me
 Timer3.Enabled = False
End If
End Sub
