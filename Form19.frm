VERSION 5.00
Begin VB.Form Form19 
   BorderStyle     =   0  'None
   Caption         =   "Form19"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LinkTopic       =   "Form19"
   Picture         =   "Form19.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command5 
      Caption         =   "打开截图工具"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "打开记事本"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打开计算器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开相机"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认执行"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "例如"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入快捷指令"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Integer
Dim xt As String

Private Sub Command1_Click()
If Not Text1.Text = "" Then
 If Text1.Text = "打开相机" Then
  Call Shell("cmd /c start tools\w" & xtbb & "\xj.lnk")
 ElseIf Text1.Text = "打开计算器" Then
  Call Shell("cmd /c start tools\w" & xtbb & "\calc.lnk")
 ElseIf Text1.Text = "打开记事本" Then
  Call Shell("cmd /c start tools\w" & xtbb & "\note.lnk")
 ElseIf Text1.Text = "打开截图工具" Then
  Call Shell("cmd /c start tools\w" & xtbb & "\jtgj.lnk")
 ElseIf Left(Text1.Text, 2) = "修改" Or Left(Text1.Text, 2) = "更改" Then
  If Mid(Text1.Text, 3, 2) = "课表" Then
   Form27.Show
  ElseIf Mid(Text1.Text, 3, 2) = "名单" Then
   Call Shell("cmd /c start data\name.txt")
  ElseIf Mid(Text1.Text, 3, 4) = "作息时间" Then
   Form28.Show
  ElseIf Mid(Text1.Text, 3, 2) = "值日" Then
   Form29.Show
  ElseIf Mid(Text1.Text, 3, 4) = "临时值日" Then
   Form31.Show
  ElseIf Mid(Text1.Text, 3, 3) = "倒计时" Then
   Form30.Show
  ElseIf Mid(Text1.Text, 3, 4) = "U盘盘符" Then
   Form32.Show
  ElseIf Mid(Text1.Text, 3, 2) = "主题" Then
   Form34.Show
  ElseIf Mid(Text1.Text, 3, 3) = "透明度" Then
   Form35.Show
  End If
 Else
  Call Shell("cmd /c start ..\..\Desktop\" & Mid(Text1.Text, 3) & ".lnk")
 End If
 Text1.Text = ""
End If
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
Call Shell("cmd /c start tools\w" & xtbb & "\xj.lnk")
End Sub

Private Sub Command3_Click()
Call Shell("cmd /c start tools\w" & xtbb & "\calc.lnk")
End Sub

Private Sub Command4_Click()
Call Shell("cmd /c start tools\w" & xtbb & "\note.lnk")
End Sub

Private Sub Command5_Click()
Call Shell("cmd /c start tools\w" & xtbb & "\jtgj.lnk")
End Sub

Private Sub Form_Load()
Open "data\xtbb.txt" For Input As #1
Line Input #1, xt
Close #1
xtbb = Val(xt)
If textc = "1" Then
 Label1.ForeColor = vbWhite
 Label2.ForeColor = vbWhite
 Text1.ForeColor = vbWhite
Else
 Label1.ForeColor = vbBlack
 Label2.ForeColor = vbBlack
 Text1.ForeColor = vbBlack
End If
Me.Picture = LoadPicture("themes\" & theme & "\bg4.jpg")
Text1.BackColor = RGB(collist("bg22r"), collist("bg22g"), collist("bg22b"))
speed = 1
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = Screen.Height + 100
End Sub

Private Sub Timer1_Timer()
If speed > 0 Then
 If Me.Top - speed >= Screen.Height - (Me.Height - 1000) / 3 Then
  Me.Top = Me.Top - speed
  speed = speed + 20
 Else
  Me.Top = Me.Top - speed
  speed = speed - 5
 End If
Else
 Me.Top = Screen.Height - Me.Height - 1000
 speed = 1
 Timer1.Enabled = False
End If
End Sub

Public Function collist(c$) As String
Dim allcolor As String
Open "themes\" & theme & "\color.txt" For Input As #1
Line Input #1, allcolor
Close #1
collist = Trim(Mid(allcolor, InStr(allcolor, CStr(c)) + 5, 3))
End Function

Private Sub Timer2_Timer()
If Me.Top < Screen.Height Then
 Me.Top = Me.Top + speed
 speed = speed + 5
Else
 Me.Top = Screen.Height + 100
 Unload Me
 speed = 1
 kjzs = False
 Timer2.Enabled = False
End If
End Sub
