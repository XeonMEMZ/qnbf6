VERSION 5.00
Begin VB.Form Form35 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改透明度"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form35.frx":0000
   LinkTopic       =   "Form35"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   2040
      Max             =   255
      Min             =   10
      TabIndex        =   2
      Top             =   1200
      Value           =   10
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "恢复默认"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "修改透明度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1380
      TabIndex        =   6
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "透明度:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldtmd As Integer
Dim re As Integer

Private Sub Command1_Click()
Open "data\tmdu.txt" For Output As #1
Print #1, HScroll1.Value
Close #1
re = MsgBox("修改透明度后重启软件生效!" & vbCrLf & "是否重启软件?", 52, "提示")
If re = 6 Then
 If dtrs = "1" Then
  Call Shell("cmd /c start killme.bat")
  Unload Me
 ElseIf dtrs = "0" Then
  Call Shell("cmd /c start rstme.bat")
  Unload Me
 End If
Else
 Unload Me
End If
End Sub

Private Sub Command2_Click()
tmdu = oldtmd
Unload Me
End Sub

Private Sub Command3_Click()
HScroll1.Value = 230
Text1.Text = 230
tmdu = 230
End Sub

Private Sub Form_Load()
oldtmd = tmdu
Dim tmd As String
Open "data\tmdu.txt" For Input As #1
Line Input #1, tmd
Close #1
Text1.Text = tmd
HScroll1.Value = tmd
End Sub

Private Sub Form_Unload(Cancel As Integer)
tmdu = oldtmd
End Sub

Private Sub HScroll1_change()
Text1.Text = HScroll1.Value
tmdu = HScroll1.Value
End Sub

Private Sub HScroll1_scroll()
Text1.Text = HScroll1.Value
tmdu = HScroll1.Value
End Sub

