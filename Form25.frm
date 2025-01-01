VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form25 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "软件设置"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   25740
   Icon            =   "Form25.frx":0000
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   25740
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   10440
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9960
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9480
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9000
      Top             =   120
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "自动化"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4695
      Left            =   17160
      TabIndex        =   9
      Top             =   1560
      Width           =   8295
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "设置系统版本"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "拍照设置"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00FFFFC0&
         Caption         =   "设置自动关机"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00FFFFC0&
         Caption         =   "禁用自动关机"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00C0FFC0&
         Caption         =   "设置上课自动执行"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00C0FFC0&
         Caption         =   "设置下课自动执行"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "禁用自动打开U盘"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00C0C0FF&
         Caption         =   "设置时间控制"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H00C0C0FF&
         Caption         =   "设置自动任务"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "窗口外观"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4695
      Left            =   8640
      TabIndex        =   8
      Top             =   1560
      Width           =   8295
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "右"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "中"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "左"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "修改主题"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFC0FF&
         Caption         =   "修改透明度"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "----------窗口位置----------"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3040
         TabIndex        =   35
         Top             =   1200
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "基本设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   8295
      Begin VB.CommandButton Command16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "修改U盘盘符"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "修改课表"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "禁用双进程保护"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2320
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "禁用语音"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4280
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00C0FFC0&
         Caption         =   "修改倒计时"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "修改作息时间"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2320
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "改为特殊作息时间"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2320
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0FF&
         Caption         =   "修改名单"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C0FFFF&
         Caption         =   "修改值日"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4280
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "临时修改值日"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "重启软件"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改时间控制密码"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   6855
   End
   Begin VB.CommandButton Command28 
      Height          =   1215
      Left            =   120
      Picture         =   "Form25.frx":6988A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      Height          =   1215
      Left            =   1680
      Picture         =   "Form25.frx":6D4E7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command30 
      Height          =   1215
      Left            =   3240
      Picture         =   "Form25.frx":70ECC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command31 
      Height          =   1215
      Left            =   6960
      Picture         =   "Form25.frx":7428A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command32 
      Height          =   1215
      Left            =   5400
      Picture         =   "Form25.frx":783B9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   8520
      Top             =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   0
      X2              =   8520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   0
      X2              =   8520
      Y1              =   6360
      Y2              =   6360
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim udatopp As String
Dim zdgj As String
Dim tcp As String
Dim speed As Integer
Dim yleft As Integer
Dim cd As Integer

Private Sub Command1_Click()
Form34.Show
End Sub

Private Sub Command10_Click()
Form31.Show
End Sub

Private Sub Command11_Click()
If sjms = False Then
 leftt = Screen.Width - Form3.Width - 1805
 sxkg = False
 Form3.Left = Screen.Width - Form3.Width
 Form4.Left = Screen.Width - Form3.Width - 1805
 Form5.Left = Screen.Width - Form3.Width - 1805
 Form6.Left = Screen.Width - Form3.Width - 1805
 Form7.Left = Screen.Width - Form3.Width - 1805
 Form8.Left = Screen.Width - Form3.Width - 1805
 Form9.Left = Screen.Width - Form3.Width - 1805
 Form10.Left = Screen.Width - Form3.Width - 1805
 Form11.Left = Screen.Width - Form3.Width - 1805
 Form12.Left = Screen.Width - Form3.Width - 1805
 Form13.Left = Screen.Width - Form3.Width - 1805
 Form14.Left = Screen.Width - Form3.Width - 1805
 Form15.Left = Screen.Width - Form3.Width - 1805
 Form16.Left = Screen.Width - Form3.Width - 1805
 Form2.Left = Screen.Width - Form2.Width
 Form16.Width = 1805
 mode = False
End If
End Sub

Private Sub Command12_Click()
If dt = "1" Then
 dt = "0"
 Command12.Caption = "启用双进程保护"
 Open "data\dt.txt" For Output As #1
 Print #1, "0"
 Close #1
 MsgBox "修改后关闭软件再次打开生效!", 48, "提示"
Else
 dt = "1"
 Command12.Caption = "禁用双进程保护"
 Open "data\dt.txt" For Output As #1
 Print #1, "1"
 Close #1
 MsgBox "修改后关闭软件再次打开生效!", 48, "提示"
End If
End Sub

Private Sub Command13_Click()
Form35.Show
End Sub

Private Sub Command14_Click()
If aud = "1" Then
 aud = "0"
 Command14.Caption = "启用语音"
 Open "data\aud.txt" For Output As #1
 Print #1, "0"
 Close #1
Else
 aud = "1"
 Command14.Caption = "禁用语音"
 Open "data\aud.txt" For Output As #1
 Print #1, "1"
 Close #1
End If
End Sub

Private Sub Command15_Click()
Form28.Show
End Sub

Private Sub Command16_Click()
Form32.Show
End Sub

Private Sub Command17_Click()
Form24.Show
End Sub

Private Sub Command18_Click()
Form29.Show
End Sub

Private Sub Command19_Click()
Open "data\pswtc.txt" For Input As #1
Line Input #1, tcp
Close #1
If tcp = "" Then
 Form36.Show
Else
 sjmm = "1"
 Form38.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command20_Click()
Open "data\udatop.txt" For Input As #1
Line Input #1, udatopp
Close #1
If udatopp = "1" Then
 Open "data\udatop.txt" For Output As #1
 Print #1, "0"
 Close #1
 Command20.Caption = "启用自动打开U盘"
ElseIf udatopp = "0" Then
 Open "data\udatop.txt" For Output As #1
 Print #1, "1"
 Close #1
 Command20.Caption = "禁用自动打开U盘"
End If
End Sub

Private Sub Command21_Click()
If fricls = 1 Then
 fricls = 0
 Command21.Caption = "改为特殊作息时间"
Else
 fricls = 1
 Command21.Caption = "改为正常作息时间"
End If
End Sub

Private Sub Command22_Click()
Form39.Show
End Sub

Private Sub Command23_Click()
Form40.Show
End Sub

Private Sub Command24_Click()
Form37.Show
End Sub

Private Sub Command25_Click()
Form41.Show
End Sub

Private Sub Command26_Click()
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
If zdgj = "1" Then
 Open "data\zdgj.txt" For Output As #1
 Print #1, "0"
 Close #1
 Command26.Caption = "启用自动关机"
ElseIf zdgj = "0" Then
 Open "data\zdgj.txt" For Output As #1
 Print #1, "1"
 Close #1
 Command26.Caption = "禁用自动关机"
End If
End Sub

Private Sub Command27_Click()
Form30.Show
End Sub

Private Sub Command28_Click()
If Timer1.Enabled = False And Timer2.Enabled = False And Timer3.Enabled = False And Timer4.Enabled = False Then
 If cd = 2 Then
  yleft = Frame2.Left
  Timer2.Enabled = True
  cd = 1
 ElseIf cd = 3 Then
  yleft = Frame2.Left
  Timer4.Enabled = True
  cd = 1
 End If
End If
End Sub

Private Sub Command29_Click()
If Timer1.Enabled = False And Timer2.Enabled = False And Timer3.Enabled = False And Timer4.Enabled = False Then
 If cd = 1 Then
  yleft = Frame2.Left
  Timer1.Enabled = True
  cd = 2
 ElseIf cd = 3 Then
  yleft = Frame2.Left
  Timer2.Enabled = True
  cd = 2
 End If
End If
End Sub

Private Sub Command3_Click()
Form27.Show
End Sub

Private Sub Command30_Click()
If Timer1.Enabled = False And Timer2.Enabled = False And Timer3.Enabled = False And Timer4.Enabled = False Then
 If cd = 2 Then
  yleft = Frame2.Left
  Timer1.Enabled = True
  cd = 3
 ElseIf cd = 1 Then
  yleft = Frame2.Left
  Timer3.Enabled = True
  cd = 3
 End If
End If
End Sub

Private Sub Command31_Click()
Unload Form26
Form26.Show
End Sub

Private Sub Command32_Click()
Open "data\pswtc.txt" For Input As #1
Line Input #1, tcp
Close #1
If tcp = "" Then
 If aud = "1" Then
  WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\tc.mp3"
 End If
 Open "data\exit.txt" For Output As #1
 Print #1, "1"
 Close #1
 Timer5.Enabled = True
Else
 sjmm = "0"
 Form38.Show
End If
End Sub

Private Sub Command4_Click()
Form42.Show
End Sub

Private Sub Command5_Click()
If sjms = False Then
 leftt = 0
 sxkg = True
 Form3.Left = 1805
 Form4.Left = 0
 Form5.Left = 0
 Form6.Left = 0
 Form7.Left = 0
 Form8.Left = 0
 Form9.Left = 0
 Form10.Left = 0
 Form11.Left = 0
 Form12.Left = 0
 Form13.Left = 0
 Form14.Left = 0
 Form15.Left = 0
 Form16.Left = 0
 Form2.Left = 0
 Form16.Width = 1805
 mode = False
End If
End Sub

Private Sub Command6_Click()
If sjms = False Then
 leftt = (Screen.Width - Form2.Width) / 2
 sxkg = True
 Form3.Left = (Screen.Width - Form2.Width) / 2 + 1805
 Form4.Left = (Screen.Width - Form2.Width) / 2
 Form5.Left = (Screen.Width - Form2.Width) / 2
 Form6.Left = (Screen.Width - Form2.Width) / 2
 Form7.Left = (Screen.Width - Form2.Width) / 2
 Form8.Left = (Screen.Width - Form2.Width) / 2
 Form9.Left = (Screen.Width - Form2.Width) / 2
 Form10.Left = (Screen.Width - Form2.Width) / 2
 Form11.Left = (Screen.Width - Form2.Width) / 2
 Form12.Left = (Screen.Width - Form2.Width) / 2
 Form13.Left = (Screen.Width - Form2.Width) / 2
 Form14.Left = (Screen.Width - Form2.Width) / 2
 Form15.Left = (Screen.Width - Form2.Width) / 2
 Form16.Left = (Screen.Width - Form2.Width) / 2
 Form2.Left = (Screen.Width - Form2.Width) / 2
 Form16.Width = 1805
 mode = False
End If
End Sub

Private Sub Command7_Click()
If dtrs = "1" Then
 Call Shell("cmd /c start killme.bat")
 Unload Me
ElseIf dtrs = "0" Then
 Call Shell("cmd /c start rstme.bat")
 Unload Me
End If
End Sub

Private Sub Command8_Click()
Form33.Show
End Sub

Private Sub Command9_Click()
Call Shell("cmd /c start data\name.txt")
End Sub

Private Sub Form_Load()
Me.Height = 7530
Me.Width = 8610
speed = 1
kbdh = True
cd = 1
Dim udatop As String
Open "data\udatop.txt" For Input As #1
Line Input #1, udatop
Close #1
If udatop = "1" Then
 Command20.Caption = "禁用自动打开U盘"
ElseIf udatop = "0" Then
 Command20.Caption = "启用自动打开U盘"
End If
If fricls = 0 Then
 Command21.Caption = "改为特殊作息时间"
Else
 Command21.Caption = "改为正常作息时间"
End If
Dim zdgj As String
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
If zdgj = "1" Then
 Command26.Caption = "禁用自动关机"
ElseIf zdgj = "0" Then
 Command26.Caption = "启用自动关机"
End If
If dt = "1" Then
 Command12.Caption = "禁用双进程保护"
Else
 Command12.Caption = "启用双进程保护"
End If
If aud = "1" Then
 Command14.Caption = "禁用语音"
Else
 Command14.Caption = "启用语音"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not Form2.Left = 0 Then
 kbdh = False
End If
End Sub

Private Sub Timer1_Timer()
If speed > 0 Then
 If Frame2.Left - speed >= yleft - (8295 / 3) Then
  Frame1.Left = Frame1.Left - speed
  Frame2.Left = Frame2.Left - speed
  Frame3.Left = Frame3.Left - speed
  speed = speed + 18
 Else
  Frame1.Left = Frame1.Left - speed
  Frame2.Left = Frame2.Left - speed
  Frame3.Left = Frame3.Left - speed
  speed = speed - 8
 End If
Else
 speed = 1
 Frame2.Left = yleft - 8520
 Frame1.Left = Frame2.Left - 8520
 Frame3.Left = Frame2.Left + 8520
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If speed > 0 Then
 If Frame2.Left + speed <= yleft + (8295 / 3) Then
  Frame1.Left = Frame1.Left + speed
  Frame2.Left = Frame2.Left + speed
  Frame3.Left = Frame3.Left + speed
  speed = speed + 18
 Else
  Frame1.Left = Frame1.Left + speed
  Frame2.Left = Frame2.Left + speed
  Frame3.Left = Frame3.Left + speed
  speed = speed - 8
 End If
Else
 speed = 1
 Frame2.Left = yleft + 8520
 Frame1.Left = Frame2.Left - 8520
 Frame3.Left = Frame2.Left + 8520
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If speed > 0 Then
 If Frame2.Left - speed >= yleft - (17040 / 3) Then
  Frame1.Left = Frame1.Left - speed
  Frame2.Left = Frame2.Left - speed
  Frame3.Left = Frame3.Left - speed
  speed = speed + 41
 Else
  Frame1.Left = Frame1.Left - speed
  Frame2.Left = Frame2.Left - speed
  Frame3.Left = Frame3.Left - speed
  speed = speed - 22
 End If
Else
 speed = 1
 Frame2.Left = yleft - 17040
 Frame1.Left = Frame2.Left - 8520
 Frame3.Left = Frame2.Left + 8520
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If speed > 0 Then
 If Frame2.Left + speed <= yleft + (17040 / 3) Then
  Frame1.Left = Frame1.Left + speed
  Frame2.Left = Frame2.Left + speed
  Frame3.Left = Frame3.Left + speed
  speed = speed + 41
 Else
  Frame1.Left = Frame1.Left + speed
  Frame2.Left = Frame2.Left + speed
  Frame3.Left = Frame3.Left + speed
  speed = speed - 22
 End If
Else
 speed = 1
 Frame2.Left = yleft + 17040
 Frame1.Left = Frame2.Left - 8520
 Frame3.Left = Frame2.Left + 8520
 Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
If Not Form2.Left > Screen.Width + 100 Then
 Form2.Left = Form2.Left + speed
 Form3.Left = Form3.Left + speed
 Form4.Left = Form4.Left + speed
 Form5.Left = Form5.Left + speed
 Form6.Left = Form6.Left + speed
 Form7.Left = Form7.Left + speed
 Form8.Left = Form8.Left + speed
 Form9.Left = Form9.Left + speed
 Form10.Left = Form10.Left + speed
 Form11.Left = Form11.Left + speed
 Form12.Left = Form12.Left + speed
 Form13.Left = Form13.Left + speed
 Form14.Left = Form14.Left + speed
 Form15.Left = Form15.Left + speed
 Form16.Left = Form16.Left + speed
 speed = speed + 7
Else
 Call Shell("cmd /c start killme.bat")
 Unload Me
End If
End Sub
