VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3375
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3120
      Top             =   2040
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   2640
      Top             =   2040
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2160
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1200
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   720
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Interval        =   2
      Left            =   1680
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   2040
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "888"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1000
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "88"
      Top             =   2520
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1380
      TabIndex        =   13
      Top             =   2640
      Width           =   285
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1380
      TabIndex        =   12
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1380
      TabIndex        =   11
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "节课上课还有"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1740
      TabIndex        =   9
      Top             =   2610
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "距第"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2610
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "分钟"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   2610
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2000/01/01"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "星期一"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ":    :"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   1800
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
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
Dim speed As Integer
Dim tomd As Integer
Dim lleft As Integer
Dim jzkb As Integer

Private Sub Form_DblClick()
If sxkg = False Then
 kbdh = True
 If Me.Left < Screen.Width - Me.Width / 2 Then
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   sjms = True
   Label9.Visible = False
   Timer7.Enabled = True
  End If
 Else
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   Label10.Visible = False
   Label11.Visible = False
   Label12.Visible = False
   Timer6.Enabled = True
  End If
 End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = Screen.Width
lleft = Screen.Width - Me.Width
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
If textc = "1" Then
 Label1.ForeColor = vbWhite
 Label2.ForeColor = vbWhite
 Label3.ForeColor = vbWhite
 Label4.ForeColor = vbWhite
 Label5.ForeColor = vbWhite
 Label6.ForeColor = vbWhite
 Label7.ForeColor = vbWhite
 Label8.ForeColor = vbWhite
 Label9.ForeColor = vbWhite
 Label10.ForeColor = vbWhite
 Label11.ForeColor = vbWhite
 Label12.ForeColor = vbWhite
 Text1.ForeColor = vbWhite
 Text4.ForeColor = vbWhite
Else
 Label1.ForeColor = vbBlack
 Label2.ForeColor = vbBlack
 Label3.ForeColor = vbBlack
 Label4.ForeColor = vbBlack
 Label5.ForeColor = vbBlack
 Label6.ForeColor = vbBlack
 Label7.ForeColor = vbBlack
 Label8.ForeColor = vbBlack
 Label9.ForeColor = vbBlack
 Label10.ForeColor = vbBlack
 Label11.ForeColor = vbBlack
 Label12.ForeColor = vbBlack
 Text1.ForeColor = vbBlack
 Text4.ForeColor = vbBlack
End If
Me.Picture = LoadPicture("themes\" & theme & "\bg1.jpg")
Text1.BackColor = RGB(collist("bg11r"), collist("bg11g"), collist("bg11b"))
Text4.BackColor = RGB(collist("bg11r"), collist("bg11g"), collist("bg11b"))
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
kbdh = False
sxkg = False
sjms = False
caidan = False
tomd = 255
speed = 1
Form3.Show
Form3.Top = Form3.Height * -1 - 50
Form3.Left = lleft + 1795
jzkb = 4
Timer3.Enabled = False
Timer4.Enabled = True
Open "data\exit.txt" For Output As #1
Print #1, "0"
Close #1
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 If Weekday(Date, 2) = 5 Then
  fricls = 1
 Else
  fricls = 0
 End If
ElseIf fxs = "6" Then
 If Weekday(Date, 2) = 6 Then
  fricls = 1
 Else
  fricls = 0
 End If
ElseIf fxs = "7" Then
 If Weekday(Date, 2) = 7 Then
  fricls = 1
 Else
  fricls = 0
 End If
End If
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
Label2.Caption = Date
Label3.Caption = "星期" & " " & Weekday(Date, 2)
If Time < CDate(zxtime("zm1s")) Then
 Text1.Text = "1"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm1s")) - Time) + Hour(CDate(zxtime("zm1s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
 Text1.Text = "1"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm1x")) - Time) + Hour(CDate(zxtime("zm1x")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
 Text1.Text = "2"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm2s")) - Time) + Hour(CDate(zxtime("zm2s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
 Text1.Text = "2"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm2x")) - Time) + Hour(CDate(zxtime("zm2x")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
 Text1.Text = "3"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm3s")) - Time) + Hour(CDate(zxtime("zm3s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
 Text1.Text = "3"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm3x")) - Time) + Hour(CDate(zxtime("zm3x")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
 Text1.Text = "4"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm4s")) - Time) + Hour(CDate(zxtime("zm4s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
 Text1.Text = "4"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm4x")) - Time) + Hour(CDate(zxtime("zm4x")) - Time) * 60 + 1
Else
 If fricls = 0 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zm5s")) Then
   Text1.Text = "5"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm5s")) - Time) + Hour(CDate(zxtime("zm5s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm5s")) <= Time And Time < CDate(zxtime("zm5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm5x")) - Time) + Hour(CDate(zxtime("zm5x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm5x")) <= Time And Time < CDate(zxtime("zm6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm6s")) - Time) + Hour(CDate(zxtime("zm6s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm6s")) <= Time And Time < CDate(zxtime("zm6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm6x")) - Time) + Hour(CDate(zxtime("zm6x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm6x")) <= Time And Time < CDate(zxtime("zm7s")) Then
   Text1.Text = "7"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm7s")) - Time) + Hour(CDate(zxtime("zm7s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm7s")) <= Time And Time < CDate(zxtime("zm7x")) Then
   Text1.Text = "7"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm7x")) - Time) + Hour(CDate(zxtime("zm7x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm7x")) <= Time And Time < CDate(zxtime("zm8s")) Then
   Text1.Text = "8"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm8s")) - Time) + Hour(CDate(zxtime("zm8s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm8s")) <= Time And Time < CDate(zxtime("zm8x")) Then
   Text1.Text = "8"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm8x")) - Time) + Hour(CDate(zxtime("zm8x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm8x")) <= Time And Time < CDate(zxtime("zm9s")) Then
   Text1.Text = "9"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm9s")) - Time) + Hour(CDate(zxtime("zm9s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm9s")) <= Time And Time < CDate(zxtime("zm9x")) Then
   Text1.Text = "9"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm9x")) - Time) + Hour(CDate(zxtime("zm9x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm9x")) <= Time And Time < CDate(zxtime("zm10s")) Then
   Text1.Text = "10"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm10s")) - Time) + Hour(CDate(zxtime("zm10s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm10s")) <= Time And Time < CDate(zxtime("zm10x")) Then
   Text1.Text = "10"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm10x")) - Time) + Hour(CDate(zxtime("zm10x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm10x")) <= Time And Time < CDate(zxtime("zm11s")) Then
   Text1.Text = "11"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm11s")) - Time) + Hour(CDate(zxtime("zm11s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm11s")) <= Time And Time < CDate(zxtime("zm11x")) Then
   Text1.Text = "11"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm11x")) - Time) + Hour(CDate(zxtime("zm11x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm11x")) <= Time And Time < CDate(zxtime("zm12s")) Then
   Text1.Text = "12"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm12s")) - Time) + Hour(CDate(zxtime("zm12s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm12s")) <= Time And Time < CDate(zxtime("zm12x")) Then
   Text1.Text = "12"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm12x")) - Time) + Hour(CDate(zxtime("zm12x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm12x")) < Time Then
   Text1.Text = "12"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 Else
  If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zf5s")) Then
   Text1.Text = "5"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zf5s")) - Time) + Hour(CDate(zxtime("zf5s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf5s")) <= Time And Time < CDate(zxtime("zf5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf5x")) - Time) + Hour(CDate(zxtime("zf5x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf5x")) <= Time And Time < CDate(zxtime("zf6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zf6s")) - Time) + Hour(CDate(zxtime("zf6s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf6s")) <= Time And Time < CDate(zxtime("zf6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf6x")) - Time) + Hour(CDate(zxtime("zf6x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf6x")) < Time Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 End If
End If
End Sub

Private Sub Label1_DblClick()
If sxkg = False Then
 kbdh = True
 If Me.Left < Screen.Width - Me.Width / 2 Then
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   sjms = True
   Label9.Visible = False
   Timer7.Enabled = True
  End If
 Else
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   Label10.Visible = False
   Label11.Visible = False
   Label12.Visible = False
   Timer6.Enabled = True
  End If
 End If
End If
End Sub

Private Sub Label5_DblClick()
If sxkg = False Then
 kbdh = True
 If Me.Left < Screen.Width - Me.Width / 2 Then
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   sjms = True
   Label9.Visible = False
   Timer7.Enabled = True
  End If
 Else
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   Label10.Visible = False
   Label11.Visible = False
   Label12.Visible = False
   Timer6.Enabled = True
  End If
 End If
End If
End Sub

Private Sub Label6_DblClick()
If sxkg = False Then
 kbdh = True
 If Me.Left < Screen.Width - Me.Width / 2 Then
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   sjms = True
   Label9.Visible = False
   Timer7.Enabled = True
  End If
 Else
  If Timer6.Enabled = False And Timer7.Enabled = False Then
   Label10.Visible = False
   Label11.Visible = False
   Label12.Visible = False
   Timer6.Enabled = True
  End If
 End If
End If
End Sub

Private Sub Timer1_Timer()
If Len(Hour(Time)) = 1 Then
 Label1.Caption = "0" & Hour(Time)
 h = "0" & Hour(Time)
Else
 Label1.Caption = Hour(Time)
 h = Hour(Time)
End If
If Len(Minute(Time)) = 1 Then
 Label5.Caption = "0" & Minute(Time)
 m = "0" & Minute(Time)
Else
 Label5.Caption = Minute(Time)
 m = Minute(Time)
End If
If Len(Second(Time)) = 1 Then
 Label6.Caption = "0" & Second(Time)
 s = "0" & Second(Time)
Else
 Label6.Caption = Second(Time)
 s = Second(Time) + 3
End If
Label2.Caption = Date
If Weekday(Date, 2) = 1 Then
 Label3.Caption = "星期一"
ElseIf Weekday(Date, 2) = 2 Then
 Label3.Caption = "星期二"
ElseIf Weekday(Date, 2) = 3 Then
 Label3.Caption = "星期三"
ElseIf Weekday(Date, 2) = 4 Then
 Label3.Caption = "星期四"
ElseIf Weekday(Date, 2) = 5 Then
 Label3.Caption = "星期五"
ElseIf Weekday(Date, 2) = 6 Then
 Label3.Caption = "星期六"
ElseIf Weekday(Date, 2) = 7 Then
 Label3.Caption = "星期日"
End If
End Sub

Private Sub Timer2_Timer()
If Time < CDate(zxtime("zm1s")) Then
 Text1.Text = "1"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm1s")) - Time) + Hour(CDate(zxtime("zm1s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
 Text1.Text = "1"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm1x")) - Time) + Hour(CDate(zxtime("zm1x")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
 Text1.Text = "2"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm2s")) - Time) + Hour(CDate(zxtime("zm2s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
 Text1.Text = "2"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm2x")) - Time) + Hour(CDate(zxtime("zm2x")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
 Text1.Text = "3"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm3s")) - Time) + Hour(CDate(zxtime("zm3s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
 Text1.Text = "3"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm3x")) - Time) + Hour(CDate(zxtime("zm3x")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
 Text1.Text = "4"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm4s")) - Time) + Hour(CDate(zxtime("zm4s")) - Time) * 60 + 1
ElseIf CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
 Text1.Text = "4"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm4x")) - Time) + Hour(CDate(zxtime("zm4x")) - Time) * 60 + 1
Else
 If fricls = 0 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zm5s")) Then
   Text1.Text = "5"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm5s")) - Time) + Hour(CDate(zxtime("zm5s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm5s")) <= Time And Time < CDate(zxtime("zm5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm5x")) - Time) + Hour(CDate(zxtime("zm5x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm5x")) <= Time And Time < CDate(zxtime("zm6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm6s")) - Time) + Hour(CDate(zxtime("zm6s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm6s")) <= Time And Time < CDate(zxtime("zm6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm6x")) - Time) + Hour(CDate(zxtime("zm6x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm6x")) <= Time And Time < CDate(zxtime("zm7s")) Then
   Text1.Text = "7"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm7s")) - Time) + Hour(CDate(zxtime("zm7s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm7s")) <= Time And Time < CDate(zxtime("zm7x")) Then
   Text1.Text = "7"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm7x")) - Time) + Hour(CDate(zxtime("zm7x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm7x")) <= Time And Time < CDate(zxtime("zm8s")) Then
   Text1.Text = "8"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm8s")) - Time) + Hour(CDate(zxtime("zm8s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm8s")) <= Time And Time < CDate(zxtime("zm8x")) Then
   Text1.Text = "8"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm8x")) - Time) + Hour(CDate(zxtime("zm8x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm8x")) <= Time And Time < CDate(zxtime("zm9s")) Then
   Text1.Text = "9"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm9s")) - Time) + Hour(CDate(zxtime("zm9s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm9s")) <= Time And Time < CDate(zxtime("zm9x")) Then
   Text1.Text = "9"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm9x")) - Time) + Hour(CDate(zxtime("zm9x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm9x")) <= Time And Time < CDate(zxtime("zm10s")) Then
   Text1.Text = "10"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm10s")) - Time) + Hour(CDate(zxtime("zm10s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm10s")) <= Time And Time < CDate(zxtime("zm10x")) Then
   Text1.Text = "10"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm10x")) - Time) + Hour(CDate(zxtime("zm10x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm10x")) <= Time And Time < CDate(zxtime("zm11s")) Then
   Text1.Text = "11"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm11s")) - Time) + Hour(CDate(zxtime("zm11s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm11s")) <= Time And Time < CDate(zxtime("zm11x")) Then
   Text1.Text = "11"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm11x")) - Time) + Hour(CDate(zxtime("zm11x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm11x")) <= Time And Time < CDate(zxtime("zm12s")) Then
   Text1.Text = "12"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm12s")) - Time) + Hour(CDate(zxtime("zm12s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm12s")) <= Time And Time < CDate(zxtime("zm12x")) Then
   Text1.Text = "12"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm12x")) - Time) + Hour(CDate(zxtime("zm12x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zm12x")) < Time Then
   Text1.Text = "12"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 Else
  If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zf5s")) Then
   Text1.Text = "5"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zf5s")) - Time) + Hour(CDate(zxtime("zf5s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf5s")) <= Time And Time < CDate(zxtime("zf5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf5x")) - Time) + Hour(CDate(zxtime("zf5x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf5x")) <= Time And Time < CDate(zxtime("zf6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zf6s")) - Time) + Hour(CDate(zxtime("zf6s")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf6s")) <= Time And Time < CDate(zxtime("zf6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf6x")) - Time) + Hour(CDate(zxtime("zf6x")) - Time) * 60 + 1
  ElseIf CDate(zxtime("zf6x")) < Time Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 End If
End If
End Sub

Private Sub Timer3_Timer()
If speed > 0 Then
 If Not Form3.Top + speed >= Me.Height / 3 + 140 Then
  Form3.Top = Form3.Top + speed
  speed = speed + 5
 Else
  Form3.Top = Form3.Top + speed
  speed = speed - 27
 End If
 If Form3.Top > 0 Then
  SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255 - Form3.Top \ (Me.Height / (256 - tmdu)), LWA_ALPHA
  toumd = 255 - Form3.Top \ (Me.Height / (256 - tmdu))
 Else
  SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
  toumd = 255
 End If
Else
 Form3.Top = Me.Height
 speed = 1
 Timer5.Enabled = True
 Timer4.Enabled = False
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If speed > 0 Then
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), (Screen.Width - Me.Left) * 0.0422, LWA_ALPHA
 tomd = (Screen.Width - Me.Left) * 0.0422
 If Me.Left - speed >= lleft + (Me.Width / 3) Then
  Me.Left = Me.Left - speed
  speed = speed + 5
 Else
  Me.Left = Me.Left - speed
  speed = speed - 10.5
 End If
Else
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 tomd = 255
 Me.Left = lleft
 leftt = Me.Left
 Form3.Top = Form3.Height * -1 - 100
 Form3.Left = lleft + 1795
 speed = 1
 Timer3.Enabled = True
 Timer4.Enabled = False
End If
End Sub

Public Function zxtime(s$) As String
Dim timelist As String
Open "data\timelist.txt" For Input As #1
Line Input #1, timelist
Close #1
If Len(s) = 4 Then
 zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 4, 9))
ElseIf Len(s) = 5 Then
 zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 5, 9))
End If
End Function

Private Sub Timer5_Timer()
If jzkb = 4 Then '
 Form4.Show
ElseIf jzkb = 5 Then
 Form5.Show
ElseIf jzkb = 6 Then
 Form6.Show
ElseIf jzkb = 7 Then
 Form7.Show
ElseIf jzkb = 8 Then
 Form8.Show
ElseIf jzkb = 9 Then
 Form9.Show
ElseIf jzkb = 10 Then
 Form10.Show
ElseIf jzkb = 11 Then
 Form11.Show
ElseIf jzkb = 12 Then
 Form12.Show
ElseIf jzkb = 13 Then
 Form13.Show
ElseIf jzkb = 14 Then
 Form14.Show
ElseIf jzkb = 15 Then
 Form15.Show
ElseIf jzkb = 16 Then
 Form16.Show
Else
 Unload Form1
 Timer3.Enabled = False
 Timer4.Enabled = False
 Timer5.Enabled = False
End If
jzkb = jzkb + 1
End Sub

Public Function collist(c$) As String
Dim allcolor As String
Open "themes\" & theme & "\color.txt" For Input As #1
Line Input #1, allcolor
Close #1
collist = Trim(Mid(allcolor, InStr(allcolor, CStr(c)) + 5, 3))
End Function

Private Sub Timer6_Timer()
If speed > 0 Then
 If Form3.Left - speed >= Screen.Width - (Form3.Width / 3) Then
  Form3.Left = Form3.Left - speed
  Form4.Left = Form4.Left - speed
  Form5.Left = Form5.Left - speed
  Form6.Left = Form6.Left - speed
  Form7.Left = Form7.Left - speed
  Form8.Left = Form8.Left - speed
  Form9.Left = Form9.Left - speed
  Form10.Left = Form10.Left - speed
  Form11.Left = Form11.Left - speed
  Form12.Left = Form12.Left - speed
  Form13.Left = Form13.Left - speed
  Form14.Left = Form14.Left - speed
  Form15.Left = Form15.Left - speed
  Form16.Left = Form16.Left - speed
  Me.Left = Me.Left - speed
  Label1.Top = Label1.Top + speed / 5.5
  Label5.Top = Label5.Top - speed / 14
  Label5.Left = Label5.Left + speed / 2.07
  Label6.Top = Label6.Top - speed / 3
  Label6.Left = Label6.Left + speed / 1.04
  Label1.FontSize = Label1.FontSize + CInt(speed / 130)
  Label5.FontSize = Label5.FontSize + CInt(speed / 130)
  Label6.FontSize = Label6.FontSize + CInt(speed / 130)
  Label2.Top = Label2.Top + speed / 5
  Label3.Top = Label3.Top + speed / 5
  Label4.Top = Label4.Top - speed / 4
  Label7.Top = Label7.Top - speed / 4
  Label8.Top = Label8.Top - speed / 4
  Text1.Top = Text1.Top - speed / 4
  Text4.Top = Text4.Top - speed / 4
  speed = speed + 4
 Else
  Form3.Left = Form3.Left - speed
  Form4.Left = Form4.Left - speed
  Form5.Left = Form5.Left - speed
  Form6.Left = Form6.Left - speed
  Form7.Left = Form7.Left - speed
  Form8.Left = Form8.Left - speed
  Form9.Left = Form9.Left - speed
  Form10.Left = Form10.Left - speed
  Form11.Left = Form11.Left - speed
  Form12.Left = Form12.Left - speed
  Form13.Left = Form13.Left - speed
  Form14.Left = Form14.Left - speed
  Form15.Left = Form15.Left - speed
  Form16.Left = Form16.Left - speed
  Me.Left = Me.Left - speed
  Label1.Top = Label1.Top + speed / 5.5
  Label5.Top = Label5.Top - speed / 14
  Label5.Left = Label5.Left + speed / 2.07
  Label6.Top = Label6.Top - speed / 3
  Label6.Left = Label6.Left + speed / 1.04
  Label1.FontSize = Label1.FontSize + CInt(speed / 130)
  Label5.FontSize = Label5.FontSize + CInt(speed / 130)
  Label6.FontSize = Label6.FontSize + CInt(speed / 130)
  Label2.Top = Label2.Top + speed / 5
  Label3.Top = Label3.Top + speed / 5
  Label4.Top = Label4.Top - speed / 4
  Label7.Top = Label7.Top - speed / 4
  Label8.Top = Label8.Top - speed / 4
  Text1.Top = Text1.Top - speed / 4
  Text4.Top = Text4.Top - speed / 4
  speed = speed - 1.5
 End If
Else
 speed = 1
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
 Me.Left = Screen.Width - Me.Width
 Label1.Top = 720
 Label5.Top = 720
 Label5.Left = 2160
 Label6.Top = 720
 Label6.Left = 4200
 Label1.FontSize = 72
 Label5.FontSize = 72
 Label6.FontSize = 72
 Label2.Top = 120
 Label3.Top = 120
 Label4.Top = 2610
 Label7.Top = 2610
 Label8.Top = 2610
 Text1.Top = 2520
 Text4.Top = 2520
 Label9.Visible = True
 Form16.Width = 1805
 mode = False
 kbdh = False
 sjms = False
 Timer6.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()
If speed > 0 Then
 If Form3.Left + speed <= Screen.Width - (Form3.Width / 3) Then
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
  Me.Left = Me.Left + speed
  Label1.Top = Label1.Top - speed / 5.5
  Label5.Top = Label5.Top + speed / 14
  Label5.Left = Label5.Left - speed / 2.07
  Label6.Top = Label6.Top + speed / 3
  Label6.Left = Label6.Left - speed / 1.04
  Label1.FontSize = Label1.FontSize - CInt(speed / 130)
  Label5.FontSize = Label5.FontSize - CInt(speed / 130)
  Label6.FontSize = Label6.FontSize - CInt(speed / 130)
  Label2.Top = Label2.Top - speed / 5
  Label3.Top = Label3.Top - speed / 5
  Label4.Top = Label4.Top + speed / 4
  Label7.Top = Label7.Top + speed / 4
  Label8.Top = Label8.Top + speed / 4
  Text1.Top = Text1.Top + speed / 4
  Text4.Top = Text4.Top + speed / 4
  speed = speed + 1.5
 Else
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
  Me.Left = Me.Left + speed
  Label1.Top = Label1.Top - speed / 5.5
  Label5.Top = Label5.Top + speed / 14
  Label5.Left = Label5.Left - speed / 2.07
  Label6.Top = Label6.Top + speed / 3
  Label6.Left = Label6.Left - speed / 1.04
  Label1.FontSize = Label1.FontSize - CInt(speed / 130)
  Label5.FontSize = Label5.FontSize - CInt(speed / 130)
  Label6.FontSize = Label6.FontSize - CInt(speed / 130)
  Label2.Top = Label2.Top - speed / 5
  Label3.Top = Label3.Top - speed / 5
  Label4.Top = Label4.Top + speed / 4
  Label7.Top = Label7.Top + speed / 4
  Label8.Top = Label8.Top + speed / 4
  Text1.Top = Text1.Top + speed / 4
  Text4.Top = Text4.Top + speed / 4
  speed = speed - 4
 End If
Else
 speed = 1
 Form3.Left = Screen.Width
 Form4.Left = Screen.Width - 1805
 Form5.Left = Screen.Width - 1805
 Form6.Left = Screen.Width - 1805
 Form7.Left = Screen.Width - 1805
 Form8.Left = Screen.Width - 1805
 Form9.Left = Screen.Width - 1805
 Form10.Left = Screen.Width - 1805
 Form11.Left = Screen.Width - 1805
 Form12.Left = Screen.Width - 1805
 Form13.Left = Screen.Width - 1805
 Form14.Left = Screen.Width - 1805
 Form15.Left = Screen.Width - 1805
 Form16.Left = Screen.Width - 1805
 Me.Left = Screen.Width - (Form2.Width - Form3.Width)
 Label1.Left = 120
 Label5.Left = 120
 Label6.Left = 120
 Label10.Visible = True
 Label11.Visible = True
 Label12.Visible = True
 Form16.Width = 1805
 mode = False
 kbdh = False
 Timer7.Enabled = False
End If
End Sub
