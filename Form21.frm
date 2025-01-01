VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H004D5E2B&
   BorderStyle     =   0  'None
   Caption         =   "Form21"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16005
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1560
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   14790
      Picture         =   "Form21.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   540
      TabIndex        =   10
      Top             =   7710
      Width           =   540
   End
   Begin VB.CommandButton Command3 
      Caption         =   "≈ƒ’’"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   7560
      TabIndex        =   5
      Top             =   6600
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÀÊª˙µ„√˚"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1080
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   4380
      TabIndex        =   3
      Text            =   "ø∆ƒø"
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   6900
      TabIndex        =   2
      Text            =   " ±º‰"
      Top             =   4800
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   4380
      TabIndex        =   1
      Text            =   "ø∆ƒø"
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   6900
      TabIndex        =   0
      Text            =   " ±º‰"
      Top             =   5640
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "∂® ±◊‘∂Ø≈ƒ’’“—∆Ù”√"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   4020
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 ≈ƒ…„"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   12420
      TabIndex        =   9
      Top             =   5400
      Width           =   3330
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 ≈ƒ…„"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   300
      TabIndex        =   8
      Top             =   5400
      Width           =   3330
   End
   Begin VB.Image Image2 
      Height          =   2205
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3915
   End
   Begin VB.Image Image1 
      Height          =   2205
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   99.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2265
      Left            =   4020
      TabIndex        =   7
      Top             =   2280
      Width           =   8055
   End
End
Attribute VB_Name = "Form21"
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
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objshell As Object
Option Explicit
Dim speed As Integer
Dim tmd As Integer
Dim ctimg As Integer
Dim sj As Integer
Dim zdpzsj As Integer
Dim zdpz As String
Dim zdpzs As String
Dim cname As String
Dim piczt As String
Dim myWindow

Private Sub Command2_Click()
Open "data\namec.txt" For Input As #1
Line Input #1, cname
Close #1
Dim sj As Integer
Randomize
sj = Int(Rnd * (CInt(cname) - 1 + 1) + 1)
Text5.Text = namelist(sj)
End Sub

Private Sub Command3_Click()
Set objshell = CreateObject("Wscript.Shell")
Call Shell("cmd /c start capt.exe")
Sleep 500
While exitproc("capt.exe")
Wend
Sleep 500
Open "data\zt.txt" For Input As #1
Line Input #1, piczt
Close #1
If piczt = "1" Then
 Open "data\picname.txt" For Input As #1
 Line Input #1, imgname
 Close #1
 If Not ctimg Mod 2 = 0 Then
  Image1.Picture = LoadPicture("capture\" & imgname)
  Label2.Caption = Label1.Caption & " ≈ƒ…„"
  ctimg = ctimg + 1
 Else
  Image2.Picture = LoadPicture("capture\" & imgname)
  Label3.Caption = Label1.Caption & " ≈ƒ…„"
  ctimg = ctimg + 1
 End If
End If
myWindow = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture("themes\" & theme & "\bg5.jpg")
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
myWindow = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
speed = 1
tmd = 0
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
Label1.Top = Me.Height / 2 - 2220
Label1.Left = (Me.Width - Label1.Width) / 2
Text1.Top = Label1.Top + 2520
Text1.Left = Label1.Left + 360
Text2.Top = Label1.Top + 2520
Text2.Left = Label1.Left + 2880
Text3.Top = Label1.Top + 3360
Text3.Left = Label1.Left + 360
Text4.Top = Label1.Top + 3360
Text4.Left = Label1.Left + 2880
Command2.Top = Label1.Top + 4440
Command2.Left = Label1.Left + 1140
Command3.Top = Label1.Top + 2880
Command3.Left = Label1.Left + 2820
Text5.Top = Label1.Top + 4320
Text5.Left = Label1.Left + 3540
Image1.Top = Me.Height / 2 - 1380
Image1.Left = 0
Image1.Height = (Label1.Left - 105) / 16 * 8
Image1.Width = Label1.Left - 105
Image2.Top = Me.Height / 2 - 1380
Image2.Left = Label1.Left + 8100
Image2.Height = (Label1.Left - 105) / 16 * 8
Image2.Width = Label1.Left - 105
Label2.Top = Image1.Top + Image1.Height + 75
Label2.Left = (Image1.Width - Label2.Width) / 2
Label3.Top = Image2.Top + Image2.Height + 75
Label3.Left = Image2.Left + (Image2.Width - Label2.Width) / 2
Label4.Top = Label1.Top - 1380
Label4.Left = Label1.Left
Picture1.Top = Me.Height - 1535
Picture1.Left = Me.Width - 1535
sj = 0
Open "data\zdpz.txt" For Input As #1
Line Input #1, zdpz
Close #1
Open "data\zdpzsj.txt" For Input As #1
Line Input #1, zdpzs
Close #1
zdpzsj = Val(zdpzs)
If zxms = True Then
 Text1.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
 If zdpz = "1" Then
  Label4.Visible = True
  Command3.Enabled = False
  Timer4.Enabled = True
 End If
ElseIf zdms = True Then
 Text1.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
 Set objshell = CreateObject("Wscript.Shell")
 Call Shell("cmd /c start capt.exe")
 Sleep 500
 While exitproc("capt.exe")
 Wend
 Sleep 500
 Open "data\zt.txt" For Input As #1
 Line Input #1, piczt
 Close #1
 If piczt = "1" Then
  Open "data\picname.txt" For Input As #1
  Line Input #1, imgname
  Close #1
  If Not ctimg Mod 2 = 0 Then
   Image1.Picture = LoadPicture("capture\" & imgname)
   Label2.Caption = Label1.Caption & " ≈ƒ…„"
   ctimg = ctimg + 1
  Else
   Image2.Picture = LoadPicture("capture\" & imgname)
   Label3.Caption = Label1.Caption & " ≈ƒ…„"
   ctimg = ctimg + 1
  End If
 End If
 myWindow = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
ElseIf bbms = True Then
 Image1.Visible = False
 Image2.Visible = False
 Label1.Visible = False
 Label2.Visible = False
 Label3.Visible = False
 Command2.Visible = False
 Command3.Visible = False
 Text1.Visible = False
 Text2.Visible = False
 Text3.Visible = False
 Text4.Visible = False
 Text5.Visible = False
ElseIf ksms = True Then
 Image1.Visible = False
 Image2.Visible = False
 Label2.Visible = False
 Label3.Visible = False
 Command2.Visible = False
 Command3.Visible = False
 Text5.Visible = False
End If
ctimg = 1
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Picture1_DblClick()
Timer4.Enabled = False
speed = 1
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Not tmd + speed >= 255 Then
 tmd = tmd + speed
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), tmd, LWA_ALPHA
 speed = speed + 5
Else
 speed = 1
 tmd = 255
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Not Me.Top < Screen.Height * -1 - 100 Then
 Me.Top = Me.Top - speed
 speed = speed + 7
Else
 speed = 1
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
 ksms = False
 zxms = False
 zdms = False
 bbms = False
 Unload Me
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
End Sub

Public Function namelist(n%) As String
Dim name As String
Open "data\name.txt" For Input As #1
Line Input #1, name
Close #1
If n < 10 Then
 namelist = Trim(Mid(name, InStr(name, CStr(n)) + 1, 4))
ElseIf n >= 10 Then
 namelist = Trim(Mid(name, InStr(name, CStr(n)) + 2, 4))
End If
End Function

Private Sub Timer4_Timer()
sj = sj + 1
If zdpzsj = sj / 2 Then
 sj = 0
 Set objshell = CreateObject("Wscript.Shell")
 Call Shell("cmd /c start capt.exe")
 Sleep 500
 While exitproc("capt.exe")
 Wend
 Sleep 500
 Open "data\zt.txt" For Input As #1
 Line Input #1, piczt
 Close #1
 If piczt = "1" Then
  Open "data\picname.txt" For Input As #1
  Line Input #1, imgname
  Close #1
  If Not ctimg Mod 2 = 0 Then
   Image1.Picture = LoadPicture("capture\" & imgname)
   Label2.Caption = Label1.Caption & " ≈ƒ…„"
   ctimg = ctimg + 1
  Else
   Image2.Picture = LoadPicture("capture\" & imgname)
   Label3.Caption = Label1.Caption & " ≈ƒ…„"
   ctimg = ctimg + 1
  End If
 End If
 myWindow = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
End If
End Sub
