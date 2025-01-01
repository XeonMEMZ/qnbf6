VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer7 
      Interval        =   2000
      Left            =   3000
      Top             =   480
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2520
      Top             =   480
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   2040
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Interval        =   2
      Left            =   1560
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1080
      Top             =   480
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "°×°å±³¾°"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "×ÔÏ°¹ÜÀí"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿ì½ÝÖúÊÖ"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ò»¼üÅÄÕÕ"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ò»ÖÜ¿Î±í"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2880
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   600
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1665
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "×÷ÒµÃæ°å"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3645
      Picture         =   "Form3.frx":2466
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3165
      Picture         =   "Form3.frx":27BB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3645
      Picture         =   "Form3.frx":2B12
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "´ò¿ªUÅÌ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ëæ»úµãÃû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3300
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3255
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1425
      TabIndex        =   3
      Top             =   5040
      Width           =   1395
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4320
      Width           =   1395
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H004E4E4E&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   885
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5760
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   480
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   3600
      TabIndex        =   26
      Top             =   120
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ÖµÈÕ:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   20
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¹¤¾ßÀ¸"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3120
      TabIndex        =   19
      Top             =   1360
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¾àÆÚÄ©¿¼»¹ÓÐ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   165
      TabIndex        =   18
      Top             =   1380
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ìì"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   17
      Top             =   2055
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Êµµ½:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ó¦µ½:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "¼Ù"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   14
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Çë"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   13
      Top             =   5760
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objshell As Object
Option Explicit
Dim speed As Integer
Dim dskop As Integer
Dim disk As Integer
Dim nc As String
Dim fxs As String
Dim zrnext As String
Dim zname As String
Dim cname As String
Dim namect As String
Dim zrname1 As String
Dim zrname2 As String
Dim znext As String
Dim nzr As String
Dim upf As String
Dim atl1 As String
Dim att1 As String
Dim djstext As String
Dim djstime As String
Dim lin1 As String
Dim tim1 As String
Dim oc1 As String
Dim lin2 As String
Dim tim2 As String
Dim oc2 As String
Dim lin3 As String
Dim tim3 As String
Dim oc3 As String
Dim lin4 As String
Dim tim4 As String
Dim oc4 As String
Dim sklin As String
Dim skoc As String
Dim xklin As String
Dim xkoc As String
Dim xkchb As String
Dim xkchbtm As String
Dim autostd As String
Dim zdgj As String
Dim tlino As String
Dim tlint As String
Dim ttimjy As String
Dim ttimjj As String
Dim tllq As String
Dim tmgr As String
Dim udatop As String
Dim piczt As String
Dim picna As String
Dim zymb As Boolean

Private Sub Command1_Click()
Open "data\upf.txt" For Input As #1
Line Input #1, upf
Close #1
Call Shell("cmd /c start " & upf)
End Sub

Private Sub Command10_Click()
If Timer3.Enabled = False Then
speed = 1
 If kjzs = False Then
  Unload Form19
  Form19.Show
  kjzs = True
 Else
  Timer3.Enabled = True
 End If
End If
End Sub

Private Sub Command11_Click()
If ksms = False And zxms = False And zdms = False And bbms = False Then
 Unload Form21
 zxms = True
 Form21.Show
End If
End Sub

Private Sub Command12_Click()
If ksms = False And zxms = False And zdms = False And bbms = False Then
 Unload Form21
 bbms = True
 Form21.Show
End If
End Sub

Private Sub Command2_Click()
If ksms = False And zxms = False And zdms = False And bbms = False Then
 Unload Form21
 ksms = True
 Form21.Show
End If
End Sub

Private Sub Command4_Click()
Call Shell("cmd /c start capture")
End Sub

Private Sub Command5_Click()
Form25.Show
End Sub

Private Sub Command6_Click()
If Not Form2.Left = 0 Then
 If Timer2.Enabled = False Then
 speed = 1
  If zymb = False Then
   Unload Form17
   Form17.Show
   zymb = True
  Else
   If Form17.Left = 0 Then
    Timer2.Enabled = True
   End If
  End If
 End If
End If
End Sub

Private Sub Command7_Click()
Open "data\namec.txt" For Input As #1
Line Input #1, cname
Close #1
Dim sj As Integer
Randomize
sj = Int(Rnd * (CInt(cname) - 1 + 1) + 1)
Text3.Text = namelist(sj)
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

Private Sub Command8_Click()
Form18.Show
End Sub

Private Sub Command9_Click()
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
 Line Input #1, picna
 Close #1
 Call Shell("cmd /c start capture\" & picna)
Else
 MsgBox "ÅÄÉãÕÕÆ¬Ê§°Ü" & vbCrLf & "Çë¼ì²éÉãÏñÍ·ÊÇ·ñÕýÈ·Á¬½Ó", 48, "ÎÞ·¨ÅÄÉã"
End If
End Sub

Private Sub Form_Load()
Form2.Show
If textc = "1" Then
 Label1.ForeColor = vbWhite
 Label2.ForeColor = vbWhite
 Label3.ForeColor = vbWhite
 Label4.ForeColor = vbWhite
 Label5.ForeColor = vbWhite
 Label6.ForeColor = vbWhite
 Label7.ForeColor = vbWhite
 Label8.ForeColor = vbWhite
 Text1.ForeColor = vbWhite
 Text2.ForeColor = vbWhite
 Text3.ForeColor = vbWhite
 Text4.ForeColor = vbWhite
 Text5.ForeColor = vbWhite
 Text6.ForeColor = vbWhite
Else
 Label1.ForeColor = vbBlack
 Label2.ForeColor = vbBlack
 Label3.ForeColor = vbBlack
 Label4.ForeColor = vbBlack
 Label5.ForeColor = vbBlack
 Label6.ForeColor = vbBlack
 Label7.ForeColor = vbBlack
 Label8.ForeColor = vbBlack
 Text1.ForeColor = vbBlack
 Text2.ForeColor = vbBlack
 Text3.ForeColor = vbBlack
 Text4.ForeColor = vbBlack
 Text5.ForeColor = vbBlack
 Text6.ForeColor = vbBlack
End If
Me.Picture = LoadPicture("themes\" & theme & "\bg2.jpg")
Text1.BackColor = RGB(collist("bg22r"), collist("bg22g"), collist("bg22b"))
Text3.BackColor = RGB(collist("bg22r"), collist("bg22g"), collist("bg22b"))
Text2.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Text4.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Text5.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Text6.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), tmdu, LWA_ALPHA
Timer6.Enabled = False
speed = 1
disk = Drive1.ListCount
If aud = "1" Then
 WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\qd.mp3"
End If
Open "data\namezr.txt" For Input As #1
Line Input #1, nzr
Close #1
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 If Not Weekday(Date, 2) = 6 And Not Weekday(Date, 2) = 7 Then
  Open "data\nextzr.txt" For Input As #1
  Line Input #1, zrnext
  Close #1
  If Not CStr(Day(Date)) = zrnext Then
   Open "data\namezr.txt" For Input As #1
   Line Input #1, zname
   Close #1
   Open "data\namec.txt" For Input As #1
   Line Input #1, namect
   Close #1
   If Not Fix(zname) >= Fix(namect) Then
    Open "data\namezr.txt" For Output As #1
    zrname1 = CStr(Fix(Val(zname)) + 1)
    Print #1, zrname1
    Close #1
   Else
    Open "data\namezr.txt" For Output As #1
    zrname2 = "1"
    Print #1, zrname2
    Close #1
   End If
   Open "data\nextzr.txt" For Output As #1
   znext = CStr(Day(Date))
   Print #1, znext
   Close #1
  End If
 End If
ElseIf fxs = "6" Then
 If Not Weekday(Date, 2) = 7 Then
  Open "data\nextzr.txt" For Input As #1
  Line Input #1, zrnext
  Close #1
  If Not CStr(Day(Date)) = zrnext Then
   Open "data\namezr.txt" For Input As #1
   Line Input #1, zname
   Close #1
   Open "data\namec.txt" For Input As #1
   Line Input #1, namect
   Close #1
   If Not Fix(zname) >= Fix(namect) Then
    Open "data\namezr.txt" For Output As #1
    zrname1 = CStr(Fix(Val(zname)) + 1)
    Print #1, zrname1
    Close #1
   Else
    Open "data\namezr.txt" For Output As #1
    zrname2 = "1"
    Print #1, zrname2
    Close #1
   End If
   Open "data\nextzr.txt" For Output As #1
   znext = CStr(Day(Date))
   Print #1, znext
   Close #1
  End If
 End If
ElseIf fxs = "7" Then
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
End If
If lszr = "" Then
 Open "data\namezr.txt" For Input As #1
 Line Input #1, nzr
 Close #1
 Text1.Text = namelist(Int(nzr))
Else
 Text1.Text = lszr
End If
Open "data\namec.txt" For Input As #1
Line Input #1, nc
Close #1
Text4.Text = nc
Text5.Text = nc
Open "data\atlink1.txt" For Input As #1
Line Input #1, atl1
Close #1
Open "data\attime1.txt" For Input As #1
Line Input #1, att1
Close #1
If Not atl1 = "" And Not att1 = "" Then
 zdrw = True
Else
 zdrw = False
End If
Open "data\djstext.txt" For Input As #1
Line Input #1, djstext
Close #1
Open "data\djstime.txt" For Input As #1
Line Input #1, djstime
Close #1
Label3.Caption = "¾à" & djstext & "»¹ÓÐ"
Text2.Text = CDate(djstime) - Date
Open "data\autostd.txt" For Input As #1
Line Input #1, autostd
Close #1
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
End Sub

Private Sub Timer1_Timer()
If lszr = "" Then
 Text1.Text = namelist(Int(nzr))
Else
 Text1.Text = lszr
End If
Open "data\djstext.txt" For Input As #1
Line Input #1, djstext
Close #1
Open "data\djstime.txt" For Input As #1
Line Input #1, djstime
Close #1
Label3.Caption = "¾à" & djstext & "»¹ÓÐ"
Text2.Text = CDate(djstime) - Date
End Sub

Private Sub Timer2_Timer()
If Form17.Left > Form17.Width * -1 Then
 Form17.Left = Form17.Left - speed
 speed = speed + 10
Else
 Unload Form17
 speed = 1
 zymb = False
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If Form19.Top < Screen.Height Then
 Form19.Top = Form19.Top + speed
 speed = speed + 5
Else
 Form19.Top = Screen.Height + 100
 Unload Form19
 speed = 1
 kjzs = False
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If Me.Top < Form2.Height Then
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), toumd, LWA_ALPHA
Else
 Timer4.Enabled = False
End If
End Sub

Public Function collist(c$) As String
Dim allcolor As String
Open "themes\" & theme & "\color.txt" For Input As #1
Line Input #1, allcolor
Close #1
collist = Trim(Mid(allcolor, InStr(allcolor, CStr(c)) + 5, 3))
End Function

Private Sub Timer5_Timer()
If lszr = "" Then
 Open "data\namezr.txt" For Input As #1
 Line Input #1, nzr
 Close #1
 Text1.Text = namelist(Int(nzr))
Else
 Text1.Text = lszr
End If
If zdrw = True Then
 Open "data\atlink1.txt" For Input As #1
 Line Input #1, lin1
 Open "data\attime1.txt" For Input As #2
 Line Input #2, tim1
 Open "data\atoc1.txt" For Input As #3
 Line Input #3, oc1
 Close #1
 Close #2
 Close #3
 If tim1 = CStr(Time) Then
  If oc1 = "1" Then
   Call Shell("cmd /c start " & lin1)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin1)
  End If
 End If
 Open "data\atlink2.txt" For Input As #1
 Line Input #1, lin2
 Open "data\attime2.txt" For Input As #2
 Line Input #2, tim2
 Open "data\atoc2.txt" For Input As #3
 Line Input #3, oc2
 Close #1
 Close #2
 Close #3
 If tim2 = CStr(Time) Then
  If oc2 = "1" Then
   Call Shell("cmd /c start " & lin2)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin2)
  End If
 End If
 Open "data\atlink3.txt" For Input As #1
 Line Input #1, lin3
 Open "data\attime3.txt" For Input As #2
 Line Input #2, tim3
 Open "data\atoc3.txt" For Input As #3
 Line Input #3, oc3
 Close #1
 Close #2
 Close #3
 If tim3 = CStr(Time) Then
  If oc3 = "1" Then
   Call Shell("cmd /c start " & lin3)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin3)
  End If
 End If
 Open "data\atlink4.txt" For Input As #1
 Line Input #1, lin4
 Open "data\attime4.txt" For Input As #2
 Line Input #2, tim4
 Open "data\atoc4.txt" For Input As #3
 Line Input #3, oc4
 Close #1
 Close #2
 Close #3
 If tim4 = CStr(Time) Then
  If oc4 = "1" Then
   Call Shell("cmd /c start " & lin4)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin4)
  End If
 End If
End If
Open "data\skatlink.txt" For Input As #1
Line Input #1, sklin
Close #1
Open "data\skatoc.txt" For Input As #1
Line Input #1, skoc
Close #1
If fricls = 0 Then
 If CDate(zxtime("zm1s")) = Time Or CDate(zxtime("zm2s")) = Time Or CDate(zxtime("zm3s")) = Time Or CDate(zxtime("zm4s")) = Time Or CDate(zxtime("zm5s")) = Time Or CDate(zxtime("zm6s")) = Time Or CDate(zxtime("zm7s")) = Time Or CDate(zxtime("zm8s")) = Time Or CDate(zxtime("zm9s")) = Time Or CDate(zxtime("zm10s")) = Time Or CDate(zxtime("zm11s")) = Time Or CDate(zxtime("zm12s")) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sk.mp3"
  End If
  If Not sklin = "" Then
   If skoc = "1" Then
    Call Shell("cmd /c start " & sklin)
   Else
    Call Shell("cmd /c taskkill /f /im " & sklin)
   End If
  End If
 End If
Else
 If CDate(zxtime("zm1s")) = Time Or CDate(zxtime("zm2s")) = Time Or CDate(zxtime("zm3s")) = Time Or CDate(zxtime("zm4s")) = Time Or CDate(zxtime("zf5s")) = Time Or CDate(zxtime("zf6s")) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sk.mp3"
  End If
  If Not sklin = "" Then
   If skoc = "1" Then
    Call Shell("cmd /c start " & sklin)
   Else
    Call Shell("cmd /c taskkill /f /im " & sklin)
   End If
  End If
 End If
End If
Open "data\xkatlink.txt" For Input As #1
Line Input #1, xklin
Close #1
Open "data\xkatoc.txt" For Input As #1
Line Input #1, xkoc
Close #1
If Not xklin = "" Then
 If fricls = 0 Then
  If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zm5x")) = Time Or CDate(zxtime("zm6x")) = Time Or CDate(zxtime("zm7x")) = Time Or CDate(zxtime("zm8x")) = Time Or CDate(zxtime("zm9x")) = Time Or CDate(zxtime("zm10x")) = Time Or CDate(zxtime("zm11x")) = Time Or CDate(zxtime("zm12x")) = Time Then
   If xkoc = "1" Then
    Call Shell("cmd /c start " & xklin)
   Else
    Call Shell("cmd /c taskkill /f /im " & xklin)
   End If
  End If
 Else
  If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zf5x")) = Time Or CDate(zxtime("zf6x")) = Time Then
   If xkoc = "1" Then
    Call Shell("cmd /c start " & xklin)
   Else
    Call Shell("cmd /c taskkill /f /im " & xklin)
   End If
  End If
 End If
End If
Open "data\xkchb.txt" For Input As #1
Line Input #1, xkchb
Close #1
Open "data\xkchbtime.txt" For Input As #1
Line Input #1, xkchbtm
Close #1
If fricls = 0 Then
 If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zm5x")) = Time Or CDate(zxtime("zm6x")) = Time Or CDate(zxtime("zm7x")) = Time Or CDate(zxtime("zm8x")) = Time Or CDate(zxtime("zm9x")) = Time Or CDate(zxtime("zm10x")) = Time Or CDate(zxtime("zm11x")) = Time Or CDate(zxtime("zm12x")) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\xk.mp3"
  End If
  If xkchb = "1" Then
   If lszr = "" Then
    fdctext = Text1.Text & ",Çë²ÁºÚ°å"
   Else
    fdctext = lszr & ",Çë²ÁºÚ°å"
    Close #1
   End If
   fdctime = Val(xkchbtm)
   If ksms = False Then
    Form22.Show
   End If
  End If
 End If
Else
 If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zf5x")) = Time Or CDate(zxtime("zf6x")) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\xk.mp3"
  End If
  If xkchb = "1" Then
   If lszr = "" Then
    fdctext = Text1.Text & ",Çë²ÁºÚ°å"
   Else
    fdctext = lszr & ",Çë²ÁºÚ°å"
    Close #1
   End If
   fdctime = Val(xkchbtm)
   If ksms = False Then
    Form22.Show
   End If
  End If
 End If
End If
Open "data\autostd.txt" For Input As #1
Line Input #1, autostd
Close #1
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
If zdgj = "1" Then
 If CDate(autostd) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\gj.mp3"
  End If
  fdctext = "¼ÆËã»ú½«ÔÚ5ÃëÄÚ¹Ø»ú"
  fdctime = 3
  Form22.Show
  Timer6.Enabled = True
 End If
End If
Open "data\tctask1.txt" For Input As #1
Line Input #1, tlino
Open "data\tctask2.txt" For Input As #2
Line Input #2, tlint
Open "data\tctimejy.txt" For Input As #3
Line Input #3, ttimjy
Open "data\tctimejj.txt" For Input As #4
Line Input #4, ttimjj
Open "data\tcllq.txt" For Input As #5
Line Input #5, tllq
Open "data\tcmgr.txt" For Input As #6
Line Input #6, tmgr
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
If Not ttimjy = "" And Not ttimjj = "" Then
 If ttimjy <= Time And Time <= ttimjj Then
  sjkz = True
 Else
  sjkz = False
 End If
Else
 sjkz = False
End If
If sjkz = True Then
 If ttimjj <= Time Then
  sjkz = False
 Else
  If exitproc(tlino) Then
   If aud = "1" Then
    WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
   End If
   Call Shell("cmd /c taskkill /f /im " & tlino)
  End If
  If exitproc(tlint) Then
   If aud = "1" Then
    WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
   End If
   Call Shell("cmd /c taskkill /f /im " & tlint)
  End If
  If tllq = "1" Then
   If exitproc("iexplore.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im iexplore.exe")
   End If
   If exitproc("msedge.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im msedge.exe")
   End If
   If exitproc("chrome.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im chrome.exe")
   End If
   If exitproc("firefox.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im firefox.exe")
   End If
  End If
  If tmgr = "1" Then
   If exitproc("Taskmgr.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & theme & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im Taskmgr.exe")
   End If
  End If
 End If
End If
Drive1.Refresh
Open "data\udatop.txt" For Input As #1
Line Input #1, udatop
Close #1
If udatop = "1" Then
 If Drive1.ListCount > disk Then
  If dskop = 0 Then
   Dim upf As String
   Open "data\upf.txt" For Input As #1
   Line Input #1, upf
   Close #1
   Call Shell("cmd /c start " & upf)
   Form23.Show
   dskop = 1
  End If
 Else
  dskop = 0
 End If
End If
End Sub

Private Sub Timer6_Timer()
Call Shell("cmd /c shutdown /s /t 0")
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
If exitproc("qnbf.exe") Then
 If dt = "1" Then
  If Not exitproc("dual.exe") Then
   Call Shell("cmd /c start dual.exe")
  End If
 End If
End If
End Sub
