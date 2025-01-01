VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form9 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   LinkTopic       =   "Form9"
   ScaleHeight     =   615
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1930
      Picture         =   "Form9.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   70
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Interval        =   2
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   660
      Left            =   1710
      Top             =   0
      Width           =   110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H004D5E2B&
      BackStyle       =   0  'Transparent
      Caption         =   "¿Î±í"
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
      Height          =   614
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   70
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
Attribute VB_Name = "Form9"
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Dim speed As Integer
Dim bg As String

Private Sub Command1_Click()
kbcout = 9
kbtext = Label1.Caption
kbcolr = Me.BackColor
Form20.Show
End Sub

Private Sub Form_Load()
Me.Width = 1805
Me.Left = leftt
Me.Top = Screen.Height
Open "data\f9color.txt" For Input As #1
Line Input #1, bg
Close #1
Dim rtn As Long
Me.BackColor = bg
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, bg, tmdu, LWA_ALPHA
speed = 250
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
If hkcout = 9 Then
 Label1.Caption = hktext
Else
 If Val(allclass("ct")) < 6 Then
  Label1.Caption = ""
 Else
  Label1.Caption = allclass("c6")
 End If
End If
End Sub

Private Sub Label1_Click()
If caidan = True Then
 WindowsMediaPlayer1.URL = "themes\cdaudio\6.mp3"
End If
End Sub

Private Sub Label1_DblClick()
If Timer2.Enabled = False And Timer3.Enabled = False And Timer4.Enabled = False And kbdh = False Then
 If Me.Left < Screen.Width - Form2.Width / 2 Then
  Timer2.Enabled = True
 End If
End If
End Sub

Private Sub Timer1_Timer()
If Me.Top + 250 >= Form2.Height + Me.Height * 5 And speed > 0 Then
 If Me.Top >= Form2.Height + Me.Height * 5 + 2000 Then
  Me.Top = Me.Top - speed
 Else
  Me.Top = Me.Top - speed
  speed = speed - 19
 End If
Else
 speed = 1
 Me.Top = Form2.Height + Me.Height * 5
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If kbdh = True Then
 speed = 1
 Me.Width = 1805
 Me.Left = Form2.Left
 Timer3.Enabled = False
 Timer4.Enabled = False
 Timer2.Enabled = False
End If
If speed > 0 Then
 If Me.Left - speed >= leftt + 1805 - 2580 + (780 / 3) Then
  Me.Width = Me.Width + speed
  Me.Left = Me.Left - speed
  speed = speed + 2
 Else
  Me.Width = Me.Width + speed
  Me.Left = Me.Left - speed
  speed = speed - 4
 End If
Else
 speed = 1
 Me.Width = 2580
 Me.Left = leftt + 1805 - 2580
 Timer3.Enabled = True
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If kbdh = True Then
 speed = 1
 Me.Width = 1805
 Me.Left = Form2.Left
 Timer2.Enabled = False
 Timer4.Enabled = False
 Timer3.Enabled = False
End If
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
If kbdh = True Then
 speed = 1
 Me.Width = 1805
 Me.Left = Form2.Left
 Timer2.Enabled = False
 Timer3.Enabled = False
 Timer4.Enabled = False
End If
If speed > 0 Then
 If Me.Left - speed <= leftt + 1805 - 2580 + (780 / 2.5) Then
  Me.Width = Me.Width - speed
  Me.Left = Me.Left + speed
  speed = speed + 2
 Else
  Me.Width = Me.Width - speed
  Me.Left = Me.Left + speed
  speed = speed - 1.5
 End If
Else
 speed = 1
 Me.Width = 1805
 Me.Left = leftt
 Timer4.Enabled = False
End If
End Sub

Public Function allclass(s$) As String
Dim classlist As String
Open "data\z" & Weekday(Date, 2) & ".txt" For Input As #1
Line Input #1, classlist
Close #1
allclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 2, 3))
End Function

Public Function zxtime(s$) As String
Dim timelist As String
Open "data\timelist.txt" For Input As #1
Line Input #1, timelist
Close #1
zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 4, 9))
End Function

