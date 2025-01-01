VERSION 5.00
Begin VB.Form Form17 
   BorderStyle     =   0  'None
   Caption         =   "Form17"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form17"
   ScaleHeight     =   6975
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form17.frx":0000
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form17.frx":0011
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form17"
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

Private Sub Form_Load()
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(240, 240, 240), 200, LWA_ALPHA
speed = 1
Me.Width = Screen.Width / 2
Me.Height = Screen.Height
Me.Left = Me.Width * -1
Me.Top = 0
Text1.Left = 120
Text1.Top = 120
Text1.Width = Me.Width / 2 - 180
Text1.Height = Me.Height - 960
Text2.Left = 240 + Text1.Width
Text2.Top = 120
Text2.Width = Me.Width / 2 - 180
Text2.Height = Me.Height - 960
Text1.Text = zy1
Text2.Text = zy2
End Sub

Private Sub Text1_Change()
zy1 = Text1.Text
End Sub

Private Sub Text2_Change()
zy2 = Text2.Text
End Sub

Private Sub Timer1_Timer()
If speed > 0 Then
 If Not Me.Left + speed >= (Me.Width * -1 / 3) Then
  Me.Left = Me.Left + speed
  speed = speed + 13
 Else
  Me.Left = Me.Left + speed
  speed = speed - 25
 End If
 SetLayeredWindowAttributes hwnd, RGB(240, 240, 240), 240 - Abs(Me.Left) * (240 / Me.Width), LWA_ALPHA
Else
 Me.Left = 0
 speed = 1
 Timer1.Enabled = False
End If
End Sub

