VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   0  'None
   Caption         =   "Form22"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   LinkTopic       =   "Form22"
   Moveable        =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   1080
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�ı��ı��ı��ı��ı�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "Form22"
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
Dim sspeed As Long
Dim n As Integer

Private Sub Form_Load()
Dim rtn As Long
Me.BackColor = RGB(205, 229, 231)
rtn = 8895
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(205, 229, 231), 220, LWA_ALPHA
Label1.Caption = fdctext
sspeed = 1
n = 0
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 0 - Me.Height
Me.Width = 1335
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub Timer1_Timer()
If sspeed > 0 Then
 If Me.Top + sspeed <= Me.Height * -1 / 3 Then
  Me.Top = Me.Top + sspeed
  Me.Width = Me.Width + sspeed * 7
  Me.Left = (Screen.Width - Me.Width) / 2
  Label1.Caption = fdctext
  sspeed = sspeed + 3
 Else
  Me.Top = Me.Top + sspeed
  Me.Width = Me.Width + sspeed * 3
  Me.Left = (Screen.Width - Me.Width) / 2
  Label1.Caption = fdctext
  sspeed = sspeed - 5
 End If
Else
 Timer2.Enabled = True
 Timer3.Enabled = False
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Not n >= fdctime * 2 Then
 n = n + 1
 Label1.Caption = fdctext
Else
 sspeed = 1
 Timer3.Enabled = True
 Timer1.Enabled = False
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If Not Me.Top < Me.Height * -1 - 100 Then
 Me.Top = Me.Top - sspeed
 Me.Width = Me.Width - sspeed * 5
 Me.Left = (Screen.Width - Me.Width) / 2
 Label1.Caption = fdctext
 sspeed = sspeed + 4
Else
 Timer1.Enabled = False
 Timer2.Enabled = False
 Timer3.Enabled = False
 Unload Me
End If
End Sub

