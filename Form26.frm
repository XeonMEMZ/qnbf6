VERSION 5.00
Begin VB.Form Form26 
   BorderStyle     =   0  'None
   Caption         =   "Form26"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form26"
   Picture         =   "Form26.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9235
      Picture         =   "Form26.frx":34114
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   -10
      Width           =   375
   End
End
Attribute VB_Name = "Form26"
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

Private Sub Form_Load()
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
qdtm = 0
qdsd = 1
End Sub

Private Sub Picture1_Click()
If Timer1.Enabled = False Then
 Timer2.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
If qdtm + 10 <= 240 Then
 qdtm = qdtm + 10
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
Else
 qdtm = 240
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), qdtm, LWA_ALPHA
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
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
 Timer2.Enabled = False
End If
End Sub
