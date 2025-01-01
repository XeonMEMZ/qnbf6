VERSION 5.00
Begin VB.Form Form38 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê±¼ä¿ØÖÆ"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form38.frx":0000
   LinkTopic       =   "Form38"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È·¶¨"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÇëÊäÈëÊ±¼ä¿ØÖÆÃÜÂë"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   660
      TabIndex        =   2
      Top             =   340
      Width           =   3240
   End
End
Attribute VB_Name = "Form38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tcpsw As String

Private Sub Command1_Click()
Open "data\pswtc.txt" For Input As #1
Line Input #1, tcpsw
Close #1
If Text1.Text = tcpsw Then
 If sjmm = "0" Then
  Unload Form25
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\tc.mp3"
  End If
  Open "data\exit.txt" For Output As #1
  Print #1, "1"
  Close #1
  Timer1.Enabled = True
 ElseIf sjmm = "1" Then
  Form36.Show
  Unload Me
 End If
Else
 MsgBox "Ê±¼ä¿ØÖÆÃÜÂë´íÎó", vbCritical, "ÌáÊ¾"
End If
End Sub

Private Sub Timer1_Timer()
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
 Unload Form1
 Unload Form2
 Unload Form4
 Unload Form5
 Unload Form6
 Unload Form7
 Unload Form8
 Unload Form9
 Unload Form10
 Unload Form11
 Unload Form12
 Unload Form13
 Unload Form14
 Unload Form15
 Unload Form16
 Unload Form17
 Unload Form18
 Unload Form19
 Unload Form20
 Unload Form21
 Unload Form22
 Unload Form23
 Unload Form24
 Unload Form26
 Unload Form27
 Unload Form28
 Unload Form29
 Unload Form30
 Unload Form31
 Unload Form32
 Unload Form33
 Unload Form34
 Unload Form35
 Unload Form3
 Unload Me
End If
End Sub
