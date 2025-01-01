VERSION 5.00
Begin VB.Form Form43 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "±à¼­Ó¦ÓÃ³éÌë"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   Icon            =   "Form43.frx":0000
   LinkTopic       =   "Form43"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5895
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È·¶¨"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "È¡Ïû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3720
      Width           =   2655
   End
End
Attribute VB_Name = "Form43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If InStr(Text1.Text, ".exe") = 0 Then
 MsgBox "ÇëÑ¡Ôñ.exeÎÄ¼þ!", vbCritical, "ÌáÊ¾"
Else
 Open "data\cy" & cyn & ".txt" For Output As #1
 Print #1, Text1.Text
 Close #1
 cyc = False
 Unload Me
End If
End Sub

Private Sub Command2_Click()
cyc = False
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text1.Text = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub File1_Click()
Text1.Text = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub Form_Load()
Dim cy As String
Open "data\cy" & cyn & ".txt" For Input As #1
Line Input #1, cy
Close #1
Me.Caption = "±à¼­Ó¦ÓÃ³éÌë" & cyn
Text1.Text = cy
End Sub

Private Sub Form_Unload(Cancel As Integer)
cyc = False
End Sub
