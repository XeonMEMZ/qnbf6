VERSION 5.00
Begin VB.Form Form36 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ʱ�����"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4830
   Icon            =   "Form36.frx":0000
   LinkTopic       =   "Form36"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4830
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��ʱ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ʱ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʱ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�������������(������Ч)"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ʱ��:"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����1:"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����2:"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��:"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1020
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text2.Text = "" And Not Text4.Text = "" And Len(Text2.Text) = 8 And Len(Text4.Text) = 8 Then
 If Not Text1.Text = "main.exe" And Not Text3.Text = "main.exe" And Not Text1.Text = "dual.exe" And Not Text3.Text = "dual.exe" And Text2.Text > Text4.Text Then
  Dim atlo As String
  Open "data\tctask1.txt" For Output As #1
  atlo = Text1.Text
  Print #1, atlo
  Close #1
  Dim atlt As String
  Open "data\tctask2.txt" For Output As #1
  atlt = Text3.Text
  Print #1, atlt
  Close #1
  Dim attjy As String
  Open "data\tctimejy.txt" For Output As #1
  attjy = Text4.Text
  Print #1, attjy
  Close #1
  Dim attjj As String
  Open "data\tctimejj.txt" For Output As #1
  attjj = Text2.Text
  Print #1, attjj
  Close #1
  Dim llq As String
  If Check1.Value = 1 Then
   Open "data\tcllq.txt" For Output As #1
   llq = "1"
   Print #1, llq
   Close #1
  Else
   Open "data\tcllq.txt" For Output As #1
   llq = "0"
   Print #1, llq
   Close #1
  End If
  Dim mgr As String
  If Check2.Value = 1 Then
   Open "data\tcmgr.txt" For Output As #1
   mgr = "1"
   Print #1, mgr
   Close #1
  Else
   Open "data\tcmgr.txt" For Output As #1
   mgr = "0"
   Print #1, mgr
   Close #1
  End If
  Unload Me
 Else
  MsgBox "������˲���", vbCritical, "����"
 End If
Else
 MsgBox "��������Ч�Ĳ���", vbCritical, "��ʾ"
End If
End Sub

Private Sub Command2_Click()
sjkz = False
Open "data\tctask1.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tctask2.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tctimejy.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tctimejj.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tcllq.txt" For Output As #1
Print #1, "0"
Close #1
Open "data\tcmgr.txt" For Output As #1
Print #1, "0"
Close #1
Unload Me
End Sub

Private Sub Command3_Click()
If sjkz = True Then
 sjkz = False
 Unload Me
Else
 MsgBox "ʱ�����δ����", vbCritical, "��ʾ"
End If
End Sub

Private Sub Form_Load()
Dim lino As String
Open "data\tctask1.txt" For Input As #1
Line Input #1, lino
Text1.Text = lino
Close #1
Dim lint As String
Open "data\tctask2.txt" For Input As #1
Line Input #1, lint
Text3.Text = lint
Close #1
Dim timjy As String
Open "data\tctimejy.txt" For Input As #1
Line Input #1, timjy
Text4.Text = timjy
Close #1
Dim timjj As String
Open "data\tctimejj.txt" For Input As #1
Line Input #1, timjj
Text2.Text = timjj
Close #1
Dim tllq As String
Open "data\tcllq.txt" For Input As #1
Line Input #1, tllq
Close #1
Check1.Value = tllq
Dim tmgr As String
Open "data\tcmgr.txt" For Input As #1
Line Input #1, tmgr
Close #1
Check2.Value = tmgr
End Sub


