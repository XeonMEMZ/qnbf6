VERSION 5.00
Begin VB.Form Form29 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸�ֵ��"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form29.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  '��Ļ����
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
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
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
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�޸�ֵ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zr As String

Private Sub Command1_Click()
If Not Text1.Text = "" Then
 If Not searchname(Text1.Text) = "0" And IsNumeric(searchname(Text1.Text)) = True Then
  Dim zname As String
  zname = searchname(Text1.Text)
  Open "data\namezr.txt" For Output As #1
  Print #1, zname
  Close #1
  Unload Me
 Else
  MsgBox "�������������в�����", vbCritical, "��ʾ"
 End If
Else
 MsgBox "ֵ�ղ���Ϊ��", vbCritical, "��ʾ"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Open "data\namezr.txt" For Input As #1
Line Input #1, zr
Close #1
Text1.Text = namelist(Int(zr))
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

Public Function searchname(n$) As String
Dim allname As String
Open "data\name.txt" For Input As #1
Line Input #1, allname
Close #1
If Not InStr(allname, CStr(n)) = 0 Then
 searchname = Trim(Mid(allname, InStr(allname, CStr(n)) - 2, 2))
Else
 searchname = "0"
End If
End Function
