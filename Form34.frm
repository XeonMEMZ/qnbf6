VERSION 5.00
Begin VB.Form Form34 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸�����"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7455
   Icon            =   "Form34.frx":0000
   LinkTopic       =   "Form34"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7455
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "ǳɫ"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ɫ"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   2175
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   975
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   2175
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l1c As String
Dim l2c As String
Dim l3c As String
Dim re As Integer

Private Sub Command1_Click()
Open "themes\themes.txt" For Output As #1
Print #1, "1"
Close #1
re = MsgBox("�޸���������������Ч!" & vbCrLf & "�Ƿ��������?", 52, "��ʾ")
If re = 6 Then
 If dtrs = "1" Then
  Call Shell("cmd /c start killme.bat")
  Unload Me
 ElseIf dtrs = "0" Then
  Call Shell("cmd /c start rstme.bat")
  Unload Me
 End If
Else
 Unload Me
End If
End Sub

Private Sub Command2_Click()
Open "themes\themes.txt" For Output As #1
Print #1, "2"
Close #1
re = MsgBox("�޸���������������Ч!" & vbCrLf & "�Ƿ��������?", 52, "��ʾ")
If re = 6 Then
 If dtrs = "1" Then
  Call Shell("cmd /c start killme.bat")
  Unload Me
 ElseIf dtrs = "0" Then
  Call Shell("cmd /c start rstme.bat")
  Unload Me
 End If
Else
 Unload Me
End If
End Sub

Private Sub Command3_Click()
Open "themes\themes.txt" For Output As #1
Print #1, "3"
Close #1
re = MsgBox("�޸���������������Ч!" & vbCrLf & "�Ƿ��������?", 52, "��ʾ")
If re = 6 Then
 If dtrs = "1" Then
  Call Shell("cmd /c start killme.bat")
  Unload Me
 ElseIf dtrs = "0" Then
  Call Shell("cmd /c start rstme.bat")
  Unload Me
 End If
Else
 Unload Me
End If
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("themes\1\bg1.jpg")
Image2.Picture = LoadPicture("themes\1\bg2.jpg")
Image4.Picture = LoadPicture("themes\2\bg1.jpg")
Image5.Picture = LoadPicture("themes\2\bg2.jpg")
Image7.Picture = LoadPicture("themes\3\bg1.jpg")
Image8.Picture = LoadPicture("themes\3\bg2.jpg")
Open "themes\1\name.txt" For Input As #1
Line Input #1, l1c
Close #1
Command1.Caption = l1c
Open "themes\2\name.txt" For Input As #1
Line Input #1, l2c
Close #1
Command2.Caption = l2c
Open "themes\3\name.txt" For Input As #1
Line Input #1, l3c
Close #1
Command3.Caption = l3c
End Sub

