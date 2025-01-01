VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÐÞ¸Äµ±Ç°¿Î±í"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   4575
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÕûÈÕ»»¿Î"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   975
      Left            =   300
      TabIndex        =   22
      Top             =   4860
      Width           =   3975
      Begin VB.CommandButton Command26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜÈÕ"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜÁù"
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
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜÎå"
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜËÄ"
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
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜÈý"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜ¶þ"
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
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÖÜÒ»"
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
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿ì½Ý»»¿Î"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2775
      Left            =   300
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
      Begin VB.CommandButton Command20 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
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
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
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
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
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
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
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
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00979E10&
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
      Height          =   765
      Left            =   1020
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2850
      Picture         =   "Form20.frx":6988A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "È·¶¨"
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
      TabIndex        =   1
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÐÞ¸Äµ±Ç°¿Î±í"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   2610
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00979E10&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2700
      Top             =   1080
      Width           =   780
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bg As String
Dim kbcount As Integer
Dim wk As Integer
Dim d As Integer
Dim c As Integer
Dim b As Integer

Private Sub Command1_Click()
 CommonDialog1.CancelError = True
 On Error GoTo cuowu
 CommonDialog1.ShowColor
 Text1.BackColor = CommonDialog1.Color
 Shape1.BackColor = CommonDialog1.Color
cuowu:
 Exit Sub
End Sub

Private Sub Command10_Click()
hkcout = kbcount
hktext = Command10.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command11_Click()
hkcout = kbcount
hktext = Command11.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command12_Click()
hkcout = kbcount
hktext = Command12.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command13_Click()
hkcout = kbcount
hktext = Command13.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command14_Click()
hkcout = kbcount
hktext = Command14.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command15_Click()
hkcout = kbcount
hktext = Command15.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command16_Click()
hkcout = kbcount
hktext = Command16.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command17_Click()
hkcout = kbcount
hktext = Command17.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command18_Click()
hkcout = kbcount
hktext = Command18.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command19_Click()
hkcout = kbcount
hktext = Command19.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command2_Click()
If kbcount = 4 Then
 Form4.BackColor = Text1.BackColor
ElseIf kbcount = 5 Then
 Form5.BackColor = Text1.BackColor
ElseIf kbcount = 6 Then
 Form6.BackColor = Text1.BackColor
ElseIf kbcount = 7 Then
 Form7.BackColor = Text1.BackColor
ElseIf kbcount = 8 Then
 Form8.BackColor = Text1.BackColor
ElseIf kbcount = 9 Then
 Form9.BackColor = Text1.BackColor
ElseIf kbcount = 10 Then
 Form10.BackColor = Text1.BackColor
ElseIf kbcount = 11 Then
 Form11.BackColor = Text1.BackColor
ElseIf kbcount = 12 Then
 Form12.BackColor = Text1.BackColor
ElseIf kbcount = 13 Then
 Form13.BackColor = Text1.BackColor
ElseIf kbcount = 14 Then
 Form14.BackColor = Text1.BackColor
ElseIf kbcount = 15 Then
 Form15.BackColor = Text1.BackColor
End If
bg = Text1.BackColor
Open "data\f" & kbcount & "color.txt" For Output As #1
Print #1, bg
Close #1
hkcout = kbcount
hktext = Text1.Text
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command20_Click()
hkcout = kbcount
hktext = Command20.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command21_Click()
wk = 2
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command22_Click()
wk = 3
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command23_Click()
wk = 4
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command24_Click()
wk = 5
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command25_Click()
wk = 6
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command26_Click()
wk = 7
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
wk = 1
hkcout = 4
hktext = allclass("c1")
Unload Form4
Form4.Show
hkcout = 5
hktext = allclass("c2")
Unload Form5
Form5.Show
hkcout = 6
hktext = allclass("c3")
Unload Form6
Form6.Show
hkcout = 7
hktext = allclass("c4")
Unload Form7
Form7.Show
hkcout = 8
If Val(allclass("ct")) < 5 Then
 hktext = ""
Else
 hktext = allclass("c5")
End If
Unload Form8
Form8.Show
hkcout = 9
If Val(allclass("ct")) < 6 Then
 hktext = ""
Else
 hktext = allclass("c6")
End If
Unload Form9
Form9.Show
hkcout = 10
If Val(allclass("ct")) < 7 Then
 hktext = ""
Else
 hktext = allclass("c7")
End If
Unload Form10
Form10.Show
hkcout = 11
If Val(allclass("ct")) < 8 Then
 hktext = ""
Else
 hktext = allclass("c8")
End If
Unload Form11
Form11.Show
hkcout = 12
If Val(allclass("ct")) < 9 Then
 hktext = ""
Else
 hktext = allclass("c9")
End If
Unload Form12
Form12.Show
hkcout = 13
If Val(allclass("ct")) < 10 Then
 hktext = ""
Else
 hktext = allclass("c10")
End If
Unload Form13
Form13.Show
hkcout = 14
If Val(allclass("ct")) < 11 Then
 hktext = ""
Else
 hktext = allclass("c11")
End If
Unload Form14
Form14.Show
hkcout = 15
If Val(allclass("ct")) < 12 Then
 hktext = ""
Else
 hktext = allclass("c12")
End If
Unload Form15
Form15.Show
Unload Me
End Sub

Private Sub Command5_Click()
hkcout = kbcount
hktext = Command5.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command6_Click()
hkcout = kbcount
hktext = Command6.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command7_Click()
hkcout = kbcount
hktext = Command7.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command8_Click()
hkcout = kbcount
hktext = Command8.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Command9_Click()
hkcout = kbcount
hktext = Command9.Caption
If kbcount = 4 Then
 Unload Form4
 Form4.Show
ElseIf kbcount = 5 Then
 Unload Form5
 Form5.Show
ElseIf kbcount = 6 Then
 Unload Form6
 Form6.Show
ElseIf kbcount = 7 Then
 Unload Form7
 Form7.Show
ElseIf kbcount = 8 Then
 Unload Form8
 Form8.Show
ElseIf kbcount = 9 Then
 Unload Form9
 Form9.Show
ElseIf kbcount = 10 Then
 Unload Form10
 Form10.Show
ElseIf kbcount = 11 Then
 Unload Form11
 Form11.Show
ElseIf kbcount = 12 Then
 Unload Form12
 Form12.Show
ElseIf kbcount = 13 Then
 Unload Form13
 Form13.Show
ElseIf kbcount = 14 Then
 Unload Form14
 Form14.Show
ElseIf kbcount = 15 Then
 Unload Form15
 Form15.Show
End If
Unload Me
End Sub

Private Sub Form_Load()
kbcount = kbcout
Text1.Text = kbtext
Text1.BackColor = kbcolr
Shape1.BackColor = kbcolr
b = 1
For d = 1 To 7
 For c = 1 To Val(alldclass("ct", d))
  If b = 1 Then
   Command5.Caption = alldclass("c1", d)
   b = b + 1
  ElseIf b = 2 Then
   If Command5.Caption <> alldclass("c" & c, d) Then
    Command6.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 3 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) Then
    Command7.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 4 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) Then
    Command8.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 5 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) Then
    Command9.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 6 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) Then
    Command10.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 7 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) Then
    Command11.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 8 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) Then
    Command12.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 9 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) Then
    Command13.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 10 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) Then
    Command14.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 11 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) And Command14.Caption <> alldclass("c" & c, d) Then
    Command15.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 12 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) And Command14.Caption <> alldclass("c" & c, d) And Command15.Caption <> alldclass("c" & c, d) Then
    Command16.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 13 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) And Command14.Caption <> alldclass("c" & c, d) And Command15.Caption <> alldclass("c" & c, d) And Command16.Caption <> alldclass("c" & c, d) Then
    Command17.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 14 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) And Command14.Caption <> alldclass("c" & c, d) And Command15.Caption <> alldclass("c" & c, d) And Command16.Caption <> alldclass("c" & c, d) And Command17.Caption <> alldclass("c" & c, d) Then
    Command18.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 15 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) And Command14.Caption <> alldclass("c" & c, d) And Command15.Caption <> alldclass("c" & c, d) And Command16.Caption <> alldclass("c" & c, d) And Command17.Caption <> alldclass("c" & c, d) And Command18.Caption <> alldclass("c" & c, d) Then
    Command19.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  ElseIf b = 16 Then
   If Command5.Caption <> alldclass("c" & c, d) And Command6.Caption <> alldclass("c" & c, d) And Command7.Caption <> alldclass("c" & c, d) And Command8.Caption <> alldclass("c" & c, d) And Command9.Caption <> alldclass("c" & c, d) And Command10.Caption <> alldclass("c" & c, d) And Command11.Caption <> alldclass("c" & c, d) And Command12.Caption <> alldclass("c" & c, d) And Command13.Caption <> alldclass("c" & c, d) And Command14.Caption <> alldclass("c" & c, d) And Command15.Caption <> alldclass("c" & c, d) And Command16.Caption <> alldclass("c" & c, d) And Command17.Caption <> alldclass("c" & c, d) And Command18.Caption <> alldclass("c" & c, d) And Command19.Caption <> alldclass("c" & c, d) Then
    Command20.Caption = alldclass("c" & c, d)
    b = b + 1
   End If
  End If
 Next c
Next d
End Sub

Public Function allclass(s$) As String
Dim classlist As String
Open "data\z" & wk & ".txt" For Input As #1
Line Input #1, classlist
Close #1
If Len(s) = 2 Then
 allclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 2, 3))
Else
 allclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 3, 3))
End If
End Function

Public Function alldclass(s$, z%) As String
Dim classlist As String
Open "data\z" & z & ".txt" For Input As #1
Line Input #1, classlist
Close #1
If Len(s) = 2 Then
 alldclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 2, 3))
Else
 alldclass = Trim(Mid(classlist, InStr(classlist, CStr(s)) + 3, 3))
End If
End Function

