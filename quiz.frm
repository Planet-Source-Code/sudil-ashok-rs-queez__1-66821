VERSION 5.00
Object = "{2E4F703B-D223-48AF-AAA5-74361825BBF9}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form quiz 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RS Queez`"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "quiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "quiz.frx":1708A
   ScaleHeight     =   6375
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   10215
      TabIndex        =   25
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   5400
         TabIndex        =   26
         Top             =   0
         Width           =   9255
         Begin VB.Image Image1 
            Height          =   1200
            Left            =   -360
            Picture         =   "quiz.frx":2BC4A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5280
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RS Queez`"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label txtusername 
         BackStyle       =   0  'Transparent
         Caption         =   "New User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   28
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2400
      Top             =   1080
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   7680
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   30
         Left            =   -240
         TabIndex        =   10
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin b8Controls4.b8GradLine b8GradLine4 
         Height          =   30
         Left            =   -240
         TabIndex        =   12
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin b8Controls4.b8GradLine b8GradLine3 
         Height          =   30
         Left            =   -240
         TabIndex        =   14
         Top             =   2760
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin b8Controls4.b8GradLine b8GradLine5 
         Height          =   30
         Left            =   -240
         TabIndex        =   16
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label score 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label wrong 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label correct 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label total 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wrong"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Corrected"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Questions Asked"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Score Card"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -240
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   7695
      Begin VB.TextBox q 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   6615
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   6480
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qno :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label ans 
         Caption         =   "Label5"
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image io2 
         Height          =   585
         Left            =   480
         MousePointer    =   99  'Custom
         Picture         =   "quiz.frx":2D0EC
         Top             =   2400
         Width           =   585
      End
      Begin VB.Image io3 
         Height          =   585
         Left            =   480
         MousePointer    =   99  'Custom
         Picture         =   "quiz.frx":2D5FF
         Top             =   2880
         Width           =   585
      End
      Begin VB.Image io4 
         Height          =   585
         Left            =   480
         MousePointer    =   99  'Custom
         Picture         =   "quiz.frx":2DB12
         Top             =   3360
         Width           =   585
      End
      Begin VB.Image io1 
         Height          =   585
         Left            =   480
         MousePointer    =   99  'Custom
         Picture         =   "quiz.frx":2E025
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label op4 
         BackStyle       =   0  'Transparent
         Caption         =   "RS Queez`"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   3600
         Width           =   4935
      End
      Begin VB.Label op3 
         BackStyle       =   0  'Transparent
         Caption         =   "RS Queez`"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   3120
         Width           =   4935
      End
      Begin VB.Label op2 
         BackStyle       =   0  'Transparent
         Caption         =   "RS Queez`"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Label op1 
         BackStyle       =   0  'Transparent
         Caption         =   "RS Queez`"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Questions                     "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   -840
         TabIndex        =   4
         Top             =   360
         Width           =   10095
      End
      Begin VB.Label qno 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
   End
   Begin b8Controls4.b8GradLine b8GradLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   24
      Top             =   -240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   53
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim que, o1, o2, o3, o4, ra As String
Dim ques, rqno As Integer

Private Sub Form_Load()
qno.Caption = 1

txtusername.Caption = signup.txtname.Text

Open App.Path & "\data\question.rsq" For Input As #1

Do Until EOF(1)
Input #1, que, o1, o2, o3, o4, ra
ques = ques + 1
Loop
Close

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To Int((Rnd * ques + 1))
Input #1, que, o1, o2, o3, o4, ra
Next
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(0) = ques
Close


End Sub
Private Sub io1_Click()
If Val(ans) = 1 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub

Private Sub io2_Click()
If Val(ans) = 2 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub

Private Sub io3_Click()
If Val(ans) = 3 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub

Private Sub io4_Click()
If Val(ans) = 4 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub


Private Sub Label3_Click()
End Sub

Private Sub op1_Click()
If Val(ans) = 1 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub

Private Sub op2_Click()
If Val(ans) = 2 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": frmresult.Show: Unload Me
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub

Private Sub op3_Click()
If Val(ans) = 3 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub

Private Sub op4_Click()
If Val(ans) = 4 Then correct.Caption = correct.Caption + 1 Else wrong.Caption = wrong.Caption + 1
aa:
If qno = ques Then MsgBox "Question finished!!", vbCritical, "Regider Queez`": Exit Sub
rqno = Int((Rnd * ques + 1))
For i = 1 To qno.Caption
If Val(List1.List(i - 1)) = rqno Then GoTo aa:
Next

Open App.Path & "\data\question.rsq" For Input As #1
For i = 1 To rqno
Input #1, que, o1, o2, o3, o4, ra
Next
Close
q = que: op1 = o1: op2 = o2: op3 = o3: op4 = o4: ans = ra
List1.List(qno) = rqno
qno.Caption = qno.Caption + 1
total.Caption = qno.Caption - 1
If qno = 11 Then frmresult.Show: Unload Me
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()
score = correct * 5
End Sub
