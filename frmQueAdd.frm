VERSION 5.00
Object = "{2E4F703B-D223-48AF-AAA5-74361825BBF9}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmQueAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Question"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmQueAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQueAdd.frx":1708A
   ScaleHeight     =   5070
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin b8Controls4.b8ToolButton b8ToolButton1 
      Height          =   855
      Left            =   1080
      TabIndex        =   12
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Picture         =   "frmQueAdd.frx":2BC4A
      BackColor       =   16119285
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   0
   End
   Begin VB.ComboBox comboAns 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3240
      Width           =   5055
   End
   Begin VB.TextBox txtOpt4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   5055
   End
   Begin VB.TextBox txtOpt3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox txtOpt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   5055
   End
   Begin VB.TextBox txtOpt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox txtQuestion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin b8Controls4.b8ToolButton b8ToolButton2 
      Height          =   855
      Left            =   3600
      TabIndex        =   13
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      Picture         =   "frmQueAdd.frx":2C6D3
      BackColor       =   16119285
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   840
   End
End
Attribute VB_Name = "frmQueAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, txta As Integer
Dim txtq, txt1, txt2, txt3, txt4 As String
Private Sub b8ToolButton1_Click()
If txtQuestion.Text = "" Or txtOpt1.Text = "" Or txtOpt2.Text = "" Or txtOpt3.Text = "" Or txtOpt4.Text = "" Then MsgBox "Invalid Entry", vbCritical, "RS Queez`": Exit Sub
Open App.Path & "\data\question.rsq" For Append As #1
Write #1, txtQuestion, txtOpt1, txtOpt2, txtOpt3, txtOpt4, (comboAns.ListIndex + 1)
Close
MsgBox "Question Saved!", vbInformation, "RS Queez`"
txtQuestion.Text = Empty
txtOpt1.Text = Empty
txtOpt2.Text = Empty
txtOpt3.Text = Empty
txtOpt4.Text = Empty

End Sub

Private Sub b8ToolButton2_Click()
Unload Me
End Sub
Private Sub Form_Load()
comboAns.Text = "Option 1"
End Sub

Private Sub Label7_Click()
Label7.Caption = comboAns.ListIndex
End Sub

Private Sub Image2_Click()
Open App.Path & "\data\question.rsq" For Input As #1
End Sub
