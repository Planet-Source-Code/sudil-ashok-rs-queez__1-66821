VERSION 5.00
Object = "{2E4F703B-D223-48AF-AAA5-74361825BBF9}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmresult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   Icon            =   "frmresult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmresult.frx":1708A
   ScaleHeight     =   5670
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin b8Controls4.b8SBCenter b8SBCenter1 
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   4048
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label email 
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
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Email id: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label address 
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
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label labname 
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
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   735
      End
   End
   Begin b8Controls4.b8SBCenter b8SBCenter2 
      Height          =   2295
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4048
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Correct : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label c 
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
         Index           =   5
         Left            =   1560
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Wrong :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label w 
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
         Index           =   4
         Left            =   1560
         TabIndex        =   13
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Score : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label s 
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
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Score Card"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
   End
   Begin b8Controls4.b8SBCenter b8SBCenter3 
      Height          =   3615
      Left            =   3480
      TabIndex        =   17
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmresult.frx":2BC4A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Index           =   8
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Thanks"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   4335
      End
   End
   Begin b8Controls4.b8ToolButton b8ToolButton1 
      Height          =   735
      Left            =   4800
      TabIndex        =   20
      Top             =   4680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Picture         =   "frmresult.frx":2BD40
      BackColor       =   16119285
      Caption         =   "OK"
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
      ForeColor       =   8421504
      BgColorHover    =   16119285
      BgColorDown     =   16119285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   -1080
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b8ToolButton1_Click()
Unload Me
main.Show
End Sub

Private Sub Form_Load()
labname(0).Caption = signup.labname
address(1).Caption = signup.labadd
email(2).Caption = signup.labemail


c(5).Caption = quiz.correct
w(4).Caption = quiz.wrong
s(3).Caption = quiz.score

End Sub
