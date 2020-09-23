VERSION 5.00
Object = "{2E4F703B-D223-48AF-AAA5-74361825BBF9}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":1708A
   ScaleHeight     =   7620
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   12135
      TabIndex        =   7
      Top             =   0
      Width           =   12135
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   5400
         TabIndex        =   11
         Top             =   0
         Width           =   9255
         Begin VB.Image Image1 
            Height          =   1200
            Left            =   0
            Picture         =   "main.frx":2BC4A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5280
         End
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
         TabIndex        =   10
         Top             =   480
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
         TabIndex        =   9
         Top             =   840
         Width           =   2055
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
         TabIndex        =   8
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   -360
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin b8Controls4.b8ToolButton btnplay 
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         Picture         =   "main.frx":2D0EC
         BackColor       =   16119285
         Caption         =   "Play Quiz"
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
      Begin b8Controls4.b8ToolButton btnqueedit 
         Height          =   855
         Left            =   600
         TabIndex        =   2
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         Picture         =   "main.frx":2D7CC
         BackColor       =   16119285
         Caption         =   "Add Questions"
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
      Begin b8Controls4.b8ToolButton btnabout 
         Height          =   855
         Left            =   600
         TabIndex        =   3
         Top             =   2880
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         Picture         =   "main.frx":2E414
         BackColor       =   16119285
         Caption         =   "About Us"
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
      Begin b8Controls4.b8ToolButton btnQuit 
         Height          =   855
         Left            =   600
         TabIndex        =   4
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         Picture         =   "main.frx":2F09F
         BackColor       =   16119285
         Caption         =   "Quit"
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
         BackColor       =   &H00404040&
         Caption         =   "Main Menu"
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
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "Copyright © Regider Software 2006  "
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
      Left            =   2640
      TabIndex        =   6
      Top             =   7200
      Width           =   9375
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnabout_Click()
frmAbout.Show vbModal
End Sub

Private Sub btnhighsc_Click()
frmHighsc.Show vbModal
End Sub

Private Sub btnplay_Click()
signup.Show
Unload Me
End Sub

Private Sub btnqueedit_Click()
frmQueAdd.Show vbModal
End Sub

Private Sub btnQuit_Click()
MsgBox "Thanks for downloading this code!  Don't forget to vote us!", vbInformation, "RS Queez`"
End
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Date & "   " & Time
End Sub

