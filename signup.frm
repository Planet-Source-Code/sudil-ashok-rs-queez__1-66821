VERSION 5.00
Object = "{2E4F703B-D223-48AF-AAA5-74361825BBF9}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form signup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RS Queez`"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "signup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "signup.frx":1708A
   ScaleHeight     =   4755
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtemail 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox txtadd 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox txtoccu 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtage 
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   3975
   End
   Begin b8Controls4.b8ToolButton btnwuit 
      Height          =   735
      Left            =   3480
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Picture         =   "signup.frx":2BC4A
      BackColor       =   16119285
      Caption         =   "Cancel"
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
   Begin b8Controls4.b8ToolButton b8ToolButton1 
      Height          =   735
      Left            =   1080
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Picture         =   "signup.frx":2C8EE
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
   Begin VB.Label labname 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labadd 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
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
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labemail 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID :"
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
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
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
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation : "
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
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Age : "
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
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
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
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Registration Area"
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
      Left            =   -600
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Number_Only(KeyAscii As Integer, Text As String)
    Dim strvalid As String * 11
    If InStr(Text, ".") = 0 Then
        strvalid = "0123456789."
    Else
        strvalid = "0123456789"
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    ElseIf InStr(strvalid, Chr(KeyAscii)) = 0 And KeyAscii > 26 Then
        KeyAscii = 0
    End If
End Sub

Private Sub b8ToolButton1_Click()
Open App.Path & "\data\info.rsq" For Append As #1
Write #1, txtname, txtage, txtoccu, txtadd, txtemail
Close
labname = txtname: labadd = txtadd: labemail = txtemail
MsgBox "!! Thanks " & txtname & ", you are sucessfully registered!", vbInformation, "RS Queez`"

Me.Hide
quiz.Show

End Sub

Private Sub btnwuit_Click()
Unload Me
main.Show
End Sub


Private Sub txtage_KeyPress(KeyAscii As Integer)
Number_Only KeyAscii, txtage
End Sub
