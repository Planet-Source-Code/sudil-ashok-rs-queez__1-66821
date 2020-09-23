VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   Icon            =   "RSQueez.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "www.regider.co.nr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   5025
      Left            =   -120
      Picture         =   "RSQueez.frx":1708A
      Top             =   -120
      Width           =   7590
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
main.Show
End Sub
