VERSION 5.00
Begin VB.Form help 
   Caption         =   "Help"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   $"help.frx":0000
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label email 
      Caption         =   "vana11886@yahoo.com"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As New CHyperlinks

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()


email = email




End Sub
Private Sub email_Click()
sendemail
End Sub
