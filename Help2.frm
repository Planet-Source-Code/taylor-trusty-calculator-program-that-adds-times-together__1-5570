VERSION 5.00
Begin VB.Form Help2 
   Caption         =   "Help"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   $"Help2.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Help2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
