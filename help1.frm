VERSION 5.00
Begin VB.Form help1 
   Caption         =   "Help "
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   Icon            =   "help1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "This allows you to extend the amount of number boxes to add.  Click and find out!"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "help1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
