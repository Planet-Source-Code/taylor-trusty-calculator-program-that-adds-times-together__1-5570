VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form eform 
   Caption         =   "Taylor's Adding Times Program"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "eform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "?"
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Less"
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton move 
      Caption         =   "Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox totalb 
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox totala 
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox twod 
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox oned 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox twoc 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox onec 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox twob 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox oneb 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox twoa 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox onea 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Caption         =   ":"
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "Total Time:"
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   3000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label8 
      Caption         =   ":"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "Fourth Time:"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   ":"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Third Time:"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   ":"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Second Time:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   ":"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "First Time:"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu itmfile 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
      End
      Begin VB.Menu itmexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu itmhelp 
      Caption         =   "&Help"
      Begin VB.Menu itmhelpme 
         Caption         =   "H&elp me"
      End
      Begin VB.Menu itmabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "eform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmMain.Show
Unload eform
End Sub

Private Sub Command2_Click()
Help2.Show
End Sub

Private Sub itmAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdAdd_Click()
    
    If onea.Text = "" Then
        onea.Text = "0"
        End If
    If twoa.Text = "" Then
        twoa.Text = "0"
        End If
    If oneb.Text = "" Then
        oneb.Text = "0"
        End If
    If twob.Text = "" Then
        twob.Text = "0"
        End If
    If onec.Text = "" Then
        onec.Text = "0"
        End If
    If oned.Text = "" Then
        oned.Text = "0"
        End If
    If twoc.Text = "" Then
        twoc.Text = "0"
        End If
    If twod.Text = "" Then
        twod.Text = "0"
        End If
    
    
    
    
        
        
    Dim x As Integer
    
    Dim y As Integer
    
    Dim a As Integer
    
    Dim b As Integer
    
    Dim c As Integer
    
    Dim d As Integer
    
    Dim e As Integer
        
    Dim f As Integer
    
    Dim z As Integer
    
    Dim p As Integer
       
    x = CInt(onea.Text)
    
    y = CInt(oneb.Text)
    
    a = CInt(twoa.Text)
    
    b = CInt(twob.Text)
    
    c = CInt(twoc.Text)
    
    d = CInt(twod.Text)
    
    e = CInt(onec.Text)
    
    f = CInt(oned.Text)
    
    z = x + y + e + f
          
    p = a + b + c + d
        If p >= 60 And p < 120 Then
            z = z + 1
        If p > 59 And p < 120 Then
            p = p - 60
        End If
            End If
        If p >= 120 And p < 180 Then
            z = z + 2
            p = p - 120
        End If
            
        If p >= 180 And p < 240 Then
            z = z + 3
            p = p - 180
        End If
        If p >= 240 And p < 300 Then
            z = z + 4
            p = p - 240
            End If
        
    totala.Text = CStr(z)
    
    totalb.Text = CStr(p)
    
End Sub

Private Sub helpextend_Click()
help1.Show
End Sub

Private Sub itmhelpme_Click()
help.Show
End Sub

Private Sub Helpmove_Click()
Help2.Show
End Sub


Private Sub itmExit_Click()
Dim Msg$               'Message box message
    Dim OpVal%             'Option value variable
    Dim RetVal%            'variable for return value
    Dim TitleMsg$          'Title message variable
    
Msg$ = "Are you sure you want to exit?"
   
OpVal% = vbExclamation + vbYesNo + vbDefaultButton2
   
TitleMsg$ = "Are you sure you want to Exit?"

RetVal% = MsgBox(Msg$, OpVal%, TitleMsg$)

If RetVal% = vbYes Then
   Unload Me
   End If

End Sub


Private Sub move_Click()
onea.Text = ""
twoa.Text = ""
twoa.Text = totalb.Text
onea.Text = totala.Text
totala.Text = ""
totalb.Text = ""
oneb.Text = ""
twob.Text = ""
onec.Text = ""
oned.Text = ""
twoc.Text = ""
twod.Text = ""
End Sub

Private Sub new_Click()
    onea.Text = ""
    oneb.Text = ""
    onec.Text = ""
    oned.Text = ""
    twoa.Text = ""
    twob.Text = ""
    twoc.Text = ""
    twod.Text = ""
    totala.Text = ""
    totalb.Text = ""
End Sub


