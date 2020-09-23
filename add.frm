VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Taylor's adding times Program"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4290
   Icon            =   "add.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Helpmove 
      Caption         =   "?"
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton helpextend 
      Caption         =   "?"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton extend 
      Caption         =   "More"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CDmain 
      Left            =   1680
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   2880
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox totalb 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox twob 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox twoa 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox totala 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox oneb 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox onea 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line2 
      X1              =   2760
      X2              =   2760
      Y1              =   240
      Y2              =   2040
   End
   Begin VB.Label jj 
      Caption         =   ":"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   ":"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   ":"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   840
      Width           =   135
   End
   Begin VB.Label lblTotla 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Time:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   2760
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblNumTwo 
      Alignment       =   1  'Right Justify
      Caption         =   "Second Time:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblNumONe 
      Alignment       =   1  'Right Justify
      Caption         =   "First Time:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu hyphen 
         Caption         =   "-"
      End
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu itmhelp 
      Caption         =   "&Help"
      Begin VB.Menu helpme 
         Caption         =   "Help"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub About_Click()
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
    
        
        
    Dim x As Integer
    
    Dim y As Integer
    
    Dim a As Integer
    
    Dim b As Integer
    
    Dim z As Integer
    
    Dim p As Integer
       
    x = CInt(onea.Text)
    
    y = CInt(oneb.Text)
    
    a = CInt(twoa.Text)
    
    b = CInt(twob.Text)
    
    z = x + y
          
    p = a + b
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
            
        
        
    totala.Text = CStr(z)
    
    totalb.Text = CStr(p)
    
End Sub

Private Sub extend_Click()
eform.Show
Unload frmMain
End Sub

Private Sub helpextend_Click()
help1.Show
End Sub

Private Sub helpme_Click()
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
End Sub

Private Sub new_Click()
    onea.Text = ""
    oneb.Text = ""
    twoa.Text = ""
    twob.Text = ""
    totala.Text = ""
    totalb.Text = ""
End Sub

