VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   4320
      Top             =   6360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   240
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Auto-resetting this in 3 seconds !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   3615
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Give no. less than 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "Fibonacci Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim disp As String
Dim n, x, y, newnum, counter, countertemp As Integer

Private Sub Command1_Click()
n = Val(Text1.Text)
If Not IsNumeric(Text1.Text) Or n > 500 Then
MsgBox "Give a number less than 500"
Text1.Text = ""
Exit Sub
End If
disp = "1 1 "
x = 1
y = 1
counter = 3
Label3.Caption = disp
Timer1.Enabled = True
Text1.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
Label4.Visible = False
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Timer1_Timer()
newnum = x + y
disp = disp & newnum & " "
Label3.Caption = disp
x = y
y = newnum
counter = counter + 1

If counter > n Then
Timer1.Enabled = False
Timer2.Enabled = True
Label4.Visible = True
End If
End Sub

Private Sub Timer2_Timer()
Label3.Caption = ""
Text1.Enabled = True
Command1.Enabled = True
Command1.Caption = "Do it Again"
Text1.Text = ""
Timer2.Enabled = False
Label4.Visible = False
End Sub
