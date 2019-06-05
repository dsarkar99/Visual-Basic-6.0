VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Calculator"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton clear 
      Caption         =   "CE"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton plus 
      BackColor       =   &H80000007&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   16
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton equalto 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2400
      TabIndex        =   15
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton zero 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1440
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton dec 
      BackColor       =   &H80000007&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   3360
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton three 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   2400
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton two 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton one 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton intu 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton six 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton five 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton four 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton divide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton nine 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton eight 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton seven 
      BackColor       =   &H80000007&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Caption1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   19
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim val1 As Double
Dim sign As String
Dim val2 As Double
Dim flag As Integer
Dim operator As Boolean
Dim val3 As Double
Dim temp As Integer



Private Sub clear_Click()
Label1.Caption = ""
Caption1.Caption = ""
End Sub

Private Sub Form_Load()
Caption1.Caption = ""
val1 = 0
val2 = 0
sign = 0
flag = 0
val3 = 0
temp = 0
operator = False

End Sub
Private Sub Command1_Click()
Caption1.Caption = ""
Label1.Caption = ""
val1 = 0
val2 = 0
val3 = 0
temp = 0
flag = 0

End Sub
Private Sub eight_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
Caption1.Caption = Caption1.Caption & 8
operator = False
End Sub
Private Sub five_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 5
End Sub
Private Sub four_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 4
End Sub


Private Sub nine_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 9
End Sub
Private Sub one_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 1
End Sub



Private Sub seven_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 7
End Sub
Private Sub six_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 6
End Sub
Private Sub three_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 3
End Sub
Private Sub two_Click(Index As Integer)
If operator = True Then Caption1.Caption = ""
operator = False
Caption1.Caption = Caption1.Caption & 2
End Sub
Private Sub dec_Click(Index As Integer)
Caption1.Caption = Caption1.Caption & "."
End Sub
Private Sub equalto_Click(Index As Integer)
val2 = Caption1.Caption
If sign = "+" Then
Label1.Caption = val1 & sign & val2
Caption1.Caption = val1 + val2
ElseIf sign = "-" Then
Label1.Caption = val1 & sign & val2
Caption1.Caption = val1 - val2
ElseIf sign = "x" Then
Label1.Caption = val1 & sign & val2
Caption1.Caption = val1 * val2
ElseIf sign = "/" Then
Label1.Caption = val1 & sign & val2
Caption1.Caption = val1 / val2
End If
End Sub
Private Sub intu_Click(Index As Integer)
If Not Caption1.Caption = Empty Then
val1 = Caption1.Caption
sign = "x"
Label1.Caption = Caption1.Caption & sign
operator = True
Else
MsgBox "Empty Click", vbCritical
End If
End Sub

Private Sub plus_Click(Index As Integer)
operator = True
val1 = Caption1.Caption
sign = "+"
Label1.Caption = val1 & sign
End Sub
Private Sub divide_Click(Index As Integer)
If Not Caption1.Caption = Empty Then
val1 = Caption1.Caption
sign = "/"
Label1.Caption = Caption1.Caption & sign
operator = True
Else
MsgBox "Empty Click", vbCritical
End If
End Sub

