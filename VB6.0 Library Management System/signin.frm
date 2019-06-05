VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LMS"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "*"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "New User , SIGN UP now?"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "User Id :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim fnam As String
Dim lnam As String


Private Sub Command1_Click()
    id = Text1.Text
    psw = Text2.Text
    c = 0
    sql = "select * from table1 where userid='" & id & "' and pass='" & psw & "'"
    Set rs = cn.Execute(sql)
    While Not rs.EOF
        c = c + 1
        fnam = rs.Fields("fname")
        lnam = rs.Fields("lname")
        Label5.Caption = fnam & " " & lnam
        rs.MoveNext
    Wend
    If c > 0 Then
     
    MsgBox "Valid user " & nam
    Form1.Visible = "False"
    Form2.Visible = "True"
   
    Else
    MsgBox "Invalid User"
    End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
    Label4.FontBold = True
    Label4.FontUnderline = True
    Label4.ForeColor = vbBlue
    Set cn = New ADODB.Connection
    cn.ConnectionString = "DSN=users"
    cn.Open
    nam = ""
    Label5.Caption = ""
End Sub

Private Sub Label4_Click()
Form4.Visible = True
Form1.Visible = False
End Sub
