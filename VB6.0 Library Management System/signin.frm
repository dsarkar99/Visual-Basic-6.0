VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LMS"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "signin.frx":0000
   ScaleHeight     =   4905
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "New User , SIGN UP now?"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   240
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
Dim act_id As String



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
        act_id = rs.Fields("act_id")
        Label5.Caption = fnam & " " & lnam
        Label6.Caption = act_id
        rs.MoveNext
    Wend
    If c > 0 Then
     
    'MsgBox "Valid user " & nam
    Form1.Visible = "False"
    Form2.Visible = "True"
    Text1.Text = ""
    Text2.Text = ""
   
    Else
    MsgBox "Invalid User"
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
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
