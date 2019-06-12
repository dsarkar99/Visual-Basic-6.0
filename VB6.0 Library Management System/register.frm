VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4815
   LinkTopic       =   "Form4"
   ScaleHeight     =   6555
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register Now !"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1260
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1260
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Repeat Password"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Last name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "First name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "::Registration Panel::"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim pasd As String
Dim rpasd As String


Private Sub Command1_Click()
fname = Text1.Text
lname = Text2.Text
userid = Text3.Text
pasd = Text4.Text
rpasd = Text5.Text
If (Len(fname) <> 0 And Len(lname) <> 0 And Len(userid) <> 0 And Len(pasd) <> 0 And Len(rpasd) <> 0) Then
If pasd = rpasd Then
    sql = "Insert into table1 Values(' " & userid & " ',' " & pasd & " ',' " & fname & " ',' " & lname & " ') "
    Set rs = cn.Execute(sql)
MsgBox "User Created" & nam
Form1.Visible = True
Form4.Visible = False
Else
MsgBox "Password did't Match!"
End If
Else
MsgBox "Please don't leave any fields empty!"
End If
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
    cn.ConnectionString = "DSN=users"
    cn.Open
End Sub

