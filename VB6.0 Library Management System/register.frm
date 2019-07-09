VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4725
   LinkTopic       =   "Form4"
   ScaleHeight     =   484.5
   ScaleMode       =   0  'User
   ScaleWidth      =   236.25
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Sex"
      Height          =   735
      Left            =   1200
      TabIndex        =   16
      Top             =   6240
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   1200
      TabIndex        =   12
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register Now !"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   7560
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Contact Number"
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
      Left            =   1260
      TabIndex        =   15
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Address"
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
      Left            =   1260
      TabIndex        =   13
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Click to get id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3480
      Width           =   2055
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
      TabIndex        =   10
      Top             =   8160
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
      TabIndex        =   7
      Top             =   7200
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
      TabIndex        =   5
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
Dim val As String
Dim fname As String



Private Sub Command1_Click()
fname = Text1.Text
lname = Text2.Text
adres = Text3.Text
phno = Text6.Text
userid = Label7.Caption
pasd = Text4.Text
rpasd = Text5.Text
If (Len(fname) <> 0 And Len(lname) <> 0 And Len(userid) <> 0 And Len(pasd) <> 0 And Len(rpasd) <> 0) Then
If Option1 = True Then
sex = "Male"
ElseIf Option2 = True Then
sex = "Female"
End If
If pasd = rpasd Then
    sql = "Insert into table1 Values('" & val & "','" & userid & "','" & pasd & "','" & fname & "','" & lname & "','" & adres & "','" & phno & "','" & sex & "') "
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
Randomize
val = Int((200000 * Rnd) + 1)
Set cn = New ADODB.Connection
    cn.ConnectionString = "DSN=users"
    cn.Open
    
    
End Sub


Private Sub Label7_Click()
Randomize
Val2 = Int((10 * Rnd) + 1)
Label7.Caption = Text1.Text & Val2
Label7.Enabled = False
End Sub

