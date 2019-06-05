VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11895
   LinkTopic       =   "Form3"
   ScaleHeight     =   4695
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "By Topic"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "By Author name"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "By Book Name"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
sql = "select * from bookTable where "
If Len(test1.Text) > 0 Then
    bname = Text1.Text
    sql = sql & "b_name = '" & bname & "'"
    Set rs = cn.Execute(sql)
    While (rs.EOF)
        Print (rs.Fields(b_author))
        rs.MoveNext
    Wend
End If
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=C:\Users\asus\Documents\bookdb.mdb"
'cn.ConnectionString = "DSN=booksdsn"
'cn.Open
End Sub
