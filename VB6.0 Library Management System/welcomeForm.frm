VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "welcomeForm.frx":0000
   ScaleHeight     =   6030
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Account"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Logout"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Book Return"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Book Issue"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form3
Form3.Visible = True
End Sub

Private Sub Command4_Click()
Form2.Visible = False
Form1.Visible = True
End Sub

Private Sub Command5_Click()
Load Form5
Form2.Visible = False
Form5.Visible = True
End Sub

Private Sub Form_Load()
Label1.Caption = "Welcome " & Form1.Label5 & ","
End Sub
