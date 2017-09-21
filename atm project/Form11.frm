VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H8000000E&
   Caption         =   "Form11"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form11"
   ScaleHeight     =   7800
   ScaleWidth      =   10140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "CONFIRM NEW PIN:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "NEW PIN:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YourBank ATM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim abc As New ADODB.Recordset
Dim sqlStr As String
Private Sub Command1_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\BANKINFO.mdb;"
If (Text1.Text <> Text2.Text) Then
d = MsgBox("Passwords Don't Match", vbCritical, "Error")
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
sqlsearch = "select * from code where pin='" & Form2.Text1.Text & "'"
rs.Open (sqlsearch), conn, adOpenStatic, adLockReadOnly
If rs.Fields(0) <> "" Then
rs.Close
sqlsearch = "select * from code where pin='" & Form2.Text1.Text & "'"
rs.Open (sqlsearch), conn, adOpenDynamic, adLockOptimistic
rs.Fields("pin") = Text1.Text
rs.Update
rs.Close
d = MsgBox("Pin Change Succesful.Thank You For Banking With Us!", vbInformation, "MESSAGE")
End
Else
d = MsgBox("Account doesn't exist!", vbCritical, "Error")
rs.Close
End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub

