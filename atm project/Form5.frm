VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000E&
   Caption         =   "Form5"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
   LinkTopic       =   "Form5"
   ScaleHeight     =   6585
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   2280
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "YOUR BALANCE IS Rs."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
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
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "PROCEED"
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NO:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "PROCEED"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form5.frx":0000
      Left            =   5400
      List            =   "Form5.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT ACCOUNT TYPE:"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
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
      Left            =   2400
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim abc As New ADODB.Recordset
Dim sqlStr As String

Private Sub Command2_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command5_Click()
d = MsgBox("Would You Like To Perform Another Transaction?", vbOKCancel, "MESSAGE")
If (d = 1) Then
Unload Me
Form3.Show
Else
d = MsgBox("Thank You For Banking With Us!", vbInformation, "MESSAGE")
End
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem ("SAVINGS")
Combo1.AddItem ("CURRENT")
Combo1.Text = "SAVINGS"

End Sub
Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub
Private Sub Command1_Click()
Frame1.Visible = True
Label5.Visible = True
Text2.SetFocus
Command3.Visible = True
Command4.Visible = True
Label3.Visible = False
End Sub
Private Sub Command4_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\BANKINFO.mdb;"
ACCTYPE = Combo1.Text

ACC = Text2.Text
sqlq = "select * from code where pin='" & Form2.Text1.Text & "' and accno='" & ACC & "'"
abc.Open (sqlq), conn, adOpenStatic, adLockReadOnly
If (abc.Fields(0) = "" And abc.Fields(1) = "") Then
d = MsgBox("Account Number Is Incorrect!Please Try Again", vbCritical, "Error")
End
End If
abc.Close
SQLSRCH = "select * from datain where accno='" & ACC & "' and acctype='" & ACCTYPE & "'"
rs.Open (SQLSRCH), conn, adOpenStatic, adLockReadOnly
If rs.Fields(4) <> "" And rs.Fields(6) <> "" Then
bal = rs.Fields("balance")
Frame2.Visible = True
Command5.Visible = True
Label6.Visible = True
Label6.Caption = Label6.Caption + bal
rs.Close
Else
    d = MsgBox("Account Doesn't Exist!", vbCritical, "Error")
    Text2.Text = ""
    Text2.SetFocus
    rs.Close
    End If
End Sub
