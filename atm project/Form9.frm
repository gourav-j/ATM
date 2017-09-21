VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H8000000E&
   Caption         =   "Form9"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   LinkTopic       =   "Form9"
   ScaleHeight     =   6255
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   7335
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2400
         TabIndex        =   19
         Text            =   "Combo2"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
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
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
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
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYEE'S ACCOUNT TYPE:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1920
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   960
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
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
         Left            =   240
         TabIndex        =   11
         Top             =   960
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Text            =   "Combo1"
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
      TabIndex        =   3
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
      Left            =   1800
      TabIndex        =   2
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
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim abc As New ADODB.Recordset
Dim abcd As New ADODB.Recordset
Dim sqlStr As String
Private Sub Command1_Click()
Frame1.Visible = True
Label4.Visible = True
Label5.Visible = True
Text1.SetFocus
Command3.Visible = True
Command4.Visible = True
End Sub

Private Sub Command2_Click()
Form3.Show
Unload Me
End Sub
Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub
Private Sub Command4_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\BANKINFO.mdb;"
ACCTYPE = Combo1.Text
amount = Val(Text1.Text)
ACC = Text2.Text
sqlq = "select * from code where pin='" & Form2.Text1.Text & "' and accno='" & ACC & "'"
abc.Open (sqlq), conn, adOpenStatic, adLockReadOnly
If (abc.Fields(0) = "" And abc.Fields(1) = "") Then
d = MsgBox("Account Number Is Incorrect!", vbCritical, "Error")
End
End If
abc.Close
SQLSRCH = "select * from datain where accno='" & ACC & "' and acctype='" & ACCTYPE & "'"
rs.Open (SQLSRCH), conn, adOpenStatic, adLockReadOnly
bal = Val(rs.Fields("balance"))
If rs.Fields(4) <> "" And rs.Fields(6) <> "" Then
    If (bal < amount) Then
    d = MsgBox("Balance Is Not Sufficient.Please Try Again", vbCritical, "Error")
    Unload Me
    Form3.Show
    End If
    rs.Close
    Frame2.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Combo2.Visible = True
    Combo2.AddItem ("SAVINGS")
    Combo2.AddItem ("CURRENT")
    Combo2.Text = "SAVINGS"
    Text3.Visible = True
    Command5.Visible = True
    Command6.Visible = True
    Else
    d = MsgBox("Account Doesn't Exist!", vbCritical, "Error")
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    rs.Close
    End If
End Sub

Private Sub Command5_Click()
d = MsgBox("Are you sure?", vbOKCancel, "CONFIRM")
If (d = 2) Then
'd = MsgBox("Your Transaction Was Cancelled.Thank You For Banking With Us!", vbInformation, "MESSAGE")
Unload Me
Form3.Show
Else
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\BANKINFO.mdb;"
ACCTYPE = Combo2.Text
ACC = Text3.Text
sqlq = "select * from code where accno='" & ACC & "'"
abc.Open (sqlq), conn, adOpenStatic, adLockReadOnly
If (abc.Fields(1) = "") Then
    d = MsgBox("Account Number Is Incorrect!Please Try Again", vbCritical, "Error")
    End
End If
abc.Close
SQLSRCH = "select * from datain where accno='" & ACC & "' and acctype='" & ACCTYPE & "'"
rs.Open (SQLSRCH), conn, adOpenStatic, adLockReadOnly
bal = Val(rs.Fields("balance"))
If rs.Fields(4) <> "" And rs.Fields(6) <> "" Then
    rs.Close
    sqlq = "select * from datain where accno='" & ACC & "' and acctype='" & ACCTYPE & "'"
    rs.Open (sqlq), conn, adOpenDynamic, adLockOptimistic
    sqlt = "select * from datain where accno='" & Text2.Text & "' and acctype='" & Combo1.Text & "'"
    abcd.Open (sqlt), conn, adOpenDynamic, adLockOptimistic
    bal2 = abcd.Fields("balance")
    amount = Val(Text1.Text)
    rs.Fields("balance") = CStr(bal + amount)
    rs.Update
    rs.Close
    abcd.Fields("balance") = CStr(bal2 - amount)
    abcd.Update
    SQLSRCH = "select * from transaction1"
    rss.Open (SQLSRCH), conn, adOpenDynamic, adLockOptimistic
    rss.AddNew
    rss.Fields("Date") = Format$(Now, "dd/mm/yyyy")
    rss.Fields("AccNo") = ACC
    rss.Fields("Description") = "Money Transfer" + Text2.Text
    rss.Fields("Deposit") = amount
    rss.Fields("Balance") = CStr(bal + amount)
    rss.Update
    rss.AddNew
    rss.Fields("Date") = Format$(Now, "dd/mm/yyyy")
    rss.Fields("AccNo") = Text2.Text
    rss.Fields("Description") = "Money Transfer" + ACC
    rss.Fields("Withdrawal") = amount
    rss.Fields("Balance") = CStr(bal2 - amount)
    rss.Update
    abcd.Close
    rss.Close
    Me.Hide
    frmSplash3.Show
Else
    d = MsgBox("Account Doesn't Exist!", vbCritical, "Error")
    Combo2.Text = "SAVINGS"
    Text3.Text = ""
    Combo2.SetFocus
    rs.Close
    End If
End If
End Sub

Private Sub Command6_Click()
Unload Me
Form3.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem ("SAVINGS")
Combo1.AddItem ("CURRENT")
Combo1.Text = "SAVINGS"
End Sub
