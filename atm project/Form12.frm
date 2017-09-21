VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H8000000E&
   Caption         =   "Form12"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   LinkTopic       =   "Form12"
   ScaleHeight     =   6255
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1920
      TabIndex        =   2
      Top             =   2400
      Width           =   7215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   840
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
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT TYPE:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1935
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
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
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
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command4_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\BANKINFO.mdb;"
ACCTYPE = Combo1.Text
ACC = Text2.Text
sqlq = "select * from code where pin='" & Form2.Text1.Text & "' and accno='" & ACC & "'"
rs.Open (sqlq), conn, adOpenStatic, adLockReadOnly
If (rs.Fields(0) = "" And rs.Fields(1) = "") Then
d = MsgBox("Account Number Is Incorrect!Please Try Again", vbCritical, "Error")
End
End If
rs.Close
DataEnvironment1.Connection1.Open
DataEnvironment1.rsCommand1.Open "select * from transaction1 where AccNo = '" & ACC & "'", conn, adOpenDynamic, adLockOptimistic
rs.Open "select * from transaction1 where AccNo = '" & ACC & "'", conn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
DataReport1.Show
End If
rs.Close
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
