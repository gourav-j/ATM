VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000E&
   Caption         =   "Form2"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14265
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   14265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Height          =   615
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Height          =   615
      Index           =   0
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   3240
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton exit 
      BackColor       =   &H8000000D&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H8000000D&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton proceed 
      BackColor       =   &H8000000D&
      Caption         =   "PROCEED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   4320
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   3240
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   2160
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   4320
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3240
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2160
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4320
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3240
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H8000000D&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2160
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   4
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   4125
      Left            =   7200
      Picture         =   "Form2.frx":0000
      Top             =   2640
      Width           =   5250
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "Please ENTER Your PIN :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
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
      Left            =   3240
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String
Private Sub clear_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub digits_Click(Index As Integer)
Text1.Text = Text1.Text + digits(Index).Caption
End Sub

Private Sub exit_Click()
d = MsgBox("Are you sure?", vbOKCancel, "Confirm")
If (d = 1) Then
d = MsgBox("Thank You For Banking With Us.Have A Nice Day!", vbInformation, "MESSAGE")
End
End If
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\BANKINFO.mdb;"

End Sub

Private Sub proceed_Click()
searchvar = Text1.Text
SQLSRCH = "select * from code where pin=" & "'" & searchvar & "'"
rs.Open (SQLSRCH), conn, adOpenStatic, adLockReadOnly
If rs.Fields(0) <> "" Then
Form3.Show
Me.Hide
Else
d = MsgBox("Pin Number Is Incorrect!", vbCritical, "Pin error")
Text1.Text = ""
Text1.SetFocus
End If
rs.Close
End Sub
