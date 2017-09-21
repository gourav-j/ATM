VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9765
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000002&
      Caption         =   "MINI STATEMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000002&
      Caption         =   "PIN CHANGE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000002&
      Caption         =   "MONEY TRANSFER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000002&
      Caption         =   "DEPOSIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000002&
      Caption         =   "BALANCE ENQUIRY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000002&
      Caption         =   "WITHDRAW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String

Private Sub Command1_Click()
Me.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Form7.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Form9.Show
Me.Hide
End Sub

Private Sub Command5_Click()
Form11.Show
Me.Hide
End Sub

Private Sub Command6_Click()
Form12.Show
Me.Hide
End Sub

