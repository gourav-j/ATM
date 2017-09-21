VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H8000000E&
   Caption         =   "Form10"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form10"
   ScaleHeight     =   5970
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Rs."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MONEY TRANSFER SUCCESSFUL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2235
      Left            =   2520
      Picture         =   "Form10.frx":0000
      Top             =   1200
      Width           =   5280
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Label2.Caption = Label2.Caption + Form9.Text1.Text
End Sub
