VERSION 5.00
Begin VB.Form frmFine 
   BackColor       =   &H8000000D&
   Caption         =   "Fine Setting"
   ClientHeight    =   1680
   ClientLeft      =   5352
   ClientTop       =   4500
   ClientWidth     =   3192
   Icon            =   "frmFine.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   3192
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   372
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox txtFine 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2652
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   372
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty per day:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2772
   End
End
Attribute VB_Name = "frmFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rs.Fields(0) = txtFine.Text
rs.Update
MsgBox "Fine per day of overdue updated.", vbInformation
unload Me
End Sub

Private Sub Command2_Click()
unload Me
End Sub

Private Sub Form_Load()
Call connect
Set rs = Nothing
rs.Open "select * from fine", conn, adOpenDynamic, adLockOptimistic
txtFine.Text = rs.Fields(0)

End Sub


Private Sub txtFine_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 13 Then
    Command1_Click
Else
    KeyAscii = 0
End If

End Sub
