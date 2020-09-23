VERSION 5.00
Begin VB.Form frmEnter 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transact"
   ClientHeight    =   2352
   ClientLeft      =   5016
   ClientTop       =   4704
   ClientWidth     =   4092
   Icon            =   "frmEnter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2352
   ScaleWidth      =   4092
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Transact"
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1212
   End
   Begin VB.TextBox txtMemberID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   612
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4092
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID:"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub returnBooks()
Set recset = Nothing
   recset.Open "Select * from ISSUED where MemberID like '%" & frmEnter.txtMemberID.Text & "%' order by bookID", conn, adOpenDynamic, adLockOptimistic
   With frmReturn
   While recset.EOF = False
    .ListIssue.ListItems.Add , , recset!bookid
    .ListIssue.ListItems(.ListIssue.ListItems.Count).ListSubItems.Add , , recset!Title
    .ListIssue.ListItems(.ListIssue.ListItems.Count).ListSubItems.Add , , recset!category
    .ListIssue.ListItems(.ListIssue.ListItems.Count).ListSubItems.Add , , recset!Date_borrowed
    .ListIssue.ListItems(.ListIssue.ListItems.Count).ListSubItems.Add , , recset!Date_due

    frmReturn.lblID.Caption = recset.Fields(3)
    frmReturn.lblFirstname.Caption = recset.Fields(5)
    frmReturn.lblLastname.Caption = recset.Fields(4)
    frmReturn.lblReturn.Caption = Format$(Now, "mm/dd/yyyy")

    recset.MoveNext
    Wend
End With
    unload Me
    frmReturn.show 1
End Sub
Private Sub issueBooks()
    frmIssue.lblID.Caption = rs.Fields(0)
    frmIssue.lblFirstname.Caption = rs.Fields(2)
    frmIssue.lblLastname.Caption = rs.Fields(1)
    frmIssue.txtIssue.Text = Format$(Now, "mm/dd/yyyy")
    frmIssue.txtReturn.Text = Format$(Now + 2, "mm/dd/yyyy")
    unload Me
    frmIssue.show 1
Exit Sub
End Sub

Private Sub cmdEnter_Click()
If txtMemberID.Text = "" Then
    MsgBox "Please enter the ID.", vbInformation
Else
Dim a
Call connect
Set rs = Nothing
    rs.Open "Select * from MEMBER where member_ID like '" & txtMemberID.Text & "'", conn, adOpenDynamic
    While rs.EOF = False
    GoTo show
    rs.MoveNext
    a = 1
    Wend

    If a <> 1 Then
    MsgBox "Member Not found!", vbInformation, "Confirmation"
    txtMemberID.SetFocus
    Exit Sub
    unload Me
    End If
    
End If


show:
If Banner.Caption = "Issue" Then
    issueBooks
End If
If Banner.Caption = "Return" Then

If rs.Fields(8) = 0 Then
    MsgBox "Nothing Borrowed.", vbInformation
    Exit Sub
Else
returnBooks
    End If
    
End If
unload:
 unload Me




End Sub

Private Sub txtMemberID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 13 Then
    cmdEnter_Click
Else
    KeyAscii = 0
End If

End Sub
