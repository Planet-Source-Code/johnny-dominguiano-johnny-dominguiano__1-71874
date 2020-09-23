VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIssue 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISSUE"
   ClientHeight    =   8880
   ClientLeft      =   3648
   ClientTop       =   732
   ClientWidth     =   6720
   Icon            =   "frmIssue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIssue.frx":058A
   ScaleHeight     =   8880
   ScaleWidth      =   6720
   Begin VB.TextBox txtTemp 
      Height          =   288
      Left            =   7320
      TabIndex        =   26
      Top             =   2760
      Width           =   972
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   3360
      MouseIcon       =   "frmIssue.frx":0AE1
      MousePointer    =   99  'Custom
      Picture         =   "frmIssue.frx":0C33
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancel"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "Issue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2040
      MouseIcon       =   "frmIssue.frx":11B3
      MousePointer    =   99  'Custom
      Picture         =   "frmIssue.frx":1305
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Issue book"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Book's Information"
      Height          =   3012
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   6132
      Begin VB.CommandButton Command1 
         Height          =   492
         Left            =   3360
         Picture         =   "frmIssue.frx":18FF
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   492
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1320
         TabIndex        =   11
         Top             =   444
         Width           =   1932
      End
      Begin VB.Label lblISBN 
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   23
         Top             =   2400
         Width           =   3492
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
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
         Left            =   1080
         TabIndex        =   22
         Top             =   1920
         Width           =   3492
      End
      Begin VB.Label lblCategory 
         BackStyle       =   0  'Transparent
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
         Left            =   1440
         TabIndex        =   21
         Top             =   1440
         Width           =   1932
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   4332
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   852
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   852
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1092
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1332
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Borrower's Information"
      Height          =   3012
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   6132
      Begin MSMask.MaskEdBox txtIssue 
         Height          =   252
         Left            =   1920
         TabIndex        =   18
         Top             =   1680
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtReturn 
         Height          =   252
         Left            =   1920
         TabIndex        =   19
         Top             =   2160
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   ","
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
         Left            =   2880
         TabIndex        =   17
         Top             =   1080
         Width           =   252
      End
      Begin VB.Label lblFirstname 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3240
         TabIndex        =   16
         Top             =   1080
         Width           =   1452
      End
      Begin VB.Label lblLastname 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1080
         TabIndex        =   15
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Issue:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1812
      End
      Begin VB.Label lblID 
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   852
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Height          =   3012
      Left            =   480
      TabIndex        =   28
      Top             =   4680
      Width           =   6132
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   3012
      Left            =   480
      TabIndex        =   27
      Top             =   960
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Issue Details"
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
      Height          =   732
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6732
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Public Sub addinHand()
Set rs = Nothing
    rs.Open "select * from member where member_ID = '" & (lblID.Caption) & "' ", conn, adOpenDynamic, adLockOptimistic
    rs.Fields(8) = val(rs.Fields(8)) + 1
    rs.Fields(9) = val(rs.Fields(9)) + 1
    rs.Update

End Sub
 Public Sub checkQuantity()
Call connect
rsSubtractBookQty.Open "Select * From BOOK where BookID = '" & (txtTemp.Text) & "'", conn, adOpenStatic, adLockOptimistic
    If rsSubtractBookQty.Fields(11) = 0 Then
    MsgBox "There are no available books.", vbInformation
    Exit Sub
    Else
    rsSubtractBookQty.Fields(12) = val(rsSubtractBookQty.Fields(12)) + 1
    rsSubtractBookQty.Fields(11) = val(rsSubtractBookQty.Fields(11)) - 1
    rsSubtractBookQty.Fields(10) = val(rsSubtractBookQty.Fields(10)) + 1
    rsSubtractBookQty.Update
   Set rsSubtractBookQty = Nothing
   End If
Exit Sub
End Sub
Private Sub cmd_cancel_Click()
unload Me
End Sub

Private Sub cmdIssue_Click()
Call connect
If txtSearch.Text = "" Then
    MsgBox "Please enter Book ID.", vbInformation
Else

Dim str As String
    Call connect
    Set rs = Nothing
    str = "Select count(*) from Issued where Bookid = '" & Trim(txtSearch.Text) & "' And Memberid = '" & Trim(lblID.Caption) & "'"
    rs.Open str, conn, adOpenStatic, adLockOptimistic
    
    If (rs(0) <> 0) Then
        MsgBox ("You have already borrowed that book."), vbInformation
    cmdIssue.Enabled = False
Exit Sub
Else
Call checkQuantity
   Call addinHand
Set recset = Nothing
    recset.Open "Select * from issued", conn, adOpenDynamic, adLockOptimistic
    With recset
    .AddNew
    .Fields(0) = txtSearch.Text
    .Fields(1) = lbltitle.Caption
    .Fields(2) = lblCategory.Caption
    .Fields(3) = lblID.Caption
    .Fields(4) = lblLastname.Caption
    .Fields(5) = lblFirstname.Caption
    .Fields(6) = txtIssue.Text
    .Fields(7) = txtReturn.Text
    .Fields(9) = lblAuthor.Caption
    .Update
End With
MsgBox "Issue Info.:MemberId=" & CDbl(lblID.Caption) & " And  BookId=" & CDbl(txtSearch.Text), vbInformation
If MsgBox("Issue another book?", vbYesNo) = vbYes Then
cmdIssue.Enabled = False
txtSearch.Text = ""
lbltitle.Caption = ""
lblCategory.Caption = ""
lblAuthor.Caption = ""
lblISBN.Caption = ""
Else
unload Me
End If
End If
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If txtSearch.Text = "" Then
    MsgBox "Please enter the Book ID.", vbInformation
    txtSearch.SetFocus
Else
Dim a
Call connect
    Set rs = Nothing
    rs.Open "Select * from BOOK where BookID like '" & txtSearch.Text & "'", conn, adOpenDynamic
    While rs.EOF = False
    MsgBox "Record Found!", vbInformation
    cmdIssue.Enabled = True
    lbltitle.Caption = rs.Fields(1)
    lblCategory.Caption = rs.Fields(3)
    lblAuthor.Caption = rs.Fields(4)
    lblISBN.Caption = rs.Fields(6)
    rs.MoveNext
    txtTemp.Text = txtSearch.Text
    a = 1
    Wend
    If a <> 1 Then
    MsgBox "Book Not found!", vbInformation
    End If
End If


End Sub

Private Sub txtIssue_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtReturn_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 13 Then
    Command1_Click
Else
    KeyAscii = 0
End If
End Sub
