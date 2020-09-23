VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReturn 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RETURN"
   ClientHeight    =   8760
   ClientLeft      =   3816
   ClientTop       =   1200
   ClientWidth     =   6852
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   6852
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   7560
      TabIndex        =   24
      Top             =   4560
      Width           =   1212
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
      Left            =   3600
      MouseIcon       =   "frmReturn.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmReturn.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancel"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
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
      Left            =   2280
      MouseIcon       =   "frmReturn.frx":0C5C
      MousePointer    =   99  'Custom
      Picture         =   "frmReturn.frx":0DAE
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Return book"
      Top             =   7920
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Borrowed Books"
      Height          =   3012
      Left            =   360
      TabIndex        =   4
      Top             =   4800
      Width           =   6132
      Begin MSComctlLib.ListView ListIssue 
         CausesValidation=   0   'False
         Height          =   2652
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   5892
         _ExtentX        =   10393
         _ExtentY        =   4678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483646
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Book ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Book Title"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Borrowed"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Due Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Borrower's Information"
      Height          =   3012
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   6132
      Begin VB.Label lblReturn 
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
         Left            =   1920
         TabIndex        =   14
         Top             =   2520
         Width           =   1692
      End
      Begin VB.Label lblDue 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   2040
         Width           =   1692
      End
      Begin VB.Label lblIssue 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   1560
         Width           =   1692
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Returned:"
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
         TabIndex        =   11
         Top             =   2520
         Width           =   1692
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
         Left            =   2760
         TabIndex        =   10
         Top             =   960
         Width           =   252
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
         TabIndex        =   9
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Issued:"
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
         Top             =   1560
         Width           =   1812
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1692
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
         TabIndex        =   6
         Top             =   960
         Width           =   1692
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
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   1452
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
         TabIndex        =   2
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Height          =   3252
      Left            =   240
      TabIndex        =   26
      Top             =   4680
      Width           =   6372
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   3012
      Left            =   600
      TabIndex        =   25
      Top             =   840
      Width           =   6012
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Label8"
      Height          =   492
      Left            =   7440
      TabIndex        =   23
      Top             =   4080
      Width           =   972
   End
   Begin VB.Label Label7 
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
      Left            =   1920
      TabIndex        =   22
      Top             =   4320
      Width           =   612
   End
   Begin VB.Label lblCategory 
      Caption         =   "Label7"
      Height          =   372
      Left            =   7560
      TabIndex        =   21
      Top             =   3360
      Width           =   972
   End
   Begin VB.Label lbltitle 
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
      Left            =   2760
      TabIndex        =   20
      Top             =   4320
      Width           =   4092
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1092
   End
   Begin VB.Label lblBookID 
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
      Left            =   1200
      TabIndex        =   18
      Top             =   4320
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6840
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Return Details"
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
      Width           =   6852
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub fine()
Call connect
Set rs = Nothing
rs.Open "Select * from fine", conn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub deleteRecord()
Call connect
Set rs = Nothing
    rs.Open "Select * from ISSUED where bookid='" & lblBookID.Caption & "'", conn, adOpenDynamic, adLockOptimistic
    rs.Delete
End Sub
Public Sub Fine_()
Set rs = Nothing
    rs.Open "select * from fine", conn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub subtractinHand()
Call connect
Set recset = Nothing
    recset.Open "select * from member where member_ID = '" & (lblID.Caption) & "' ", conn, adOpenDynamic, adLockOptimistic
    recset.Fields(8) = val(recset.Fields(8)) - 1
    recset.Update
End Sub
Public Sub checkQuantity()
Dim rsSubtractBookQty As New ADODB.Recordset
Call connect
Set rsSubtractBookQty = Nothing
rsSubtractBookQty.Open "Select * From BOOK where BookID = '" & (lblBookID.Caption) & "'", conn, adOpenStatic, adLockOptimistic
    rsSubtractBookQty.Fields(11) = val(rsSubtractBookQty.Fields(11)) + 1
    rsSubtractBookQty.Fields(10) = val(rsSubtractBookQty.Fields(10)) - 1
    
    rsSubtractBookQty.Update
Exit Sub
End Sub
Private Sub cmd_cancel_Click()
unload Me
End Sub


Private Sub cmdReturn_Click()
Call connect
    Set rs = Nothing
    rs.Open "select * from RETURNED", conn, adOpenDynamic, adLockOptimistic
    With rs
    .AddNew
    .Fields(0) = lblBookID.Caption
    .Fields(1) = lbltitle.Caption
    .Fields(2) = lblCategory.Caption
    .Fields(3) = lblID.Caption
    .Fields(4) = lblLastname.Caption
    .Fields(5) = lblFirstname.Caption
    .Fields(6) = lblIssue.Caption
    .Fields(7) = lblDue.Caption
    .Fields(8) = lblReturn.Caption
    .Fields(9) = lblAuthor.Caption
    .Update
    End With
    ListIssue.ListItems.Remove (ListIssue.SelectedItem.Index)
    ' CDate(lblReturn.Caption) > CDate(lblDue.Caption)
    MsgBox "Book has been returned.", vbInformation
    If Text1.Text > 0 Then
        MsgBox "Print Penalty Receipt", vbInformation
        Set rs = Nothing
        rs.Open "select * from ISSUED where bookid = '" & lblBookID.Caption & "' and memberid = '" & lblID.Caption & "'", conn, adOpenDynamic, adLockOptimistic
        While rs.EOF = False
        Set rptReceipt.DataSource = rs
            rptReceipt.Sections("Section1").Controls.Item("lblMemberID").Caption = lblID.Caption
            rptReceipt.Sections("Section1").Controls.Item("lblbookID").Caption = lblBookID.Caption
            rptReceipt.Sections("Section1").Controls.Item("lblTitle").Caption = lbltitle.Caption
            rs.MoveNext
        Wend
            Call Fine_
            rptReceipt.Sections("Section1").Controls.Item("lblDays").Caption = Text1.Text
            rptReceipt.Sections("Section1").Controls.Item("lblAmount").Caption = Text1.Text * rs.Fields(0)
            rptReceipt.show 1
        
    End If
    Call subtractinHand
    Call checkQuantity
    Call deleteRecord
   
End Sub

Private Sub ListIssue_Click()
On Error Resume Next
lblBookID.Caption = ListIssue.SelectedItem
Call connect
Set rs = Nothing
rs.Open "Select * from ISSUED where bookid = '" & (lblBookID.Caption) & "'", conn, adOpenDynamic, adLockOptimistic
lbltitle.Caption = rs.Fields(1)
lblCategory.Caption = rs.Fields(2)
lblIssue.Caption = rs.Fields(6)
lblDue.Caption = rs.Fields(7)
lblAuthor.Caption = rs.Fields(9)
Text1.Text = CDate(lblReturn.Caption) - CDate(lblDue.Caption)

End Sub

