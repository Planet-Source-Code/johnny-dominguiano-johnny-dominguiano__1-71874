VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000009&
   Caption         =   "Library Management System"
   ClientHeight    =   9888
   ClientLeft      =   360
   ClientTop       =   588
   ClientWidth     =   13728
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   9888
   ScaleWidth      =   13728
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView tv 
      Height          =   6852
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   2892
      _ExtentX        =   5101
      _ExtentY        =   12086
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "/"
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imlIcon"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcon 
      Left            =   1800
      Top             =   720
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E3E6
            Key             =   "folders"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E980
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EF1A
            Key             =   "folderClose"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F4B4
            Key             =   "acrobat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FA4E
            Key             =   "access"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FFE8
            Key             =   "exel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10582
            Key             =   "word"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10B1C
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E36
            Key             =   "bitmap"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113D0
            Key             =   "database"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1196A
            Key             =   "user"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F04
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1249E
            Key             =   "warning"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12A38
            Key             =   "check"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12FD2
            Key             =   "group"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":139E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13EF7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   732
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   14052
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderWidth     =   3
      X1              =   0
      X2              =   13800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "International Christian School "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   732
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14052
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function initTreeview()
With tv
    .Nodes.Clear
    .Nodes.Add , , "Library", "Library Management", 12
    .Nodes.Add "Library", tvwChild, "Members", "Members", 2
    .Nodes.Add "Library", tvwChild, "Books", "Books", 2
    .Nodes.Add "Library", tvwChild, "Transactions", "Transactions", 2
    .Nodes.Add "Library", tvwChild, "Reports", "Reports", 2
    .Nodes.Add "Library", tvwChild, "Other", "Other", 2

    .Nodes.Add "Members", tvwChild, "MemberAdd", "Add", 14
    .Nodes.Add "Members", tvwChild, "MemberEdit", "Edit", 15
    .Nodes.Add "Members", tvwChild, "MemberDelete", "Delete", 13
 
    .Nodes.Add "Books", tvwChild, "BookAdd", "Add", 14
    .Nodes.Add "Books", tvwChild, "BookEdit", "Edit", 15
    .Nodes.Add "Books", tvwChild, "BookDelete", "Delete", 13

    
    .Nodes.Add "Transactions", tvwChild, "Issue", "Issue", 17
    .Nodes.Add "Transactions", tvwChild, "Return", "Return", 16
    
    .Nodes.Add "Reports", tvwChild, "ReportBook", "Books", 1
    .Nodes.Add "Reports", tvwChild, "ReportMember", "Members", 15
    .Nodes.Add "Reports", tvwChild, "ReportBorrowed", "Borrowed Books", 17
    .Nodes.Add "Reports", tvwChild, "ReportReturned", "Returned Books", 16
    
    .Nodes.Add "Other", tvwChild, "Search", "Search", 12
    .Nodes.Add "Other", tvwChild, "Summary", "Summary", 8
    .Nodes.Add "Other", tvwChild, "Fine", "Fine Setting", 10
    .Nodes.Add "Other", tvwChild, "Statistics", "Statistics", 1
    
    .Nodes.Add "Statistics", tvwChild, "FrequentBooks", "Frequently Borrowed", 14
    .Nodes.Add "Statistics", tvwChild, "FrequentBorrowers", "Frequent Borrowers", 15
   End With
End Function

Private Sub Form_Load()
initTreeview
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
Dim id As String
Dim bid As String
Call connect
Set rs = Nothing
rs.Open "select * from member", conn, adOpenDynamic, adLockOptimistic
If rs.RecordCount = 0 Then
id = "0001"
End If
rs.MoveLast
id = Format(val(Right(rs!Member_ID, 4)) + 1, "000#")
rs.Close

'Members======================================

If Node.Key = "MemberAdd" Then
Call setlock1(False)
Call setbutton1(False)
frmMembers.Adodc1.Recordset.AddNew
frmMembers.txtID.Text = id
frmMembers.DataGrid1.Enabled = False
frmMembers.show vbModal
End If

If Node.Key = "MemberEdit" Then
If frmMembers.Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "The record is empty!", vbOKOnly, "Delete"
Else
frmMembers.cmdDelete.Enabled = False
frmMembers.cmdSave.Enabled = False
frmMembers.show 1
End If
End If


If Node.Key = "MemberDelete" Then
If frmMembers.Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "The record is empty!", vbOKOnly, "Delete"
Else
frmMembers.cmdSave.Enabled = False
frmMembers.cmdEdit.Enabled = False
frmMembers.show 1

End If
End If


'Books====================================
If Node.Key = "BookAdd" Then
Call setlock2(False)
Call setbutton2(False)
Set rs = Nothing
rs.Open "select * from book", conn, adOpenDynamic, adLockOptimistic
If rs.RecordCount = 0 Then
bid = "0001"
End If
rs.MoveLast
bid = Format(val(Right(rs!bookid, 4)) + 1, "000#")
rs.Close

With frmBooks
Set .DataCombo1.RowSource = .Adodc2
    .DataCombo1.ListField = "Category_name"
    .Adodc1.Recordset.AddNew
    frmBooks.txtID.Text = bid
frmBooks.DataGrid1.Enabled = False

.show vbModal
End With
End If

If Node.Key = "BookEdit" Then
If frmBooks.Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "The record is empty!", vbOKOnly, "Delete"
Else
frmBooks.cmdDelete.Enabled = False
frmBooks.cmdCancel.Enabled = True
frmBooks.txtID.Locked = True
frmBooks.cmdAuthor.Enabled = False
frmBooks.show 1
End If
End If

If Node.Key = "BookDelete" Then
If frmBooks.Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "The record is empty!", vbOKOnly, "Delete"
Else
frmBooks.cmdCancel.Enabled = True
frmBooks.cmdEdit.Enabled = False
frmBooks.cmdAuthor.Enabled = False
frmBooks.show 1
End If
End If

'Transactions===========================================
If Node.Key = "Issue" Then
frmEnter.Banner.Caption = "Issue"
frmEnter.show 1
End If

If Node.Key = "Return" Then
frmEnter.Banner.Caption = "Return"
frmEnter.show 1
End If

If Node.Key = "Book" Then
frmSearch.show 1
End If

'Reports=================================================
If Node.Key = "ReportBook" Then
frmBooklist.show 1
End If

If Node.Key = "ReportMember" Then
frmMemberlist.show 1
End If

If Node.Key = "ReportBorrowed" Then
frmListOfBooks.Banner.Caption = "ISSUED"
frmListOfBooks.show 1
End If

If Node.Key = "ReportReturned" Then
frmListOfBooks.Banner.Caption = "RETURNED"
frmListOfBooks.show 1
End If

If Node.Key = "Search" Then
frmSearch.show 1
End If

If Node.Key = "Summary" Then
frmSummary.show 1
End If

If Node.Key = "Fine" Then
frmFine.show 1
End If

If Node.Key = "FrequentBooks" Then
frmStatBooks.show 1
End If

If Node.Key = "FrequentBorrowers" Then
frmStatMembers.show 1
End If


cancelbutton:
Exit Sub

End Sub
