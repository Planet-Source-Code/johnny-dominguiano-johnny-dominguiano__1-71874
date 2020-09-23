VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListOfBooks 
   BackColor       =   &H8000000D&
   Caption         =   "Book Information"
   ClientHeight    =   6936
   ClientLeft      =   4452
   ClientTop       =   1656
   ClientWidth     =   4080
   Icon            =   "frmListOfBooks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6936
   ScaleWidth      =   4080
   Begin VB.TextBox txtID 
      Height          =   288
      Left            =   1920
      TabIndex        =   17
      Top             =   3720
      Width           =   1572
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   252
      Left            =   1440
      TabIndex        =   14
      Top             =   3720
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   445
      _Version        =   393216
      CustomFormat    =   "mm/dd/yyyy"
      Format          =   65798145
      CurrentDate     =   39877
   End
   Begin VB.CommandButton cmdShow 
      Height          =   492
      Left            =   1680
      Picture         =   "frmListOfBooks.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   612
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H0080FFFF&
      Caption         =   "All"
      Height          =   492
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   972
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0080FFFF&
      Caption         =   "By Member ID"
      Height          =   492
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   1332
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080FFFF&
      Caption         =   "By Author"
      Height          =   492
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   972
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FFFF&
      Caption         =   "By Category"
      Height          =   492
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   1212
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080FFFF&
      Caption         =   "By Date"
      Height          =   492
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   972
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Options:"
      Height          =   2292
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   3252
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmListOfBooks.frx":0866
      DataMember      =   "cmdCategory"
      DataSource      =   "DE"
      Height          =   288
      Left            =   1920
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   508
      _Version        =   393216
      ListField       =   "Category_name"
      Text            =   ""
      Object.DataMember      =   "cmdCategory"
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3132
      Left            =   4560
      TabIndex        =   8
      Top             =   960
      Width           =   2652
      _ExtentX        =   4678
      _ExtentY        =   5525
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Author Name"
         Object.Width           =   4464
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Height          =   3372
      Left            =   4440
      TabIndex        =   21
      Top             =   840
      Width           =   2892
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   2292
      Left            =   480
      TabIndex        =   20
      Top             =   1200
      Width           =   3252
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   4080
      Y1              =   720
      Y2              =   6960
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
      Height          =   732
      Left            =   -360
      TabIndex        =   19
      Top             =   0
      Width           =   5292
   End
   Begin VB.Label qwe 
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
      Height          =   732
      Left            =   4920
      TabIndex        =   18
      Top             =   0
      Width           =   5292
   End
   Begin VB.Label Label5 
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
      Left            =   480
      TabIndex        =   16
      Top             =   3720
      Width           =   1452
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "List of Books:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   480
      TabIndex        =   13
      Top             =   3720
      Width           =   732
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      TabIndex        =   12
      Top             =   4800
      Width           =   1332
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
      Height          =   372
      Left            =   4560
      TabIndex        =   11
      Top             =   4680
      Width           =   2652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   4440
      TabIndex        =   10
      Top             =   4320
      Width           =   852
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Width           =   1212
   End
End
Attribute VB_Name = "frmListOfBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset

Private Sub showLabel(val As Boolean)
frmListOfBooks.Width = 4176
Label9.Visible = val
DataCombo1.Visible = val
Label3.Visible = val
DTPicker1.Visible = val
Label5.Visible = val
txtID.Visible = val
End Sub

Public Sub showReport()
If Banner.Caption = "ISSUED" Then
    rptBorrowed.show 1
    Exit Sub
End If

If Banner.Caption = "RETURNED" Then
    rptReturned.show 1
    Exit Sub
End If

End Sub

Private Sub cmdShow_Click()

If Option1.Value = True Then

    Call connect
        Set rs1 = Nothing
        rs1.Open "Select * from " & (Banner.Caption) & " where Date_borrowed ='" & (DTPicker1.Value) & "'", conn, adOpenDynamic, adLockOptimistic
        Set rptBorrowed.DataSource = rs1
        Set rptReturned.DataSource = rs1
        showReport
Exit Sub
End If

If Option2.Value = True Then
If DataCombo1.Text = "" Then
        MsgBox "Please select a Category", vbInformation
    Else
    Call connect
        Set rs2 = Nothing
        rs2.Open "Select * from " & (Banner.Caption) & " where category = '" & (DataCombo1.Text) & "'", conn, adOpenDynamic, adLockOptimistic
        Set rptBorrowed.DataSource = rs2
        Set rptReturned.DataSource = rs2
        showReport
    End If
Exit Sub
End If

If Option3.Value = True Then
        Call connect
        Set rs3 = Nothing
        rs3.Open "Select * from " & (Banner.Caption) & " where author_name = '" & lblAuthor.Caption & "'", conn, adOpenDynamic, adLockOptimistic
        Set rptBorrowed.DataSource = rs3
        Set rptReturned.DataSource = rs3
        showReport
Exit Sub
End If

If Option4.Value = True Then
        Call connect
        Set rs4 = Nothing
        rs4.Open "Select * from " & (Banner.Caption) & " where memberid = '" & (txtID.Text) & "'", conn, adOpenDynamic, adLockOptimistic
        Set rptBorrowed.DataSource = rs4
        Set rptReturned.DataSource = rs4
        showReport
Exit Sub
End If

If Option5.Value = True Then
        Call connect
        Set rs5 = Nothing
        rs5.Open "Select * from " & (Banner.Caption) & "", conn, adOpenDynamic, adLockOptimistic
        Set rptBorrowed.DataSource = rs5
        Set rptReturned.DataSource = rs5
        showReport
Exit Sub
End If



End Sub

Private Sub Form_Load()
Call connect
Set rs = Nothing
rs.Open "select Author_name from AUTHOR order by author_name", conn, adOpenDynamic, adLockOptimistic
While rs.EOF = False
ListView1.ListItems.Add , , rs!author_name
rs.MoveNext
Wend

DTPicker1.MaxDate = Date
End Sub

Private Sub ListView1_Click()
lblAuthor.Caption = ListView1.SelectedItem
End Sub

Private Sub Option1_Click()
Call showLabel(False)
Label3.Visible = True
DTPicker1.Visible = True
End Sub

Private Sub Option2_Click()
Call showLabel(False)
Label9.Visible = True
DataCombo1.Visible = True
End Sub

Private Sub Option3_Click()
Call showLabel(False)
frmListOfBooks.Width = 7500
End Sub

Private Sub Option4_Click()
Call showLabel(False)
Label5.Visible = True
txtID.Visible = True
txtID.SetFocus
End Sub

Private Sub Option5_Click()
Call showLabel(False)
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 13 Then
    cmdShow_Click
Else
    KeyAscii = 0
End If
End Sub
