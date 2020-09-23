VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBooklist 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Books"
   ClientHeight    =   4416
   ClientLeft      =   4848
   ClientTop       =   3084
   ClientWidth     =   2988
   Icon            =   "frmBooklist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   2988
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmBooklist.frx":058A
      DataMember      =   "cmdCategory"
      DataSource      =   "DE"
      Height          =   288
      Left            =   720
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   508
      _Version        =   393216
      ListField       =   "Category_name"
      Text            =   ""
      Object.DataMember      =   "cmdCategory"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   2532
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2052
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Author"
         Height          =   372
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   972
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "All"
         Height          =   372
         Left            =   480
         TabIndex        =   0
         Top             =   1680
         Width           =   972
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Category"
         Height          =   372
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.CommandButton cmdShow 
      Height          =   492
      Left            =   1200
      Picture         =   "frmBooklist.frx":05A7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   612
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3132
      Left            =   3360
      TabIndex        =   8
      Top             =   240
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
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Height          =   3372
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   2892
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   2532
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   2052
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   4440
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
      Left            =   3360
      TabIndex        =   10
      Top             =   3600
      Width           =   852
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
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
      Left            =   4200
      TabIndex        =   9
      Top             =   3600
      Width           =   2172
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
      Left            =   840
      TabIndex        =   5
      Top             =   4080
      Width           =   1332
   End
End
Attribute VB_Name = "frmBooklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

Private Sub cmdShow_Click()
Call connect
If Option1.Value = True Then
     
    If DataCombo1.Text = "" Then
        MsgBox "Please select a Category", vbInformation
    Else
     
        Set rs1 = Nothing
        rs1.Open "Select * from book where category_name='" & DataCombo1.Text & "'", conn, adOpenStatic, adLockOptimistic
        Set rptBooks.DataSource = rs1
        rptBooks.show 1
    End If
Exit Sub
End If

If Option2.Value = True Then
    
        Set rs2 = Nothing
        rs2.Open "Select * from book where author_name='" & lblAuthor.Caption & "'", conn, adOpenStatic, adLockOptimistic
        Set rptBooks.DataSource = rs2
        rptBooks.show 1
Exit Sub
End If
If Option3.Value = True Then
      
        Set rs3 = Nothing
        rs3.Open "Select * from book", conn, adOpenStatic, adLockOptimistic
        Set rptBooks.DataSource = rs3
        rptBooks.show 1
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

End Sub

Private Sub ListView1_Click()
lblAuthor.Caption = ListView1.SelectedItem
End Sub

Private Sub Option1_Click()
frmBooklist.Width = 3060
DataCombo1.Visible = True
End Sub

Private Sub Option2_Click()
frmBooklist.Width = 6516
DataCombo1.Visible = False

End Sub

Private Sub Option3_Click()
frmBooklist.Width = 3060
DataCombo1.Visible = False

End Sub
