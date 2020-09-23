VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAuthor 
   BackColor       =   &H8000000D&
   Caption         =   "AUTHOR"
   ClientHeight    =   4752
   ClientLeft      =   4500
   ClientTop       =   2664
   ClientWidth     =   5508
   Icon            =   "frmAuthor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4752
   ScaleWidth      =   5508
   Begin MSComctlLib.ListView ListView1 
      Height          =   2532
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   4466
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Author Name"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   1560
      Top             =   4440
      Visible         =   0   'False
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   550
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Name from AUTHOR order by Name"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   372
      Left            =   2040
      TabIndex        =   3
      Top             =   3960
      Width           =   1212
   End
   Begin VB.TextBox txtAuthor 
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   2772
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3360
      TabIndex        =   1
      Top             =   3960
      Width           =   1212
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   372
      Left            =   720
      TabIndex        =   0
      Top             =   3960
      Width           =   1212
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   2772
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click the Author name below to select:"
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
      TabIndex        =   4
      Top             =   480
      Width           =   5172
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdAdd_Click()
txtAuthor.SetFocus
cmdAdd.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
Call connect
Set rs = Nothing
rs.Open "select * from AUTHOR", conn, adOpenDynamic, adLockOptimistic
If txtAuthor.Text <> "" Then
    rs.AddNew
    rs.Fields(1) = txtAuthor.Text
    rs.Update
    MsgBox "Author added with an ID of: " & rs.Fields(0), vbInformation
    unload Me
    Me.show 1

Else
MsgBox "Please input a valid name", vbInformation

End If


End Sub

Private Sub Command2_Click()
On Error Resume Next
unload Me
End Sub

Private Sub DataGrid1_DblClick()

End Sub

Private Sub Form_Load()
Call connect
Set rs = Nothing
rs.Open "select author_name from AUTHOR order by author_name", conn, adOpenDynamic, adLockOptimistic
While rs.EOF = False
ListView1.ListItems.Add , , rs!author_name
rs.MoveNext
Wend
cmdSave.Enabled = False
End Sub

Private Sub ListView1_DblClick()
frmBooks.txtAuthor.Text = ListView1.SelectedItem
unload Me

End Sub

