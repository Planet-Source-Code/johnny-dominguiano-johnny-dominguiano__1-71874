VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H8000000D&
   Caption         =   "Search"
   ClientHeight    =   7620
   ClientLeft      =   3036
   ClientTop       =   1056
   ClientWidth     =   9972
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9972
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "frmSearch.frx":1642
      Left            =   3000
      List            =   "frmSearch.frx":165E
      TabIndex        =   4
      Top             =   360
      Width           =   2652
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   735
      Left            =   7320
      MouseIcon       =   "frmSearch.frx":16B1
      MousePointer    =   99  'Custom
      Picture         =   "frmSearch.frx":1803
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Search"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Height          =   288
      Left            =   3000
      TabIndex        =   1
      Top             =   804
      Width           =   4092
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5892
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   9612
      _ExtentX        =   16955
      _ExtentY        =   10393
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Book ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ISBN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Call Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   6132
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   9852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By:"
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
      TabIndex        =   5
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Type your search here:"
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
      TabIndex        =   2
      Top             =   840
      Width           =   2772
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSearch_Click()
If txtSearch = "" Then
    MsgBox "Type your query in the textbox!", vbCritical, "System Message"
    txtSearch.SetFocus
ElseIf Combo1 = "" Then
    MsgBox "Please select search by!", vbCritical, "System Message"

Else
    Dim X
    Call connect
    Set recset = Nothing
    recset.Open "Select * from BOOK where " & (Combo1) & "  like '%" & (txtSearch.Text) & "%'", conn, adOpenDynamic, adLockOptimistic
    ListView1.ListItems.Clear
    While recset.EOF = False
    ListView1.ListItems.Add , , recset!bookid
    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , recset!Title
    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , recset!Category_name
    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , recset!author_name
    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , recset!ISBN
    ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , recset!Call_num
   recset.MoveNext
   X = 1
   Wend
    MsgBox "Search Complete!", vbInformation, "Confirmation"
    ListView1.Enabled = True
    If X <> 1 Then
    MsgBox "No results!", vbInformation, "Confirmation"
    txtSearch.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
Dim a As String
a = Combo1.Text
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSearch_Click
End If
End Sub
