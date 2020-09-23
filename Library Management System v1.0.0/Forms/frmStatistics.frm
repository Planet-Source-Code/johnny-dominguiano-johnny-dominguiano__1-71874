VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatBooks 
   BackColor       =   &H8000000D&
   Caption         =   "Library Statistics"
   ClientHeight    =   7200
   ClientLeft      =   3390
   ClientTop       =   1785
   ClientWidth     =   7020
   Icon            =   "frmStatistics.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7020
   Begin MSComctlLib.ListView ListView1 
      Height          =   5292
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6732
      _ExtentX        =   11880
      _ExtentY        =   9340
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TItle"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Times Borrowed"
         Object.Width           =   2999
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderWidth     =   3
      X1              =   0
      X2              =   8280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Frequently Borrowed"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   25.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   732
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7212
   End
End
Attribute VB_Name = "frmStatBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call connect
Set rs = Nothing
rs.Open "select * from book where counter <> 0  order by counter desc", conn, adOpenDynamic, adLockOptimistic
While rs.EOF = False
ListView1.ListItems.Add , , rs!Title
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rs!author_name
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rs!Category_name
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rs!Counter
rs.MoveNext
Wend

End Sub
