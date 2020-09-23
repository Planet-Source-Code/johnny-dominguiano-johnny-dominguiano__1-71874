VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCategory 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CATEGORY"
   ClientHeight    =   4224
   ClientLeft      =   2544
   ClientTop       =   2616
   ClientWidth     =   7812
   Icon            =   "frmCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4224
   ScaleWidth      =   7812
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   1692
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   5292
      Begin VB.TextBox txtName 
         DataField       =   "Category_name"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2532
      End
      Begin VB.TextBox txtDesc 
         DataField       =   "Description"
         DataSource      =   "Adodc1"
         Height          =   732
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   2532
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name:"
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
         Top             =   240
         Width           =   1932
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   372
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   1212
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   2280
      Top             =   3720
      Visible         =   0   'False
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   550
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
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
      RecordSource    =   "CATEGORY"
      Caption         =   "Navigator"
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4200
      TabIndex        =   2
      Top             =   3000
      Width           =   1212
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   3000
      Width           =   1212
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   1692
      Left            =   1080
      TabIndex        =   9
      Top             =   960
      Width           =   5292
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Category Details"
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
      Width           =   7812
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew
txtName.SetFocus
txtName.Text = ""
txtDesc.Text = ""
txtName.Locked = False
txtDesc.Locked = False

End Sub

Private Sub cmdCancel_Click()
unload Me
End Sub

Private Sub cmdDelete_Click()
Adodc1.Recordset.Delete adAffectCurrent
MsgBox "Category Deleted!", vbOKOnly, "Category"
Adodc1.Refresh
End Sub

Private Sub cmdSave_Click()
On Error GoTo a
Adodc1.Recordset.Update
MsgBox "Category has been successfully Added!", vbOKOnly, "Category"
frmBooks.DataCombo1.Refresh
unload Me
Adodc1.Refresh
Exit Sub
a:
If Err.Number = -2147467259 Then
MsgBox ("Category already exist,please enter another category."), vbCritical, "Category exist"
frmBooks.DataCombo1.Refresh
txtName.SetFocus
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub

