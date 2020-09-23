VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBooks 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOOKS"
   ClientHeight    =   7512
   ClientLeft      =   36
   ClientTop       =   1368
   ClientWidth     =   13176
   Icon            =   "frmBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7512
   ScaleWidth      =   13176
   Begin VB.CommandButton cmdAuthor 
      Caption         =   "Go to Authors"
      Height          =   372
      Left            =   10320
      TabIndex        =   31
      Top             =   3480
      Width           =   1332
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Commands"
      Height          =   1572
      Left            =   10320
      TabIndex        =   30
      Top             =   1320
      Width           =   2772
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   372
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   372
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   972
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   372
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   372
         Left            =   1560
         TabIndex        =   12
         Top             =   960
         Width           =   972
      End
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Go to Add Category"
      Height          =   372
      Left            =   10560
      TabIndex        =   28
      Top             =   6960
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBooks.frx":058A
      Height          =   5532
      Left            =   240
      TabIndex        =   27
      Top             =   1200
      Width           =   4932
      _ExtentX        =   8700
      _ExtentY        =   9758
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483635
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "BookID"
         Caption         =   "BookID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Title"
         Caption         =   "Title"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Edition"
         Caption         =   "Edition"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Category_name"
         Caption         =   "Category_name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Author_name"
         Caption         =   "Author_name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Publisher"
         Caption         =   "Publisher"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "ISBN"
         Caption         =   "ISBN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Pages"
         Caption         =   "Pages"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Total_copy"
         Caption         =   "Total_copy"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Call_num"
         Caption         =   "Call_num"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Borrowed"
         Caption         =   "Borrowed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Remaining"
         Caption         =   "Remaining"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   828.284
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1944
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   708.095
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   792
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      DataField       =   "BookID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   7320
      TabIndex        =   26
      Top             =   840
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   312
      Left            =   10440
      Top             =   6480
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
      RecordSource    =   "CATEGORY"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   6000
      Top             =   6840
      Visible         =   0   'False
      Width           =   3732
      _ExtentX        =   6583
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
      RecordSource    =   "BOOK"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Height          =   1812
      Left            =   5280
      TabIndex        =   22
      Top             =   5040
      Width           =   4812
      Begin VB.TextBox txtCallnum 
         DataField       =   "Call_num"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   8
         Top             =   1320
         Width           =   2532
      End
      Begin VB.TextBox txtPages 
         DataField       =   "Pages"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   7
         Top             =   840
         Width           =   2532
      End
      Begin VB.TextBox txtCopies 
         DataField       =   "Total_copy"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   6
         Top             =   360
         Width           =   2532
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Call Number:"
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
         TabIndex        =   25
         Top             =   1320
         Width           =   1452
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Pages:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Copies:"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Height          =   1812
      Left            =   5280
      TabIndex        =   18
      Top             =   3120
      Width           =   4812
      Begin VB.TextBox txtISBN 
         DataField       =   "ISBN"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   2532
      End
      Begin VB.TextBox txtPublisher 
         DataField       =   "Publisher"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   2532
      End
      Begin VB.TextBox txtAuthor 
         DataField       =   "Author_name"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   2532
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN #:"
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
         TabIndex        =   21
         Top             =   1320
         Width           =   1452
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher:"
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
         TabIndex        =   20
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Author Name:"
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
         Top             =   360
         Width           =   1572
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   1692
      Left            =   5280
      TabIndex        =   14
      Top             =   1320
      Width           =   4812
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "Category_name"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   2532
         _ExtentX        =   4466
         _ExtentY        =   508
         _Version        =   393216
         ListField       =   ""
         Text            =   ""
      End
      Begin VB.TextBox txtEdition 
         DataField       =   "Edition"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   2532
      End
      Begin VB.TextBox txtTitle 
         DataField       =   "Title"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   2532
      End
      Begin VB.Label Label5 
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
         Height          =   372
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition:"
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
         TabIndex        =   16
         Top             =   720
         Width           =   1452
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Label12"
      Height          =   1692
      Left            =   10440
      TabIndex        =   33
      Top             =   1200
      Width           =   2772
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   5412
      Left            =   5400
      TabIndex        =   32
      Top             =   1200
      Width           =   4812
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID #:"
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
      Left            =   5880
      TabIndex        =   29
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Book Details"
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
      TabIndex        =   13
      Top             =   0
      Width           =   13212
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function check() As Boolean
Dim status As Boolean
status = False
    If txtTitle.Text = "" Then
        MsgBox "Please enter the Book Title.", vbInformation, "Information Required"
    ElseIf txtEdition.Text = "" Then
        MsgBox "Please enter the Edition.", vbInformation, "Information Required"
   ElseIf DataCombo1.Text = "" Then
        MsgBox "Select a Category.", vbInformation, "Information Required"
   ElseIf txtAuthor.Text = "" Then
        MsgBox "Please enter the author name.", vbInformation, "Information Required"
   ElseIf txtPublisher.Text = "" Then
        MsgBox "Please enter the publisher.", vbInformation, "Information Required"
   ElseIf txtISBN.Text = "" Then
        MsgBox "Please enter the ISBN number.", vbInformation, "Information Required"
   ElseIf txtCopies.Text = "" Then
        MsgBox "Please enter the quantity.", vbInformation, "Information Required"
   ElseIf Not IsNumeric(txtCopies.Text) Then
        MsgBox "Quantity must be in numeric form.", vbInformation, "Numeric"
   ElseIf txtPages.Text = "" Then
        MsgBox "Please enter the number of pages.", vbInformation, "Information Required"
   ElseIf Not IsNumeric(txtPages.Text) Then
        MsgBox "Pages must be in numeric form.", vbInformation, "Numeric"
   ElseIf txtCallnum.Text = "" Then
        MsgBox "Please enter the call number.", vbInformation, "Information Required"
   Else
   status = True
End If
check = status
End Function



Private Sub cmdAuthor_Click()
frmAuthor.show 1
End Sub

Private Sub cmdCancel_Click()
Adodc1.Refresh
cmdCancel.Enabled = False
Call setlock2(True)
unload Me
End Sub

Private Sub cmdClose_Click()
unload Me
End Sub

Private Sub cmdDelete_Click()
Dim del As Integer
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "The record is empty!", vbOKOnly, "Delete"
Else
del = MsgBox("Are you sure?", vbYesNo, "Confirm")
If del = vbYes Then
    Adodc1.Recordset.Delete adAffectCurrent
    MsgBox "Record has been Deleted!", vbOKOnly, "Delete"
    Adodc1.Refresh
    unload Me
End If
End If
End Sub

Private Sub cmdEdit_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "The record is empty!", vbOKOnly, "Delete"
Else
cmdAuthor.Enabled = True
Call setlock2(False)
Call setbutton2(False)
txtID.Locked = True
End If
End Sub

Private Sub cmdGoto_Click()
frmCategory.show 1
End Sub

Private Sub cmdNew_Click()

End Sub

Private Sub cmdSave_Click()
If check = True Then
Adodc1.Recordset.Fields(11) = Adodc1.Recordset.Fields(8)
Adodc1.Recordset.Update
MsgBox "Record has been successfully Saved!", vbOKOnly, "Books"
MsgBox "Insert book index?", vbYesNo, "Insert Index"
If MsgBox("Insert book index?", vbYesNo, "Insert Index") = vbYes Then
frmIndex.show 1




Else
Adodc1.Refresh
Call setlock2(True)
Call setbutton2(True)
unload Me
End If
End Sub

Private Sub Form_Load()
DataCombo1.Refresh
Call setlock2(True)
Call setbutton2(True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Refresh
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub txtPages_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then

Else
    KeyAscii = 0
End If
End Sub
