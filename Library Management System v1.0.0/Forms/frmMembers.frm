VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMembers 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MEMBERS"
   ClientHeight    =   8580
   ClientLeft      =   3645
   ClientTop       =   1050
   ClientWidth     =   8355
   Icon            =   "frmMembers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8355
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMembers.frx":058A
      Height          =   2412
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   8052
      _ExtentX        =   14208
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   25
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Member_ID"
         Caption         =   "Member_ID"
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
         DataField       =   "First_name"
         Caption         =   "First_name"
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
         DataField       =   "Last_name"
         Caption         =   "Last_name"
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
         DataField       =   "MI"
         Caption         =   "MI"
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
         DataField       =   "Sex"
         Caption         =   "Sex"
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
         DataField       =   "Contact_num"
         Caption         =   "Contact_num"
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
         DataField       =   "Level"
         Caption         =   "Level"
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
         DataField       =   "Year"
         Caption         =   "Year"
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
         DataField       =   "Bookinhand"
         Caption         =   "Bookinhand"
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
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1019.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Commands"
      Height          =   1452
      Left            =   1320
      TabIndex        =   24
      Top             =   3720
      Width           =   2892
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   372
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   972
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   372
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   972
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   372
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   372
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   972
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   2160
      Top             =   5400
      Visible         =   0   'False
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   582
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
      RecordSource    =   "MEMBER"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      DataField       =   "Member_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   720
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Student's Information"
      Height          =   2292
      Left            =   1320
      TabIndex        =   14
      Top             =   1200
      Width           =   6132
      Begin VB.ComboBox cboLevel 
         DataField       =   "Level"
         DataSource      =   "Adodc1"
         Height          =   288
         ItemData        =   "frmMembers.frx":059F
         Left            =   1080
         List            =   "frmMembers.frx":05A9
         TabIndex        =   5
         Top             =   1680
         Width           =   1332
      End
      Begin VB.ComboBox cboYear 
         DataField       =   "Year"
         DataSource      =   "Adodc1"
         Height          =   288
         ItemData        =   "frmMembers.frx":05C4
         Left            =   3600
         List            =   "frmMembers.frx":05C6
         TabIndex        =   6
         Top             =   1680
         Width           =   732
      End
      Begin VB.TextBox txtLastname 
         DataField       =   "Last_name"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   2292
      End
      Begin VB.TextBox txtFirstname 
         DataField       =   "First_name"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   1692
      End
      Begin VB.TextBox txtMI 
         DataField       =   "MI"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   5280
         MaxLength       =   1
         TabIndex        =   2
         Top             =   360
         Width           =   612
      End
      Begin VB.ComboBox cboSex 
         DataField       =   "Sex"
         DataSource      =   "Adodc1"
         Height          =   288
         ItemData        =   "frmMembers.frx":05C8
         Left            =   1080
         List            =   "frmMembers.frx":05D2
         TabIndex        =   3
         Top             =   1080
         Width           =   1332
      End
      Begin VB.TextBox txtContactnum 
         DataField       =   "Contact_num"
         DataSource      =   "Adodc1"
         Height          =   288
         Left            =   4080
         TabIndex        =   4
         Top             =   1080
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1920
         TabIndex        =   21
         Top             =   720
         Width           =   612
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   4080
         TabIndex        =   20
         Top             =   720
         Width           =   732
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "MI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5400
         TabIndex        =   19
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   732
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   852
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   16
         Top             =   1680
         Width           =   732
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   15
         Top             =   1080
         Width           =   1332
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Height          =   1452
      Left            =   1440
      TabIndex        =   26
      Top             =   3600
      Width           =   2892
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   2292
      Left            =   1440
      TabIndex        =   25
      Top             =   1080
      Width           =   6132
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderWidth     =   3
      X1              =   0
      X2              =   8280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID #:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1440
      TabIndex        =   13
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Member Details"
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
      TabIndex        =   11
      Top             =   0
      Width           =   8412
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function check() As Boolean
Dim status As Boolean
status = False

    If txtLastname.Text = "" Then
        MsgBox "Please enter the last name.", vbInformation, "Information Required"
    ElseIf txtFirstname.Text = "" Then
        MsgBox "Please enter the first name.", vbInformation, "Information Required"
     ElseIf txtMI.Text = "" Then
        MsgBox "Please enter the middle initial.", vbInformation, "Information Required"
     ElseIf cboSex.Text = "" Then
        MsgBox "Please select a gender.", vbInformation, "Information Required"
    ElseIf txtContactnum.Text = "" Then
        MsgBox "Please enter a valid contact.", vbInformation, "Information Required"
    ElseIf Not IsNumeric(txtContactnum.Text) Then
        MsgBox "Please enter a valid contact", vbInformation, "Information Required"
    ElseIf cboLevel.Text = "" Then
        MsgBox "Please select a level.", vbInformation, "Information Required"
     ElseIf cboYear.Text = "" Then
        MsgBox "Please select a current year.", vbInformation, "Information Required"
     Else
    status = True
End If
check = status
End Function


Private Sub cboLevel_Click()
If cboLevel = "Elementary" Then
    cboYear.Clear
    cboYear.AddItem "1"
    cboYear.AddItem "2"
    cboYear.AddItem "3"
    cboYear.AddItem "4"
    cboYear.AddItem "5"
    cboYear.AddItem "6"
ElseIf cboLevel = "Secondary" Then
    cboYear.Clear
    cboYear.AddItem "1"
    cboYear.AddItem "2"
    cboYear.AddItem "3"
    cboYear.AddItem "4"
End If
End Sub


Private Sub cmdCancel_Click()
Adodc1.Refresh
unload Me
Call setlock1(True)
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

Private Sub cmdNew_Click()
Call setlock1(False)
Call setbutton1(False)

End Sub

Private Sub cmdEdit_Click()
cmdSave.Enabled = True
Call setlock1(False)
End Sub

Private Sub cmdSave_Click()
If check = True Then
Adodc1.Recordset.Update
MsgBox "Record has been successfully Saved!", vbOKOnly, "Member"
Adodc1.Refresh
Call setlock1(True)
Call setbutton1(True)
unload Me
End If
End Sub


Private Sub Form_Load()
Call setlock1(True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Refresh

End Sub

Private Sub txtContactnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 122 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 122 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub txtMI_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 122 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub cboLevel_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 122 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 122 Then

Else
    KeyAscii = 0
End If
End Sub


Private Sub cboYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then

Else
    KeyAscii = 0
End If
End Sub

