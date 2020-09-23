VERSION 5.00
Begin VB.Form frmMemberlist 
   BackColor       =   &H8000000D&
   Caption         =   "List of Members"
   ClientHeight    =   4344
   ClientLeft      =   5436
   ClientTop       =   2868
   ClientWidth     =   2928
   Icon            =   "frmMemberlist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4344
   ScaleWidth      =   2928
   Begin VB.CommandButton cmdShow 
      Height          =   492
      Left            =   1080
      Picture         =   "frmMemberlist.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   612
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   2172
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2052
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Member ID"
         Height          =   372
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   1212
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Year/Grade       Level"
         Height          =   372
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   1212
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "All"
         Height          =   372
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   972
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
         TabIndex        =   6
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.TextBox txtID 
      Height          =   288
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.ComboBox cboLevel 
      DataField       =   "Level"
      DataSource      =   "Adodc1"
      Height          =   288
      ItemData        =   "frmMemberlist.frx":0AE6
      Left            =   1200
      List            =   "frmMemberlist.frx":0AF0
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   2172
      Left            =   480
      TabIndex        =   11
      Top             =   0
      Width           =   2052
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
      Left            =   720
      TabIndex        =   10
      Top             =   3960
      Width           =   1332
   End
   Begin VB.Label Label1 
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
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
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
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   852
   End
End
Attribute VB_Name = "frmMemberlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset


Private Function level(val As Boolean)
Label6.Visible = val
cboLevel.Visible = val
End Function

Private Sub cboLevel_Click()
'If cboLevel = "Elementary" Then
'    cboYear.AddItem "1"
'    cboYear.AddItem "2"
'    cboYear.AddItem "3"
'    cboYear.AddItem "4"
'    cboYear.AddItem "5"
'    cboYear.AddItem "6"
'ElseIf cboLevel = "Secondary" Then
'    cboYear.Clear
'    cboYear.AddItem "1"
'    cboYear.AddItem "2"
'    cboYear.AddItem "3"
'    cboYear.AddItem "4"
'End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdShow_Click()
Call connect
If Option1.Value = True Then
        
    If txtID.Text = "" Then
    
        MsgBox "Please enter a Member ID", vbInformation
        txtID.SetFocus
    Else
       Set rs1 = Nothing
       rs1.Open "select * from MEMBER where Member_ID='" & txtID.Text & "'", conn, adOpenStatic, adLockOptimistic
       Set rptMember.DataSource = rs1
       rptMember.show 1
       txtID.Text = ""
    End If
   Exit Sub
End If
        
If Option2.Value = True Then
    If cboLevel.Text = "" Then
    MsgBox "Please select a Year Level", vbInformation
    Else
    Set rs2 = Nothing
       rs2.Open "select * from MEMBER where Level='" & cboLevel.Text & "'", conn, adOpenStatic, adLockOptimistic
    Set rptMember.DataSource = rs2
    rptMember.show 1
    End If
 Exit Sub
 End If
    
If Option3.Value = True Then
Set rs3 = Nothing
       rs3.Open "select * from MEMBER", conn, adOpenStatic, adLockOptimistic
       Set rptMember.DataSource = rs3
       rptMember.show 1
       Exit Sub
End If


    


End Sub

Private Sub Option1_Click()
txtID.Visible = True
Label1.Visible = True
Call level(False)
txtID.SetFocus
End Sub

Private Sub Option2_Click()
Call level(True)
txtID.Visible = False
Label1.Visible = False
End Sub

Private Sub Option3_Click()
Call level(False)
txtID.Visible = False
Label1.Visible = False

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii >= 48 And KeyAscii <= 57 Then
ElseIf KeyAscii = 13 Then
    cmdShow_Click
Else
    KeyAscii = 0
End If

End Sub
