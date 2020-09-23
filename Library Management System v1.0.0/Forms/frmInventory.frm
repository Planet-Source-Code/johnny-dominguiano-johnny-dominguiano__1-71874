VERSION 5.00
Begin VB.Form frmSummary 
   BackColor       =   &H8000000D&
   Caption         =   "Library Summary"
   ClientHeight    =   4890
   ClientLeft      =   3210
   ClientTop       =   2850
   ClientWidth     =   3975
   Icon            =   "frmInventory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   3975
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   3612
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3732
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   972
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   972
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   972
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   972
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   972
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3120
         Width           =   972
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   372
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total #of Books:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2412
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total types of Books:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2532
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total # of Borrowed:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2532
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total # of Remaining:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   2412
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total # of Members:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   2292
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Authors:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   2412
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Categories:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   2412
      End
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
      Caption         =   "Summary"
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
      TabIndex        =   15
      Top             =   0
      Width           =   3972
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim books As New ADODB.Recordset
Dim category As New ADODB.Recordset
Dim issued As New ADODB.Recordset
Dim members As New ADODB.Recordset
Dim authors As New ADODB.Recordset

Public Sub showdata()

If books.Fields(0) <> 0 Then
Text1.Text = books.Fields(0)
Text2.Text = books.Fields(3)
Text3.Text = books.Fields(1)
Text4.Text = books.Fields(2)
Else
Text1.Text = 0
Text2.Text = 0
Text3.Text = 0
Text4.Text = 0
End If

If authors.Fields(0) <> 0 Then
Text5.Text = authors.Fields(0)
Else
Text5.Text = 0
End If

If category.Fields(0) <> 0 Then
Text6.Text = category.Fields(0)
Else
Text6.Text = 0
End If

If members.Fields(0) <> 0 Then
Text7.Text = members.Fields(0)
Else
Text7.Text = 0
End If

End Sub


Public Sub records()
Call connect
Set books = Nothing
books.Open "Select sum(total_copy),sum(Borrowed),sum(remaining),count(*) from book", conn, adOpenDynamic, adLockOptimistic

Set category = Nothing
category.Open "Select count(*) from category", conn, adOpenDynamic, adLockOptimistic

Set members = Nothing
members.Open "Select count(*) from member", conn, adOpenDynamic, adLockOptimistic

Set authors = Nothing
authors.Open "Select count(*) from author", conn, adOpenDynamic, adLockOptimistic




End Sub

Private Sub Form_Load()
Call records
Call showdata

End Sub

