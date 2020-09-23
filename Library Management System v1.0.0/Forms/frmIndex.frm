VERSION 5.00
Begin VB.Form frmIndex 
   Caption         =   "Form1"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5172
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      DataSource      =   "Adodc1"
      Height          =   288
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   2532
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
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Banner 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Book Index"
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
      Height          =   612
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5652
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
