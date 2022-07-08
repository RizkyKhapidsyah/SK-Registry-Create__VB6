VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Thingy (Registry Create)"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "There Was No Command Line"
      Top             =   360
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text Box (Shows Command Line)"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rich Text Box"
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
rtfCommand RichTextBox1
txtCommand Text1
MakeDefault ".tty", "Text Thingy", "Text Thingy Document"
End Sub

