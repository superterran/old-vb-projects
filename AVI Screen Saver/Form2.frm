VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "AVI Screen Saver Properties"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8730
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8493
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      FileName        =   "staviss.rtf"
      TextRTF         =   $"Form2.frx":0000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ok"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   8760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Superterran AVI Screen Saver 1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()
End

End Sub
