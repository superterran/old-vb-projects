VERSION 5.00
Begin VB.Form Wizard2 
   Caption         =   "Windows XP Professional Activation Control Wizard"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   -120
      Picture         =   "Wizard2.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Professional Activation Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   930
         Width           =   6435
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   5520
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   5880
      Picture         =   "Wizard2.frx":261A
      Top             =   5520
      Width           =   360
   End
   Begin VB.Label Label3 
      Caption         =   $"Wizard2.frx":2D1C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   405
      TabIndex        =   3
      Top             =   2760
      Width           =   7155
   End
   Begin VB.Label Label2 
      Caption         =   "Change your Windows XP Product ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   4770
   End
End
Attribute VB_Name = "Wizard2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()

wizard3.Show

Unload Wizard2

End Sub
