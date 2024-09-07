VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "No Switch Found..."
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select...  "
      Height          =   780
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3390
      Begin VB.CommandButton Command2 
         Caption         =   "Options"
         Height          =   420
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1545
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Screen Saver"
         Height          =   420
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
BlankForm.Show

End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me

End Sub
