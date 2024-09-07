VERSION 5.00
Begin VB.Form Form1 
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
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   -45
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
   Begin VB.Image Image4 
      Height          =   360
      Left            =   1650
      Picture         =   "Form1.frx":261A
      Top             =   4875
      Width           =   360
   End
   Begin VB.Label Label6 
      Caption         =   "About this application"
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
      Left            =   2085
      TabIndex        =   6
      Top             =   4905
      Width           =   4770
   End
   Begin VB.Label Label5 
      Caption         =   "What do you want to do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   465
      TabIndex        =   5
      Top             =   1800
      Width           =   5760
   End
   Begin VB.Label Label4 
      Caption         =   "View Windows XP Product ID's"
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
      Left            =   2085
      TabIndex        =   4
      Top             =   4290
      Width           =   4770
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   1650
      Picture         =   "Form1.frx":2D1C
      Top             =   4260
      Width           =   360
   End
   Begin VB.Label Label3 
      Caption         =   "De-Activate Windows XP"
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
      Left            =   2085
      TabIndex        =   3
      Top             =   3675
      Width           =   4770
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1650
      Picture         =   "Form1.frx":341E
      Top             =   3645
      Width           =   360
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
      Left            =   2070
      TabIndex        =   2
      Top             =   3060
      Width           =   4770
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   1650
      Picture         =   "Form1.frx":3B20
      Top             =   3030
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
