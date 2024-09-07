VERSION 5.00
Object = "{AAC8DFAF-8A34-11D3-B327-000021C5C8A9}#1.0#0"; "SysTray.ocx"
Begin VB.Form Loader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Superterran QuicNote"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin SysTrayCtl.cSysTray cSysTray1 
         Left            =   2880
         Top             =   3720
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   -1  'True
         TrayIcon        =   "Loader.frx":0000
         TrayTip         =   "VB 5 - SysTray Control."
      End
      Begin VB.PictureBox Picture1 
         Height          =   2295
         Left            =   120
         Picture         =   "Loader.frx":01DA
         ScaleHeight     =   2235
         ScaleWidth      =   2955
         TabIndex        =   5
         Top             =   600
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "superterran.com"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"Loader.frx":161AC
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " QuicNote 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -840
         TabIndex        =   1
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.Menu mnuimenu 
      Caption         =   "mnuIMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRelocate 
         Caption         =   "Relocate"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About..."
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit "
      End
   End
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
On Error Resume Next



If Button = 1 Then Form1.Show
If Button = 2 Then Form1.PopupMenu mnuimenu



End Sub

Private Sub cSysTray2_MouseMove(Id As Long)

End Sub

Private Sub Label3_Click()

Loader.Visible = False


End Sub

Private Sub mnuabout_Click()
Loader.Visible = True

End Sub

Private Sub MnuExit_Click()
End

End Sub

Private Sub mnuFont_Click()

Dialog.ShowFont
End Sub

Private Sub mnuRelocate_Click()

Form2.Show

End Sub

