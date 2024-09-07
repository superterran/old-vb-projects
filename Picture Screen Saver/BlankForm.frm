VERSION 5.00
Begin VB.Form BlankForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   900
      Left            =   840
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   480
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2880
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "10:35:00 PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Superterran Picture Slide Show 1.0  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   600
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "BlankForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mx As Integer
Public my As Integer
Public FileRight As Boolean

Public apath As String
Public PathLoc As String



Private Sub Form_Load()

PathLoc = GetSetting("Superterran Software", "Slide Show Screen Saver", "Path")
Timer1.interval = GetSetting("Superterran Software", "Slide Show Screen Saver", "Interval")

Timer1.interval = Timer1.interval * 1000

apath = PathLoc

apath1 = Right(apath, 1)

If apath1 <> "\" Then apath = apath & "\"



Module1.HideMouse
'Module1.ShowMouse


With BlankForm
.Top = 0
.Left = 0
.Width = Screen.Width
.Height = Screen.Height

End With
File1.path = PathLoc
File1.ListIndex = 0

Dim Extention

starttest:

FileRight = False
Do Until FileRight = True

Extention = Right(File1.FileName, 3)

'MsgBox Extention

If Extention = "jpg" Then FileRight = True
If Extention = "bmp" Then FileRight = True
If Extention = "gif" Then FileRight = True

Loop

If FileRight = False Then
    On Error GoTo errorhandler:
    File1.ListIndex = File1.ListIndex + 1
    GoTo starttest
End If



'Image1.Picture = LoadPicture(apath & File1.FileName)
'Image1.Refresh

Image1.Top = 0
Image1.Left = 0
Image1.Width = Screen.Width
Image1.Height = Screen.Height
Label1.Top = Screen.Height - Label1.Height - 70
Label1.Left = Screen.Width - Label1.Width
Label2.Top = Screen.Height - Label2.Height - 70
Label2.Left = 0
Shape1.Width = Screen.Width
Shape1.Top = Screen.Height - Shape1.Height
Shape1.Left = 0


Exit Sub

errorhandler:
File1.ListIndex = 0
Return
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If mx = 0 Then
    mx = X
    my = y
    Exit Sub
End If

If mx = X Then
    Exit Sub
        Else
    Module1.ShowMouse
    Unload BlankForm

    End
End If

If my = y Then
    Exit Sub
        Else
    Module1.ShowMouse
    Unload BlankForm
    
    End
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If mx = 0 Then
    mx = X
    my = y
    Exit Sub
End If

If mx = X Then
    Exit Sub
        Else
    Module1.ShowMouse
    Unload BlankForm

    End
End If

If my = y Then
    Exit Sub
        Else
    Module1.ShowMouse
    Unload BlankForm
    
    End
End If


End Sub

Private Sub SystemMonitor1_OnCounterSelected(ByVal iIndex As Long)

End Sub

Private Sub Timer1_Timer()

Image1.Visible = False

Dim Extention

starttest:

FileRight = False
Do Until FileRight = True

Extention = Right(File1.FileName, 3)
Extention = LCase(Extention)

'MsgBox Extention '-there to see if the extention monitor works

If Extention = "jpg" Then FileRight = True
If Extention = "bmp" Then FileRight = True
If Extention = "gif" Then FileRight = True



    On Error GoTo errorhandler:

If FileRight = False Then

    File1.ListIndex = File1.ListIndex + 1
    GoTo starttest
End If

Loop

Dim shapetop, shapeh

shapetop = Shape1.Top
shapeh = Shape1.Height

Shape1.Top = 0
Shape1.Height = Screen.Height

Image1.Picture = LoadPicture(apath & File1.FileName)
Image1.Refresh



Shape1.Top = shapetop
Shape1.Height = shapeh


File1.ListIndex = File1.ListIndex + 1

Dim bob As Integer

bob = 0

Do Until bob = 20000

bob = bob + 1

Loop

Image1.Visible = True


Exit Sub

errorhandler:
File1.ListIndex = 0
If FileRight = True Then Exit Sub Else GoTo starttest
'Resume

End Sub

Private Sub Timer2_Timer()
Label2.Caption = "  " & Time & " - " & Date


End Sub
