VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9480
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   240
         Top             =   720
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "GAME OVER"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   40.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   2160
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PLAYER 2 WINS"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5040
         TabIndex        =   4
         Top             =   5280
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PLAYER 1 WINS"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   5280
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   40.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   6480
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0 "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   40.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1800
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Ball 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   14  'Copy Pen
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4440
         Shape           =   1  'Square
         Top             =   2640
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         Height          =   5895
         Left            =   4560
         Top             =   0
         Width           =   15
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   8640
         Top             =   2160
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   240
         Top             =   2160
         Width           =   135
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   9360
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   9360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9360
      X2              =   9360
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Menu mnutest 
      Caption         =   "Test"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Randomx As Integer
Public Randomy As Integer
Public Randomall As Integer
Public Randomnum As Integer

Public Score1 As Integer
Public Score2 As Integer

Public Gameon As Boolean

Public xway As Boolean
Public yway As Boolean


Public Framex As Integer
Public Framey As Integer
Public Framex1 As Integer
Public Framey1 As Integer


Private Sub buttonpress_Timer()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
    
    If Shape1.Top < Frame1.Top + 10 Then Exit Sub Else
    
        Shape1.Top = Shape1.Top - 280
    
    End If


If KeyCode = vbKeyDown Then

   If Shape1.Top > Frame1.Height - Shape1.Width - 830 Then Exit Sub Else
   
        Shape1.Top = Shape1.Top + 280
        
   End If
   

Shape2.Top = Shape1.Top

End Sub

Private Sub Form_Load()
KeyPreview = True

Score1 = 0
Score2 = 0

Label1.Caption = Score1
Label2.Caption = Score2


End Sub

Private Sub testplay_Click()


End Sub


Private Sub mnutest_Click()



Randomnum = 300

xway = False
yway = False


   Randomnum = Randomnum + 10
       
    Randomize Timer
    Randomy = Int((10 ^ Rnd) + 1)
    
    Randomize Timer
    Randomx = Int((10 ^ Rnd) + 1)
    
    Randomall = Randomx + Randomy
    
   Do While Randomall < Randomnum
           
        If Randomx + Randomy < Randomnum Then Randomx = Randomx + 1
        If Randomx + Randomy < Randomnum Then Randomy = Randomy + 1
        Randomall = Randomx + Randomy
        
    Loop

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
  

'test to see if permimeter has been reached by ball

If Ball.Left < Shape1.Left + Shape1.Width Then
    If Ball.Top > Shape1.Top Then
    If Ball.Top + Ball.Height < Shape1.Top + Shape1.Height Then
        xway = True
    
   Randomnum = Randomnum + 10
       
    Randomize Timer
    Randomy = Int((10 ^ Rnd) + 1)
    
    Randomize Timer
    Randomx = Int((10 ^ Rnd) + 1)
    
    Randomall = Randomx + Randomy
    
   Do While Randomall < Randomnum
           
        If Randomx + Randomy < Randomnum Then Randomx = Randomx + 1
        If Randomx + Randomy < Randomnum Then Randomy = Randomy + 1
        Randomall = Randomx + Randomy
        
    Loop
            

    
    End If
End If
End If

If Ball.Left + Ball.Width > Shape2.Left Then
    If Ball.Top > Shape2.Top Then
    If Ball.Top + Ball.Height < Shape2.Top + Shape2.Height Then
        xway = False

  
   Randomnum = Randomnum + 10
       
    Randomize Timer
    Randomy = Int((10 ^ Rnd) + 1)
    
    Randomize Timer
    Randomx = Int((10 ^ Rnd) + 1)
    
    Randomall = Randomx + Randomy
    
   Do While Randomall < Randomnum
           
        If Randomx + Randomy < Randomnum Then Randomx = Randomx + 1
        If Randomx + Randomy < Randomnum Then Randomy = Randomy + 1
        Randomall = Randomx + Randomy
        
    Loop
            

    End If
End If
End If



If Ball.Top < 70 Then yway = True
If Ball.Left < Frame1.Left - 50 Then

xway = True
Score1 = Score1 + 1
Label2.Caption = Score1

End If



If Ball.Top > Frame1.Height - Ball.Height - 50 Then
    Ball.Top = Frame1.Height - Ball.Height - 50
        yway = False
        
        
        
End If


If Ball.Left > Frame1.Width - Ball.Width - 50 Then
    xway = False
    Ball.Left = Frame1.Width - Ball.Width - 50
    Score2 = Score2 + 1
    Label1.Caption = Score2
      
    End If

If xway = True Then Ball.Left = Ball.Left + Randomx
If xway = False Then Ball.Left = Ball.Left - Randomx
If yway = True Then Ball.Top = Ball.Top + Randomy
If yway = False Then Ball.Top = Ball.Top - Randomy

If Score1 = 10 Then
    Label4.Visible = True
    Label5.Visible = True
    Timer1.Enabled = False
    Ball.Left = (Frame1.Width / 2 - (Ball.Width / 2))
    Ball.Top = (Frame1.Height / 2 - (Ball.Height / 2))
End If
   
If Score2 = 10 Then
    Label5.Visible = True
    Label3.Visible = True
    Timer1.Enabled = False
    Ball.Left = (Frame1.Width / 2 - (Ball.Width / 2))
    Ball.Top = (Frame1.Height / 2 - (Ball.Height / 2))
End If
   
End Sub


