VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   10635
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   ScaleHeight     =   7335
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Superterran's AVI Screen Saver 1.0"
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin WMPLibCtl.WindowsMediaPlayer Player 
      Height          =   5760
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   8700
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   -1  'True
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   15346
      _cy             =   10160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mx As Integer
Public AVIfilename As String
Public Position As Integer
Public Msense As Boolean
Public mute As Boolean
Public first As Boolean


Private Sub Form_Load()
mx = 0
Msense = True
mute = False
first = False




Select Case LCase(Left(Command, 2))
       
        Case "/p"
            End
           Exit Sub
       Case "/s"
            'Proceed
       Case Else
           'Show settings
 Form2.Show
 Unload Form1
 
            Exit Sub
End Select


On Error GoTo 5

AVIfilename = GetSetting("superterran", "avi screensaver", "filename")
Position = GetSetting("superterran", "avi screensaver", "playpos")
mute = GetSetting("superterran", "avi screensaver", "mute")

If AVIfilename = "" Then

5 Dialog.ShowOpen
AVIfilename = Dialog.FileName
SaveSetting "superterran", "avi screensaver", "filename", AVIfilename
SaveSetting "superterran", "avi screensaver", "playpos", Position
SaveSetting "superterran", "avi screensaver", "mute", mute
End If



'sets the screen to full : ' D


Form1.Top = 0
Form1.Left = 0
Player.Top = 0
Player.Left = 0

Form1.Width = Screen.Width
Form1.Height = Screen.Height
Player.Height = Screen.Height
Player.Width = Screen.Width

'Player.fullScreen = True
Player.URL = AVIfilename
Player.Controls.currentPosition = Position

Player.settings.mute = mute

End Sub

Private Sub Player_KeyPress(ByVal nKeyAscii As Integer)

'MsgBox nKeyAscii

If nKeyAscii = 51 Then

Player.Controls.currentPosition = Player.Controls.currentPosition - 10

End If

If nKeyAscii = 52 Then

Player.Controls.currentPosition = Player.Controls.currentPosition + 10

End If


If nKeyAscii = 50 Then
Msense = False

If mute = False Then
   Player.settings.mute = True
   mute = True

Else

    Player.settings.mute = False
    mute = False
    
End If

SaveSetting "superterran", "avi screensaver", "mute", mute
   
Msense = True

End If


If nKeyAscii = 49 Then
Msense = False

Dialog.ShowOpen
AVIfilename = Dialog.FileName
SaveSetting "superterran", "avi screensaver", "filename", AVIfilename

Player.URL = AVIfilename

Msense = True


End If

End Sub

Private Sub Player_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
If Msense = False Then Exit Sub

If first = False Then
    mx = fX
    first = True
End If



If mx <> fX Then

Position = Player.Controls.currentPosition
SaveSetting "superterran", "avi screensaver", "playpos", Position
End
End If

mx = fX

End Sub

