VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Superterran Alarm Clock 2.0 Demo"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   3480
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fade In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox sSec 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   690
      Left            =   4440
      TabIndex        =   7
      Text            =   "AM"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox sMin 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   690
      Left            =   4440
      TabIndex        =   6
      Text            =   "00"
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox sHour 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1290
      Left            =   2640
      TabIndex        =   5
      Text            =   "12"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   3000
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4320
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   3000
   End
   Begin VB.ListBox sPlaylist 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3150
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      BorderWidth     =   5
      X1              =   0
      X2              =   7920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Caption         =   "stop"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Add Song"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer Player 
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
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
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   1085
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      BorderWidth     =   5
      X1              =   0
      X2              =   7920
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   5
      X1              =   0
      X2              =   7920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape ind1 
      BackStyle       =   1  'Opaque
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1680
      Top             =   1800
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   5
      X1              =   0
      X2              =   7920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label dSec 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label dMin 
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label dHour 
      Alignment       =   1  'Right Justify
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "No file Selected"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'superterran MP3 Alarm Clock 2.0 - July 10'th 2004

Public alarmisplaying As Boolean
Public DefaultDir As String
Public rTime As String
Public rHour As Integer
Public rMin As Integer
Public rMOp As String
Public filename As String
Public bob As Integer
Public Alarm As Boolean
Public Playlist As String
Public AlarmFade As Boolean
Public FadeInc As Integer


Public Sub CurTime()




End Sub


Private Sub Check1_Click()
If Check1.Value = True Then
    AlarmFade = True
    Player.settings.volume = 0
End If

If Check1.Value = False Then
    AlarmFade = False
    Player.settings.volume = 100
End If

End Sub

Private Sub Label3_Click()
sHour.Text = "12"
sMin.Text = "00"
sSec.Text = "AM"
Alarm = True
Player.Close
Timer2.Enabled = False
ind1.Width = 0

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command1_Click()
sPlaylist.ListIndex = 1


End Sub

Public Sub SoundAlarm()
On Error GoTo errorsa:

If filename = "" Then filename = "default.wav"

Player.URL = filename
Timer2.Enabled = True
ind1.Visible = True

errorsa:
Exit Sub

End Sub


Private Sub Command3_Click()
MsgBox Player.currentMedia.duration
End Sub

Private Sub Form_Load()

'get the exact dir to the folder this is stored in.




STAPlaylist.Load
Form1.LoadDefaults
AlarmFade = True
Player.settings.volume = 0




With sHour

Do Until anum = 13
.AddItem anum
anum = anum + 1
Loop

End With

anum = 0

With sMin
Do Until anum = 60
If anum < 10 Then .AddItem "0" & anum Else .AddItem anum
anum = anum + 1

Loop
End With


sSec.AddItem "AM"
sSec.AddItem "PM"



End Sub

Private Sub Label2_Click()
Dialog.filename = "*.mp3"
Dialog.ShowOpen
filename = Dialog.filename
Label1.Caption = "..." & Right(filename, 30)
sPlaylist.AddItem (filename)

STAPlaylist.Save


End Sub

Private Sub List1_Click()

End Sub

Private Sub sPlaylist_Click()
filename = sPlaylist.Text
Label1.Caption = "..." & Right(filename, 30)
End Sub

Private Sub Timer1_Timer()

   

'If The Alarm is playing, this will govern if it's time to play the next song.

If alarmisplaying = True Then
'MsgBox Player.Controls.currentPosition
If Player.Controls.currentPosition = 0 Then

    Player.Controls.stop
    
    If sPlaylist.ListIndex = sPlaylist.ListCount - 1 Then sPlaylist.ListIndex = 0 Else sPlaylist.ListIndex = sPlaylist.ListIndex + 1
    filename = sPlaylist.Text
STAlarm.Start

       
    End If
End If







'checks to see if now is a good time to play the music

If Alarm = True Then
If sSec = dSec Then
    If sHour = dHour Then
        If sMin = dMin Then

Alarm = False
alarmisplaying = True

'checks to see if alarm fading is enabled, and if it is, it starts the process

If AlarmFade = True Then
    Player.settings.volume = 0
    Timer3.Enabled = True
    Timer3.Interval = FadeInc
    ind1.Height = 0
End If


STAlarm.Start

        End If
    End If
End If
End If

STAlarm.CurTime

End Sub

Private Sub Timer2_Timer()
'Dictates Indicator Bar Function


If jim = Form1.Width Then Exit Sub
On Error GoTo me1

jim = Form1.Width / Player.currentMedia.duration
ind1.Width = ind1.Width + jim


me1:
Exit Sub


End Sub

Public Function LoadDefaults()
AlarmFade = True
FadeInc = 1000


DefaultDir = VB.App.Path
dirtest = Right(DefaultDir, 1)
If dirtest <> "\" Then DefaultDir = DefaultDir & "\"

Alarm = True
ind1.Width = 0
anum = 1

ind1.Left = 0

sPlaylist.ListIndex = 0
filename = Form1.sPlaylist.Text

End Function

Private Sub Timer3_Timer()


If alarmisplaying = True Then
    If AlarmFade = True Then
       If Player.settings.volume < 100 Then Player.settings.volume = Player.settings.volume + 1
        If ind1.Height < 495 Then ind1.Height = ind1.Height + 4.95
        
        
        End If
    End If
    
End Sub
