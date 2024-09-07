VERSION 5.00
Begin VB.Form wizard3 
   Caption         =   "Windows XP Professional Activation Control Wizard"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Columns         =   2
      Height          =   2400
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   7215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   -120
      Picture         =   "wizard3.frx":0000
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
   Begin VB.Label Label3 
      Caption         =   "Select the Product ID you would like to use to activate Windows."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   7095
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
      TabIndex        =   3
      Top             =   1920
      Width           =   4770
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   5880
      Picture         =   "wizard3.frx":261A
      Top             =   5520
      Width           =   360
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
      TabIndex        =   2
      Top             =   5520
      Width           =   1410
   End
End
Attribute VB_Name = "wizard3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With List1

.AddItem "VFWGM-3GRYG-QB43V-MVP84-XB3G9"
.AddItem "VFWGM-3GRYG-QB43V-MVP84-XB3G9"
.AddItem "HHBCG-3V2QR-72KPV-YQCXW-KHGJ7"
.AddItem "YVGKQ-B7QYX-WWPBR-M7T2R-PB8BM"
.AddItem "XW6Q2-MP4HK-GXFK3-KPGG4-GM36T"
.AddItem "X4682-4BKMH-M4MRW-MY32H-FC8BD"
.AddItem "DX87K-XKPJ3-2H4D6-TX47Q-J46B6"
.AddItem "FYWR6-8KPYK-YM63K-FWWRV-MT3DP"
.AddItem "TVBWC-BMG9W-JG688-WFDQQ-DP7D2"
.AddItem "Q36CG-GVYDT-JQH4W-67XGR-99D8C"
.AddItem "3M37B-3HDFR-2KYQJ-YJ7VT-JYKR2"
.AddItem "M4KFB-K3TFW-BMG6B-622C4-RW7HW"
.AddItem "TWHCC-DC33F-G4JJP-BTR2B-RKRYT"
.AddItem "4X7WM-GTH3D-DWVCV-H382J-HPMRD"
.AddItem "VQD7P-3KK7H-M7VV2-CTXM4-MC7FW"
.AddItem "C34VY-TJYXD-3BG2V-HYX8T-76CY6"
.AddItem "3D2W3-8DJM6-YKQRB-B2XDB-TVQHF"
.AddItem "YXF2Y-QRRKR-BFKVQ-RHQ7H-DJPKD"
.AddItem "BMYY7-WH8QJ-6MTWG-MXXVQ-MD97B"
.AddItem "CRBH4-MXB2P-HP7V6-8YTMD-CBHJR"
.AddItem "G2JMP-2PC7G-RYBYX-PPF38-3KKTY"
.AddItem "HBJFW-XJ7K3-34JDX-VPPTW-227G6"
.AddItem "RXKFJ-67HBV-84TD2-RMDK8-9BDMT"
.AddItem "4FWCC-M3XVT-GQVVC-MKQYG-HP7YB"
.AddItem "VV2JP-HCKYQ-DMYB8-MQ733-6CHGC"
.AddItem "V8KG7-FRF6Y-WWRRB-G7KYY-TD4B7"
.AddItem "MTTXT-YX8JQ-6PC2M-TTXDT-WDM8K"
.AddItem "8V678-K66HP-GH28R-PTHKH-98PWP"
.AddItem "4BR3X-4CP6X-2DTXP-FFDHT-7Q298"
.AddItem "CFYHY-FQPJR-RWPC6-PWHKB-MXVKH"
.AddItem "YC62K-W8FW7-7BGVV-PYXD4-R679J"
.AddItem "KC4BB-2JHWW-VKCD6-2MXFV-98VH6"
.AddItem "27GY6-MPPMH-MJ43B-MPP2T-8WQ6Y"
.AddItem "8BCD7-WRTCW-JB6X6-XQF6J-2GCB2"
.AddItem "HVFK6-XQR33-PTW2H-VK6CX-TT738"
.AddItem "QGB7C-8VJ6F-WWHQB-VPVTD-KCPK4"
.AddItem "2P3K7-Y2CRK-T23MH-CR247-KT222"
.AddItem "KXWRG-72G83-P3J32-WB6MT-93JDR"
.AddItem "2KJ6K-BPRYY-6DQYR-C6HB6-FWD26"
.AddItem "BCX44-G46Y6-XBWTV-8QKHB-2VXJP"
.AddItem "8GV67-QRPTM-P6YMB-G2T6Y-D27X8"
.AddItem "2T7C7-3VTRV-2CFFB-2JHDD-QCBJ9"
.AddItem "G4XWB-YQX7Y-WHPW6-4BBBQ-YBBMY"
.AddItem "7XRRM-FQP8R-MW847-XMF6Q-XHYXK"
.AddItem "TBHXM-H6W74-4D8GM-B6XX4-M29T8"
.AddItem "47YK2-D8R6C-BPQBY-F4R3R-TVBTH"


End With

End Sub

Private Sub Label4_Click()
wizard4.Show

wizard4.Label3.Caption = List1.Text
Unload wizard3

End Sub
