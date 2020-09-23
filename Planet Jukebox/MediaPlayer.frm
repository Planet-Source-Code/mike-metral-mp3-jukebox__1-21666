VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmPlanetJukebox 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planet Jukebox"
   ClientHeight    =   7020
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   7695
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "MediaPlayer.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF80&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   4080
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFF80&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FFFF80&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   720
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   720
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00FFFF00&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   720
   End
   Begin VB.CheckBox chkLoadMusic 
      BackColor       =   &H00000000&
      Caption         =   "Load Music"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Songs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H0080FF80&
      Caption         =   "Remove Song"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080FF80&
      Caption         =   "Add Song"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1260
      Left            =   4080
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.ListBox lstMusic 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2595
      ItemData        =   "MediaPlayer.frx":CF25
      Left            =   240
      List            =   "MediaPlayer.frx":CF27
      TabIndex        =   0
      Top             =   4560
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   29
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   28
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   27
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   26
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   25
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   24
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   600
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   960
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   960
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   240
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   120
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   20
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   4
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   20
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   3
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   20
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   2
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   20
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   1
      Left            =   240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblSongLength 
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSongs 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of songs in playlist: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   20
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   0
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadPlaylist 
         Caption         =   "&Load Playlist"
      End
      Begin VB.Menu mnuSavePlaylist 
         Caption         =   "&Save Playlist"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSkins 
      Caption         =   "S&kins"
      Begin VB.Menu mnuAcidEtching 
         Caption         =   "Acid &Etching"
      End
      Begin VB.Menu mnuAlien 
         Caption         =   "Alien &Plasma"
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "&Armor"
      End
      Begin VB.Menu mnuDancingSpirits 
         Caption         =   "&Dancing Spirits"
      End
      Begin VB.Menu mnuRockyEdge 
         Caption         =   "&Rocky Edge"
      End
      Begin VB.Menu mnuSteelMachine 
         Caption         =   "Steel &Machine"
      End
   End
End
Attribute VB_Name = "frmPlanetJukebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intSongs As Integer
Dim strFilename As String
Dim strFilename1 As String
Dim intcount As Integer
Dim strLocation As String
Dim intCircleColor As Integer
Dim intBorderColor As Integer

Private Sub chkLoadMusic_Click()
If chkLoadMusic.Value = 1 Then
cmdAdd.Enabled = True
Drive1.Visible = True
Dir1.Visible = True
File1.Visible = True
lblSongs.Top = 3720
cmdAdd.Top = 3840
cmdRemove.Top = 3840
cmdClear.Top = 3840
lstMusic.Top = 4320
lstMusic.Height = 2400
frmPlanetJukebox.Height = 7530
Else
cmdAdd.Enabled = False
Drive1.Visible = False
Dir1.Visible = False
File1.Visible = False
lblSongs.Top = 2400
cmdAdd.Top = 2400
cmdRemove.Top = 2400
cmdClear.Top = 2400
lstMusic.Top = 3240
frmPlanetJukebox.Height = 6900

End If
End Sub

Private Sub chkLoadMusic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 0 Then
chkLoadMusic.BackColor = vbBlue
End If
End Sub

Private Sub cmdPause_Click()
MMControl1.Command = "Pause"
End Sub

Private Sub cmdPlay_Click()

lblSongLength.Caption = MMControl1.Length
Timer2.Enabled = True
MMControl1.FileName = lstMusic.Text
MMControl1.Command = "Open"
MMControl1.Command = "Play"

End Sub

Private Sub cmdRemove_Click()
MMControl1.Command = "Close"

If lstMusic.ListIndex = -1 Then
 MsgBox ("Choose a song to remove"), , "Remove Song"
Else
 lstMusic.Tag = lstMusic.ListIndex
 lstMusic.RemoveItem lstMusic.Tag
 intSongs = intSongs - 1
 MsgBox ("Song Removed"), , "Remove Song"
End If

lblSongs.Caption = "Number of songs in playlist: " & intSongs

End Sub

Private Sub cmdStop_Click()
MMControl1.Command = "Close"
MMControl1.Command = "Open"
End Sub

Private Sub cmdAdd_Click()
MMControl1.Command = "Close"


If Dir1.Path = Drive1.Drive & "\" Then
strFilename = Dir1.Path & File1.FileName
Else
strFilename = Dir1.Path & "\" & File1.FileName
End If

If Right(strFilename, 3) = "mp3" Or Right(strFilename, 3) = "MP3" Or Right(strFilename, 3) = "mP3" Or Right(strFilename, 3) = "Mp3" Then
  lstMusic.AddItem strFilename
  MMControl1.DeviceType = "MPEGVideo"
  intSongs = intSongs + 1
Else
  MsgBox ("Enter in a appropriate music file"), , "File cannot be played"
End If
  
lblSongs.Caption = "Number of songs in playlist: " & intSongs

End Sub

Private Sub cmdClear_Click()

MMControl1.Command = "Pause"

If MsgBox("Are you sure you want to delete your entire playlist?", vbQuestion Or vbYesNo, "Delete Playlist") = vbYes Then
 lstMusic.Clear
 lblSongs.Caption = "Number of songs in playlist: 0"
 intSongs = 0
 MsgBox ("Playlist Deleted"), , "Media Player Notice"
 MMControl1.Command = "Close"
Else
 MsgBox ("Playlist Not Deleted"), , "Media Player Notice"
 MMControl1.Command = "Play"
End If

End Sub

Private Sub cmdNext_Click()
If lstMusic.ListCount = 1 Then
MsgBox ("Only 1 File in Playlist"), , "Next Command"
Else
lblSongLength.Caption = MMControl1.Length
Timer2.Enabled = True
MMControl1.Command = "Close"
If lstMusic.ListIndex <> intSongs - 1 Then
lstMusic.ListIndex = lstMusic.ListIndex + 1
ElseIf lstMusic.ListIndex = intSongs - 1 And lstMusic.ListIndex <> -1 Then
lstMusic.ListIndex = 0
End If
MMControl1.FileName = lstMusic.Text
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End If
End Sub

Private Sub cmdBack_Click()
If lstMusic.ListCount = 1 Then
MsgBox ("Only 1 File in Playlist"), , "Back Command"
Else
lblSongLength.Caption = MMControl1.Length
Timer2.Enabled = True
MMControl1.Command = "Close"
If lstMusic.ListIndex <> intSongs - 1 And lstMusic.ListIndex <> 0 Then
lstMusic.ListIndex = lstMusic.ListIndex - 1
ElseIf lstMusic.ListIndex = intSongs - 1 Then
lstMusic.ListIndex = lstMusic.ListIndex - 1
ElseIf lstMusic.ListIndex = 0 Then
MsgBox ("Can't Go Back Farther"), , "Back Command"
End If
MMControl1.FileName = lstMusic.Text
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End If
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim retries As Integer
retries = 0

On Error GoTo error

error:
If Err.Number = 68 Then
MsgBox ("Drive Not Ready"), , "Drive Error"
If retries >= 0 Then
Resume Next
End If
Else
Timer1.Enabled = True
Dir1.Path = Drive1.Drive
End If
End Sub

Private Sub Form_Load()
Randomize
mnuAlien.Checked = True
mnuRockyEdge.Checked = False
mnuSteelMachine.Checked = False
mnuArmor.Checked = False
mnuAcidEtching.Checked = False
mnuDancingSpirits.Checked = False

If mnuAlien.Checked = True Then
mnuAlien_Click
End If

intSongs = 0
lblSongs.Top = 2400
cmdAdd.Top = 2400
cmdRemove.Top = 2400
cmdClear.Top = 2400
lstMusic.Top = 3240
frmPlanetJukebox.Height = 7275

If MMControl1.Command = "Play" Then
 cmdAdd.Enabled = False
 cmdRemove.Enabled = False
 cmdClear.Enabled = False
ElseIf MMControl1.Command = "Stop" Then
 cmdAdd.Enabled = True
 cmdRemove.Enabled = True
 cmdClear.Enabled = True
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If chkLoadMusic.BackColor = vbBlue And Button = 0 Then
chkLoadMusic.BackColor = vbBlack
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MMControl1.Command = "Close"
End Sub

Private Sub lstMusic_DblClick()
MMControl1.Command = "Close"
MMControl1.FileName = lstMusic.Text
MMControl1.Command = "Open"
MMControl1.Command = "Play"
lblSongLength.Caption = MMControl1.Length
Timer2.Enabled = True
End Sub

Private Sub mnuAcidEtching_Click()
mnuAlien.Checked = False
mnuRockyEdge.Checked = False
mnuSteelMachine.Checked = False
mnuArmor.Checked = False
mnuAcidEtching.Checked = True
mnuDancingSpirits.Checked = False

frmPlanetJukebox.Picture = LoadPicture(App.Path & "\" & "Skins\acid etching.jpg")

Drive1.ForeColor = &H808000
Dir1.ForeColor = &H808000
File1.ForeColor = &H808000
lblSongs.ForeColor = &HFFFFFF
cmdAdd.BackColor = &HFFFF00
cmdRemove.BackColor = &HFFFF00
cmdClear.BackColor = &HFFFF00
cmdPlay.BackColor = &HFF&
cmdPause.BackColor = &HFF&
cmdBack.BackColor = &HFF&
cmdNext.BackColor = &HFF&
cmdStop.BackColor = &HFF&
Shape1.Item(1).FillColor = &H80FF80
Shape1.Item(2).FillColor = &H80FF80
Shape1.Item(4).FillColor = &H80FF80
Shape1.Item(0).FillColor = &HFFFF80
Shape1.Item(3).FillColor = &HFFFF80
lstMusic.ForeColor = &H808000

For intBorderColor = 0 To 4
Shape1.Item(intBorderColor).BorderColor = &H404000
Next intBorderColor

For intCircleColor = 1 To 40
Shape5.Item(intCircleColor).FillColor = &H4000&
Next intCircleColor

End Sub

Private Sub mnuAlien_Click()
mnuAlien.Checked = True
mnuRockyEdge.Checked = False
mnuSteelMachine.Checked = False
mnuArmor.Checked = False
mnuAcidEtching.Checked = False
mnuDancingSpirits.Checked = False

frmPlanetJukebox.Picture = LoadPicture(App.Path & "\" & "Skins\alien.jpg")

Drive1.ForeColor = &HFFFF&
Dir1.ForeColor = &HFFFF&
File1.ForeColor = &HFFFF&
lblSongs.ForeColor = &HFF00&
cmdAdd.BackColor = &H80FF80
cmdRemove.BackColor = &H80FF80
cmdClear.BackColor = &H80FF80
cmdPlay.BackColor = &HFFFF80
cmdPause.BackColor = &HFFFF80
cmdBack.BackColor = &HFFFF80
cmdNext.BackColor = &HFFFF80
cmdStop.BackColor = &HFFFF80
Shape1.Item(1).FillColor = &H800000
Shape1.Item(2).FillColor = &H800000
Shape1.Item(4).FillColor = &H800000
Shape1.Item(0).FillColor = &H8000&
Shape1.Item(3).FillColor = &H8000&
lstMusic.ForeColor = &HFFFF&

For intBorderColor = 0 To 4
Shape1.Item(intBorderColor).BorderColor = &H80FFFF
Next intBorderColor

For intCircleColor = 1 To 40
Shape5.Item(intCircleColor).FillColor = &H80&
Next intCircleColor

End Sub

Private Sub mnuArmor_Click()
mnuAlien.Checked = False
mnuRockyEdge.Checked = False
mnuSteelMachine.Checked = False
mnuArmor.Checked = True
mnuAcidEtching.Checked = False
mnuDancingSpirits.Checked = False

frmPlanetJukebox.Picture = LoadPicture(App.Path & "\" & "Skins\armor.jpg")

Drive1.ForeColor = &HFF&
Dir1.ForeColor = &HFF&
File1.ForeColor = &HFF&
lblSongs.ForeColor = &H80FF&
cmdAdd.BackColor = &H808080
cmdRemove.BackColor = &H808080
cmdClear.BackColor = &H808080
cmdPlay.BackColor = &HFF&
cmdPause.BackColor = &HFF&
cmdBack.BackColor = &HFF&
cmdNext.BackColor = &HFF&
cmdStop.BackColor = &HFF&
Shape1.Item(1).FillColor = &HFFFFFF
Shape1.Item(2).FillColor = &HFFFFFF
Shape1.Item(4).FillColor = &HFFFFFF
Shape1.Item(0).FillColor = &HFFFFFF
Shape1.Item(3).FillColor = &HFFFFFF
lstMusic.ForeColor = &HFF&

For intBorderColor = 0 To 4
Shape1.Item(intBorderColor).BorderColor = &H808080
Next intBorderColor

For intCircleColor = 1 To 40
Shape5.Item(intCircleColor).FillColor = &H800000
Next intCircleColor

End Sub

Private Sub mnuDancingSpirits_Click()
mnuAlien.Checked = False
mnuRockyEdge.Checked = False
mnuSteelMachine.Checked = False
mnuArmor.Checked = False
mnuAcidEtching.Checked = False
mnuDancingSpirits.Checked = True

frmPlanetJukebox.Picture = LoadPicture(App.Path & "\" & "Skins\spirit.jpg")

Drive1.ForeColor = &HFF0000
Dir1.ForeColor = &HFF0000
File1.ForeColor = &HFF0000
lblSongs.ForeColor = &HFF&
cmdAdd.BackColor = &H80FFFF
cmdRemove.BackColor = &H80FFFF
cmdClear.BackColor = &H80FFFF
cmdPlay.BackColor = &H80FF&
cmdPause.BackColor = &H80FF&
cmdBack.BackColor = &H80FF&
cmdNext.BackColor = &H80FF&
cmdStop.BackColor = &H80FF&
Shape1.Item(1).FillColor = &H80C0FF
Shape1.Item(2).FillColor = &H80C0FF
Shape1.Item(4).FillColor = &H80C0FF
Shape1.Item(0).FillColor = &HFFFF&
Shape1.Item(3).FillColor = &HFFFF&
lstMusic.ForeColor = &HFF0000

For intBorderColor = 0 To 4
Shape1.Item(intBorderColor).BorderColor = &H80FFFF
Next intBorderColor

For intCircleColor = 1 To 40
Shape5.Item(intCircleColor).FillColor = &H80&
Next intCircleColor

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLoadPlaylist_Click()
Dim File As String
CommonDialog1.DialogTitle = "Load your list."
  CommonDialog1.MaxFileSize = 16384
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Playlist Files|*.lst"
  CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
File = CommonDialog1.FileName
Dim a As String
Dim x As String
On Error GoTo error
Open File For Input As #1
Do Until EOF(1)
Input #1, a$
lstMusic.AddItem a$
Dim intSongs1 As Integer
intSongs1 = lstMusic.ListCount
intSongs = intSongs + 1
lblSongs.Caption = "Number of songs in playlist: " & intSongs1
Loop
Close 1
Exit Sub
MMControl1.Command = "Close"
error:
x = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Private Sub mnuRockyEdge_Click()
mnuAlien.Checked = False
mnuRockyEdge.Checked = True
mnuSteelMachine.Checked = False
mnuArmor.Checked = False
mnuAcidEtching.Checked = False
mnuDancingSpirits.Checked = False

frmPlanetJukebox.Picture = LoadPicture(App.Path & "\" & "Skins\rocky edge.jpg")

Drive1.ForeColor = &HFF00&
Dir1.ForeColor = &HFF00&
File1.ForeColor = &HFF00&
lblSongs.ForeColor = &HFFFFFF
cmdAdd.BackColor = &H808080
cmdRemove.BackColor = &H808080
cmdClear.BackColor = &H808080
cmdPlay.BackColor = &HFF&
cmdPause.BackColor = &HFF&
cmdBack.BackColor = &HFF&
cmdNext.BackColor = &HFF&
cmdStop.BackColor = &HFF&
Shape1.Item(1).FillColor = &HFF0000
Shape1.Item(2).FillColor = &HFF0000
Shape1.Item(4).FillColor = &HFF0000
Shape1.Item(0).FillColor = &HFF0000
Shape1.Item(3).FillColor = &HFF0000
lstMusic.ForeColor = &HFF00&

For intBorderColor = 0 To 4
Shape1.Item(intBorderColor).BorderColor = &H404040
Next intBorderColor

For intCircleColor = 1 To 40
Shape5.Item(intCircleColor).FillColor = &H808000
Next intCircleColor

End Sub

Private Sub mnuSavePlaylist_Click()
Dim PlaylistName As String
PlaylistName = InputBox("Name of the list?", "ListName")
PlaylistName = PlaylistName & ".lst"
CommonDialog2.DialogTitle = "Save your list."
CommonDialog2.MaxFileSize = 16384
CommonDialog2.FileName = PlaylistName
CommonDialog2.Filter = "Playlist Files|*.lst"
CommonDialog2.Filter = PlaylistName
CommonDialog2.DefaultExt = ".lst"
CommonDialog2.ShowSave
Open (PlaylistName) For Output As #1
      Dim i%
      For i = 0 To lstMusic.ListCount - 1
      Print #1, lstMusic.List(i)
      Next
      Close #1

End Sub


Private Sub mnuSteelMachine_Click()
mnuAlien.Checked = False
mnuRockyEdge.Checked = False
mnuSteelMachine.Checked = True
mnuArmor.Checked = False
mnuAcidEtching.Checked = False
mnuDancingSpirits.Checked = False

frmPlanetJukebox.Picture = LoadPicture(App.Path & "\" & "Skins\steel machine.jpg")

Drive1.ForeColor = &HC0&
Dir1.ForeColor = &HC0&
File1.ForeColor = &HC0&
lblSongs.ForeColor = &HFFFF00
cmdAdd.BackColor = &HFF0000
cmdRemove.BackColor = &HFF0000
cmdClear.BackColor = &HFF0000
cmdPlay.BackColor = &HFF0000
cmdPause.BackColor = &HFF0000
cmdBack.BackColor = &HFF0000
cmdNext.BackColor = &HFF0000
cmdStop.BackColor = &HFF0000
Shape1.Item(1).FillColor = &HFFFFC0
Shape1.Item(2).FillColor = &HFFFFC0
Shape1.Item(4).FillColor = &HFFFFC0
Shape1.Item(0).FillColor = &HFFFFC0
Shape1.Item(3).FillColor = &HFFFFC0
lstMusic.ForeColor = &HC0&

For intBorderColor = 0 To 4
Shape1.Item(intBorderColor).BorderColor = &H808080
Next intBorderColor

For intCircleColor = 1 To 40
Shape5.Item(intCircleColor).FillColor = &HC0&
Next intCircleColor

End Sub

Private Sub Timer2_Timer()
If MMControl1.Position = MMControl1.Length Then
Timer2.Enabled = False
cmdNext_Click
End If
End Sub

