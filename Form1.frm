VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Pictures 2 B&W"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton sw 
      Caption         =   "Change Picture 2 B&&W"
      Height          =   375
      Left            =   1230
      TabIndex        =   5
      Top             =   3900
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Picture"
      Height          =   375
      Left            =   30
      TabIndex        =   4
      Top             =   3420
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3270
      TabIndex        =   3
      Top             =   3420
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3180
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Max             =   2945
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3360
      ScaleHeight     =   2955
      ScaleWidth      =   2925
      TabIndex        =   1
      Top             =   120
      Width           =   2985
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin MSComDlg.CommonDialog cmd 
         Left            =   960
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Picture2.Cls
End Sub

Private Sub Command2_Click()

cmd.ShowOpen
Picture1.Picture = LoadPicture(cmd.FileName, , vbLPVGAColor)

Command1_Click

End Sub

Private Sub Form_Load()
gr = 15
End Sub

Private Sub sw_Click()
Dim xx As Long, yy As Long, fa As Long, ds As Long, gr As Byte, r As Long, g As Long, b As Long, color As Long
    
gr = 15
xx = 1
yy = 1
Do
Do
    color = Picture1.Point(xx, yy)
    r = color Mod 256
    color = color \ 256
    g = color Mod 256
    color = color \ 256
    b = color Mod 256
    ds = (r + g + b) / 3
    fa = RGB(ds, ds, ds)
    Picture2.PSet (xx, yy), fa
    
    xx = xx + gr
Loop Until xx >= 128 * 23
bar.Value = yy
yy = yy + gr
xx = 1

Loop Until yy >= 128 * 23
yy = 1
xx = 1
bar.Value = 0.1

End Sub
