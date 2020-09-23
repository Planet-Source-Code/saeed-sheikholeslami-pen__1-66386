VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Pen"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   DrawWidth       =   3
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Galam.frx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   6945
   Begin VB.CommandButton Command4 
      Caption         =   "SavePic"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cds 
      Left            =   1200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.jpg"
   End
   Begin MSComDlg.CommonDialog cdo 
      Left            =   720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.jpg"
      FilterIndex     =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load BackGroundPic"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ClearScreen"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   5280
      Width           =   1935
   End
   Begin Project1.Transparent Transparent1 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   2566
      _ExtentY        =   1720
      MaskColor       =   255
   End
   Begin VB.PictureBox D 
      AutoRedraw      =   -1  'True
      DrawWidth       =   3
      Height          =   4215
      Left            =   720
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.HScrollBar h 
      Height          =   375
      Left            =   4680
      Max             =   40
      Min             =   5
      TabIndex        =   2
      Top             =   3600
      Value           =   20
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Pen"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4200
      Width           =   1935
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   3
      Height          =   1935
      Left            =   4440
      Picture         =   "Galam.frx":8E122
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   " -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3600
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim z
Dim xx, yy As Single
Dim lX As Single, lY As Single
Dim moving As Boolean, makeChangesToRegistry, settingsAltered, settings_saved
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
D.Cls
End Sub

Private Sub Command3_Click()
On Error GoTo i:
cdo.ShowOpen
D.Picture = LoadPicture(cdo.FileName)
i:
End Sub

Private Sub Command4_Click()
On Error GoTo i:
cds.ShowSave
SavePicture D.Image, cds.FileName & ".jpg"
i:
End Sub

Private Sub Form_Load()
h.Value = (h.Max + h.Min) / 2
    Show
    Set Transparent1.MaskPicture = Form1.Picture
End Sub

Private Sub d_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo i:
If Button = 1 Then
z = h.Value / 10
For xx = X To X + (p.Width - 100) / z Step 15
For yy = Y To Y + (p.Height - 100) / z Step 15
If Not p.Point((xx - X) * z, (yy - Y) * z) = p.BackColor Then
D.Line (xx, yy)-(xx, yy), p.Point((xx - X) * z, (yy - Y) * z)
End If
Next
Next
End If
If Button = 2 Then
z = h.Value / 10
For xx = X To X + (p.Width - 100) / z Step 15
For yy = Y To Y + (p.Height - 100) / z Step 15
If Not p.Point((xx - X) * z, (yy - Y) * z) = p.BackColor Then
D.Line (xx, yy)-(xx, yy), D.BackColor
End If
Next
Next
End If
i:
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = True
    lX = X
    lY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If moving Then
        Me.Move (Me.Left + X - lX), (Me.Top + Y - lY)
        DoEvents
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
Me.WindowState = 1
End Sub

Private Sub p_DblClick()
p.Visible = False
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
p.Line (X, Y)-(X, Y)
End If
End Sub

