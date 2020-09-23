VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox l 
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UpGrade"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
      Width           =   4095
   End
   Begin VB.PictureBox T 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   3
      Left            =   5760
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox T 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   2
      Left            =   3840
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox T 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   1
      Left            =   1920
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   7800
      Pattern         =   "*.BMP"
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar v 
      Height          =   1935
      Left            =   7800
      Max             =   10
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox T 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   0
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim DData() As Byte
Dim TOPEN As String
Sub Download(Address As String, Save As String)
    DData() = Inet1.OpenURL(Address, icByteArray)
    Open Save For Binary Access Write As #1
    Put #1, , DData()
    Close #1
End Sub
Sub LoadT(Address As String, Text As String)
Open Address For Input As #1
Text = Input(LOF(1), 1)
Close #1
End Sub
Sub Text2List(Text As String, Listt As ListBox)
Dim ad, part As String
For a = 0 To Len(Text)
part = Right$(Left$(Text, a), 1)
If part = Chr(13) Then
Listt.AddItem Right$(ad, Len(ad) - 1)
ad = Empty
Else
If Not part = Chr(13) Then
ad = ad & part
End If
End If
Next
If Not Listt.List(Listt.ListCount) = ad Then
Listt.AddItem ad
End If
End Sub
Sub FindUrl(Listt As ListBox)
Dim name As String
For i = 0 To Listt.ListCount - 1
If Left$(Listt.List(i), Len("http://")) = "http://" Then
Download Listt.List(i), App.Path & "\Pen\" & name & i & ".bmp"
Else
name = Empty
name = Listt.List(i)
End If
Next
End Sub
Private Sub Command1_Click()
On Error GoTo i:
Download "http://saeedsheikh.741.com/UPPEN.dat", App.Path & "\Pen\" & "UPPEN.dat"
LoadT App.Path & "\Pen\" & "UPPEN.dat", TOPEN
Text2List TOPEN, l
FindUrl l
Form2.Refresh
i:
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\Pen"
For i = 0 To T.Count - 1
File1.ListIndex = i
T(i).Picture = LoadPicture(File1.Path & "\" & File1.FileName)
v.Max = File1.ListCount
Next
End Sub

Private Sub T_Click(Index As Integer)
Form1.p.Picture = T(Index).Picture
End Sub
Private Sub T_DblClick(Index As Integer)
Form1.p.Picture = T(Index).Picture
Form2.Hide
End Sub

Private Sub v_Change()
v.Max = File1.ListCount - T.Count
For i = 0 To T.Count - 1
File1.ListIndex = i + (v.Value)
T(i).Picture = LoadPicture(File1.Path & "\" & File1.FileName)
Next
End Sub
