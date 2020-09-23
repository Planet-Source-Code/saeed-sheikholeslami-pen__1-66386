VERSION 5.00
Begin VB.UserControl Transparent 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Transparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRectRgn Lib _
    "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib _
    "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowRgn Lib _
    "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib _
    "gdi32" (ByVal hObject As Long) As Long
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get MaskPicture() As Picture
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal p As Picture)
    Set UserControl.MaskPicture = p
    makeTransparent
    Set UserControl.Picture = p
    PropertyChanged "MaskPicture"
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal c As OLE_COLOR)
    UserControl.MaskColor = c
    makeTransparent
    PropertyChanged "MaskColor"
End Property
Private Sub makeTransparent()
    Dim hRgn As Long, X As Long
    If UserControl.MaskPicture <> 0 Then
        hRgn = CreateRectRgn(0, 0, 0, 0)
        UserControl.Width = ScaleX(UserControl.MaskPicture.Width)
        UserControl.Height = ScaleY(UserControl.MaskPicture.Height)
        UserControl.Extender.Container.Width = UserControl.Width
        UserControl.Extender.Container.Height = UserControl.Height
        UserControl.Extender.Move 0, 0
        DoEvents
        X = GetWindowRgn(UserControl.hwnd, hRgn)
        X = SetWindowRgn(UserControl.Extender.Container.hwnd, hRgn, True)
        X = DeleteObject(hRgn)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.MaskColor = PropBag.ReadProperty("MaskColor", &H8000000F)
    Set Me.MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MaskColor", Me.MaskColor, &H8000000F
    PropBag.WriteProperty "MaskPicture", Me.MaskPicture, Nothing
End Sub




