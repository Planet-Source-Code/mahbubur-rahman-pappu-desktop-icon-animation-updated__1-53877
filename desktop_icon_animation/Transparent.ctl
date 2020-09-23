VERSION 5.00
Begin VB.UserControl Transparent 
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
' Declare functions obtained from the API Text Viewer
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

' Event Declarations
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event DblClick()

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
    
    ' If there's no picture, don't do anything
    If UserControl.MaskPicture <> 0 Then
        hRgn = CreateRectRgn(0, 0, 0, 0)
        ' Set the width and height for the transparent control
        UserControl.Width = ScaleX(UserControl.MaskPicture.Width)
        UserControl.Height = ScaleY(UserControl.MaskPicture.Height)
        ' Adjust the size of the form to accomodate the control
        UserControl.Extender.Container.Width = UserControl.Width
        UserControl.Extender.Container.Height = UserControl.Height
        ' Position the Control on the container to top left
        UserControl.Extender.Move 0, 0
        
        ' Give Windows time to finish setting everything up
        DoEvents
        ' Get the window region for the Transparent Control,
        ' ignores anything else on the Form
        X = GetWindowRgn(UserControl.hwnd, hRgn)
        ' Clip the form to the region
        X = SetWindowRgn(UserControl.Extender.Container.hwnd, hRgn, True)
        ' Tidy up
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub



