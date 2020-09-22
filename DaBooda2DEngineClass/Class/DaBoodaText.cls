VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ZOrder As Integer
Private Visible As Boolean
Private Color As Long

Private fText As String
Private XPosition As Single, YPosition As Single
Private Width As Single, Height As Single

Public Sub SetZOrder(Order As Integer)
    ZOrder = Order
End Sub
Public Function GetZOrder() As Integer
    GetZOrder = ZOrder
End Function

Public Sub SetVisible(Toggle As Boolean)
    Visible = Toggle
End Sub
Public Function GetVisible() As Boolean
    GetVisible = Visible
End Function

Public Sub SetColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    Color = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub
Public Function GetColor() As Long
    GetColor = Color
End Function

Public Sub SetText(Text As String)
    fText = Text
End Sub
Public Function GetText() As String
    GetText = fText
End Function

Public Sub QM_SetPutRect(Xpos As Single, Ypos As Single, W As Single, H As Single)
    XPosition = Xpos
    YPosition = Ypos
    Width = W
    Height = H
End Sub

Public Sub SetXPosition(Value As Single)
    XPosition = Value
End Sub
Public Function GetXPosition() As Single
    GetXPosition = XPosition
End Function

Public Sub SetYPosition(Value As Single)
    YPosition = Value
End Sub
Public Function GetYPosition() As Single
    GetYPosition = YPosition
End Function

Public Sub SetWidth(Value As Single)
    Width = Value
End Sub
Public Function GetWidth() As Single
    GetWidth = Width
End Function

Public Sub SetHeight(Value As Single)
    Height = Value
End Sub
Public Function GetHeight() As Single
    GetHeight = Height
End Function

