VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private TextureReference As Variant
Private UpperLeftColor As Long
Private UpperRightColor As Long
Private LowerLeftColor As Long
Private LowerRightColor As Long
Private UpperLeftSpecular As Long
Private UpperRightSpecular As Long
Private LowerLeftSpecular As Long
Private LowerRightSpecular As Long
Private GetRect As RECT
Private PutRect As RECT
Private Visible As Boolean
Private ZOrder As Integer
Private RotationAngle As Single
Private XOrigin As Single
Private YOrigin As Single

Private Sub Class_Initialize()
    TextureReference = 0
    UpperLeftColor = D3DColorARGB(255, 255, 255, 255)
    UpperRightColor = D3DColorARGB(255, 255, 255, 255)
    LowerLeftColor = D3DColorARGB(255, 255, 255, 255)
    LowerRightColor = D3DColorARGB(255, 255, 255, 255)
    With GetRect
        .Left = 0
        .Top = 0
        .Bottom = 256
        .Right = 256
    End With
    With PutRect
        .Left = 0
        .Top = 0
        .Bottom = 256
        .Right = 256
    End With
    Visible = False
    ZOrder = 0
    RotationAngle = 0
    XOrigin = 0
    YOrigin = 0
End Sub

'Subs to set Properties
Public Sub SetTextureReference(Index As Variant)
    TextureReference = Index
End Sub

Public Sub SetUpperLeftColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    UpperLeftColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetUpperRightColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    UpperRightColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetLowerLeftColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    LowerLeftColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetLowerRightColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    LowerRightColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetUpperLeftSpecular(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    UpperLeftSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetUpperRightSpecular(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    UpperRightSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetLowerLeftSpecular(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    LowerLeftSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetLowerRightSpecular(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    LowerRightSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetPutRect(Left As Single, Top As Single, Right As Single, Bottom As Single)
    With PutRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
End Sub

Public Sub SetGetRect(Left As Single, Top As Single, Right As Single, Bottom As Single)
    With GetRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
End Sub

Public Sub SetVisible(Value As Boolean)
    Visible = Value
End Sub

Public Sub SetZOrder(Value As Integer)
    ZOrder = Value
End Sub

Public Sub SetRotationAngle(Value As Single)
    RotationAngle = Value
End Sub

Public Sub SetXOrigin(Value As Single)
    XOrigin = Value
End Sub

Public Sub SetYOrigin(Value As Single)
    YOrigin = Value
End Sub

Public Sub QM_SetColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    Dim Col As Long
    Col = D3DColorARGB(Alpha, Red, Green, Blue)
    UpperLeftColor = Col
    UpperRightColor = Col
    LowerLeftColor = Col
    LowerRightColor = Col
End Sub

Public Sub QM_SetSpecular(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    Dim Col As Long
    Col = D3DColorARGB(Alpha, Red, Green, Blue)
    UpperLeftSpecular = Col
    UpperRightSpecular = Col
    LowerLeftSpecular = Col
    LowerRightSpecular = Col
End Sub

'Functions to Retrieve Properties
Public Function GetTextureReference() As Variant
    GetTextureReference = TextureReference
End Function

Public Function GetUpperLeftColor() As Long
    GetUpperLeftColor = UpperLeftColor
End Function

Public Function GetUpperRightColor() As Long
    GetUpperRightColor = UpperRightColor
End Function

Public Function GetLowerLeftColor() As Long
    GetLowerLeftColor = LowerLeftColor
End Function

Public Function GetLowerRightColor() As Long
    GetLowerRightColor = LowerRightColor
End Function

Public Function GetUpperLeftSpecular() As Long
    GetUpperLeftSpecular = UpperLeftSpecular
End Function

Public Function GetUpperRightSpecular() As Long
    GetUpperRightSpecular = UpperRightSpecular
End Function

Public Function GetLowerLeftSpecular() As Long
    GetLowerLeftSpecular = LowerLeftSpecular
End Function

Public Function GetLowerRightSpecular() As Long
    GetLowerRightSpecular = LowerRightSpecular
End Function

Public Function GetGetRectLeft() As Single
    GetGetRectLeft = GetRect.Left
End Function

Public Function GetGetRectTop() As Single
    GetGetRectTop = GetRect.Top
End Function

Public Function GetGetRectRight() As Single
    GetGetRectRight = GetRect.Right
End Function

Public Function GetGetRectBottom() As Single
    GetGetRectBottom = GetRect.Bottom
End Function

Public Function GetPutRectLeft() As Single
    GetPutRectLeft = PutRect.Left
End Function

Public Function GetPutRectTop() As Single
    GetPutRectTop = PutRect.Top
End Function

Public Function GetPutRectRight() As Single
    GetPutRectRight = PutRect.Right
End Function

Public Function GetPutRectBottom() As Single
    GetPutRectBottom = PutRect.Bottom
End Function

Public Function GetVisible() As Boolean
    GetVisible = Visible
End Function

Public Function GetZOrder() As Integer
    GetZOrder = ZOrder
End Function

Public Function GetRotationAngle() As Single
    GetRotationAngle = RotationAngle
End Function

Public Function GetXOrigin() As Single
    GetXOrigin = XOrigin
End Function

Public Function GetYOrigin() As Single
    GetYOrigin = YOrigin
End Function
