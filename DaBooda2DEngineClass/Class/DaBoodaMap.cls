VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ReferenceKey As Variant

Private XCount As Integer
Private YCount As Integer
Private XPosition As Single
Private YPosition As Single
Private XIncrement As Single
Private YIncrement As Single
Private Width As Single
Private Height As Single
Private XReference As Variant
Private YReference As Variant
Private MoveUp As Boolean
Private MoveDown As Boolean
Private MoveLeft As Boolean
Private MoveRight As Boolean
Private Looping As Boolean
Private SubMapWidth As Single
Private SubMapHeight As Single
Private ZOrder As Integer

Private Type SubMapType
        TextureReference As Variant
        ULColor As Long
        URColor As Long
        LLColor As Long
        LRColor As Long
        ULSpecular As Long
        URSpecular As Long
        LLSpecular As Long
        LRSpecular As Long
        Visible As Boolean
        LimitUp As Boolean
        LimitDown As Boolean
        LimitLeft As Boolean
        LimitRight As Boolean
End Type

Private SubMap(64, 64) As SubMapType

'Set Map Properties
Public Sub SetReferenceKey(sKey As Variant)
    ReferenceKey = sKey
End Sub

Public Sub SetXCount(Value As Integer)
    If Value < 1 Then Value = 1
    If Value > 64 Then Value = 64
    XCount = Value
    Width = XCount * SubMapWidth
End Sub

Public Sub SetYCount(Value As Integer)
    If Value < 1 Then Value = 1
    If Value > 64 Then Value = 64
    YCount = Value
    Height = YCount * SubMapHeight
End Sub

Public Sub SetXPosition(Value As Single)
    XPosition = Value
End Sub

Public Sub SetYPosition(Value As Single)
    YPosition = Value
End Sub

Public Sub SetXIncrement(Value As Single)
    XIncrement = Value
End Sub

Public Sub SetYIncrement(Value As Single)
    YIncrement = Value
End Sub

Public Sub SetXReference(Value As Variant)
    XReference = Value
End Sub

Public Sub SetYReference(Value As Variant)
    YReference = Value
End Sub

Public Sub SetMoveUp(Value As Boolean)
    MoveUp = Value
End Sub

Public Sub SetMoveDown(Value As Boolean)
    MoveDown = Value
End Sub

Public Sub SetMoveLeft(Value As Boolean)
    MoveLeft = Value
End Sub

Public Sub SetMoveRight(Value As Boolean)
    MoveRight = Value
End Sub

Public Sub SetLooping(Value As Boolean)
    Looping = Value
End Sub

Public Sub SetSubMapWidth(Value As Single)
    SubMapWidth = Value
    Width = XCount * SubMapWidth
End Sub

Public Sub SetSubMapHeight(Value As Single)
    SubMapHeight = Value
    Height = YCount * SubMapHeight
End Sub

Public Sub SetZOrder(Order As Integer)
    ZOrder = Order
End Sub

Public Sub SetSubMapTextureReference(XIndex As Integer, YIndex As Integer, Value As Variant)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).TextureReference = Value
End Sub

Public Sub SetSubMapUpperLeftColor(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).ULColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapUpperRightColor(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).URColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapLowerLeftColor(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LLColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapLowerRightColor(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LRColor = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapUpperLeftSpecular(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).ULSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapUpperRightSpecular(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).URSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapLowerLeftSpecular(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LLSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapLowerRightSpecular(XIndex As Integer, YIndex As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LRSpecular = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub SetSubMapVisible(XIndex As Integer, YIndex As Integer, Value As Boolean)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).Visible = Value
End Sub

Public Sub SetSubMapLimitUp(XIndex As Integer, YIndex As Integer, Value As Boolean)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LimitUp = Value
End Sub

Public Sub SetSubMapLimitDown(XIndex As Integer, YIndex As Integer, Value As Boolean)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LimitDown = Value
End Sub

Public Sub SetSubMapLimitLeft(XIndex As Integer, YIndex As Integer, Value As Boolean)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LimitLeft = Value
End Sub

Public Sub SetSubMapLimitRight(XIndex As Integer, YIndex As Integer, Value As Boolean)
    If CheckIndex(XIndex) = False Then Exit Sub
    If CheckIndex(YIndex) = False Then Exit Sub
    SubMap(XIndex, YIndex).LimitRight = Value
End Sub

'functions for Getting Values
Public Function GetReferenceKey() As Variant
    GetReferenceKey = ReferenceKey
End Function

Public Function GetXCount() As Integer
    GetXCount = XCount
End Function

Public Function GetYCount() As Integer
    GetYCount = YCount
End Function

Public Function GetXPosition() As Single
    GetXPosition = XPosition
End Function

Public Function GetYPosition() As Single
    GetYPosition = YPosition
End Function

Public Function GetXIncrement() As Single
    GetXIncrement = XIncrement
End Function

Public Function GetYIncrement() As Single
    GetYIncrement = YIncrement
End Function

Public Function GetWidth() As Single
    GetWidth = Width
End Function

Public Function GetHeight() As Single
    GetHeight = Height
End Function

Public Function GetXReference() As Variant
    GetXReference = XReference
End Function

Public Function GetYReference() As Variant
    GetYReference = YReference
End Function

Public Function GetMoveUp() As Boolean
    GetMoveUp = MoveUp
End Function

Public Function GetMoveDown() As Boolean
    GetMoveDown = MoveDown
End Function

Public Function GetMoveLeft() As Boolean
    GetMoveLeft = MoveLeft
End Function

Public Function GetMoveRight() As Boolean
    GetMoveRight = MoveRight
End Function

Public Function GetLooping() As Boolean
    GetLooping = Looping
End Function

Public Function GetZOrder() As Integer
    GetZOrder = ZOrder
End Function

Public Function GetSubMapWidth() As Single
    GetSubMapWidth = SubMapWidth
End Function

Public Function GetSubMapHeight() As Single
    GetSubMapHeight = SubMapHeight
End Function

Public Function GetSubMapTextureReference(XIndex As Integer, YIndex As Integer) As Variant
    GetSubMapTextureReference = SubMap(XIndex, YIndex).TextureReference
End Function

Public Function GetSubMapUpperLeftColor(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapUpperLeftColor = SubMap(XIndex, YIndex).ULColor
End Function

Public Function GetSubMapUpperRightColor(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapUpperRightColor = SubMap(XIndex, YIndex).URColor
End Function

Public Function GetSubMapLowerLeftColor(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapLowerLeftColor = SubMap(XIndex, YIndex).LLColor
End Function

Public Function GetSubMapLowerRightColor(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapLowerRightColor = SubMap(XIndex, YIndex).LRColor
End Function

Public Function GetSubMapUpperLeftSpecular(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapUpperLeftSpecular = SubMap(XIndex, YIndex).ULSpecular
End Function

Public Function GetSubMapUpperRightSpecular(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapUpperRightSpecular = SubMap(XIndex, YIndex).URSpecular
End Function

Public Function GetSubMapLowerLeftSpecular(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapLowerLeftSpecular = SubMap(XIndex, YIndex).LLSpecular
End Function

Public Function GetSubMapLowerRightSpecular(XIndex As Integer, YIndex As Integer) As Long
    GetSubMapLowerRightSpecular = SubMap(XIndex, YIndex).LRSpecular
End Function

Public Function GetSubMapVisible(XIndex As Integer, YIndex As Integer) As Boolean
    GetSubMapVisible = SubMap(XIndex, YIndex).Visible
End Function

Public Function GetSubMapLimitUp(XIndex As Integer, YIndex As Integer) As Boolean
    GetSubMapLimitUp = SubMap(XIndex, YIndex).LimitUp
End Function

Public Function GetSubMapLimitDown(XIndex As Integer, YIndex As Integer) As Boolean
    GetSubMapLimitDown = SubMap(XIndex, YIndex).LimitDown
End Function

Public Function GetSubMapLimitLeft(XIndex As Integer, YIndex As Integer) As Boolean
    GetSubMapLimitLeft = SubMap(XIndex, YIndex).LimitLeft
End Function

Public Function GetSubMapLimitRight(XIndex As Integer, YIndex As Integer) As Boolean
    GetSubMapLimitRight = SubMap(XIndex, YIndex).LimitRight
End Function

'QuickMacros
Public Sub QM_SetColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    Dim Col As Long, XIndex As Integer, YIndex As Integer
    Col = D3DColorARGB(Alpha, Red, Green, Blue)
    For XIndex = 1 To XCount
        For YIndex = 1 To YCount
            With SubMap(XIndex, YIndex)
                .ULColor = Col
                .URColor = Col
                .LLColor = Col
                .LRColor = Col
            End With
        Next YIndex
    Next XIndex
End Sub

Public Sub QM_SetSpecular(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    Dim Col As Long, XIndex As Integer, YIndex As Integer
    Col = D3DColorARGB(Alpha, Red, Green, Blue)
    For XIndex = 1 To XCount
        For YIndex = 1 To YCount
            With SubMap(XIndex, YIndex)
                .ULSpecular = Col
                .URSpecular = Col
                .LLSpecular = Col
                .LRSpecular = Col
            End With
        Next YIndex
    Next XIndex
End Sub

Public Sub QM_SetSubMapLimits(XIndex As Integer, YIndex As Variant, Up As Boolean, Right As Boolean, Down As Boolean, Left As Boolean)
    With SubMap(XIndex, YIndex)
        .LimitUp = Up
        .LimitRight = Right
        .LimitDown = Down
        .LimitLeft = Left
    End With
End Sub

'Initialization for default variables
Private Sub Class_Initialize()
    Dim a As Integer, b As Integer
    
    XCount = 64
    YCount = 64
    XPosition = 0
    YPosition = 0
    XIncrement = 1
    YIncrement = 1
    Width = 256 * 64
    Height = 256 * 64
    XReference = 0
    YReference = 0
    MoveUp = True
    MoveDown = True
    MoveLeft = True
    MoveRight = True
    Looping = True
    SubMapWidth = 256
    SubMapHeight = 256
    
    For a = 1 To 64
        For b = 1 To 64
            With SubMap(a, b)
                .TextureReference = 0
                .LimitDown = False
                .LimitLeft = False
                .LimitRight = False
                .LimitUp = False
                .Visible = True
                .ULColor = D3DColorARGB(255, 255, 255, 255)
                .URColor = D3DColorARGB(255, 255, 255, 255)
                .LLColor = D3DColorARGB(255, 255, 255, 255)
                .LRColor = D3DColorARGB(255, 255, 255, 255)
                .ULSpecular = D3DColorARGB(0, 0, 0, 0)
                .URSpecular = D3DColorARGB(0, 0, 0, 0)
                .LLSpecular = D3DColorARGB(0, 0, 0, 0)
                .LRSpecular = D3DColorARGB(0, 0, 0, 0)
            End With
        Next b
    Next a
End Sub

Private Function CheckIndex(Index As Integer) As Boolean
    If Index > 0 And Index < 65 Then
        CheckIndex = True
    Else:
        CheckIndex = False
    End If
End Function
