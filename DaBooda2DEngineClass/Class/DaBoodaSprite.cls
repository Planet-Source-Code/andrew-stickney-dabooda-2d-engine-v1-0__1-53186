VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ReferenceKey As Variant

Private XPosition As Single
Private YPosition As Single
Private TempXPosition As Single
Private TempYPosition As Single
Private Width As Single
Private Height As Single
Private MapReference As Variant
Private SpriteReference As Variant
Private GetRect As RECT
Private TextureReference As Variant
Private RotationAngle As Single
Private TempRotationAngle As Single
Private XOrigin As Single
Private YOrigin As Single
Private UpperLeftColor As Long
Private UpperRightColor As Long
Private LowerLeftColor As Long
Private LowerRightColor As Long
Private UpperLeftSpecular As Long
Private UpperRightSpecular As Long
Private LowerLeftSpecular As Long
Private LowerRightSpecular As Long
Private Id As String
Private Id2 As Integer
Private ZOrder As Integer
Private Visible As Boolean
Private Drawn As Boolean
Private Counter As Single
Private DrawnX As Single
Private DrawnY As Single

'Subs to Set Properties
Public Sub SetReferenceKey(sKey As Variant)
    ReferenceKey = sKey
End Sub

Public Sub SetXPosition(Value As Single, Optional ResetTemp As Boolean)
    XPosition = Value
    If ResetTemp = False Then
        TempXPosition = Value
    End If
End Sub

Public Sub SetYPosition(Value As Single, Optional ResetTemp As Boolean)
    YPosition = Value
    If ResetTemp = False Then
        TempYPosition = Value
    End If
End Sub

Public Sub SetTempXPosition(Value As Single)
    TempXPosition = Value
End Sub

Public Sub SetTempYPosition(Value As Single)
    TempYPosition = Value
End Sub

Public Sub SetWidth(Value As Single)
    Width = Value
End Sub

Public Sub SetHeight(Value As Single)
    Height = Value
End Sub

Public Sub SetMapReference(Value As Variant)
    MapReference = Value
End Sub

Public Sub SetSpriteReference(Value As Variant)
    SpriteReference = Value
End Sub

Public Sub SetTextureReference(Value As Variant)
    TextureReference = Value
End Sub

Public Sub SetGetRect(Left As Single, Top As Single, Width As Single, Height As Single)
    With GetRect
        .Left = Left
        .Top = Top
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
End Sub

Public Sub SetRotationAngle(Value As Single, Optional ResetTemp As Boolean)
    RotationAngle = Value
    If ResetTemp = False Then
        TempRotationAngle = Value
    End If
End Sub

Public Sub SetXOrigin(Value As Single)
    XOrigin = Value
End Sub

Public Sub SetYOrigin(Value As Single)
    YOrigin = Value
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

Public Sub SetId(Value As String)
    Id = Value
End Sub

Public Sub SetId2(Value As Integer)
    Id2 = Value
End Sub

Public Sub SetZOrder(Value As Integer)
    ZOrder = Value
End Sub

Public Sub SetVisible(Value As Boolean)
    Visible = Value
End Sub

Public Sub SetCounter(Value As Single)
    Counter = Value
End Sub

Public Sub SetDrawn(Value As Boolean)
    Drawn = Value
End Sub

Public Sub SetDrawnX(Value As Single)
    DrawnX = Value
End Sub

Public Sub SetDrawnY(Value As Single)
    DrawnY = Value
End Sub

Public Sub IncrementCounter()
    Counter = Counter + 1
End Sub

Public Sub QM_SetPutRect(pLeft As Single, pTop As Single, pWidth As Single, pHeight As Single)
    XPosition = pLeft
    TempXPosition = pLeft
    YPosition = pTop
    TempYPosition = pTop
    Width = pWidth
    Height = pHeight
End Sub

Public Sub QM_YAddition(Value As Single)
    YPosition = YPosition + Value
    TempYPosition = YPosition
End Sub

Public Sub QM_XAddition(Value As Single)
    XPosition = XPosition + Value
    TempXPosition = XPosition
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

Public Sub QM_SetSpriteGetByAnimation(AniPos As Single, Width As Single, Height As Single, TextureSize As Single, Xpos As Single, Ypos As Single)
    Dim Total As Single, x As Single, y As Single
    
    x = Xpos
    y = Ypos
    
    If AniPos > 1 Then
    For Total = 2 To AniPos
        x = x + Width
            If x >= TextureSize Then
                x = 0
                y = y + Height
            End If
    Next Total
    End If
    
    With GetRect
        .Left = x
        .Top = y
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
End Sub

Public Sub QM_Expand(Optional Offset As Single = 1)
    XPosition = XPosition - Offset
    YPosition = YPosition - Offset
    Width = Width + (Offset * 2)
    Height = Height + (Offset * 2)
End Sub

Public Sub QM_ExpandWidth(Optional Offset As Single = 1)
    XPosition = XPosition - Offset
    Width = Width + (Offset * 2)
End Sub

Public Sub QM_ExpandHeight(Optional Offset As Single = 1)
    YPosition = YPosition - Offset
    Height = Height + (Offset * 2)
End Sub

Public Sub QM_Contract(Optional Offset As Single = 1)
    XPosition = XPosition + Offset
    YPosition = YPosition + Offset
    Width = Width - (Offset * 2)
    Height = Height - (Offset * 2)
End Sub

Public Sub QM_ContractWidth(Optional Offset As Single = 1)
    XPosition = XPosition + Offset
    Width = Width - (Offset * 2)
End Sub

Public Sub QM_ContractHeight(Optional Offset As Single = 1)
    YPosition = YPosition + Offset
    Height = Height - (Offset * 2)
End Sub

'Functions to Get Properties
Public Function GetReferenceKey() As Variant
    GetReferenceKey = ReferenceKey
End Function

Public Function GetXPosition() As Single
    GetXPosition = XPosition
End Function

Public Function GetYPosition() As Single
    GetYPosition = YPosition
End Function

Public Function GetWidth() As Single
    GetWidth = Width
End Function

Public Function GetHeight() As Single
    GetHeight = Height
End Function

Public Function GetMapReference() As Variant
    GetMapReference = MapReference
End Function

Public Function GetSpriteReference() As Variant
    GetSpriteReference = SpriteReference
End Function

Public Function GetTextureReference() As Variant
    GetTextureReference = TextureReference
End Function

Public Function GetTempXPosition() As Single
    GetTempXPosition = TempXPosition
End Function

Public Function GetTempYPosition() As Single
    GetTempYPosition = TempYPosition
End Function

Public Function GetGetRectLeft() As Single
    GetGetRectLeft = GetRect.Left
End Function

Public Function GetGetRectTop() As Single
    GetGetRectTop = GetRect.Top
End Function

Public Function GetGetRectWidth() As Single
    GetGetRectWidth = GetRect.Right - GetRect.Left
End Function

Public Function GetGetRectHeight() As Single
    GetGetRectHeight = GetRect.Bottom - GetRect.Top
End Function

Public Function GetRotationAngle() As Single
    GetRotationAngle = RotationAngle
End Function

Public Function GetTempRotationAngle() As Single
    GetTempRotationAngle = TempRotationAngle
End Function

Public Function GetXOrigin() As Single
    GetXOrigin = XOrigin
End Function

Public Function GetYOrigin() As Single
    GetYOrigin = YOrigin
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

Public Function GetId() As String
    GetId = Id
End Function

Public Function GetId2() As Integer
    GetId2 = Id2
End Function

Public Function GetZOrder() As Integer
    GetZOrder = ZOrder
End Function

Public Function GetVisible() As Boolean
    GetVisible = Visible
End Function

Public Function GetDrawn() As Boolean
    GetDrawn = Drawn
End Function

Public Function GetCounter() As Single
    GetCounter = Counter
End Function

Public Function GetDrawnX() As Single
    GetDrawnX = DrawnX
End Function

Public Function GetDrawnY() As Single
    GetDrawnY = DrawnY
End Function

'Set Initialized Properties

Private Sub Class_Initialize()
    UpperLeftColor = D3DColorARGB(255, 255, 255, 255)
    UpperRightColor = D3DColorARGB(255, 255, 255, 255)
    LowerLeftColor = D3DColorARGB(255, 255, 255, 255)
    LowerRightColor = D3DColorARGB(255, 255, 255, 255)
    
End Sub
