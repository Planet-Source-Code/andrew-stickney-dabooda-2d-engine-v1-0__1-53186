VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaAutoMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

    Private AutoMoveWidth As Single
    Private AutoMoveHeight As Single
    Private AutoMoveMapReference As Variant
    Private AutoMoveSpriteReference As Variant
    Private AutoMoveXIncrement As Single
    Private AutoMoveYIncrement As Single
    Private AutoMoveOn As Boolean
    Private AutoMoveMoved As Boolean

'Subs to Set Up AutoMove
Public Sub SetWidth(Value As Single)
    AutoMoveWidth = Value
End Sub

Public Sub SetHeight(Value As Single)
    AutoMoveHeight = Value
End Sub

Public Sub SetMapReference(Value As Variant)
    AutoMoveMapReference = Value
End Sub

Public Sub SetSpriteReference(Value As Variant)
    AutoMoveSpriteReference = Value
End Sub

Public Sub SetXIncrement(Value As Single)
    AutoMoveXIncrement = Value
End Sub

Public Sub SetYIncrement(Value As Single)
    AutoMoveYIncrement = Value
End Sub

Public Sub SetOn(Value As Boolean)
    AutoMoveOn = Value
End Sub

Public Sub SetMoved(Value As Boolean)
    AutoMoveMoved = Value
End Sub

'functions for getting automove
Public Function GetWidth() As Single
    GetWidth = AutoMoveWidth
End Function

Public Function GetHeight() As Single
    GetHeight = AutoMoveHeight
End Function

Public Function GetMapReference() As Variant
    GetMapReference = AutoMoveMapReference
End Function

Public Function GetSpriteReference() As Variant
    GetSpriteReference = AutoMoveSpriteReference
End Function

Public Function GetXIncrement() As Single
    GetXIncrement = AutoMoveXIncrement
End Function

Public Function GetYIncrement() As Single
    GetYIncrement = AutoMoveYIncrement
End Function

Public Function GetOn() As Boolean
    GetOn = AutoMoveOn
End Function

Public Function GetMoved() As Boolean
    GetMoved = AutoMoveMoved
End Function
