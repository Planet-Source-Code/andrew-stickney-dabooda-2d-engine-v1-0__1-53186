VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaKeyInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Up As Boolean
Private Down As Boolean
Private Reference As Integer
Private AutoFire As Boolean
Private Pressed As Boolean

Public Sub SetKey(KeyCode As Integer)
    Reference = KeyCode
End Sub

Public Function GetKey() As Integer
    GetKey = Reference
End Function

Public Sub SetAutoFire(Value As Boolean)
    AutoFire = Value
End Sub

Public Function GetAutoFire() As Boolean
    GetAutoFire = AutoFire
End Function

Public Function ReturnKeyState() As Boolean

    'Check KeyDown if AutoFire = true
        If DIState1.Key(Reference) = 128 And AutoFire = True Then
            Down = True
            Up = False
        End If
        
    'Check KeyUp if AutoFire = true
        If DIState1.Key(Reference) = 0 And AutoFire = True Then
            Down = False
            Up = True
        End If
        
    'Check KeyPressed if autoFire is Off
        If AutoFire = False And Down = True And Pressed = True Then
            Down = False
            Up = True
        End If
        
    'Check KeyDown if AutoFire = false
        If DIState1.Key(Reference) = 128 And AutoFire = False And Pressed = False Then
            Down = True
            Up = False
            Pressed = True
        End If
        
    'CheckKeyUp if autofire = false
        If DIState1.Key(Reference) = 0 And AutoFire = False And Pressed = True Then
            Down = False
            Up = True
            Pressed = False
        End If

    'Return value
    ReturnKeyState = Down
    
End Function

Private Sub Class_Initialize()
    Up = True
    Down = False
    AutoFire = False
    Pressed = False
End Sub

