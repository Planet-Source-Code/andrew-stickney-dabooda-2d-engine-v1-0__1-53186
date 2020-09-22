VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaJoyStick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ButtonCount As Integer
Private AutoFire() As Boolean
Private ButtonState() As Boolean
Private Pressed() As Boolean

Private Up As Boolean
Private Down As Boolean
Private Left As Boolean
Private Right As Boolean

Private Center As Integer
Private Max As Integer, Min As Integer
Private PreMax As Integer, PreMin As Integer
Private PDeadZone As Integer

'Initialize
Public Sub Initialize(Hwnd&, Optional Range As Integer = 10000, Optional DeadZone As Integer = 1000)

    On Error GoTo NotInitialized
    
    Set DirectInput = DirectX.DirectInputCreate()
    
    Set JoyStickDevEnum = DirectInput.GetDIDevices(DI8DEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    
    If JoyStickDevEnum.GetCount = 0 Then
    Set JoyStickDevEnum = DirectInput.GetDIDevices(DI8DEVTYPE_GAMEPAD, DIEDFL_ATTACHEDONLY)
    End If
    
    If JoyStickDevEnum.GetCount = 0 Then
    Set JoyStickDevEnum = DirectInput.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
    End If
    
    'Cycled through all different types of controllers none found
    If JoyStickDevEnum.GetCount = 0 Then GoTo NotInitialized
        
    Set JoyStickDevice = DirectInput.CreateDevice(JoyStickDevEnum.GetItem(1).GetGuidInstance)
    JoyStickDevice.SetCommonDataFormat DIFORMAT_JOYSTICK
    JoyStickDevice.SetCooperativeLevel Hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
    
    JoyStickDevice.GetCapabilities JoyStickCaps
    
    With JoyStickRange
        .lHow = DIPH_DEVICE
        .lMin = 0
        .lMax = Range
    End With
    
    JoyStickDevice.SetProperty "DIPROP_RANGE", JoyStickRange
    
    'Get Range
    Center = Range / 2
    Max = Range
    Min = 0
    PreMax = Center + DeadZone
    PreMin = Center - DeadZone
    PDeadZone = DeadZone
    
    'Buttons
    ButtonCount = JoyStickCaps.lButtons
    ReDim AutoFire(ButtonCount)
    ReDim ButtonState(ButtonCount)
    ReDim Pressed(ButtonCount)
    
    'Accuire device
    JoyStickDevice.Acquire
    
    DaBoodaJoyStickOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaJoyStickOn = False

End Sub

Public Sub SetRange(Range As Integer, DeadZone As Integer)
    With JoyStickRange
        .lHow = DIPH_DEVICE
        .lMin = 0
        .lMax = Range
    End With
    
    JoyStickDevice.SetProperty "DIPROP_RANGE", JoyStickRange
    
    'Get Range
    Center = Range / 2
    Max = Range
    Min = 0
    PreMax = Center + DeadZone
    PreMin = Center - DeadZone
    PDeadZone = DeadZone
    
End Sub

Public Sub SetAutoFire(Button As Integer, Value As Boolean)
    If Button > ButtonCount - 1 Or Button < 0 Then
        Button = 0
    End If
    AutoFire(Button) = Value
End Sub

Public Sub PollJoyStick()
    Dim Index As Integer
    
    JoyStickDevice.GetDeviceStateJoystick JoyStickState
    
    For Index = 0 To ButtonCount - 1
    'Check KeyDown if AutoFire = true
        If JoyStickState.Buttons(Index) > 0 And AutoFire(Index) = True Then
            ButtonState(Index) = True
        End If
        
    'Check KeyUp if AutoFire = true
        If JoyStickState.Buttons(Index) = 0 And AutoFire(Index) = True Then
            ButtonState(Index) = False
        End If
        
    'Check KeyPressed if autoFire is Off
        If AutoFire(Index) = False And ButtonState(Index) = True And Pressed(Index) = True Then
            ButtonState(Index) = False
        End If
        
    'Check KeyDown if AutoFire = false
        If JoyStickState.Buttons(Index) > 0 And AutoFire(Index) = False And Pressed(Index) = False Then
            ButtonState(Index) = True
            Pressed(Index) = True
        End If
        
    'CheckKeyUp if autofire = false
        If JoyStickState.Buttons(Index) = 0 And AutoFire(Index) = False And Pressed(Index) = True Then
            ButtonState(Index) = False
            Pressed(Index) = False
        End If
        
    Next Index
    
    'Reset Checks
    Up = False
    Down = False
    Left = False
    Right = False
    
    'Check Up
    If JoyStickState.y < PreMin Then
        Up = True
    End If
    
    'Check Down
    If JoyStickState.y > PreMax Then
        Down = True
    End If
    
    'Check Left
    If JoyStickState.x < PreMin Then
        Left = True
    End If
    
    'Check Right
    If JoyStickState.x > PreMax Then
        Right = True
    End If
    
End Sub

Public Function GetUp() As Boolean
    GetUp = Up
    
End Function

Public Function GetDown() As Boolean
    GetDown = Down
    
End Function

Public Function GetLeft() As Boolean
    GetLeft = Left
    
End Function

Public Function GetRight() As Boolean
    GetRight = Right
    
End Function

Public Function GetX() As Integer
    Dim x As Integer
    x = JoyStickState.x - Center
    
    If x <= -PDeadZone Then x = x + PDeadZone
    If x < 0 And x > -PDeadZone Then x = 0
    If x >= PDeadZone Then x = x - PDeadZone
    If x > 0 And x < PDeadZone Then x = 0
    
    GetX = x
End Function

Public Function GetY() As Integer
    Dim y As Integer
    y = JoyStickState.y - Center
    
    If y <= -PDeadZone Then y = y + PDeadZone
    If y < 0 And y > -PDeadZone Then y = 0
    If y >= PDeadZone Then y = y - PDeadZone
    If y > 0 And y < PDeadZone Then y = 0
    
    GetY = y
End Function

Public Function GetButton(Index As Integer) As Boolean
    If Index < 0 Or Index > ButtonCount - 1 Then
        Index = 0
    End If
    GetButton = ButtonState(Index)
    
End Function

Public Function GetButtonCount() As Integer
    GetButtonCount = ButtonCount
End Function
