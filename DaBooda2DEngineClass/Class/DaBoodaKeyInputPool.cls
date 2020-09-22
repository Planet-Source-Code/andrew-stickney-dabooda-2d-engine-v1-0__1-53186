VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaKeyInputPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DaBoodaKeyInput"
Attribute VB_Ext_KEY = "Member0" ,"DaBoodaKeyInput"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Public Function Add(Optional sKey As String, Optional KeyCode As Integer) As DaBoodaKeyInput
    'create a new object
    Dim objNewMember As DaBoodaKeyInput
    Set objNewMember = New DaBoodaKeyInput


    'set the properties passed into the method
    
    If Len(sKey) = 0 Then
        objNewMember.SetKey KeyCode
        mCol.Add objNewMember
    Else
        objNewMember.SetKey KeyCode
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As DaBoodaKeyInput
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Sub Initialize(Hwnd&, Optional Buffer& = 10)
On Error GoTo NotInitialized

'Initialize DirectInput
Set DirectInput = DirectX.DirectInputCreate
Set kDirectInputDevice = DirectInput.CreateDevice("GUID_SysKeyboard")

'Set InputDevice1
kDirectInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
kDirectInputDevice.SetCooperativeLevel Hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

'SetUp Buffer
DevProp.lHow = DIPH_DEVICE
DevProp.lData = Buffer
kDirectInputDevice.SetProperty DIPROP_BUFFERSIZE, DevProp

'Acquire Device
kDirectInputDevice.Acquire

    DaBoodaKeyInputOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaKeyInputOn = False
    
End Sub

Public Sub GetKeyState()

    'Get Key States
    kDirectInputDevice.GetDeviceStateKeyboard DIState1
    
End Sub
