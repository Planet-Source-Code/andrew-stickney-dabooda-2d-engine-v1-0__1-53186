VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaSoundPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DaBoodaSound"
Attribute VB_Ext_KEY = "Member0" ,"DaBoodaSound"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Public Function Add(Optional sKey As String) As DaBoodaSound
    'create a new object
    Dim objNewMember As DaBoodaSound
    Set objNewMember = New DaBoodaSound


    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As DaBoodaSound
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

Public Sub SetSoundFolder(Path As String)
    SoundFolder = Path
End Sub

Public Function GetsoundFolder() As String
    GetsoundFolder = SoundFolder
End Function

Public Sub Initialize(Hwnd&, Optional Device As Integer = 1)
    'Set up DirectSound
    Set DirectSoundEnum = DirectX.GetDSEnum
    Set DirectSound = DirectX.DirectSoundCreate(DirectSoundEnum.GetGuid(Device))
    DirectSound.SetCooperativeLevel Hwnd, DSSCL_NORMAL
    
    DaBoodaSoundOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaSoundOn = False
End Sub

Public Sub SetUpDirectionalSound(MaxVolume As Long, FieldSize As Long, Decay As Long)
    DSMaxVolume = MaxVolume
    DSFieldSize = FieldSize
    DSDecay = Decay
    
    If DSMaxVolume < 0 Then DSMaxVolume = 1
    If DSMaxVolume > 100 Then DSMaxVolume = 100
    
    Dim val As Long
    val = 10000 / 100
    val = val * DSMaxVolume
    DSMaxVolume = -10000 + val
    
End Sub

