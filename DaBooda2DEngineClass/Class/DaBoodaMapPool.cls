VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaMapPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DaBoodaMap"
Attribute VB_Ext_KEY = "Member0" ,"DaBoodaMap"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable to hold collection
Private mCol As Collection
Public Function Add(Optional sKey As String) As DaBoodaMap
    'create a new object
    Dim objNewMember As DaBoodaMap
    Set objNewMember = New DaBoodaMap


    'set the properties passed into the method
    If Len(sKey) = 0 Then
        objNewMember.SetReferenceKey mCol.Count + 1
        mCol.Add objNewMember
    Else
        objNewMember.SetReferenceKey sKey
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As DaBoodaMap
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

Public Sub Clear()
    If mCol.Count = 0 Then Exit Sub
    Dim Index As Integer
    
    For Index = mCol.Count To 1 Step -1
        mCol.Remove (Index)
    Next Index
    
End Sub
