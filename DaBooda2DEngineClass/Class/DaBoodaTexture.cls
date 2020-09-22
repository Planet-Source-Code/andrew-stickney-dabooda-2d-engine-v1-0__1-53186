VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Texture As Direct3DTexture8
Private Surface As Direct3DSurface8

Private Size As Single
Private Format As CONST_D3DFORMAT
Private FileType As Integer

Private Initialized As Boolean

Public Sub CreateTexture(tSize As Single, Optional tFile As String, Optional tFormat As CONST_D3DFORMAT, Optional tFileType As Integer)
    Size = tSize
    If tFormat = 0 Then Format = D3DFMT_A8R8G8B8
    FileType = tFileType
    If tFileType < 1 Or tFileType > 3 Then FileType = 1
    Initialized = True

    Select Case FileType
        Case 1
            Set Texture = D3DX.CreateTexture(Direct3DDevice, Size, Size, 0, 0, Format, D3DPOOL_MANAGED)
        Case 2
            tFile = TextureFolder & tFile & ".bmp"
            Set Texture = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, Size, Size, 0, 0, Format, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
        Case 3
            tFile = TextureFolder & tFile & ".dds"
            Set Texture = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, Size, Size, 0, 0, Format, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
    End Select
    
    Set Surface = Texture.GetSurfaceLevel(0)
    
End Sub

Public Function GetTexture() As Direct3DTexture8
    Set GetTexture = Texture
End Function

Public Sub SetTexture(rTexture As Direct3DTexture8)
    Set Texture = rTexture
    Set Surface = rTexture.GetSurfaceLevel(0)
End Sub

Public Function GetSize() As Single
    GetSize = Size
End Function

Public Function GetFormat() As CONST_D3DFORMAT
    GetFormat = Format
End Function

Public Function GetFileType() As Integer
    GetFileType = FileType
End Function

Private Sub Class_Initialize()
    Initialized = False
End Sub
