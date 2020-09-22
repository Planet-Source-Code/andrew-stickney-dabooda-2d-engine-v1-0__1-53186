Attribute VB_Name = "DaBoodaGlobal"
Option Explicit

'These Variables are To Store The on or off values of initialized components
    Public DaBoodaDisplayOn As Boolean
    Public DaBoodaSoundOn As Boolean
    Public DaBoodaMusicOn As Boolean
    Public DaBoodaKeyInputOn As Boolean
    Public DaBoodaMouseOn As Boolean
    Public DaBoodaJoyStickOn As Boolean
    
'These Variables are the core of DirectX8
    Public DirectX As New DirectX8              'Main DirectX
    Public Direct3D As Direct3D8                'Direct3d Core
    Public Direct3DDevice As Direct3DDevice8    'Device to render to
    Public D3DX As New D3DX8                    'Helper Library to createTextures
    
'Variables for DirectInput
    Public DirectInput As DirectInput8
    Public kDirectInputDevice As DirectInputDevice8
    Public DIState1 As DIKEYBOARDSTATE
    Public DevProp As DIPROPLONG
    
'Variables for Mouse
    Public mDirectInputDevice As DirectInputDevice8
    Public mDevData() As DIDEVICEOBJECTDATA
    Public mEvents As Long
    Public mDevProp As DIPROPLONG
    
'Variables for DirectJoyStick
    Public JoyStickDevice As DirectInputDevice8
    Public JoyStickCaps As DIDEVCAPS
    Public JoyStickState As DIJOYSTATE
    Public JoyStickRange As DIPROPRANGE
    Public JoyStickDevEnum As DirectInputEnumDevices8
    
'Variables for DirectMusic
    Public DirectMusicPerformance As DirectMusicPerformance8
    Public DirectMusicLoader As DirectMusicLoader8
    Public DirectMusicSegment As DirectMusicSegment8
    
'Variables for DirectShow Objects
    Public DSAudio  As IBasicAudio         'Basic Audio Objectt
    Public DSEvent As IMediaEvent        'MediaEvent Object
    Public DSControl As IMediaControl    'MediaControl Object
    Public DSPosition As IMediaPosition 'MediaPosition Object
    Public MusicRepeat As Boolean
    
'Variables for DirectSound
    Public DirectSound As DirectSound8
    Public DirectSoundEnum As DirectSoundEnum8
    'for Directional Sound
    Public DSMaxVolume As Long
    Public DSFieldSize As Long
    Public DSDecay As Long
    
'Variables for font and Text
    Public ScreenFont As D3DXFont
    Public ScreenFontDesc As IFont
    Public TextRect As RECT
    Public StandardFont As New StdFont
    
'variables to keep class counts in to speed up render process
    Public SpriteMax As Integer, SpriteCount As Integer
    Public MapMax As Integer, MapCount As Integer
    Public OverLayMax As Integer, OverLayCount As Integer
    Public TextMax As Integer, TextCount As Integer
    
'Variables for Geometry
    'This represents the vertex format we will use
    Public Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    'This is your vertex structure,identical to directx7
    Public Type TLVertex
        x As Single
        y As Single
        z As Single
        rhw As Single
        Color As Long
        specular As Long
        tu As Single
        tv As Single
    End Type

'Globals for Folders
    Public TextureFolder As String
    Public SoundFolder As String
    Public MusicFolder As String
    
'Globals that are used in more than one class
    Public TextureColorKey As Long
    
    'This is just a simple Point Type
    Public Type Point
        x As Single
        y As Single
    End Type
    
    'This Type is for a Rect 2 point
    Public Type PointDouble
        X1 As Single
        Y1 As Single
        X2 As Single
        Y2 As Single
    End Type
    
    'This is a strip, where your geometry is stored, as a tlvertex type
    'This one strip will be used over and over again as needed
    'all four corners are needed
    Public TextureStrip(0 To 3) As TLVertex

'Globals for simple Rects
    Public PRect As RECT
    Public GRect As RECT
    
'Globals for special Sprite Collision Routines
    Public cSpriteHit As Boolean
    Public cSpriteX As Single
    Public cSpriteY As Single
    
'Globals for Returning Reference Points
    Public rSpriteX As Single
    Public rSpriteY As Single
    Public rSpriteDiffX As Single
    Public rSpriteDiffY As Single

'Globals for Copy Method
    Public CopySprite As New DaBoodaSprite
    
