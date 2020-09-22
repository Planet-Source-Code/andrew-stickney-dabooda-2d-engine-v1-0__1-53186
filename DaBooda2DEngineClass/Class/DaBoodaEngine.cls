VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Variables for Screen Mapping and display
    Private ScreenRect As RECT
    Private DisplayWidth As Single
    Private DisplayHeight As Single
    Private DisplayColor As Long

'This Rect Variable is Where the Map and Sprites are actual Drawn
    Private MapView As RECT
    Private MapXCount As Integer
    Private MapYCount As Integer
    Private MapClip As Single

'this Variable Determines how Many Levels are on the screen
'It is is used to determine zorder
    Private MaxLevel As Integer
    
'These Are Instances of all the classes
Public Sprite As New DaBoodaSpritePool
Public Map As New DaBoodaMapPool
Public Texture As New DaBoodaTexturePool
Public Overlay As New DaBoodaOverlayPool
Public FPS As New DaBoodaFRC
Public AutoMove As New DaBoodaAutoMove
Public KeyInput As New DaBoodaKeyInputPool
Public Music As New DaBoodaMusic
Public Sound As New DaBoodaSoundPool
Public TextEx As New DaBoodaTextPool
Public Mouse As New DaBoodaMouse
Public JoyStick As New DaBoodaJoyStick

Public Sub InitializeDisplay(Hwnd&, dWindowed As Boolean, Optional dWidth& = 320, Optional dHeight& = 240, Optional FramePS& = 60, Optional Color& = 0, Optional dAdapter As Long = 0, Optional DevType As CONST_D3DDEVTYPE, Optional dVsync As Boolean = True)
    On Error GoTo NotInitialized
    
'Initialize the Classes
Set Sprite = New DaBoodaSpritePool
Set Map = New DaBoodaMapPool
Set Texture = New DaBoodaTexturePool
Set Overlay = New DaBoodaOverlayPool
Set FPS = New DaBoodaFRC
Set AutoMove = New DaBoodaAutoMove
Set KeyInput = New DaBoodaKeyInputPool
Set Music = New DaBoodaMusic
Set Sound = New DaBoodaSoundPool
Set TextEx = New DaBoodaTextPool
Set Mouse = New DaBoodaMouse
Set JoyStick = New DaBoodaJoyStick

'Variables for displayMode
    Dim dParams As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
    Dim Temp As Single
    
'Set up Screen Variables
    DisplayWidth = dWidth
    DisplayHeight = dHeight
    DisplayColor = Color
    With ScreenRect
        .Left = 0
        .Top = 0
        .Right = dWidth
        .Bottom = dWidth
    End With
    Temp = FramePS
    FPS.SetFrameRate Temp
    
'Initialize Direct3d
    
    Set DirectX = New DirectX8
    Set Direct3D = DirectX.Direct3DCreate
    Set D3DX = New D3DX8

'Get Current DisplayMode
    Direct3D.GetAdapterDisplayMode dAdapter, DispMode
    
'Set the Device to Hardware if empty
    If DevType = 0 Then DevType = D3DDEVTYPE_HAL

'Set Up Parameters for displaymode
    If dWindowed = True Then
        
        With dParams
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = dWidth
        .BackBufferHeight = dHeight
        .EnableAutoDepthStencil = False
        .Windowed = 1
        End With
        If dVsync = True Then
            dParams.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else:
            dParams.SwapEffect = D3DSWAPEFFECT_COPY
        End If
    Else:
        With dParams
        .BackBufferFormat = DispMode.Format
        .EnableAutoDepthStencil = False
        .BackBufferWidth = dWidth
        .BackBufferHeight = dHeight
        .Windowed = 0
        .SwapEffect = D3DSWAPEFFECT_FLIP
        End With
    End If
    
'Set up DisplayMode
    Set Direct3DDevice = Direct3D.CreateDevice(dAdapter, DevType, Hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, dParams)
        
'Set Up Vertex Format
    Direct3DDevice.SetVertexShader FVF

'Turn Off Lighting
Direct3DDevice.SetRenderState D3DRS_LIGHTING, False
Direct3DDevice.SetRenderState D3DRS_SPECULARENABLE, True

'Set Render State
    Direct3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    Direct3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    Direct3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

'Set Up Texture Stages
    Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE

    DaBoodaDisplayOn = True
    
    Exit Sub
NotInitialized:
    MsgBox "Display Cannot Be Initialized", vbExclamation
End Sub

'Sub to render Screen
'Probably the mose important sub in engine
'next to the initialization sub
Public Sub Render()
'Set up Variable for Loop
    Dim Count As Integer
    Dim CountIndex As Integer
    
'Do Screen Limits
    If AutoMove.GetOn = True Then
        CheckAutoMoveSprite
        UpdateMapToLimit
    End If

    CheckMapLimit
    CheckSpritePosition
    CheckSpriteReferences
    
    'clear screen
    Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, DisplayColor, 1#, 0

'Begin scene
    Direct3DDevice.BeginScene

'Set Counts to zero
SpriteMax = Sprite.Count: SpriteCount = 0
MapMax = Map.Count: MapCount = 0
OverLayMax = Overlay.Count: OverLayCount = 0
TextMax = TextEx.Count: TextCount = 0

'Start Loop
    For Count = 0 To MaxLevel
'The Graphics
    'Do map
    If Map.Count > 0 And MapCount < MapMax Then
    For CountIndex = 1 To Map.Count
        If Map(CountIndex).GetZOrder = Count Then
            RenderMap CountIndex
            MapCount = MapCount + 1
        End If
    Next CountIndex
    End If
    'Do Overlays
    If Overlay.Count > 0 And OverLayCount < OverLayMax Then
    For CountIndex = 1 To Overlay.Count
        If Overlay(CountIndex).GetZOrder = Count Then
            RenderOverlay CountIndex
            OverLayCount = OverLayCount + 1
        End If
    Next CountIndex
    End If
    'do Sprites
    If Sprite.Count > 0 And SpriteCount < SpriteMax Then
    For CountIndex = 1 To Sprite.Count
        If Sprite(CountIndex).GetZOrder = Count Then
            RenderSprite CountIndex
            SpriteCount = SpriteCount + 1
        End If
    Next CountIndex
    End If
    'Do Text
    If TextEx.Count > 0 And TextCount < TextMax Then
    For CountIndex = 1 To TextEx.Count
        If TextEx(CountIndex).GetZOrder = Count Then
            RenderText CountIndex
            TextCount = TextCount + 1
        End If
    Next CountIndex
    End If
'Increment loop
    Next Count
    
'End Scene
    Direct3DDevice.EndScene

'Present to display
    Direct3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
'DoDelay
'do frame delay, to reach FrameRate
    FPS.UpdateFPS
    
'do events........keeps program from going into infinite loop
    DoEvents

'CheckMusic
If MusicRepeat = True Then
     Music.CheckMusicRepeat
End If

End Sub

'Subs to Set LevelMax
Public Sub SetMaxLevel(Value As Integer)
    MaxLevel = Value
End Sub

Public Function GetMaxLevel() As Integer
    GetMaxLevel = MaxLevel
End Function

'Set Color key
Public Sub SetColorKey(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    TextureColorKey = D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Function GetColorKey() As Long
    GetColorKey = TextureColorKey
End Function

'This sets up the view rect, where the maps and sprites are actually drawn
Public Sub SetUpMapView(mLeft As Single, mTop As Single, mRight As Single, mBottom As Single, mClip As Single)
'place Values into mapview rect
With MapView
    .Left = mLeft
    .Right = mRight
    .Top = mTop
    .Bottom = mBottom
End With

MapClip = mClip

End Sub

'Sub to copy rect from one texture to another
Public Sub TextureCopyRects(SourceIndex As Variant, DestIndex As Variant, GetX As Variant, GetY As Variant, PutX As Variant, PutY As Variant, Width As Variant, Height As Variant)
'temp surfaces and Textures
    Dim DestTexture As Direct3DTexture8
    Dim DestSurface As Direct3DSurface8
    Dim SourceTexture As Direct3DTexture8
    Dim SourceSurface As Direct3DSurface8
    
'Set textures to temps
    Set DestTexture = Texture(DestIndex).GetTexture
    Set DestSurface = DestTexture.GetSurfaceLevel(0)
    
    Set SourceTexture = Texture(SourceIndex).GetTexture
    Set SourceSurface = SourceTexture.GetSurfaceLevel(0)
    
'Set Up Rects
    With GRect
        .Left = GetX
        .Top = GetY
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    With PRect
        .Left = PutX
        .Top = PutY
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
'Blt from texture to texture
Direct3DDevice.CopyRects SourceSurface, GRect, 1, DestSurface, PRect

Texture(DestIndex).SetTexture DestTexture

End Sub

Public Function GetDisplayOn() As Boolean
    GetDisplayOn = DaBoodaDisplayOn
End Function

Public Function GetSoundOn() As Boolean
    GetSoundOn = DaBoodaSoundOn
End Function

Public Function GetMusicOn() As Boolean
    GetMusicOn = DaBoodaMusicOn
End Function

Public Function GetKeyInputOn() As Boolean
    GetKeyInputOn = DaBoodaKeyInputOn
End Function

Public Function GetMouseOn() As Boolean
    GetMouseOn = DaBoodaMouseOn
End Function

Public Function GetJoyStickOn() As Boolean
    GetJoyStickOn = DaBoodaJoyStickOn
End Function

'This Sub is important
'It takes a pixel value for texture and returns it into a value
'from 0 to 1
Private Function TextureValue(Total As Single, Current As Single) As Single
Dim PixelTex As Single
PixelTex = 1 / (Total)
TextureValue = PixelTex * (Current)
End Function

'This Function Makes the Strips
Private Function MakeStrip(x As Single, y As Single, z As Single, rhw As Single, Color As Long, specular As Long, tu As Single, tv As Single) As TLVertex
MakeStrip.x = x
MakeStrip.y = y
MakeStrip.z = z
MakeStrip.rhw = rhw
MakeStrip.Color = Color
MakeStrip.specular = specular
MakeStrip.tu = tu
MakeStrip.tv = tv

End Function

'Rotation Functions
Public Function RotateX(x As Single, y As Single, angle As Single) As Single
Dim TX As Single, TAngle As Single
'turns the angle degrees into radians
TAngle = angle * (3.141592654 / 180)

'Rotate on axis
TX = x * Cos(TAngle) - y * Sin(TAngle)

RotateX = TX

End Function

Public Function RotateY(x As Single, y As Single, angle As Single) As Single
Dim TY As Single, TAngle As Single
'turns the angle degrees into radians
TAngle = angle * (3.141592654 / 180)

'Rotate on axis
TY = y * Cos(TAngle) + x * Sin(TAngle)

RotateY = TY

End Function

'This Sub is Perhaps my neatest invention...it sets a sprites angle according to another
'this is good for tracking
Public Sub SetAutoAngle(FromSprite As Variant, ToSprite As Variant)
    Dim X1 As Single, Y1 As Single
    Dim X2 As Single, Y2 As Single
    Dim FSprite As DaBoodaSprite, tSprite As DaBoodaSprite
    Dim fAngle As Single
    
    Set FSprite = Sprite(FromSprite)
    Set tSprite = Sprite(ToSprite)
    
    X1 = FSprite.GetXPosition + (FSprite.GetWidth / 2)
    X2 = tSprite.GetXPosition + (tSprite.GetWidth / 2)
    Y1 = FSprite.GetYPosition + (FSprite.GetHeight / 2)
    Y2 = tSprite.GetYPosition + (tSprite.GetHeight / 2)
    
    fAngle = RetrieveAngle(X1, Y1, X2, Y2)
        
    Sprite(FromSprite).SetRotationAngle fAngle
    
End Sub

'this sub Returns an angle according to two positions
Public Function RetrieveAngle(FromX As Single, FromY As Single, ToX As Single, ToY As Single) As Single
    Dim Result As Single
    Dim DiffX As Single, DiffY As Single
    Dim FromPoint As Point, ToPoint As Point
    
    FromPoint.x = FromX: FromPoint.y = FromY
    ToPoint.x = ToX: ToPoint.y = ToY
    
    DiffX = ToPoint.x - FromPoint.x
    DiffY = ToPoint.y - FromPoint.y
    
    If DiffX = 0 And FromPoint.y > ToPoint.y Then
        RetrieveAngle = 0
        Exit Function
    End If
    
    If DiffX = 0 And ToPoint.y > FromPoint.y Then
        RetrieveAngle = 180
        Exit Function
    End If
    
    If DiffX = 0 And DiffY = 0 Then
        RetrieveAngle = 0
        Exit Function
    End If
    
    Result = Atn(DiffY / DiffX) * 180 / 3.14
    
    If FromPoint.x > ToPoint.x Then Result = Result - 90
    If FromPoint.x < ToPoint.x Then Result = Result + 90
    
    If Result < 0 Then
        Result = Result + 360
    End If
    
    RetrieveAngle = Result
    
End Function

'Simple Function Returns Random Number
Public Function GetRandomNumber(Random As Single, Optional Offset As Single = 1) As Single
    Randomize Timer
    GetRandomNumber = Int(Rnd * Random) + Offset
End Function

'This is a function to check to sprites for collision
Public Function CheckSpriteCollision(Sprite1 As Variant, Sprite2 As Variant) As Boolean
    Dim FSprite As DaBoodaSprite
    Dim tSprite As DaBoodaSprite
    
    Set FSprite = Sprite(Sprite1)
    Set tSprite = Sprite(Sprite2)
    
    Dim fRect As RECT
    Dim tRect As RECT
    
    With fRect
        .Left = FSprite.GetXPosition
        .Right = .Left + FSprite.GetWidth
        .Top = FSprite.GetYPosition
        .Bottom = .Top + FSprite.GetHeight
    End With
    
    With tRect
        .Left = tSprite.GetXPosition
        .Right = .Left + tSprite.GetWidth
        .Top = tSprite.GetYPosition
        .Bottom = .Top + tSprite.GetHeight
    End With

    CheckSpriteCollision = RectInRect(fRect, tRect)
    
End Function

'These subs are for collision Detection
Private Function RectInRect(Rect1 As RECT, Rect2 As RECT) As Boolean
Dim CPoint(4) As Point
    RectInRect = False
    
    CPoint(1).x = Rect1.Left: CPoint(1).y = Rect1.Top
    CPoint(2).x = Rect1.Right: CPoint(2).y = Rect1.Top
    CPoint(3).x = Rect1.Left: CPoint(3).y = Rect1.Bottom
    CPoint(4).x = Rect1.Right: CPoint(4).y = Rect1.Bottom
    
    Dim a As Integer
    a = 1
    Do While a < 5
    If PointInRect(CPoint(a).x, CPoint(a).y, Rect2) = True Then
        RectInRect = True
        Exit Function
    End If
    a = a + 1
    Loop
    
End Function

Private Function PointInRect(XPosition As Variant, YPosition As Variant, CRect As RECT) As Boolean
If XPosition < CRect.Left Then PointInRect = False: Exit Function
If XPosition > CRect.Right Then PointInRect = False: Exit Function
If YPosition < CRect.Top Then PointInRect = False: Exit Function
If YPosition > CRect.Bottom Then PointInRect = False: Exit Function

cSpriteX = XPosition - CRect.Left + 1
cSpriteY = YPosition - CRect.Top + 1
PointInRect = True

End Function

'This next Group of subs is for special Collision Detection,
'Basically the corner is checked and values are returned into variables
'This is good for finding where your sprites collide
Public Sub CheckUpperLeftSCollision(PointSprite As Variant, CheckSprite As Variant)
    Dim pSprite As DaBoodaSprite, cSprite As DaBoodaSprite
    Dim CheckRect As RECT

    Set pSprite = Sprite(PointSprite)
    Set cSprite = Sprite(CheckSprite)
    
    With CheckRect
        .Left = cSprite.GetXPosition
        .Top = cSprite.GetYPosition
        .Right = .Left + cSprite.GetWidth
        .Bottom = .Top + cSprite.GetHeight
    End With
    
    cSpriteHit = PointInRect(pSprite.GetXPosition, pSprite.GetYPosition, CheckRect)
    
End Sub

Public Sub CheckUpperRightSCollision(PointSprite As Variant, CheckSprite As Variant)
    Dim pSprite As DaBoodaSprite, cSprite As DaBoodaSprite
    Dim CheckRect As RECT

    Set pSprite = Sprite(PointSprite)
    Set cSprite = Sprite(CheckSprite)
    
    With CheckRect
        .Left = cSprite.GetXPosition
        .Top = cSprite.GetYPosition
        .Right = .Left + cSprite.GetWidth
        .Bottom = .Top + cSprite.GetHeight
    End With
    
    cSpriteHit = PointInRect(pSprite.GetXPosition + pSprite.GetWidth, pSprite.GetYPosition, CheckRect)
    
End Sub

Public Sub CheckLowerLeftSCollision(PointSprite As Variant, CheckSprite As Variant)
    Dim pSprite As DaBoodaSprite, cSprite As DaBoodaSprite
    Dim CheckRect As RECT

    Set pSprite = Sprite(PointSprite)
    Set cSprite = Sprite(CheckSprite)
    
    With CheckRect
        .Left = cSprite.GetXPosition
        .Top = cSprite.GetYPosition
        .Right = .Left + cSprite.GetWidth
        .Bottom = .Top + cSprite.GetHeight
    End With
    
    cSpriteHit = PointInRect(pSprite.GetXPosition, pSprite.GetYPosition + pSprite.GetHeight, CheckRect)
    
End Sub

Public Sub CheckLowerRightSCollision(PointSprite As Variant, CheckSprite As Variant)
    Dim pSprite As DaBoodaSprite, cSprite As DaBoodaSprite
    Dim CheckRect As RECT

    Set pSprite = Sprite(PointSprite)
    Set cSprite = Sprite(CheckSprite)
    
    With CheckRect
        .Left = cSprite.GetXPosition
        .Top = cSprite.GetYPosition
        .Right = .Left + cSprite.GetWidth
        .Bottom = .Top + cSprite.GetHeight
    End With
    
    cSpriteHit = PointInRect(pSprite.GetXPosition + pSprite.GetWidth, pSprite.GetYPosition + pSprite.GetHeight, CheckRect)
    
End Sub

'This sub Plays a directional sound Based on the Center of a sprite
Public Sub PlayDirectionalSound(tSprite As Variant, tSound As Variant)
    Dim Temp As New DaBoodaSprite
    Set Temp = Sprite.Item(tSprite)
    
    Dim cX As Single, cY As Single
    
    cX = Temp.GetWidth / 2
    cY = Temp.GetHeight / 2
    cX = Temp.GetDrawnX + cX
    cY = Temp.GetDrawnY + cY
    
    Sound(tSound).DSPlaySound cX, cY
    
End Sub
'These next subs actually are used if the programmer is using independent maps
'the values returned are self explanatory
Public Sub GetSpriteReferencePoint(rSprite As Variant, Width As Single, Height As Single, Optional XOffset As Single, Optional YOffSet As Single)
    Dim Xpos As Single, Ypos As Single
    Xpos = Sprite(rSprite).GetTempXPosition + XOffset
    Ypos = Sprite(rSprite).GetTempYPosition + YOffSet
    
    rSpriteX = Int(Xpos / Width)
    rSpriteY = Int(Ypos / Height)
    
    rSpriteDiffX = Xpos - (rSpriteX * Width) + 1
    rSpriteDiffY = Ypos - (rSpriteY * Height) + 1
    
    rSpriteX = rSpriteX + 1
    rSpriteY = rSpriteY + 1

End Sub

Public Function GetRSpriteX() As Single
    GetRSpriteX = rSpriteX
End Function

Public Function GetRSpriteY() As Single
    GetRSpriteY = rSpriteY
End Function

Public Function GetRSpriteDiffX() As Single
    GetRSpriteDiffX = rSpriteDiffX
End Function

Public Function GetRSpriteDiffY() As Single
    GetRSpriteDiffY = rSpriteDiffY
End Function

'Subs to get values from specialchecks
Public Function GetCSpriteHit() As Boolean
    GetCSpriteHit = cSpriteHit
End Function
Public Function GetCSpriteX() As Single
    GetCSpriteX = cSpriteX
End Function
Public Function GetCSpriteY() As Single
    GetCSpriteY = cSpriteY
End Function

'This Subs Check Character to Screen and moves accordingly
Private Sub CheckAutoMoveSprite()
    Dim StartX As Single, StartY As Single
    Dim DiffLeft As Single, DiffRight As Single
    Dim DiffUp As Single, DiffDown As Single
    Dim CLeft As Single, CRight As Single
    Dim CUp As Single, CDown As Single
    Dim cRef As Variant, mRef As Variant
    
    AutoMove.SetMoved False
    
    'Set up References
    cRef = AutoMove.GetSpriteReference
    mRef = AutoMove.GetMapReference
    
    'Set Up StartX and y
    StartX = Map(mRef).GetXPosition
    StartY = Map(mRef).GetYPosition
    
    'Set up Screen
    DiffLeft = (MapView.Right - MapView.Left)
    DiffLeft = (DiffLeft - AutoMove.GetWidth) / 2
    DiffLeft = DiffLeft - StartX
    DiffRight = DiffLeft + AutoMove.GetWidth
    
    DiffUp = (MapView.Bottom - MapView.Top)
    DiffUp = (DiffUp - AutoMove.GetHeight) / 2
    DiffUp = DiffUp - StartY
    DiffDown = DiffUp + AutoMove.GetHeight

    'Set up Container
    CLeft = Sprite(cRef).GetXPosition
    CRight = CLeft + Sprite(cRef).GetWidth
    CUp = Sprite(cRef).GetYPosition
    CDown = Sprite(cRef).GetHeight + CUp
    
    'Check and move accordingly
    If CLeft < DiffLeft Then MoveMapRight mRef
    If CRight > DiffRight Then MoveMapLeft mRef
    If CUp < DiffUp Then MoveMapDown mRef
    If CDown > DiffDown Then MoveMapUp mRef
    
End Sub

Private Sub UpdateMapToLimit()
    Dim LayerX As Single, LayerY As Single
    Dim IndexX As Integer, IndexY As Integer
    Dim LX As Single, LY As Single
    Dim LimitLeft As Boolean, LimitRight As Boolean, LimitTop As Boolean, LimitBottom As Boolean
    Dim LayerRect As RECT
    Dim LimitRect As RECT
    Dim ScrollX As Single, ScrollY As Single
    Dim AutoMRef As Variant
    Dim AutoSRef As Variant
    Dim XCount As Integer, YCount As Integer
    Dim SubMapWidth As Single
    Dim SubMapHeight As Single
    
    'Get AutoMoveReference
    AutoMRef = AutoMove.GetMapReference
    AutoSRef = AutoMove.GetSpriteReference
    
    'Get SubMapSize
    SubMapWidth = Map(AutoMRef).GetSubMapWidth
    SubMapHeight = Map(AutoMRef).GetSubMapHeight
    
    'GetScrollX and ScrollY
    ScrollX = Map(AutoMRef).GetXPosition
    ScrollY = Map(AutoMRef).GetYPosition
    
    'GetCounts
    XCount = Map(AutoMRef).GetXCount
    YCount = Map(AutoMRef).GetYCount
    
    'Set Layer
    LayerX = (Sprite(AutoSRef).GetWidth) / 2
    LayerY = (Sprite(AutoSRef).GetHeight) / 2
    LayerX = Int((Sprite(AutoSRef).GetXPosition + LayerX) / SubMapWidth) + 1
    LayerY = Int((Sprite(AutoSRef).GetYPosition + LayerY) / SubMapHeight) + 1
    
    LX = LayerX
    LY = LayerY
    
    If LayerX < 1 Then LayerX = XCount - LayerX
    If LayerY < 1 Then LayerY = YCount - LayerY
    If LayerX > XCount Then LayerX = LayerX - XCount
    If LayerY > YCount Then LayerY = LayerY - YCount
    
    IndexX = LayerX
    IndexY = LayerY
    
    'Get Limit
    LimitLeft = Map(AutoMRef).GetSubMapLimitLeft(IndexX, IndexY)
    LimitRight = Map(AutoMRef).GetSubMapLimitRight(IndexX, IndexY)
    LimitTop = Map(AutoMRef).GetSubMapLimitUp(IndexX, IndexY)
    LimitBottom = Map(AutoMRef).GetSubMapLimitDown(IndexX, IndexY)
    
    
    'Set Up Layer Rect
    With LayerRect
        .Left = ((LX - 1) * SubMapWidth) + ScrollX
        .Top = ((LY - 1) * SubMapHeight) + ScrollY
        .Right = .Left + SubMapWidth
        .Bottom = .Top + SubMapHeight
    End With
    
    
    'Set Up LimitRect
    With LimitRect
        .Left = 0
        .Top = 0
        .Right = MapView.Right - MapView.Left
        .Bottom = MapView.Bottom - MapView.Top
    End With
    
    Map(AutoMRef).SetMoveLeft True
    Map(AutoMRef).SetMoveUp True
    Map(AutoMRef).SetMoveRight True
    Map(AutoMRef).SetMoveDown True
    
    'Check For Limits
    If LimitTop = True And LayerRect.Top > LimitRect.Top Then
        ScrollY = ScrollY - AutoMove.GetYIncrement
        Map(AutoMRef).SetMoveDown False
    End If
    
    If LimitBottom = True And LayerRect.Bottom < LimitRect.Bottom Then
        ScrollY = ScrollY + AutoMove.GetYIncrement
        Map(AutoMRef).SetMoveUp False
    End If
    
    If LimitLeft = True And LayerRect.Left > LimitRect.Left Then
        ScrollX = ScrollX - AutoMove.GetXIncrement
        Map(AutoMRef).SetMoveRight False
    End If
    
    If LimitRight = True And LayerRect.Right < LimitRect.Right Then
        ScrollX = ScrollX + AutoMove.GetXIncrement
        Map(AutoMRef).SetMoveLeft False
    End If
    
    If LimitTop = True And LayerRect.Top = LimitRect.Top Then
        Map(AutoMRef).SetMoveDown False
    End If
    If LimitLeft = True And LayerRect.Left = LimitRect.Left Then
        Map(AutoMRef).SetMoveRight False
    End If
    If LimitBottom = True And LayerRect.Bottom = LimitRect.Bottom Then
        Map(AutoMRef).SetMoveUp False
    End If
    If LimitRight = True And LayerRect.Right = LimitRect.Right Then
        Map(AutoMRef).SetMoveLeft False
    End If
    
    'Set Changes into Scrolls
    Map(AutoMRef).SetXPosition ScrollX
    Map(AutoMRef).SetYPosition ScrollY
    
End Sub

Private Sub CheckMapLimit()
Dim Index As Integer
Dim Diff As Single, Size As Single
    
If Map.Count = 0 Then Exit Sub

    For Index = 1 To Map.Count
        If Map(Index).GetXPosition >= Map(Index).GetWidth Then
            Size = Int(Map(Index).GetXPosition / Map(Index).GetSubMapWidth)
            Diff = Map(Index).GetXPosition - (Size * Map(Index).GetSubMapWidth)
            Map(Index).SetXPosition Diff
        End If
        If Map(Index).GetXPosition <= -Map(Index).GetWidth Then
            Size = Int(Abs(Map(Index).GetXPosition) / Map(Index).GetSubMapWidth)
            Diff = Map(Index).GetXPosition + (Size * Map(Index).GetSubMapWidth)
            Map(Index).SetXPosition Diff
        End If
        If Map(Index).GetYPosition >= Map(Index).GetHeight Then
            Size = Int(Map(Index).GetYPosition / Map(Index).GetSubMapHeight)
            Diff = Map(Index).GetYPosition - (Size * Map(Index).GetSubMapHeight)
            Map(Index).SetYPosition Diff
        End If
        If Map(Index).GetYPosition <= -Map(Index).GetHeight Then
            Size = Int(Abs(Map(Index).GetYPosition) / Map(Index).GetSubMapHeight)
            Diff = Map(Index).GetYPosition + (Size * Map(Index).GetSubMapHeight)
            Map(Index).SetYPosition Diff
        End If
DontUpdate:
    Next Index

End Sub

Private Sub CheckSpritePosition()

Dim Index As Integer
Dim CRect As RECT
Dim Ref As Variant
Dim X1 As Single, Y1 As Single
    For Index = 1 To Sprite.Count
        If Sprite(Index).GetMapReference = "" Then GoTo DontOffset
        Ref = Sprite(Index).GetMapReference
        If Map(Ref).GetLooping = False Then GoTo DontOffset
        
        X1 = Sprite(Index).GetXPosition
        Y1 = Sprite(Index).GetYPosition

        With CRect
            .Left = -(Map(Ref).GetXPosition) - MapClip
            .Top = -(Map(Ref).GetYPosition) - MapClip
            .Right = .Left + Map(Ref).GetWidth - 1
            .Bottom = .Top + Map(Ref).GetHeight - 1
        End With
        
    'Check Farthest Right
        If X1 >= CRect.Left + Map(Ref).GetWidth And X1 <= CRect.Right + Map(Ref).GetWidth Then X1 = X1 - Map(Ref).GetWidth
    'Check Farthest Left
        If X1 >= CRect.Left - Map(Ref).GetWidth And X1 <= CRect.Right - Map(Ref).GetWidth Then X1 = X1 + Map(Ref).GetWidth
    'Check Farthest Down
        If Y1 >= CRect.Top + Map(Ref).GetHeight And Y1 <= CRect.Bottom + Map(Ref).GetHeight Then Y1 = Y1 - Map(Ref).GetHeight
    'Check Farthest up
        If Y1 >= CRect.Top - Map(Ref).GetHeight And Y1 <= CRect.Bottom - Map(Ref).GetHeight Then Y1 = Y1 + Map(Ref).GetHeight
       
    Sprite(Index).SetXPosition X1
    Sprite(Index).SetYPosition Y1

DontOffset:

    Next Index
End Sub

Private Sub CheckSpriteReferences()
    Dim x As Single, y As Single
    Dim Rx As Single, Ry As Single
    Dim XOrigin As Single, YOrigin As Single
    Dim Index As Integer
    Dim Ref As Variant
    
    For Index = 1 To Sprite.Count
        If Sprite(Index).GetSpriteReference = "" Then GoTo DontReference
        Ref = Sprite(Index).GetSpriteReference
        
        'Get Offsets
        XOrigin = Sprite(Ref).GetXPosition + Sprite(Ref).GetXOrigin
        YOrigin = Sprite(Ref).GetYPosition + Sprite(Ref).GetYOrigin
        
        x = Sprite(Ref).GetXPosition + Sprite(Index).GetTempXPosition - XOrigin
        y = Sprite(Ref).GetYPosition + Sprite(Index).GetTempYPosition - YOrigin
        
        Rx = RotateX(x, y, Sprite(Ref).GetRotationAngle)
        Ry = RotateY(x, y, Sprite(Ref).GetRotationAngle)
        
        Sprite(Index).SetXPosition XOrigin + Rx, True
        Sprite(Index).SetYPosition YOrigin + Ry, True
        
        'Set Rotation
        Sprite(Index).SetRotationAngle Sprite(Ref).GetRotationAngle + Sprite(Index).GetTempRotationAngle, True
        
DontReference:
    Next Index
    
End Sub

Public Sub MoveMapUp(Index As Variant)
    If Map(Index).GetMoveUp = True Then
        AutoMove.SetMoved True
        Map(Index).SetYPosition Map(Index).GetYPosition - Map(Index).GetYIncrement
    Dim Index2 As Integer, MapRef As Variant
    MapRef = Map(Index).GetReferenceKey
    For Index2 = 1 To Map.Count
        If Map(Index2).GetYReference = MapRef Then
            Map(Index2).SetYPosition Map(Index2).GetYPosition - Map(Index2).GetYIncrement
        End If
    Next Index2
    End If
End Sub

Public Sub MoveMapDown(Index As Variant)
    If Map(Index).GetMoveDown = True Then
        AutoMove.SetMoved True
        Map(Index).SetYPosition Map(Index).GetYPosition + Map(Index).GetYIncrement
    Dim Index2 As Integer, MapRef As Variant
    MapRef = Map(Index).GetReferenceKey
    For Index2 = 1 To Map.Count
        If Map(Index2).GetYReference = MapRef Then
            Map(Index2).SetYPosition Map(Index2).GetYPosition + Map(Index2).GetYIncrement
        End If
    Next Index2
    End If
End Sub

Public Sub MoveMapLeft(Index As Variant)
    If Map(Index).GetMoveLeft = True Then
        AutoMove.SetMoved True
        Map(Index).SetXPosition Map(Index).GetXPosition - Map(Index).GetXIncrement
    Dim Index2 As Integer, MapRef As Variant
    MapRef = Map(Index).GetReferenceKey
    For Index2 = 1 To Map.Count
        If Map(Index2).GetXReference = MapRef Then
            Map(Index2).SetXPosition Map(Index2).GetXPosition - Map(Index2).GetXIncrement
        End If
    Next Index2
    End If
End Sub

Public Sub MoveMapRight(Index As Variant)
    If Map(Index).GetMoveRight = True Then
        AutoMove.SetMoved True
        Map(Index).SetXPosition Map(Index).GetXPosition + Map(Index).GetXIncrement
    Dim Index2 As Integer, MapRef As Variant
    MapRef = Map(Index).GetReferenceKey
    For Index2 = 1 To Map.Count
        If Map(Index2).GetXReference = MapRef Then
            Map(Index2).SetXPosition Map(Index2).GetXPosition + Map(Index2).GetXIncrement
        End If
    Next Index2
    End If
End Sub

Public Sub SpriteAngleIncrementUp(Index As Variant, Optional Increment& = 1)
    Dim x As Single, y As Single
    x = RotateX(0, -Increment, Sprite(Index).GetTempRotationAngle)
    y = RotateY(0, -Increment, Sprite(Index).GetTempRotationAngle)
    Sprite(Index).SetXPosition Sprite(Index).GetTempXPosition + x
    Sprite(Index).SetYPosition Sprite(Index).GetTempYPosition + y
End Sub

Public Sub SpriteAngleIncrementDown(Index As Variant, Optional Increment As Single)
    Dim x As Single, y As Single
    If Increment = 0 Then Increment = 1
    x = RotateX(0, Increment, Sprite(Index).GetTempRotationAngle)
    y = RotateY(0, Increment, Sprite(Index).GetTempRotationAngle)
    Sprite(Index).SetXPosition Sprite(Index).GetTempXPosition + x
    Sprite(Index).SetYPosition Sprite(Index).GetTempYPosition + y
End Sub

Public Sub SpriteAngleIncrementLeft(Index As Variant, Optional Increment As Single)
    Dim x As Single, y As Single
    If Increment = 0 Then Increment = 1
    x = RotateX(-Increment, 0, Sprite(Index).GetTempRotationAngle)
    y = RotateY(-Increment, 0, Sprite(Index).GetTempRotationAngle)
    Sprite(Index).SetXPosition Sprite(Index).GetTempXPosition + x
    Sprite(Index).SetYPosition Sprite(Index).GetTempYPosition + y
End Sub

Public Sub SpriteAngleIncrementRight(Index As Variant, Optional Increment As Single)
    Dim x As Single, y As Single
    If Increment = 0 Then Increment = 1
    x = RotateX(Increment, 0, Sprite(Index).GetTempRotationAngle)
    y = RotateY(Increment, 0, Sprite(Index).GetTempRotationAngle)
    Sprite(Index).SetXPosition Sprite(Index).GetTempXPosition + x
    Sprite(Index).SetYPosition Sprite(Index).GetTempYPosition + y
End Sub

Public Sub SpriteIncrementUp(Index As Variant, Optional Increment& = 1)
    If Increment = 0 Then Increment = 1
    Sprite(Index).SetYPosition Sprite(Index).GetTempYPosition - Increment
End Sub

Public Sub SpriteIncrementDown(Index As Variant, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).SetYPosition Sprite(Index).GetTempYPosition + Increment
End Sub

Public Sub SpriteIncrementLeft(Index As Variant, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).SetXPosition Sprite(Index).GetTempXPosition - Increment
End Sub

Public Sub SpriteIncrementRight(Index As Variant, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).SetXPosition Sprite(Index).GetTempXPosition + Increment
End Sub

'This is the Render Overlay Sub
Private Sub RenderOverlay(Order As Integer)
'Variables for Positions
    Dim ULX As Single, ULY As Single
    Dim URX As Single, URY As Single
    Dim LLX As Single, LLY As Single
    Dim LRX As Single, LRY As Single
    
'Variables for Textures
    Dim TULX As Single, TULY As Single
    Dim TURX As Single, TURY As Single
    Dim TLLX As Single, TLLY As Single
    Dim TLRX As Single, TLRY As Single

'Variables for Colors
    Dim ULColor As Long
    Dim URColor As Long
    Dim LLColor As Long
    Dim LRColor As Long
    Dim ULSpecular As Long
    Dim URSpecular As Long
    Dim LLSpecular As Long
    Dim LRSpecular As Long
    
    
'Variables for Rotation
    Dim Rx As Single
    Dim Ry As Single
    Dim XOrigin As Single
    Dim YOrigin As Single
    Dim angle As Single

'Variables for Texture
    Dim OverlayTextureSize As Single
    Dim OverlayTexture As Direct3DTexture8
    Dim TOverlay As DaBoodaOverlay
    
'Variable for loop
    Dim Count As Long
'Loop through Overlays
    Count = Order
    Set TOverlay = Overlay.Item(Count)
    
'Check to see if it is the right order and on
    If TOverlay.GetVisible = False Then GoTo DontDraw

'Get Texture Information
    Set OverlayTexture = Texture(TOverlay.GetTextureReference).GetTexture
    OverlayTextureSize = Texture(TOverlay.GetTextureReference).GetSize
    
'Place Positions and rotate in variables
'Rotates them too
    XOrigin = TOverlay.GetPutRectLeft + TOverlay.GetXOrigin
    YOrigin = TOverlay.GetPutRectTop + TOverlay.GetYOrigin
    angle = TOverlay.GetRotationAngle
    
    ULX = TOverlay.GetPutRectLeft: ULY = TOverlay.GetPutRectTop
        Rx = ULX - XOrigin: Ry = ULY - YOrigin
            ULX = XOrigin + RotateX(Rx, Ry, angle): ULY = YOrigin + RotateY(Rx, Ry, angle)
    
    URX = TOverlay.GetPutRectRight: URY = TOverlay.GetPutRectTop
        Rx = URX - XOrigin: Ry = URY - YOrigin
            URX = XOrigin + RotateX(Rx, Ry, angle): URY = YOrigin + RotateY(Rx, Ry, angle)
    
    LLX = TOverlay.GetPutRectLeft: LLY = TOverlay.GetPutRectBottom
        Rx = LLX - XOrigin: Ry = LLY - YOrigin
            LLX = XOrigin + RotateX(Rx, Ry, angle): LLY = YOrigin + RotateY(Rx, Ry, angle)
    
    LRX = TOverlay.GetPutRectRight: LRY = TOverlay.GetPutRectBottom
        Rx = LRX - XOrigin: Ry = LRY - YOrigin
            LRX = XOrigin + RotateX(Rx, Ry, angle): LRY = YOrigin + RotateY(Rx, Ry, angle)

'Setup Colors
    ULColor = TOverlay.GetUpperLeftColor
    URColor = TOverlay.GetUpperRightColor
    LLColor = TOverlay.GetLowerLeftColor
    LRColor = TOverlay.GetLowerRightColor
    ULSpecular = TOverlay.GetUpperLeftSpecular
    URSpecular = TOverlay.GetUpperRightSpecular
    LLSpecular = TOverlay.GetLowerLeftSpecular
    LRSpecular = TOverlay.GetLowerRightSpecular

'setup Textures
    TULX = TextureValue(OverlayTextureSize, TOverlay.GetGetRectLeft)
    TULY = TextureValue(OverlayTextureSize, TOverlay.GetGetRectTop)
    
    TURX = TextureValue(OverlayTextureSize, TOverlay.GetGetRectRight)
    TURY = TextureValue(OverlayTextureSize, TOverlay.GetGetRectTop)
    
    TLLX = TextureValue(OverlayTextureSize, TOverlay.GetGetRectLeft)
    TLLY = TextureValue(OverlayTextureSize, TOverlay.GetGetRectBottom)
    
    TLRX = TextureValue(OverlayTextureSize, TOverlay.GetGetRectRight)
    TLRY = TextureValue(OverlayTextureSize, TOverlay.GetGetRectBottom)

'MakeStrip........Finally
    TextureStrip(0) = MakeStrip(ULX, ULY, 0, 1, ULColor, ULSpecular, TULX, TULY)
    TextureStrip(1) = MakeStrip(URX, URY, 0, 1, URColor, URSpecular, TURX, TURY)
    TextureStrip(2) = MakeStrip(LLX, LLY, 0, 1, LLColor, LLSpecular, TLLX, TLLY)
    TextureStrip(3) = MakeStrip(LRX, LRY, 0, 1, LRColor, LRSpecular, TLRX, TLRY)
    
'Set texture
            Direct3DDevice.SetTexture 0, OverlayTexture
            
'DrawStrip
            Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TextureStrip(0), Len(TextureStrip(0))

'Exit for not Drawn
DontDraw:

End Sub

'This is the RenderMap Sub
Private Sub RenderMap(Index As Variant)

'Variables for sub
    Dim X1 As Integer, X2 As Single, Y1 As Integer, Y2 As Single     'These are for positions
    Dim TopX As Single, TopY As Single                             'These are the top map positions
    Dim XCount As Integer, YCount As Integer
    Dim XLayer As Integer, YLayer As Integer
    Dim XStart As Single, YStart As Single
    Dim Num As Single
    
    Dim ColorUL As Long
    Dim ColorUR As Long
    Dim ColorLL As Long
    Dim ColorLR As Long
    
    Dim SpecularUL As Long
    Dim SpecularUR As Long
    Dim SpecularLL As Long
    Dim SpecularLR As Long

    Dim SubMapWidth As Single
    Dim SubMapHeight As Single
    Dim TMap As DaBoodaMap
    
'Put Index into Variable
    Num = Index
    Set TMap = Map.Item(Num)
    
'Get Sub Map Width and Height
    SubMapWidth = TMap.GetSubMapWidth
    SubMapHeight = TMap.GetSubMapHeight
    
'Get MapCounts...How many submaps are viewed on screen
Dim MapXCount As Single, MapYCount As Single
MapXCount = MapView.Right - MapView.Left
MapYCount = MapView.Bottom - MapView.Top
MapXCount = Int(MapXCount / SubMapWidth)
MapYCount = Int(MapYCount / SubMapHeight)

If (MapView.Right - MapView.Left) - (MapXCount * SubMapWidth) > 0 Then MapXCount = MapXCount + 1
If (MapView.Bottom - MapView.Top) - (MapYCount * SubMapHeight) > 0 Then MapYCount = MapYCount + 1

MapXCount = MapXCount + 1
MapYCount = MapYCount + 1

'Place Value for offset
    TopX = TMap.GetXPosition
    TopY = TMap.GetYPosition

'Find StartPosition
    If TopX > 0 Then
        X1 = Int(TopX / SubMapWidth) + 1
        XLayer = 1 - (X1)
        XStart = TopX - (SubMapWidth * X1)
    End If
    If TopY > 0 Then
        Y1 = Int(TopY / SubMapHeight) + 1
        YLayer = 1 - (Y1)
        YStart = TopY - (SubMapHeight * Y1)
    End If
    If TopX < 0 Then
        X1 = Int(Abs(TopX) / SubMapWidth)
        XLayer = 1 + X1
        XStart = TopX + (SubMapWidth * X1)
    End If
    If TopY < 0 Then
        Y1 = Int(Abs(TopY) / SubMapHeight)
        YLayer = 1 + Y1
        YStart = TopY + (SubMapHeight * Y1)
    End If
    If TopX = 0 Then
        XLayer = 1
        XStart = 0
    End If
    If TopY = 0 Then
        YLayer = 1
        YStart = 0
    End If
    
'Offset these by mapview Values
            XStart = XStart + MapView.Left
            YStart = YStart + MapView.Top

'Loop through all Maps
    For XCount = 1 To MapXCount
        For YCount = 1 To MapYCount
        
'Place position into temp variables
            X1 = XStart + (SubMapWidth * (XCount - 1))
            Y1 = YStart + (SubMapHeight * (YCount - 1))

'Set x and y into different variables
    TopX = X1: TopY = Y1
    
'SetLayer Indexes
    X1 = XLayer + (XCount - 1)
    Y1 = YLayer + (YCount - 1)
    If X1 < 1 And TMap.GetLooping = False Then GoTo DontDraw
    If X1 > TMap.GetXCount And TMap.GetLooping = False Then GoTo DontDraw
    If Y1 < 1 And TMap.GetLooping = False Then GoTo DontDraw
    If Y1 > TMap.GetYCount And TMap.GetLooping = False Then GoTo DontDraw
    
    If X1 < 1 Then X1 = TMap.GetXCount - Abs(X1)
    If X1 > TMap.GetXCount Then X1 = (X1 - TMap.GetXCount)
    If Y1 < 1 Then Y1 = TMap.GetYCount - Abs(Y1)
    If Y1 > TMap.GetYCount Then Y1 = (Y1 - TMap.GetYCount)
    
    If TMap.GetSubMapVisible(X1, Y1) = False Then GoTo DontDraw
    

'GetColors
    ColorUL = TMap.GetSubMapUpperLeftColor(X1, Y1)
    ColorUR = TMap.GetSubMapUpperRightColor(X1, Y1)
    ColorLL = TMap.GetSubMapLowerLeftColor(X1, Y1)
    ColorLR = TMap.GetSubMapLowerRightColor(X1, Y1)
    SpecularUL = TMap.GetSubMapUpperLeftSpecular(X1, Y1)
    SpecularUR = TMap.GetSubMapUpperRightSpecular(X1, Y1)
    SpecularLL = TMap.GetSubMapLowerLeftSpecular(X1, Y1)
    SpecularLR = TMap.GetSubMapLowerRightSpecular(X1, Y1)
    
'Set up Strip
            TextureStrip(0) = MakeStrip(TopX, TopY, 0, 1, ColorUL, SpecularUL, 0, 0)
            TextureStrip(1) = MakeStrip(TopX + SubMapWidth, TopY, 0, 1, ColorUR, SpecularUR, 1, 0)
            TextureStrip(2) = MakeStrip(TopX, TopY + SubMapHeight, 0, 1, ColorLL, SpecularLL, 0, 1)
            TextureStrip(3) = MakeStrip(TopX + SubMapWidth, TopY + SubMapHeight, 0, 1, ColorLR, SpecularLR, 1, 1)

'Set texture
            Direct3DDevice.SetTexture 0, Texture(TMap.GetSubMapTextureReference(X1, Y1)).GetTexture
            
'DrawStrip
            Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TextureStrip(0), Len(TextureStrip(0))

'A goto point to goto if layer isn't Drawn
DontDraw:
        Next YCount
    Next XCount

End Sub
    


'This is the Render Sub, Zorder is put in to see which ones are drawn
Private Sub RenderSprite(Order As Integer)
'Variables for Positions
    Dim sPut(4) As Point
    Dim SpriteRect As RECT
    Dim CheckRect As RECT
    
'Variables for Textures
    Dim sGet(4) As Point
    Dim TrueCount As Integer
    Dim TextureRef As Variant
    Dim TextureSize As Single
    Dim keyRef As Variant
    
'Variables for Colors
    Dim Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long
    Dim Specular1 As Long, Specular2 As Long, Specular3 As Long, Specular4 As Long
    
'Variables for Rotation
    Dim Rx As Single
    Dim Ry As Single
    Dim XOrigin As Single
    Dim YOrigin As Single
    Dim angle As Single
    
'Variable for loop
    Dim Count As Long
    Dim LX As Single, LY As Single
    Dim a As Single
    Dim mLx As Single
    Dim mLy As Single
    
'Variables for offscreen Check
    Dim Rect1 As RECT
    Dim Rect2 As RECT
    Dim Rect3 As RECT
    Dim StartX As Single, StartY As Single
    Dim MaxX As Single, MaxY As Single
    Dim TempPoint As Point
    Dim OffX As Single, OffY As Single
    Dim tSprite As DaBoodaSprite
    
'Loop through Sprites
    Count = Order
        Set tSprite = Sprite.Item(Count)

'Check to see if it is the right order and on
    If tSprite.GetVisible = False Then GoTo DontDraw
    If tSprite.GetTextureReference = "" Then GoTo DontDraw
    
    'Get StartPoint
    If tSprite.GetMapReference <> "" Then
    StartX = Map(tSprite.GetMapReference).GetXPosition
    StartY = Map(tSprite.GetMapReference).GetYPosition
    Else:
    StartX = 0
    StartY = 0
    End If
    
    'Find Initial Positions
    With SpriteRect
        .Left = tSprite.GetXPosition + StartX + MapView.Left
        .Right = .Left + tSprite.GetWidth
        .Top = tSprite.GetYPosition + StartY + MapView.Top
        .Bottom = .Top + tSprite.GetHeight
    End With

    'Set Drawn X and Y
    Sprite(Count).SetDrawnX SpriteRect.Left - MapView.Left
    Sprite(Count).SetDrawnY SpriteRect.Top - MapView.Top
    
    'Set Check Rect with Clip
    With CheckRect
        .Left = MapView.Left - MapClip
        .Top = MapView.Top - MapClip
        .Right = MapView.Right + MapClip
        .Bottom = MapView.Bottom + MapClip
    End With
    
    'Check to see if drawn
    Sprite(Count).SetDrawn False
    If RectInRect(SpriteRect, CheckRect) = True Then Sprite(Count).SetDrawn True
    If Sprite(Count).GetDrawn = False Then GoTo DontDraw

'Place Positions and rotate in variables
'Rotates them too
    XOrigin = tSprite.GetXOrigin
    YOrigin = tSprite.GetYOrigin
    angle = tSprite.GetRotationAngle
    
    XOrigin = SpriteRect.Left + XOrigin: YOrigin = SpriteRect.Top + YOrigin
    
        Rx = SpriteRect.Left - XOrigin: Ry = SpriteRect.Top - YOrigin
            sPut(1).x = XOrigin + RotateX(Rx, Ry, angle): sPut(1).y = YOrigin + RotateY(Rx, Ry, angle)
    
        Rx = SpriteRect.Right - XOrigin: Ry = SpriteRect.Top - YOrigin
            sPut(2).x = XOrigin + RotateX(Rx, Ry, angle): sPut(2).y = YOrigin + RotateY(Rx, Ry, angle)
    
        Rx = SpriteRect.Left - XOrigin: Ry = SpriteRect.Bottom - YOrigin
            sPut(3).x = XOrigin + RotateX(Rx, Ry, angle): sPut(3).y = YOrigin + RotateY(Rx, Ry, angle)
    
        Rx = SpriteRect.Right - XOrigin: Ry = SpriteRect.Bottom - YOrigin
            sPut(4).x = XOrigin + RotateX(Rx, Ry, angle): sPut(4).y = YOrigin + RotateY(Rx, Ry, angle)

'Setup Colors
    Color1 = tSprite.GetUpperLeftColor
    Color2 = tSprite.GetUpperRightColor
    Color3 = tSprite.GetLowerLeftColor
    Color4 = tSprite.GetLowerRightColor
    Specular1 = tSprite.GetUpperLeftSpecular
    Specular2 = tSprite.GetUpperRightSpecular
    Specular3 = tSprite.GetLowerLeftSpecular
    Specular4 = tSprite.GetLowerRightSpecular

'Texture Reference
    TextureRef = tSprite.GetTextureReference
    TextureSize = Texture(TextureRef).GetSize
    
'setup Textures
    sGet(1).x = TextureValue(TextureSize, tSprite.GetGetRectLeft)
    sGet(1).y = TextureValue(TextureSize, tSprite.GetGetRectTop)
    
    sGet(2).x = TextureValue(TextureSize, tSprite.GetGetRectLeft + tSprite.GetGetRectWidth)
    sGet(2).y = TextureValue(TextureSize, tSprite.GetGetRectTop)
    
    sGet(3).x = TextureValue(TextureSize, tSprite.GetGetRectLeft)
    sGet(3).y = TextureValue(TextureSize, tSprite.GetGetRectTop + tSprite.GetGetRectHeight)
    
    sGet(4).x = TextureValue(TextureSize, tSprite.GetGetRectLeft + tSprite.GetGetRectWidth)
    sGet(4).y = TextureValue(TextureSize, tSprite.GetGetRectTop + tSprite.GetGetRectHeight)

'MakeStrip........Finally.........yeah definitly
    TextureStrip(0) = MakeStrip(sPut(1).x, sPut(1).y, 0, 1, Color1, Specular1, sGet(1).x, sGet(1).y)
    TextureStrip(1) = MakeStrip(sPut(2).x, sPut(2).y, 0, 1, Color2, Specular2, sGet(2).x, sGet(2).y)
    TextureStrip(2) = MakeStrip(sPut(3).x, sPut(3).y, 0, 1, Color3, Specular3, sGet(3).x, sGet(3).y)
    TextureStrip(3) = MakeStrip(sPut(4).x, sPut(4).y, 0, 1, Color4, Specular4, sGet(4).x, sGet(4).y)
    
'Set texture
            Direct3DDevice.SetTexture 0, Texture(TextureRef).GetTexture
            
'DrawStrip
            Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TextureStrip(0), Len(TextureStrip(0))

'Exit for not Drawn
DontDraw:

End Sub

Private Sub RenderText(Order As Integer)
    Dim TText As DaBoodaText
    Set TText = TextEx.Item(Order)
    If TText.GetVisible = False Then Exit Sub
    
    With TextRect
        .Left = TText.GetXPosition
        .Top = TText.GetYPosition
        .Right = .Left + TText.GetWidth
        .Bottom = .Top + TText.GetHeight
    End With
    
    D3DX.DrawText ScreenFont, TText.GetColor, TText.GetText, TextRect, DT_TOP Or DT_LEFT
    
End Sub

Public Sub CreateFont(Name As String, Size As Single, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional StrikeThru As Boolean)
    With StandardFont
        .Name = Name
        .Size = Size
        .Bold = Bold
        .Italic = Italic
        .Underline = Underline
        .Strikethrough = StrikeThru
    End With
    
    Set ScreenFontDesc = StandardFont
    Set ScreenFont = D3DX.CreateFont(Direct3DDevice, ScreenFontDesc.hFont)
    
End Sub

Public Sub SetScreenToAutoMoveSprite(Optional Speed As Single)
    Dim tWidth As Single, tHeight As Single
    Dim tXIncrement As Single, tYIncrement As Single
    Dim tmXincrement As Single, tmYIncrement As Single
    Dim Index As Integer, OutLoop As Boolean
    
    If Speed = 0 Then Speed = 1
    tWidth = AutoMove.GetWidth
    tHeight = AutoMove.GetHeight
    tXIncrement = AutoMove.GetXIncrement
    tYIncrement = AutoMove.GetYIncrement
    tmXincrement = Map(AutoMove.GetMapReference).GetXIncrement
    tmYIncrement = Map(AutoMove.GetMapReference).GetYIncrement
    
    AutoMove.SetWidth Sprite(AutoMove.GetSpriteReference).GetWidth
    AutoMove.SetHeight Sprite(AutoMove.GetSpriteReference).GetHeight
    AutoMove.SetXIncrement Speed
    AutoMove.SetYIncrement Speed
    Map(AutoMove.GetMapReference).SetXIncrement Speed
    Map(AutoMove.GetMapReference).SetYIncrement Speed
    
        AutoMove.SetMoved False
        OutLoop = False
        
    Do While OutLoop = False
        CheckAutoMoveSprite
        UpdateMapToLimit
        If AutoMove.GetMoved = False Then OutLoop = True
        DoEvents
    Loop
    
    AutoMove.SetWidth tWidth
    AutoMove.SetHeight tHeight
    AutoMove.SetXIncrement tXIncrement
    AutoMove.SetYIncrement tYIncrement
    Map(AutoMove.GetMapReference).SetXIncrement tmXincrement
    Map(AutoMove.GetMapReference).SetYIncrement tmYIncrement
    
End Sub

Private Sub Class_Terminate()
    'Terminate All
        Set Sprite = Nothing
        Set Texture = Nothing
        Set Map = Nothing
        Set Overlay = Nothing
        Set FPS = Nothing
        Set AutoMove = Nothing
        Set Sound = Nothing
        Set KeyInput = Nothing
        Set Music = Nothing
        Set TextEx = Nothing
        Set Mouse = Nothing
        
        Set D3DX = Nothing
        Set Direct3DDevice = Nothing
        Set Direct3D = Nothing
        Set DirectX = Nothing
        
        
    If DaBoodaKeyInputOn = True Then
        kDirectInputDevice.Unacquire
        Set kDirectInputDevice = Nothing
        Set DirectInput = Nothing
    End If
    
    If DaBoodaMouseOn = True Then
        mDirectInputDevice.Unacquire
        Set mDirectInputDevice = Nothing
        Set DirectInput = Nothing
    End If
    
    If DaBoodaJoyStickOn = True Then
        JoyStickDevice.Unacquire
        Set JoyStickDevice = Nothing
        Set DirectInput = Nothing
    End If
    
    If DaBoodaMusicOn = True Then
        DSControl.Stop
        Set DSAudio = Nothing
        Set DSEvent = Nothing
        Set DSControl = Nothing
        Set DSPosition = Nothing
    End If
    
    If DaBoodaSoundOn = True Then
        Set DirectSound = Nothing
        Set DirectSoundEnum = Nothing
    End If
    
        Set ScreenFontDesc = Nothing
        Set ScreenFont = Nothing
        
End Sub
