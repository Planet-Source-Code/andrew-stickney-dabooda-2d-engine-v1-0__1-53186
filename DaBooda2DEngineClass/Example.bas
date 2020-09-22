Attribute VB_Name = "Example"
'This is the only thing you have to create with the DaBooda Engine
'This class does everything......
Public Engine As New DaBoodaEngine

'This variable is to hold max fps
Public MaxFPS As Single

'These variables are for the background
Public BackColorDest(4) As Integer      'Stores where the color is going
Public BackColor(4) As Integer

'These Variables are for the lighted sideBars...which are just average sprites
Public SideBarDest As Integer
Public SideBarColor As Integer

'these variables are for moving the map...and keeping it in bounds
Public MapX As Single, MapY As Single

'These are for the Sprites
Public SpriteColor(200) As Integer, SpriteAni(200) As Single
Public SpriteCount As Integer, SpriteState(200) As Integer


'this sub Initializes everything that we use in this example
Public Sub Initialize()

'initialize Display...This needs to be done first......very important...
'or you wont be able to make an executable out of the proggy
'1.Inputs are the container we are going to draw in....hwnd
'2.Windowed are not windowed
'3.Width and height.......a quick note.....if not windowed must use regular_
'  Resolution....for example 640x480
'4.FrameRate to achieve
'5.Back or Clear Color
'6.Ignore...unless your smarter than me.....which is pretty easy
'7.Device type like hardware or software processing.....hardware is faster_
'  But people with old video cards might not be able to use this
'8.VSync.....Just for windowed mode....it only draws to screen when the monitor does_
'  I Dont like to use it........causes artifacts(chopping)........whew
    Engine.InitializeDisplay Main.DPic.hWnd, True, 480, 480, 200, 0, , D3DDEVTYPE_HAL, False
   
'Set up Music next for it takes awhile for it to boot up and play
'This is kind of different......music doesn't have to be initialized....just loaded
'the last value is for type........1 mid........2..mp3.....yes this plays mp3's whoohoo
    Engine.Music.SetMusicFolder App.Path + "/music/"
    Engine.Music.LoadMusic "Music", 1
    Engine.Music.Volume 60              'kind of a loud midi
    Engine.Music.PlayMusic True         'play music and loop it
'note if your file extensions dont match the predefined type then change them or
'they wont load......this goes for everything that loads

'Set up the mapview........this is where the map is referenced
'for example if you want the topleft of the map to be off screen
'negative values......if still confused play around with it
'Also the last value is the clip......this sets how far offscreen sprites go
'befor not being drawn......set lower to speed up
    Engine.SetUpMapView 0, 0, 480, 480, 1024
'another thing to note is the clip is defined whether or not a sprite point is in
'the clip.......so set the clip to your largest sprite...or it will not be drawn sometimes
'i am working on a workaround.....

'Set Max Levels...these are used for zordering
'also note that the drawing order is maps..overlays..sprites..text
    Engine.SetMaxLevel 3
    
'Set up DirectSound
'Ignore second value unless you know for sure the person has more than one sound device
    Engine.Sound.Initialize Main.DPic.hWnd

'A note about collections.....when the engine ask you to add
'You can either declare a string or a number.....
'if you use a string you can use a number later...but if you use a number...you cant use a string later
'Set up Sound Buffers.....our actual sounds
    Engine.Sound.Add "Lightning"
    Engine.Sound.Add "SpriteIn"
    Engine.Sound.Add "SpriteOut"
    Engine.Sound.Add "SpriteHit"

'Set up Sound folder......this is where your sounds are.....you don't have to do this but it makes it easier
'just note.....if you dont do this you must use the whole path to load
    Engine.Sound.SetSoundFolder App.Path + "/sfx/"  'Remember to put the / at the end
    
'Now we must load the sounds into the buffers......do this before playing are manipulation
    Engine.Sound("Lightning").LoadSound "Lightning"     'Dont add the .wav extension...engine does that for you
    Engine.Sound("SpriteIn").LoadSound "SpriteIn"
    Engine.Sound("SpriteOut").LoadSound "SpriteOut"
    Engine.Sound("SpriteHit").LoadSound "SpriteHit"
    
'Set the lightning volume down a little bit.....it is in the background
    Engine.Sound("Lightning").SetVolume 85
    
'note sound volume is set max after adding sound

'Set up Directional sound......this lets you give class sound a sprite....and it
'plays it by where its at on screen.....the closer to the left it plays more out
'the left speaker......this also sets a decay....which decreases the volume..the more
'out of screen it is.....max volume is the highest volume to achieve
    Engine.Sound.SetUpDirectionalSound 100, 240, 900
'Set up directinput
'Just put in the hwnd.....the buffer is unimportant....i think...it is...
    Engine.KeyInput.Initialize Main.hWnd
    
'Set up Keys
'Must use dik for keycodes........vb keycodes will not work
    Engine.KeyInput.Add "Up", DIK_UP
    Engine.KeyInput.Add "Down", DIK_DOWN
    Engine.KeyInput.Add "Left", DIK_LEFT
    Engine.KeyInput.Add "Right", DIK_RIGHT
    Engine.KeyInput.Add "Add", DIK_A
    Engine.KeyInput.Add "Remove", DIK_R
    Engine.KeyInput.Add "Escape", DIK_ESCAPE
    Engine.KeyInput.Add "FPSUp", DIK_NUMPADPLUS
    Engine.KeyInput.Add "FPSDown", DIK_NUMPADMINUS
    
    
'there is a autofire option with keys...but it is false when they are added
'This keeps the key from repeating
'Set move keys autofire as true
    Engine.KeyInput("Up").SetAutoFire True
    Engine.KeyInput("Down").SetAutoFire True
    Engine.KeyInput("Left").SetAutoFire True
    Engine.KeyInput("Right").SetAutoFire True

'Lets set up some textures
'This is a collection class so it acts like the sound and keyinput class
    Engine.Texture.Add "BackTex"
    Engine.Texture.Add "SpriteTex"
    Engine.Texture.Add "MapTex"
    Engine.Texture.Add "ULMapTex"
    Engine.Texture.Add "URMapTex"
    Engine.Texture.Add "LLMapTex"
    Engine.Texture.Add "LRMapTex"
    Engine.Texture.Add "BackScrollTex"
    Engine.Texture.Add "ScrollTex"
    
'Set Color Key.........this key is the color that is transparent on all
'bitmaps.......dds doesnt' use this
'also declare this before creating the textures
    Engine.SetColorKey 255, 0, 0, 255   'alpha is always 255........the color key is blue
    
'Set texture folder
    Engine.Texture.SetTextureFolder App.Path + "/graphics/"
    
'Texture must be loaded(Created)
'The first value size is important to get it right...for it is used later in sprite gets and overlay gets
'the second value is the name of the file
'Next is color format......just ignore and leave default
'Lastly is the type......1 is clear.....2 is bmp......3 is dds
    Engine.Texture("BackTex").CreateTexture 512, "backtex", , 2
    Engine.Texture("SpriteTex").CreateTexture 128, "Sprite", , 3
    Engine.Texture("MapTex").CreateTexture 128, "maptex", , 3
    Engine.Texture("ULMapTex").CreateTexture 512, , , 1
    Engine.Texture("URMapTex").CreateTexture 512, , , 1
    Engine.Texture("LLMapTex").CreateTexture 512, , , 1
    Engine.Texture("LRMapTex").CreateTexture 512, , , 1
    Engine.Texture("ULMapTex").CreateTexture 512, , , 1
    Engine.Texture("BackScrollTex").CreateTexture 512, , , 1
    Engine.Texture("ScrollTex").CreateTexture 64, "BackScroll", , 2
    
'Set up Overlay.......this is where my spiffy spinning background is seen
    Engine.Overlay.Add "BackLightning"
    With Engine.Overlay("BackLightning")
        .SetPutRect 0, 0, 512, 512  'This is where the overlay is drawn
        .QM_SetColor 255, 0, 0, 0       'a quick macro to set all corners the same color
        .SetZOrder 0                    'the order in which this is drawn
        .SetTextureReference "BackTex"  'This is the name of the texture to use
        .SetGetRect 0, 0, 512, 512      'This is the position of the graphic to grab from texture
        .SetVisible True                'Make it visible....
    End With
    
    
'Now for some fun...we will set up some maps
'I will Explain these somewhat...but its extensive
'please refer to my html help for more information
'First the BackScroll......its doesn't matter in what order
'you create your maps.....just change the zorder
    Engine.Map.Add "BackScroll"
    With Engine.Map("BackScroll")
        .SetXCount 1            'These values are how many submaps are in map....
        .SetYCount 1            'you can only have up to 64x64 submaps
        .SetLooping True        'This lets you only have one submap...and it will draw it over and over again
        .SetXIncrement 1        'this sets how much it moves when you update mapmove...it will be clearer later
        .SetYIncrement 1
        .SetZOrder 1            'this is drawn after backlightning
        .SetXReference "MainMap"    'this sets the reference to mainmap...
        .SetYReference "MainMap"    'basically this makes it move with the other map...this way you only have to move the main
                                'Also there are an x and y...so it can move on either axis
        .SetSubMapWidth 512     'Sets the submapsheights and widths...
        .SetSubMapHeight 512
        .SetSubMapTextureReference 1, 1, "BackScrollTex" 'Texture to use
        'note submap calls take references...which are a number between 1 and 64
    End With
    
'Next we will create MainMap
'I wont go into to much detail......for its about the same as the one above
    Engine.Map.Add "MainMap"
    With Engine.Map("MainMap")
        .SetXCount 2
        .SetYCount 2
        .SetLooping False
        .SetXIncrement 2
        .SetYIncrement 2
        .SetZOrder 2
        .SetSubMapWidth 512
        .SetSubMapHeight 512
        .SetSubMapTextureReference 1, 1, "ULMapTex"
        .SetSubMapTextureReference 2, 1, "URMapTex"
        .SetSubMapTextureReference 1, 2, "LLMapTex"
        .SetSubMapTextureReference 2, 2, "LRMapTex"
    End With
'Again....i couldn't possibly explain everything about maps......so i deplore you
'refer to the html help..........maps are an extensive part of this engine
'Also the maps come initialized visible with color normal

'Now for some sprites
'Again these are the most extensive part of this engine.....for they have many
'settings........i will only go over a few....again...look at the html help file

'These are the sidebar sprites....im going to initialize these with just numbers instead
'of strings
'TopSideBar
    Engine.Sprite.Add           'Notice i didnt put a name.....this is the first sprite so this is number 1
    With Engine.Sprite(1)       'i put 1 instead of a name.....just don't try to use a name later.....
        .SetGetRect 12, 12, 8, 8 'I just grabbed a solid piece from the maptex
        .SetXPosition 32        'Ignore the ResetTemp Value........its for engine use
        .SetYPosition 32
        .SetWidth 960
        .SetHeight 8
        .SetZOrder 1            'This zorder is the same as the back scroll...but map is drawn then sprite...so it will be above the map
        .SetVisible True        'Make sure it is seen
        .SetTextureReference "MapTex"   'Very Important...must have a texture or it wont be drawn
        .SetMapReference "MainMap"  'This is very important and will not be drawn without a reference...
                                    'This basically lets the engine know that the sprite x,y is offset of the map x,y
                                    'Same goes with all other objects
    End With
    
'These Next Three are similar to the one above...so im not going to go into any detail
Dim index As Integer
For index = 2 To 4
    Engine.Sprite.Add
    With Engine.Sprite(index)
        .SetGetRect 12, 12, 8, 8
        .SetZOrder 1
        .SetVisible True
        .SetTextureReference "MapTex"
        .SetMapReference "MainMap"
    End With
Next index
    With Engine.Sprite(2)
        .QM_SetPutRect 32, 32, 8, 960   'A quick macro to set position
                                        'when you set properties...just type qm and the list will scroll to all available qms
    End With
    With Engine.Sprite(3)
        .QM_SetPutRect 984, 32, 8, 960
    End With
    With Engine.Sprite(4)
        .QM_SetPutRect 32, 984, 960, 8
    End With
    
'Create a font to use for textex's......this engine cane only have one font
'For creating the font creates alot of overhead......i am look for a way around this
    Engine.CreateFont "timenewroman", 8
'Add a text String
    Engine.TextEx.Add "FPSCount"
    With Engine.TextEx("FPSCount")
        .SetColor 255, 255, 255, 255    'Yes the text can have an alpha value...good for shadows
        .SetXPosition 4                 'These next four are just to set where it is drawn
        .SetYPosition 4
        .SetWidth 100                   'Note on the width, make sure to set it high enough or your text will be clipped
        .SetHeight 16                   'Same goes with height
        .SetZOrder 3                    'Make it last to be drawn...remember the order
        .SetVisible True                'make sure to do this or it wont be seen....and you will spend an hour debugging code.....because im an idiot
    End With

'Draw maps
    DrawMaps
'Set initial maxfps
    MaxFPS = 200
'Add first Sprite
    AddSprite
'Goto main Loop
    ExampleLoop
End Sub

'This is the main Loop
Public Sub ExampleLoop()

Do
'Update FpsCounter...
Engine.TextEx("FPSCount").SetText "FPS - " + Format(Engine.FPS.GetFPS, "000.0") + " / " + Format(MaxFPS, "000") + " / " + Format(SpriteCount, "000")

CheckKeys       'CheckKeysFirst
UpdateLightning 'Do background
UpdateSideBars  'do SideBars
UpdateSprites   'Do sprites
Engine.Render   'This renders all
Loop

End Sub

'this sub Updates the Sprites
Public Sub UpdateSprites()
    Dim index As Integer, xpos As Single, ypos As Single, angle As Single
    For index = 1 To SpriteCount
    'Get sprite values to use for later
        xpos = Engine.Sprite(index + 4).GetXPosition
        ypos = Engine.Sprite(index + 4).GetYPosition
        angle = Engine.Sprite(index + 4).GetRotationAngle
    'first change colors and alpha to state
        Select Case SpriteState(index)
            Case 1      'Phasing in
                If SpriteColor(index) = 10 Then  'directional sound is based on where the sprite appears on screen...so you have to give the sprite a chance to be drawn
                    'this next statement......plays a sound according to the sprites position
                    Engine.PlayDirectionalSound 4 + index, "SpriteIn"
                End If
                Engine.Sprite(index + 4).QM_SetColor SpriteColor(index), 255, 255, 255
                SpriteColor(index) = SpriteColor(index) + 10
                If SpriteColor(index) = 260 Then
                    SpriteState(index) = 0
                    SpriteColor(index) = 250
                End If
            Case 2      'hit going up
                Engine.Sprite(index + 4).QM_SetSpecular 255, SpriteColor(index), SpriteColor(index), SpriteColor(index)
                Engine.Sprite(index + 4).QM_SetColor 255, 255, 255, 255
                SpriteColor(index) = SpriteColor(index) + 10
                If SpriteColor(index) = 210 Then
                    SpriteColor(index) = 200
                    SpriteState(index) = 3
                End If
            Case 3      'hit going down
                Engine.Sprite(index + 4).QM_SetSpecular 255, SpriteColor(index), SpriteColor(index), SpriteColor(index)
                SpriteColor(index) = SpriteColor(index) - 10
                If SpriteColor(index) = -10 Then
                    SpriteColor(index) = 0
                    SpriteState(index) = 0
                End If
            Case 4      'Phase out
                Engine.Sprite(index + 4).QM_SetColor SpriteColor(index), 255, 255, 255
                SpriteColor(index) = SpriteColor(index) - 10
                If SpriteColor(index) = -10 Then
                    SpriteState(index) = 5
                    SpriteColor(index) = 0
                End If
        End Select
    'Next update the animation
    'This nifty little call actually takes a count and determines the height and width of the texture and sets up the texture get values
        Engine.Sprite(index + 4).QM_SetSpriteGetByAnimation Int(SpriteAni(index)), 32, 32, 128, 0, 0
    'That is all there is........please experiment
        SpriteAni(index) = SpriteAni(index) + 0.25
        If SpriteAni(index) > 16.75 Then SpriteAni(index) = 1
        
    'Now increment the sprite........move it up by angle
    'this line does this........0 degrees is up.....so basically it moves
    'forward by whatever the angle is
        Engine.SpriteAngleIncrementUp index + 4, 2
        
    'Create a tempangle to see if ball hit
        Dim tempangle As Single
        tempangle = angle
    'check to see if boundaries are hit, this is a routine i actualy used in directx8 pong.....
        'LeftWall
        If xpos <= 28 And angle > 270 Then angle = (360 - angle)
        If xpos <= 28 And angle > 180 And angle < 270 Then angle = 90 + 270 - angle
        If xpos <= 28 And angle = 270 Then angle = 90
        
        'right wall
        If xpos >= 964 And angle > 0 And angle < 90 Then angle = 360 - angle
        If xpos >= 964 And angle > 90 And angle < 180 Then angle = 180 - angle + 180
        If xpos >= 964 And angle = 90 Then angle = 270
        
        'Top Wall
        If ypos <= 28 And angle > 270 Then angle = 180 + 360 - angle
        If ypos <= 28 And angle > 0 And angle < 90 Then angle = 180 - angle
        If ypos <= 28 And angle = 0 Then angle = 180
        
        'BottomWall
        If ypos >= 964 And angle < 270 And angle > 180 Then angle = 270 + 270 - angle
        If ypos >= 964 And angle < 180 And angle > 90 Then angle = 180 - angle
        If ypos >= 964 And angle = 180 Then angle = 0
        
        'Whew, probably a better way to do that......but my math sucks
        'ok now set the angle into the sprite
        Engine.Sprite(index + 4).SetRotationAngle angle
        
        'Check to see if it hit and change state accordingly
        If tempangle <> angle And SpriteState(index) <> 4 Then
            SpriteState(index) = 2      'Hit up
            SpriteColor(index) = 0
            Engine.PlayDirectionalSound index + 4, "SpriteHit"
        End If
        
        'Thats it
        'for further information like sprite collision and so on.....refer to the html help
        
    Next index
    
    'one more thing must remove the removed sprites
    'couldn't do this in the loop for it increments the other sprites down one
    For index = 1 To SpriteCount
        If SpriteState(index) = 5 Then
            Engine.Sprite.Remove index + 4      'This removes a sprite, all classes have this remove option
            SpriteCount = SpriteCount - 1
        End If
    Next index
    
End Sub
'This sub Adds a sprite
Public Sub AddSprite()
    If SpriteCount = 200 Then Exit Sub
    
    SpriteCount = SpriteCount + 1
    Engine.Sprite.Add                   'Add a spriteClass
    With Engine.Sprite(4 + SpriteCount) 'I add 4 to the sprite count, because there are already 4 classes of sprites created....the sidebars
        .QM_SetPutRect Engine.GetRandomNumber(936, 28), Engine.GetRandomNumber(936, 28), 32, 32
        .QM_SetColor 0, 255, 255, 255
        .SetXOrigin 16
        .SetYOrigin 16
        .SetRotationAngle Engine.GetRandomNumber(360)
        .SetVisible True
        .SetMapReference "MainMap"
        .SetTextureReference "SpriteTex"
        .SetZOrder 3                                    'place it to be drawn after mainmap
    End With
    
    SpriteState(SpriteCount) = 1        'Phasing in
    SpriteColor(SpriteCount) = 0
    SpriteAni(SpriteCount) = 2
    
    
End Sub

'This sub removes a sprite.....or rather puts the removal into motion
Public Sub RemoveSprite()
    If SpriteCount = 1 Then Exit Sub
    
    If SpriteState(SpriteCount) <> 4 Then
    SpriteState(SpriteCount) = 4
    SpriteColor(SpriteCount) = 250
    'play the sound
    Engine.PlayDirectionalSound SpriteCount + 4, "SpriteOut"
    End If
    
    
    
End Sub

'THis sub ChecksKeys and Responds accordingly
Public Sub CheckKeys()
'Before you can check any returnkeystates...you have to getkeys...basically poll them
    Engine.KeyInput.GetKeyState
'Check Escape key and unload if true
    If Engine.KeyInput("Escape").ReturnKeyState = True Then
        Set Engine = Nothing        'That is all it takes to terminate everything
        End
    End If
    
    'Check to see if maxfps is changed
    If Engine.KeyInput("FPSDown").ReturnKeyState = True Then
        MaxFPS = MaxFPS - 10
    End If
    If Engine.KeyInput("FPSUp").ReturnKeyState = True Then
        MaxFPS = MaxFPS + 10
    End If
    If MaxFPS > 990 Then MaxFPS = 990
    If MaxFPS < 10 Then MaxFPS = 10
    Engine.FPS.SetFrameRate MaxFPS
    
    'Check And see if map moves
    'Whats great about this is that the scrollmap moves automatically with mainmap
    'and the movemapdown...and so forth makes it real easy to move the maps
    'just remember.......to move left is to actually move the map right
    'its confusing at first but makes sense later
    If Engine.KeyInput("Up").ReturnKeyState = True And MapY <> 0 Then
        MapY = MapY - 2
        Engine.MoveMapDown "MainMap"
    End If
    If Engine.KeyInput("Down").ReturnKeyState = True And MapY <> 1024 - 480 Then
        MapY = MapY + 2
        Engine.MoveMapUp "MainMap"
    End If
    If Engine.KeyInput("Left").ReturnKeyState = True And MapX <> 0 Then
        MapX = MapX - 2
        Engine.MoveMapRight "MainMap"
    End If
    If Engine.KeyInput("Right").ReturnKeyState = True And MapX <> 1024 - 480 Then
        MapX = MapX + 2
        Engine.MoveMapLeft "MainMap"
    End If

    If Engine.KeyInput("Remove").ReturnKeyState = True Then RemoveSprite
    If Engine.KeyInput("Add").ReturnKeyState = True Then AddSprite
    
End Sub
'This sub Updates the backGround and causes lightning strikes
Public Sub UpdateLightning()
    Dim RandNum As Integer, index As Integer
'This next line is pretty nifty.....daboodaengine will return a random number for you
    RandNum = Engine.GetRandomNumber(5000, 1)    'The offset is just an additive...because random always returns a 0 to num-1 value
    
'Check to see if lightning strikes
    If RandNum > 4980 Then        'Lightning struck
        For index = 1 To 4
        BackColorDest(index) = 50 + (Engine.GetRandomNumber(40, 1) * 5)
        Next index
        Engine.Sound("Lightning").PlaySound False   'Plays lightning sound once
    End If
    
'Update color to destination
    For index = 1 To 4
    If BackColor(index) < BackColorDest(index) Then BackColor(index) = BackColor(index) + 5
    If BackColor(index) > BackColorDest(index) Then BackColor(index) = BackColor(index) - 5
    If BackColor(index) = BackColorDest(index) Then BackColorDest(index) = 0
    Next index
    
    Engine.Overlay("BackLightning").SetUpperLeftColor 255, BackColor(1), BackColor(1), BackColor(1)
    Engine.Overlay("BackLightning").SetUpperRightColor 255, BackColor(2), BackColor(2), BackColor(2)
    Engine.Overlay("BackLightning").SetLowerLeftColor 255, BackColor(3), BackColor(3), BackColor(3)
    Engine.Overlay("BackLightning").SetLowerRightColor 255, BackColor(4), BackColor(4), BackColor(4)
    Engine.Overlay("BackLightning").SetUpperLeftSpecular 255, BackColor(1), BackColor(1), BackColor(1)
    Engine.Overlay("BackLightning").SetUpperRightSpecular 255, BackColor(2), BackColor(2), BackColor(2)
    Engine.Overlay("BackLightning").SetLowerLeftSpecular 255, BackColor(3), BackColor(3), BackColor(3)
    Engine.Overlay("BackLightning").SetLowerRightSpecular 255, BackColor(4), BackColor(4), BackColor(4)
    
End Sub

'This Sub updates the sidebars
Public Sub UpdateSideBars()
    Dim index As Integer
    For index = 1 To 4          'Set values to all 4 sidebar sprites
        With Engine.Sprite(index)
            .QM_SetColor SideBarColor, 255, 0, 255  'This sets the color to a purple color and alpha to the value
                                                    'This causes a fading in and out look
            .QM_SetSpecular 255, SideBarColor, SideBarColor, SideBarColor 'Specular is pretty neat......it takes the color all the way
                                                                          'up to white......the alpha doesnt really effect this......unless value of 0
        End With
    Next index
    
    If SideBarColor = 0 And SideBarDest = 0 Then
        SideBarDest = 255
    End If
    If SideBarColor = 255 And SideBarDest = 255 Then
        SideBarDest = 0
    End If
    If SideBarColor > SideBarDest Then SideBarColor = SideBarColor - 5
    If SideBarColor < SideBarDest Then SideBarColor = SideBarColor + 5
    
End Sub

'These next subs are written just so i can make up some maps....
'I wouldn't pay to much attention to these.....hence there at the end
'the only important thing to note is the texturecopyrects.......this is a great tool
'for copying from one layer to the next
'i would suggest all maps be drawn during runtime...is saves on space
Public Sub DrawMaps()

Dim EmString As String

EmString = ""

'First map
EmString = EmString + "AFFFFFFFFFFFFFFF"
EmString = EmString + "EHLLLLLLLLLLLLLL"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
DrawtoMap EmString, "ULMapTex"

EmString = ""
EmString = EmString + "FFFFFFFFFFFFFFFB"
EmString = EmString + "LLLLLLLLLLLLLLIE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
DrawtoMap EmString, "URMapTex"

EmString = ""
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EMGGGGGGGGGGGGGG"
EmString = EmString + "EKNNNNNNNNNNNNNN"
EmString = EmString + "CFFFFFFFFFFFFFFF"
DrawtoMap EmString, "LLMapTex"

EmString = ""
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "GGGGGGGGGGGGGGOE"
EmString = EmString + "NNNNNNNNNNNNNNJE"
EmString = EmString + "FFFFFFFFFFFFFFFD"
DrawtoMap EmString, "LRMapTex"

'Create the scrollmap
Dim XIndex As Single, YIndex As Single
    For XIndex = 0 To 7
        For YIndex = 0 To 7
            Engine.TextureCopyRects "ScrollTex", "BackScrollTex", 0, 0, XIndex * 64, YIndex * 64, 64, 64
        Next YIndex
    Next XIndex
    
End Sub

Public Sub DrawtoMap(Template As String, sTexture As String)
    Dim xpos As Single, ypos As Single
    Dim XGet As Single, YGet As Single
    Dim Count As Integer
    Dim index As Integer
    
    For Count = 1 To 256
    index = Asc(Mid$(Template, Count)) - 65
    YGet = Int(index / 4)
    XGet = index - (YGet * 4)
    YGet = YGet * 32
    XGet = XGet * 32
    
    'This is an important tool to copy from one texture to another
    'for dest and source enter the key or number representing the textures
    Engine.TextureCopyRects "MapTex", sTexture, XGet, YGet, xpos, ypos, 32, 32
        
    xpos = xpos + 32
    If xpos > 480 Then
        xpos = 0
        ypos = ypos + 32
    End If
    
    Next Count
    
    'whew dont worry im working on a multi purpose map maker with loader so
    'you can load premade maps......its in the works......
End Sub
