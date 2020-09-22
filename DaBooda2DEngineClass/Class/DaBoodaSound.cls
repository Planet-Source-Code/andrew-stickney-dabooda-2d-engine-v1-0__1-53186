VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private DirectSoundBuffer As DirectSoundSecondaryBuffer8

Public Sub LoadSound(File As String)
    'Set up BufferDesc
    Dim DirectSoundBufferDesc As DSBUFFERDESC
    
    'Set Desc Flags
    DirectSoundBufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Set sound into buffer
    Set DirectSoundBuffer = DirectSound.CreateSoundBufferFromFile(SoundFolder + File + ".wav", DirectSoundBufferDesc)
    
End Sub

Public Sub PlaySound(Optional Looping As Boolean)
    'Set buffer to stop and reset position
        DirectSoundBuffer.Stop
        DirectSoundBuffer.SetCurrentPosition 0
    'play sound looping or not
    If Looping = True Then
        DirectSoundBuffer.Play DSBPLAY_LOOPING
    Else:
        DirectSoundBuffer.Play DSBPLAY_DEFAULT
    End If
    
End Sub

Public Sub StopSound()
    'Stop and reset position
        DirectSoundBuffer.Stop
        DirectSoundBuffer.SetCurrentPosition 0
        
End Sub

Public Sub SetVolume(Percent As Long)
    'Check Percent
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    
    'Set up Percentage to actual Value
    Dim Vol As Long
    Vol = 10000 / 100
    Vol = Percent * Vol
    Vol = -10000 + Vol
    
    'Set Volume
    DirectSoundBuffer.SetVolume Vol
    
End Sub

Public Sub SetPan(Percent As Long)
    'Check Percent
    If Percent < -100 Then Percent = -100
    If Percent > 100 Then Percent = 100
    
    'Set up Percentage to actual value
    Dim Pan As Long
    Pan = 10000 / 100
    Pan = Pan * Percent
    
    'Set Pan
    DirectSoundBuffer.SetPan Pan
    
End Sub

Public Sub SetFrequency(Hrtz As Long)
    'Check Value
    If Hrtz < 100 Then Hrtz = 100
    If Hrtz > 100000 Then Hrtz = 100000
    
    'Set Frequency
    DirectSoundBuffer.SetFrequency Hrtz
    
End Sub

Public Sub DSUpdate(Xpos As Single, Ypos As Single)
    DirectSoundBuffer.SetPan GetDirectionalPan(Xpos)
    DirectSoundBuffer.SetVolume GetVolumeDecay(Xpos, Ypos)
End Sub

Public Sub DSPlaySound(Xpos As Single, Ypos As Single)
'Set Buffer up
    DirectSoundBuffer.SetPan GetDirectionalPan(Xpos)
    DirectSoundBuffer.SetVolume GetVolumeDecay(Xpos, Ypos)

'Stop Sound
    DirectSoundBuffer.Stop
    DirectSoundBuffer.SetCurrentPosition 0
    
'PlaySound
    DirectSoundBuffer.Play DSBPLAY_DEFAULT
    
End Sub

Private Function GetDirectionalPan(Xpos As Single) As Long
Dim HalfPoint As Long
Dim IDP As Long
Dim XDef As Long
Dim FinalPan As Long
If Xpos > -1 And Xpos < DSFieldSize + 1 Then
HalfPoint = DSFieldSize / 2
IDP = 1000 / HalfPoint
XDef = HalfPoint - Xpos
FinalPan = XDef * IDP
FinalPan = FinalPan * -1
End If
If Xpos < 0 Then FinalPan = -1000
If Xpos > DSFieldSize Then FinalPan = 1000
GetDirectionalPan = FinalPan
End Function

Private Function GetVolumeDecay(Xpos As Single, Ypos As Single) As Long
Dim Xres As Long, Yres As Long
Dim DecRes As Long
Dim IDV As Long
Dim FinalRes As Long

If Xpos > -1 And Xpos < DSFieldSize + 1 And Ypos > -1 And Ypos < DSFieldSize + 1 Then
    GetVolumeDecay = DSMaxVolume
End If

'Check X
If Xpos < 0 Then Xres = Xpos * -1
If Xpos > DSFieldSize Then Xres = Xpos - DSFieldSize
If Xpos > -1 And Xpos < DSFieldSize + 1 Then Xres = 0

'checky
If Ypos < 0 Then Yres = Ypos * -1
If Ypos > DSFieldSize Then Yres = Ypos - DSFieldSize
If Ypos > -1 And Ypos < DSFieldSize + 1 Then Yres = 0

'check to see wich is larger
If Xres > Yres Then
DecRes = Xres
Else
DecRes = Yres
End If
If DecRes > DSDecay Then DecRes = DSDecay

'Find IncrementalDecayVolume
IDV = 5000 / DSDecay
DecRes = DecRes * IDV
FinalRes = DSMaxVolume - DecRes
GetVolumeDecay = FinalRes

End Function

