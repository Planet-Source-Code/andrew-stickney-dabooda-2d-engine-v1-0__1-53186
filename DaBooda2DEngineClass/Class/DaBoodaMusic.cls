VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'This is courtesy of Jack Hoxley

Public Sub SetMusicFolder(Path As String)
    MusicFolder = Path
End Sub

Public Function GetMusicFolder() As String
    GetMusicFolder = MusicFolder
End Function

Public Sub LoadMusic(File As String, Optional Format As Integer = 1)
    'What goes on in here is magic to me(..Im too stupid to understand it)
    
        On Error GoTo NotInitialized
    'Load music
        Set DSControl = New FilgraphManager
        If Format = 1 Then
        Call DSControl.RenderFile(MusicFolder + File + ".mid")
        End If
        
        If Format = 2 Then
        Call DSControl.RenderFile(MusicFolder + File + ".mp3")
        End If
        
        Set DSAudio = DSControl
        DSAudio.Volume = 0
        DSAudio.Balance = 0

        Set DSEvent = DSControl
        Set DSPosition = DSControl
        If ObjPtr(DSPosition) Then DSPosition.Rate = 1#
        DSPosition.CurrentPosition = 0
    
    DaBoodaMusicOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaMusicOn = False
    
End Sub

Public Sub PlayMusic(Optional Repeat As Boolean, Optional Start As Long)
    'Set start
        DSPosition.CurrentPosition = Start
        
    'Play Music First
        DSControl.Run
        
    'Set up Repeats
        MusicRepeat = Repeat
        
End Sub

Public Sub StopMusic()
    'Stop Music
        DSControl.Stop
        DSPosition.CurrentPosition = 0
        
End Sub

Public Sub Pause()
    DSControl.Stop
End Sub

Public Sub UnPause()
    DSControl.Run
End Sub

Public Sub Volume(Percent As Single)
    'invert volume
    Dim Volume As Long
    Volume = Percent
    Volume = 100 - Volume
    '//Set the new volume
    If ObjPtr(DSControl) > 0 Then
        DSAudio.Volume = Volume * (-Volume * 0.25) ' makes it in the 0 to -4000 range (-4000 is almost silent)
    End If
            
End Sub

Public Sub Balance(Pan As Single)
    Dim Bal As Long
    Bal = 5000 / 100
    Bal = Pan * Bal
      
    If ObjPtr(DSControl) > 0 Then
        DSAudio.Balance = Bal
    End If

End Sub

Public Sub Tempo(Speed As Single)
    If ObjPtr(DSPosition) > 0 Then
        DSPosition.Rate = Speed
    End If
End Sub

Public Sub CheckMusicRepeat()
If DSPosition.CurrentPosition >= DSPosition.Duration Then
DSPosition.CurrentPosition = 0
PlayMusic
End If

End Sub

