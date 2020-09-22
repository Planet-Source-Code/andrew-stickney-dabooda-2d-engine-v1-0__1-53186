VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaBoodaFRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables for Timing Loop
    Private Frames As Single     'Frame Count
    Private FPS As Single        'Actual FPS
    Private FrameDelay As Single 'Frame Delay to set drawing to 60 FPS
    Private LastFrame As Single  'Last Frame Count to keep time concurrent
    Private Delay As Single      'This is to initiate the delay
    Private Delay2 As Single     'This is another delay
    Private FrameRate As Single  'This is Framerate to achieve
    Private Declare Function GetTickCount Lib "kernel32" () As Long 'Get TickCount

Public Sub SetFrameRate(Value As Single)
    FrameRate = Value
End Sub

Public Function GetFPS() As Single
    GetFPS = FPS
End Function

Public Sub UpdateFPS()
'Do timer and check delay
    For Delay = 0 To FrameDelay
        For Delay2 = 0 To 2000: Next Delay2
    Next Delay

    Frames = Frames + 1
    If Frames = 30 Then
        FPS = 1000 * 30 / (GetTime - LastFrame) + 1
        Frames = 0
        LastFrame = GetTime
        'check delay and alter
        If FPS > FrameRate Then FrameDelay = FrameDelay + 1
        If FPS < FrameRate Then FrameDelay = FrameDelay - 1
        If FrameDelay < 0 Then FrameDelay = 0
    End If

End Sub

Private Function GetTime() As Single
'This Function is just a simple Timer, you call on it to get the current tick count
GetTime = GetTickCount()
End Function

