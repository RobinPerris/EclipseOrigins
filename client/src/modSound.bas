Attribute VB_Name = "modSound"
Option Explicit

' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_ButtonClick As String = "Decision1.wav"

' sound/music engine
Public Performance As DirectMusicPerformance
Public Segment As DirectMusicSegment
Public Loader As DirectMusicLoader

Public DS As DirectSound

Public Const SOUND_BUFFERS = 50

Private Type BufferCaps
    Volume As Boolean
    Frequency As Boolean
    Pan As Boolean
End Type

Private Type SoundArray
    DSBuffer As DirectSoundBuffer
    DSCaps As BufferCaps
    DSSourceName As String
End Type

Private Sound(1 To SOUND_BUFFERS) As SoundArray

' Contains the current sound index.
Public SoundIndex As Long

Public Music_On As Boolean
Public Music_Playing As String

Public Sound_On As Boolean
Private SEngineRestart As Boolean

Private Const DefaultVolume As Long = 80

Public Sub InitMusic()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Loader = DX7.DirectMusicLoaderCreate
    Set Performance = DX7.DirectMusicPerformanceCreate
   
    Performance.Init Nothing, frmMain.hWnd
    Performance.SetPort -1, 80
   
    ' adjust volume 0-100
    Performance.SetMasterVolume DefaultVolume * 42 - 3000
    Performance.SetMasterAutoDownload True
   
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMusic", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitSound()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Make the DirectSound object
    Set DS = DX7.DirectSoundCreate(vbNullString)
   
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    DS.SetCooperativeLevel frmMain.hWnd, DSSCL_PRIORITY
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitSound", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function GetState(ByVal Index As Integer) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Returns the current state of the given sound
    GetState = Sound(Index).DSBuffer.GetStatus
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetState", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SoundStop(ByVal Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Stop the buffer and reset to the beginning
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSBuffer.SetCurrentPosition 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SoundStop", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub SoundLoad(ByVal File As String)
Dim DSBufferDescription As DSBUFFERDESC
Dim DSFormat As WAVEFORMATEX

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Set the sound index one higher for each sound.
    SoundIndex = SoundIndex + 1
   
    ' Reset the sound array if the array height is reached.
    If SoundIndex > UBound(Sound) Then
        SEngineRestart = True
        SoundIndex = 1
    End If
   
    ' Remove the sound if it exists (needed for re-loop).
    If SEngineRestart Then
        If GetState(SoundIndex) = DSBSTATUS_PLAYING Then
            SoundStop SoundIndex
            SoundRemove SoundIndex
        End If
    End If
   
    ' Load the sound array with the data given.
    With Sound(SoundIndex)
        .DSSourceName = File            'What is the name of the source?
        .DSCaps.Pan = True              'Is this sound to have Left and Right panning capabilities?
        .DSCaps.Volume = True           'Is this sound capable of altered volume settings?
    End With
   
    'Set the buffer description according to the data provided
    With DSBufferDescription
        If Sound(SoundIndex).DSCaps.Pan Then
            .lFlags = .lFlags Or DSBCAPS_CTRLPAN
        End If
        If Sound(SoundIndex).DSCaps.Volume Then
            .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
        End If
    End With
   
    'Set the Wave Format
    With DSFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 22050
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
   
    Set Sound(SoundIndex).DSBuffer = DS.CreateSoundBufferFromFile(App.Path & SOUND_PATH & Sound(SoundIndex).DSSourceName, DSBufferDescription, DSFormat)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SoundLoad", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SoundRemove(ByVal Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Reset all the variables in the sound array
    With Sound(Index)
        Set .DSBuffer = Nothing
        .DSCaps.Frequency = False
        .DSCaps.Pan = False
        .DSCaps.Volume = False
        .DSSourceName = vbNullString
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SoundRemove", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub SetVolume(ByVal Index As Integer, ByVal Vol As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Check to make sure that the buffer has the capability of altering its volume
    If Not Sound(Index).DSCaps.Volume Then Exit Sub

    'Alter the volume according to the Vol provided
    If Vol > 0 Then
        Sound(Index).DSBuffer.SetVolume (60 * Vol) - 6000
    Else
        Sound(Index).DSBuffer.SetVolume -6000
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetVolume", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub SetPan(ByVal Index As Integer, ByVal Pan As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Check to make sure that the buffer has the capability of altering its pan
    If Not Sound(Index).DSCaps.Pan Then Exit Sub

    'Alter the pan according to the pan provided
    Select Case Pan
        Case 0
            Sound(Index).DSBuffer.SetPan -10000
        Case 100
            Sound(Index).DSBuffer.SetPan 10000
        Case Else
            Sound(Index).DSBuffer.SetPan (100 * Pan) - 5000
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPan", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayMidi(ByVal fileName As String)
Dim Splitmusic() As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Options.Music = 0 Then Exit Sub

    Splitmusic = Split(fileName, ".", , vbTextCompare)
   
    If Performance Is Nothing Then Exit Sub
    If LenB(Trim$(fileName)) < 1 Then Exit Sub
    If UBound(Splitmusic) <> 1 Then Exit Sub
    If Splitmusic(1) <> "mid" Then Exit Sub
    If Not FileExist(App.Path & MUSIC_PATH & fileName, True) Then Exit Sub
   
    If Not Music_On Then Exit Sub
   
    If Music_Playing = fileName Then Exit Sub
   
    Set Segment = Nothing
    Set Segment = Loader.LoadSegment(App.Path & MUSIC_PATH & fileName)
   
    ' repeat midi file
    Segment.SetLoopPoints 0, 0
    Segment.SetRepeats 100
    Segment.SetStandardMidiFile
   
    Performance.PlaySegment Segment, 0, 0
   
    Music_Playing = fileName
   
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMidi", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub StopMidi()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not (Performance Is Nothing) Then Performance.Stop Segment, Nothing, 0, 0
    Music_Playing = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "StopMidi", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlaySound(ByVal File As String, Optional ByVal Volume As Long = 100, Optional ByVal Pan As Long = 50)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Options.Sound = 0 Then Exit Sub
    
    ' Check to see if DirectSound was successfully initalized.
    If Not Sound_On Or Not FileExist(App.Path & SOUND_PATH & File, True) Then Exit Sub
    
    ' make sure it's a valid file
    If Not Right$(File, 4) = ".wav" Then Exit Sub
   
    ' Loads our sound into memory.
    SoundLoad File
   
    ' Sets the volume for the sound.
    SetVolume SoundIndex, Volume
   
    ' Sets the pan for the sound.
    SetPan SoundIndex, Pan
   
    ' Play the sound.
    Sound(SoundIndex).DSBuffer.Play DSBPLAY_DEFAULT
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlaySound", "modSound", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
