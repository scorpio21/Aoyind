VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NumSoundBuffers        As Byte = 30

Private Const MAX_DISTANCE_TO_SOURCE As Integer = 200

Private Const DELTA_FQ               As Integer = 75

Private Const MIN_MUSIC_VOLUME       As Integer = -2500

Private Const INCREASE_MUSIC_VOLUME  As Integer = 20

Private Const MIN_SOUND_VOLUME       As Integer = -2000

Private Const INCREASE_SOUND_VOLUME  As Integer = 20

Private Const VOLUME_DISTANCE_MOD    As Long = 5

Private Const PAN_MOD                As Long = 600

Private Type SoundBuffer

    FileName As String
    looping As Boolean
    x As Integer
    y As Integer
    normalFq As Long
    buffer As DirectSoundSecondaryBuffer8

End Type

Public Enum LoopStyle

    Default = 0
    Disabled = 1
    Enabled = 2

End Enum

Public Enum MusicTypes

    Mp3
    Midi

End Enum

Dim inited                          As Boolean

Dim DirectSound                     As DirectSound8

Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer

Dim SoundPath                       As String

Dim SndVolume                       As Long

Dim Performance                     As DirectMusicPerformance8

Dim Segment                         As DirectMusicSegment8

Dim Loader                          As DirectMusicLoader8

Dim SegState                        As DirectMusicSegmentState8

Dim Music                           As Boolean

Dim MusicPath                       As String

Dim MusicMp3Path                    As String

Dim CurrentMusicFile                As String

Dim CurrentMusicFileType            As MusicTypes

Dim mActivated                      As Boolean

Dim sActivated                      As Boolean

Dim lastPosX                        As Integer

Dim lastPosY                        As Integer

Dim sEffectsActivated               As Boolean

'**** Used By MP3 Playing. *****
Dim IMC                             As IMediaControl

Dim IBA                             As IBasicAudio

Dim IME                             As IMediaEvent

Dim IMPos                           As IMediaPosition

Private Type WAVETYPE

    strHead As String * 12
    strFormatID As String * 4
    lngChunkSize As Long
    intFormat As Integer
    intChannels As Integer
    lngSamplesPerSec As Long
    lngAvgBytesPerSec As Long
    intBlockAlign As Integer
    intBitsPerSample As Integer

End Type

'**** AMBIENT ********
Dim ambient_buffer          As DirectSoundSecondaryBuffer8

Dim ambient_rain_buffer     As DirectSoundSecondaryBuffer8

Dim ambient_rain_int_buffer As DirectSoundSecondaryBuffer8

Dim ambient_fire_buffer     As DirectSoundSecondaryBuffer8

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                       ByRef Source As Any, _
                                       ByVal Length As Long)

''
'Destructor. Releases all created objects assuring no memory-leaks.

Private Sub Class_Terminate()
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/31/09
    'Destroy's all DirectMusic and DirectSound objects
    '12/31/09 - Pato: Use StopWave method to prevent errors.
    '***************************************************
    On Local Error Resume Next

    Dim LoopC As Long
   
    'Stop every channel being used and destroy the buffer
    For LoopC = 1 To NumSoundBuffers

        If LenB(DSBuffers(LoopC).FileName) <> 0 Then
            Call DSBuffers(LoopC).buffer.Stop

        End If

        Set DSBuffers(LoopC).buffer = Nothing
    Next LoopC
   
    'Stop every channel being used and destroy the buffer
    Call StopWave
    
    'Stop music
    If Not Segment Is Nothing Then Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    
    Call Performance.CloseDown
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
    
    inited = False

End Sub

''
'Initializes the clsAudio object.
'
'   @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
'   @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
'   @param    sound_path  Directory in wich to search for sound files.
'   @param    music_path  Directory in wich to search for music files.
'   @param    music_mp3_path  Directory in wich to search for music mp3 files.
'
'   @return   True if no errors occured, false otherwise.

Public Function Initialize(ByRef dX As DirectX8, _
                           ByVal hwnd As Long, _
                           ByVal sound_path As String, _
                           ByVal music_path As String, _
                           ByVal music_mp3_path As String) As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/01/20
    'Creates all DS and DM objects
    'Se agrego soporte para la carpeta de mp3 (Recox)
    '***************************************************
    On Error GoTo ErrHandler
    
    '   Prevent multiple creations
    If inited Then Exit Function
    
    'Validate DX object
    If dX Is Nothing Then GoTo ErrHandler
    
    'Initialize DirectSound
    Set DirectSound = dX.DirectSoundCreate("")
    
    Call DirectSound.SetCooperativeLevel(hwnd, DSSCL_PRIORITY)
    
    'Set sound path
    SoundPath = sound_path
    
    'Set default volume
    SndVolume = 0
    
    'Initialize DirectMusic
    Music = InitializedirectMusic(dX, hwnd)
    
    'Set music path
    MusicPath = music_path
    MusicMp3Path = music_mp3_path
    
    mActivated = True
    sActivated = True
    
    CurrentMusicFile = vbNullString
    
    'Set tempo to 0 and volume of music
    MusicTempo = 0
    Call Performance.SetMasterVolume(0)
    
    inited = True
    
    Initialize = True
    Exit Function

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.Initialize")

End Function

''
'   Creates and configures all DirectMusic objects.
'
'   @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
'   @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
'
'   @return   True if no errors occured, false otherwise.

Private Function InitializedirectMusic(ByRef dX_ As DirectX8, _
                                       ByVal hwnd As Long) As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Initialize DM. Returns false if music isn't supported by HW
    '***************************************************
    On Error GoTo ErrHandler

    Dim mus_Params As DMUS_AUDIOPARAMS

    Set Loader = dX.DirectMusicLoaderCreate()
    
    Set Performance = dX_.DirectMusicPerformanceCreate()
    Call Performance.InitAudio(hwnd, DMUS_AUDIOF_ALL, mus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
    
    'Set to port default, 1 channel group is enough for MIDIs
    Performance.SetMasterTempo 1
    Performance.SetMasterVolume 1
    
    InitializedirectMusic = True
    Exit Function

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.InitializedirectMusic")

End Function

''
'   Plays a new WAVE file.
'
'   @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'   @param    srcX        The X position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
'   @param    srcY        The Y position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
'   @param    LoopSound   The loop state for the sound.
'
'   @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Function PlayWave(ByVal FileName As String, _
                         Optional ByVal srcX As Integer = 0, _
                         Optional ByVal srcY As Integer = 0, _
                         Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/07/08
    'Last Modified by: Franco Zeoli (Noich)
    'Added support for 3D sound effects.
    '***************************************************
    On Error GoTo ErrHandler

    Dim bufferIndex As Long
   
    If Not SoundActivated Then Exit Function
    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(FileName, LoopSound)

    If bufferIndex = 0 Then Exit Function   'If an error ocurred abort
    
    With DSBuffers(bufferIndex)

        If .looping Then
            .buffer.Play DSBPLAY_LOOPING
        Else
            .buffer.Play DSBPLAY_DEFAULT

        End If
        
        'Apply volume
        Call .buffer.SetVolume(SndVolume)
        'Call .Buffer.SetVolume(Calculate_Volume(srcX, srcY))
        'Call .Buffer.SetPan(Calculate_Pan(srcX, srcY))
        
        'Store position
        .x = srcX
        .y = srcY
        
        'Get standard frequency
        .normalFq = .buffer.GetFrequency()

    End With
    
    PlayWave = bufferIndex
    
    If mOpciones.SoundEffects = True Then
        If srcX <> 0 And srcY <> 0 Then
            Call Update3DSound(bufferIndex, 0, 0)

        End If

    End If

    Exit Function

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.PlayWave")

End Function

Private Sub ResetSoundValues()

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 03/04/10
    'Reset the frequency, volume and pan.
    '***************************************************
    Dim I As Long
    
    For I = 1 To NumSoundBuffers

        If Not DSBuffers(I).buffer Is Nothing Then

            With DSBuffers(I)
                Call .buffer.SetFrequency(.normalFq)
                Call .buffer.SetVolume(SndVolume)
                Call .buffer.SetPan(0)

            End With

        End If

    Next I

End Sub

Private Sub Update3DSound(ByVal bufferIndex As Long, _
                          ByVal deltaX As Integer, _
                          ByVal deltaY As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/09/09
    'Last Modified by: Marco Vanotti (MarKoxX)
    'Fixed Dopplers formula
    'Refreshes the sound's arguments.
    'Fixed the sndVolume formula
    'Fixed 3D Volume
    '***************************************************
    Dim linearDistanceOld As Single

    Dim linearDistanceNew As Single

    Dim deltaDistance     As Single

    Dim distanceXOld      As Long

    Dim distanceYOld      As Long

    Dim distanceXNew      As Long

    Dim distanceYNew      As Long

    Dim perceivedFq       As Long
    
    With DSBuffers(bufferIndex)
        distanceXOld = .x - lastPosX
        distanceYOld = .y - lastPosY
        
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
        
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
        
        deltaDistance = linearDistanceNew - linearDistanceOld
        
        'Update frequency emulating Doppler effect
        perceivedFq = .normalFq * (1 - deltaDistance / DELTA_FQ)

        If perceivedFq < DSBFREQUENCY_MIN Then perceivedFq = DSBFREQUENCY_MIN
        
        Call .buffer.SetFrequency(perceivedFq)
        
        'Set volumen amortiguation according to distance
        If linearDistanceNew > MAX_DISTANCE_TO_SOURCE Then
            Call .buffer.SetVolume(SndVolume + (DSBVOLUME_MIN - SndVolume))
        Else
            Call .buffer.SetVolume(SndVolume + (linearDistanceNew / MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - SndVolume))

        End If
        
        'Prevent division by zero
        If linearDistanceNew = 0 Then linearDistanceNew = 1
        
        'Set panning according to relative position of the source to the listener
        Call .buffer.SetPan((distanceXNew / linearDistanceNew) * 2000)

    End With

End Sub

''
'   Updates 3D sounds based on the movement of the listener.
'
'   @param    x   The new coordinate of the listener in the x axis (consider 0 as the furthest point to the left).
'   @param    y   The new coordinate of the listener in the y axis (consider 0 as the furthest point to the front).
'
'   @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Sub MoveListener(ByVal x As Integer, ByVal y As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/07/08
    'Last Modification by: Franco Zeoli (Noich)
    'Moves the position of the listener updating all 3D sounds.
    '***************************************************
    Dim I      As Long

    Dim deltaX As Integer

    Dim deltaY As Integer
    
    If sEffectsActivated = False Then Exit Sub
    
    deltaX = x - lastPosX
    deltaY = y - lastPosY
    
    For I = 1 To NumSoundBuffers

        If Not DSBuffers(I).buffer Is Nothing Then
            If DSBuffers(I).buffer.GetStatus And DSBSTATUS_PLAYING Then
                If DSBuffers(I).x <> 0 And DSBuffers(I).y <> 0 Then
                    Call Update3DSound(I, deltaX, deltaY)

                End If

            End If

        End If

    Next I
    
    lastPosX = x
    lastPosY = y

End Sub

''
'   Loads a new WAVE file.
'
'   @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'   @param    looping     The loop state for the sound.
'
'   @return   The buffer index in which the sound is playing, or 0 in case of error.

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) As Long

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/28/09
    'Last Modification by: Marco Vanotti (MarKoxX)
    'Loads a new wave into memory (if it's not already loaded) and returns
    'the buffer index
    '***************************************************
    On Error GoTo ErrHandler

    Dim bufferDesc As DSBUFFERDESC

    Dim waveFormat As WAVEFORMATEX

    Dim I          As Long
    
    FileName = UCase$(FileName)
    
    'If Not FileExist(SoundPath & FileName, vbArchive) Then _
    '    Exit Function
    
    'Check if the buffer is in memory and not playing
    For I = 1 To NumSoundBuffers

        If DSBuffers(I).FileName = FileName Then
            If DSBuffers(I).buffer.GetStatus = 0 Or DSBuffers(I).buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                LoadWave = I
                
                'Set looping if needed
                If looping <> LoopStyle.Default Then DSBuffers(I).looping = (looping = LoopStyle.Enabled)
                
                'Reset default sound frecuency to avoid issues.
                Call DSBuffers(I).buffer.SetFrequency(DSBuffers(I).normalFq)
                
                Exit Function

            End If

        End If

    Next I
    
    'Not in memory, search for an empty buffer
    For I = 1 To NumSoundBuffers

        If DSBuffers(I).buffer Is Nothing Then
            Exit For

        End If

    Next I
    
    'No empty buffers?? Search for one not being played and use that slot
    If I > NumSoundBuffers Then

        For I = 1 To NumSoundBuffers

            If DSBuffers(I).buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(I).buffer.GetStatus() = 0 Then
                Exit For

            End If

        Next I

    End If
    
    'Still no free buffers?? Get the first non-looping sound
    If I > NumSoundBuffers Then

        For I = 1 To NumSoundBuffers

            If Not (DSBuffers(I).buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(I).buffer.Stop
                Exit For

            End If

        Next I

    End If
    
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If I > NumSoundBuffers Then
        If Not looping Then Exit Function
        I = 1

    End If
    
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    With DSBuffers(I)
        Set .buffer = Nothing   'Get rid of any previous data
        
        .FileName = FileName
        .looping = (looping = LoopStyle.Enabled)
        
        Set .buffer = DirectSound.CreateSoundBufferFromFile(PathWav & "\" & FileName & ".wav", bufferDesc)

        'Set .Buffer = DirectSound.CreateSoundBuffer(udtBufferDesc)
        'Call CreateBufferFromFile_Ex(FileName, i)
    End With
    
    LoadWave = I
    Exit Function

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.LoadWave")

End Function

''
'   Stops a given sound or all of them.
'
'   @param    bufferIndex     The index of the sound file to stop playing. If 0, all sounds are stopped.

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Stops the sound in the given buffer index, or every sound if
    'index is 0
    '***************************************************
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
        If DSBuffers(bufferIndex).buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).buffer.Stop

        End If

    ElseIf bufferIndex = 0 Then

        Dim I As Long

        For I = 1 To NumSoundBuffers

            If Not DSBuffers(I).buffer Is Nothing Then
                If DSBuffers(I).buffer.GetStatus() <> 0 And DSBuffers(I).buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(I).buffer.Stop

                End If

            End If

        Next I

    End If

End Sub

''
'   Retrieves wether there are sounds currentyl playing or not.
'
'   @return   true if there is at least one sound playing, false otherwise.

Public Property Get PlayingSound() As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns true if there is at least one sound playing
    '***************************************************
    Dim I As Long
    
    For I = 1 To NumSoundBuffers

        If Not DSBuffers(I).buffer Is Nothing Then
            If DSBuffers(I).buffer.GetStatus() And DSBSTATUS_PLAYING Then
                PlayingSound = True
                Exit Property

            End If

        End If

    Next I

End Property

''
'   Sets the volume of sound.
'
'   @param    Volume  New value to be set, in percentage (0 to 100).

Public Property Let SoundVolume(ByVal Volume As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the sound's volume as a percentage
    '***************************************************
    Dim I As Long
    
    If Volume < 0 Or Volume > 100 Then Exit Property
    
    'Take percentage to actual value
    SndVolume = Volume * INCREASE_SOUND_VOLUME + MIN_SOUND_VOLUME
    
    For I = 1 To NumSoundBuffers

        If Not DSBuffers(I).buffer Is Nothing Then
            If DSBuffers(I).buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(I).buffer.SetVolume(SndVolume)

            End If

        End If

    Next I

End Property

''
'   Gets the volume of sound.
'
'   @return   The current sound volume in percentage.

Public Property Get SoundVolume() As Byte
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the music's volume as a percentage
    '***************************************************
    SoundVolume = CByte((SndVolume - MIN_SOUND_VOLUME) / INCREASE_SOUND_VOLUME)

End Property

''
'   Plays a new MIDI file.
'
'   @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'   @param    Loops   Number of times to loop sound. For eternal loops set it to -1.
'
'   @return   True if file was properly loaded and started playing, False otherwise.

Public Function PlayMIDI(Optional ByVal File As String = vbNullString, _
                         Optional ByVal Loops As Long = -1) As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Loads a midi and start to play it
    '***************************************************
    On Error GoTo ErrHandler

    If LenB(File) > 0 Then CurrentMusicFile = File
    CurrentMusicFileType = MusicTypes.Midi
    
    If Not MusicActivated Then Exit Function
    
    'If Not Music Then Exit Function
    
    Call StopMidi

    'In case we are listening MP3 we first stop it.
    Call StopMp3
    
    If LenB(File) > 0 Then
        If Not LoadMidi(File) Then Exit Function
    Else

        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function

    End If
    
    'Play it
    Call Segment.SetRepeats(Loops)

    Set SegState = Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)
   
    PlayMIDI = True
    Exit Function

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.PlayMIDI")

End Function

''
'   Loads a new MIDI file.
'
'   @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'
'   @return   True if file was properly loaded and started playing, False otherwise.

Private Function LoadMidi(ByVal File As String) As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Loads a midi to memory
    '***************************************************
    On Error GoTo ErrHandler

    If Not FileExist(MusicPath & File, vbArchive) Then Exit Function
    
    Call StopMidi
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(MusicPath & File)
    
    If Segment Is Nothing Then
        Exit Function

    End If
    
    Call Segment.SetStandardMidiFile
    
    LoadMidi = True
    Exit Function

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.LoadMidi")

End Function

''
'   Stops playing the currently loaded MIDI file.

Public Sub StopMidi()

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Stops playing a midi
    '***************************************************
    On Error GoTo ErrHandler

    If PlayingMusic Then
        Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)

    End If

    Exit Sub

ErrHandler:
    Call LogError(err.Number, err.Description, "clsAudio.StopMidi")

End Sub

''
'   Sets the music's tempo.
'
'   @param    Tempo   New Tempo value. Must be in the range of 0.25 to 2.0

Public Property Let MusicTempo(ByVal Tempo As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the music's tempo
    '***************************************************
    If Not Music Then Exit Property
    
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
    
    Call Performance.SetMasterTempo(Tempo)

End Property

''
'   Sets the music's volume.
'
'   @param    Volume  Percentage to wich to set the music volume.

Public Property Let MusicVolume(ByVal Volume As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Sets the music's volume as a percentage
    '***************************************************
    If Not Music Then Exit Property
    
    If Volume < 0 Or Volume > 100 Then Exit Property
    
    '   Volume ranges from -10000 to 10000
    Call Performance.SetMasterVolume(Volume * INCREASE_MUSIC_VOLUME + MIN_MUSIC_VOLUME)

    If Not IBA Is Nothing Then
        IBA.Volume = Volume * INCREASE_SOUND_VOLUME + MIN_SOUND_VOLUME

    End If

End Property

''
'   Gets the music's tempo.
'
'   @return   Current music'tempo.

Public Property Get MusicTempo() As Single

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns the music's tempo
    '***************************************************
    If Not Music Then Exit Property
    
    MusicTempo = Performance.GetMasterTempo()

End Property

''
'   Gets the music's volume.
'
'   @return   Current music's volume as a percentage.

Public Property Get MusicVolume() As Long

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns the music's volume as a percentage
    '***************************************************
    If Not Music Then Exit Property
    
    MusicVolume = (Performance.GetMasterVolume() - MIN_MUSIC_VOLUME) / INCREASE_MUSIC_VOLUME

End Property

''
'   Checks wether there is music currently playing.
'
'   @return   True if there is a MIDI playing, false otherwise.

Public Property Get PlayingMusic() As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns True if a midi is being played
    '***************************************************
    If Not Music Or Segment Is Nothing Then Exit Property
    
    PlayingMusic = Performance.IsPlaying(Segment, SegState)

End Property

''
'   Retrieves the music's length.
'
'   @return   Current music's length.

Public Property Get MusicLenght() As Long

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 16/03/04
    'Returns the lenght of the current midi
    '***************************************************
    If Not Music Then Exit Property
    
    MusicLenght = Segment.GetLength()

End Property

''
'   Retrieves wether music is activated or not.
'
'   @return   True if music is activated, false otherwise.

Public Property Get MusicActivated() As Boolean
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Retrieves wether the music is activated or not
    '***************************************************
    MusicActivated = mActivated

End Property

''
'   Sets wether music is activated or not.
'
'   @param    nValue  True if music should be activated, false otherwise.

Public Property Let MusicActivated(ByVal nValue As Boolean)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Sets wether the music is activated or not
    '***************************************************
    If mActivated = nValue Then Exit Property
    
    mActivated = nValue
    
    If nValue Then

        Select Case CurrentMusicFileType

            Case MusicTypes.Midi
                mActivated = PlayMIDI(CurrentMusicFile)

            Case MusicTypes.Mp3
                mActivated = PlayMp3(CurrentMusicFile)

        End Select

    Else

        Select Case CurrentMusicFileType

            Case MusicTypes.Midi
                Call StopMidi

            Case MusicTypes.Mp3
                Call StopMp3

        End Select

    End If

End Property

''
'   Retrieves wether sound is activated or not.
'
'   @return   True if sound is activated, false otherwise.

Public Property Get SoundActivated() As Boolean
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Retrieves wether the sound is activated or not
    '***************************************************
    SoundActivated = sActivated

End Property

''
'   Sets wether sound is activated or not.
'
'   @param    nValue  True if sound should be activated, false otherwise.

Public Property Let SoundActivated(ByVal nValue As Boolean)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/03/07
    'Sets wether the sound is activated or not
    '***************************************************
    If sActivated = nValue Then Exit Property
    
    sActivated = nValue
    
    If Not nValue Then
        Call StopWave

    End If

End Property

''
'   Sets wether sound effects are activated or not.
'
'   @param nValue True if sound effects should be activated, false otherwise.

Public Property Let SoundEffectsActivated(ByVal nValue As Boolean)
    '***************************************************
    'Author: Franco Zeoli (Noich)
    'Last Modification: 03/04/10
    'Sets wether the sound effects are activated or not
    '03/04/10 - Pato: If sound effects are disable, then reset the frequency, volume and pan.
    '***************************************************
    sEffectsActivated = nValue
    
    If Not nValue Then
        Call ResetSoundValues

    End If

End Property

''
'   Retrieves wether sound effects are activated or not.
'
'   @return   True if sound effects are activated, false otherwise.

Public Property Get SoundEffectsActivated() As Boolean
    '***************************************************
    'Author: Franco Zeoli (Noich)
    'Last Modification: 28/07/08
    'Retrieves wether the sound effects are activated or not
    '***************************************************
    SoundEffectsActivated = sEffectsActivated

End Property

''
' Plays a new MP3 file.
'
' @param    FilePath    The path to the file to be loaded. If the path is null then play the loaded MP3.
'
' @return   True    If file was properly loaded and started playing, False otherwise.

Public Function PlayMp3(Optional ByVal File As String = vbNullString) As Boolean

    '**************************************
    '* Created by: Fredrik Alexandersson  *
    '* Modify Date: 20/5/2003             *
    '**************************************
    On Error GoTo ErrorHandl

    If LenB(File) > 0 Then CurrentMusicFile = File
    CurrentMusicFileType = MusicTypes.Mp3
    
    If Not MusicActivated Then Exit Function
    
    If Not Music Then Exit Function
    
    Call StopMp3
    'In case we are listening to a MIDI we first stop it.
    Call StopMidi

    If LenB(File) > 0 Then
        If Not LoadMp3(File, MusicVolume * INCREASE_SOUND_VOLUME + MIN_SOUND_VOLUME) Then Exit Function
    Else

        If IMC Is Nothing Then Exit Function

    End If

    Call IMC.Run
    
    PlayMp3 = True
    Exit Function

ErrorHandl:
    PlayMp3 = False
    Call LogError(err.Number, err.Description, "clsAudio.PlayMp3")

End Function

''
' Loads a new MP3 file.
'
' @param    FilePath    The path to the file to be loaded.
' @return   True    If file was properly loaded and started playing, False otherwise.

Public Function LoadMp3(ByVal FilePath As String, _
                        Optional ByVal Volume As Long = 0, _
                        Optional ByVal Balance As Long = 0) As Boolean '**** Loads a MP3 *****

    '**************************************
    '* Created by: Fredrik Alexandersson  *
    '* Modify Date: 20/5/2003             *
    '**************************************
    On Error GoTo ErrorHandl
    
    If Not FileExist(MusicMp3Path & FilePath, vbArchive) Then Exit Function
    
    If Not MP3Empty() = True Then GoTo ErrorHandl
    
    Set IMC = New FilgraphManager
    Call IMC.RenderFile(MusicMp3Path & FilePath)
    
    Set IBA = IMC
    
    'If SoundVolume > 0 Then SoundVolume = 0
    If Volume < -4000 Then Volume = -4000
    
    IBA.Volume = Volume
    
    If Balance > 5000 Then Balance = 5000
    If Balance < -5000 Then Balance = -5000
    
    IBA.Balance = Balance
    
    Set IME = IMC
    
    Set IMPos = IMC

    If ObjPtr(IMPos) Then IMPos.Rate = 1#
    IMPos.CurrentPosition = 0
    
    LoadMp3 = True
    Exit Function

ErrorHandl:
    LoadMp3 = False
    Call LogError(err.Number, err.Description, "clsAudio.LoadMp3")

End Function

''
' Stop a the current MP3 file.
'
' @return   True    If file was properly stop, False otherwise.

Public Function StopMp3() As Boolean

    On Error GoTo Error_Handl
    
    If IMC Is Nothing Then Exit Function
    
    IMC.Stop
    IMPos.CurrentPosition = 0
    
    StopMp3 = True
    Exit Function

Error_Handl:
    StopMp3 = False
    Call LogError(err.Number, err.Description, "clsAudio.StopMp3")

End Function

''
' Destroy the MP3 objects.
'
' @return   True If the objects were destroy.

Public Function MP3Empty() As Boolean

    '**************************************
    '* Created by: Fredrik Alexandersson  *
    '* Modify Date: 20/5/2003             *
    '**************************************
    On Error GoTo ErrorHandl
    
    If ObjPtr(IMC) > 0 Then
        Call StopMp3

    End If
    
    Set IBA = Nothing
    Set IME = Nothing
    Set IMPos = Nothing
    Set IMC = Nothing
    
    MP3Empty = True
    Exit Function

ErrorHandl:
    MP3Empty = False
    Call LogError(err.Number, err.Description, "clsAudio.MP3Empty")
    End

End Function

''
'   Plays a new MIDI or MP3 file.
'
'   @param              FileName    File name to be played
'   @param              FileType   File Type : mp3 or mid.
'   @param Optional     Loops   Loops for midi.
'
Public Sub PlayBackgroundMusic(ByVal FileName As String, _
                               ByVal FileType As Byte, _
                               Optional ByVal Loops As Long = -1)

    '***************************************************
    'Author: Recox
    'Last Modification: 06/01/20
    'Reproduce un mp3 o midi segun lo que elijamos.
    '***************************************************
    'A el himno del AO no lo saque nadie, este es un parche para que en el frmConectar no salga la musica cuando
    'Clickeo un server en la lista (Recox)
    If CurrentMusicFile = "6.mp3" And FileName = "2" Then Exit Sub

    Select Case FileType

        Case MusicTypes.Midi
            FileName = FileName & ".mid"

            'Si la musica es la misma, la dejamos que siga y no hacemos nada (Recox)
            If FileName = CurrentMusicFile Then Exit Sub

            Call PlayMIDI(FileName, Loops)

        Case MusicTypes.Mp3
            FileName = FileName & ".mp3"

            'Si la musica es la misma, la dejamos que siga y no hacemos nada (Recox)
            If FileName = CurrentMusicFile Then Exit Sub
            
            Call PlayMp3(FileName)

    End Select

End Sub

Private Function CreateBufferFromFile_Ex(ByVal file_name As Integer, _
                                         ByVal buffer_index) As Boolean
    '**************************************************************
    'Author: Augusto Jos? Rando
    'Last Modify Date: 3/04/2007
    'Creaci?n de buffers desde memoria
    '**************************************************************

    On Error GoTo ErrorHandler

    Dim btArr()       As Byte, gudtHeader As WAVETYPE

    Dim I             As Long, lngChunkSize As Long

    Dim tB(1 To 4)    As Byte

    Dim glngChunkSize As Long

    Dim lngOffset     As Long

    Dim udtBufferDesc As DSBUFFERDESC

    If Get_File_Data("Wav", file_name & ".wav", btArr) Then

        lngOffset = LenB(gudtHeader)

        Call CopyMemory(gudtHeader, btArr(0), lngOffset)

        'Only PCM
        If gudtHeader.intFormat <> WAVE_FORMAT_PCM Then Exit Function

        lngOffset = 36

        For I = lngOffset To UBound(btArr)
            tB(1) = btArr(I)
            tB(2) = btArr(I + 1)
            tB(3) = btArr(I + 2)
            tB(4) = btArr(I + 3)

            If StrConv(tB(), vbUnicode) = "data" Then
                Exit For

            End If

        Next I

        lngOffset = I + 4

        Call CopyMemory(ByVal VarPtr(glngChunkSize), btArr(lngOffset), ByVal 4)

        lngOffset = lngOffset + 4

        'Set the Wave Format
        With udtBufferDesc.fxFormat
            .nFormatTag = gudtHeader.intFormat
            .nChannels = gudtHeader.intChannels
            .lSamplesPerSec = gudtHeader.lngSamplesPerSec
            .nBitsPerSample = gudtHeader.intBitsPerSample
            .nBlockAlign = gudtHeader.intBlockAlign '/ 8 * gudtHeader.intChannels
            .lAvgBytesPerSec = gudtHeader.lngAvgBytesPerSec '* .nBlockAlign
            .nsize = gudtHeader.lngChunkSize

        End With

        'Create the buffer
        udtBufferDesc.lBufferBytes = glngChunkSize
        udtBufferDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME

        Set DSBuffers(buffer_index).buffer = DirectSound.CreateSoundBuffer(udtBufferDesc)

        'Load the buffer with data
        DSBuffers(buffer_index).buffer.WriteBuffer 0, glngChunkSize, btArr(lngOffset), DSBLOCK_ENTIREBUFFER

        CreateBufferFromFile_Ex = True

    End If

    Exit Function

ErrorHandler:
    'Ocurri? un error, la causa m?s probable es que el archivo no exista
    'If General_File_Exists(Windows_Temp_Dir & file_name & ".wav", vbNormal) Then
    '    Delete_File Windows_Temp_Dir & file_name & ".wav"
    'End If

End Function

