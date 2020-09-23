Attribute VB_Name = "modDS"
'The Main Direct Sound Object
Public ds As DirectSound

'The woohoobuffer and its description (Homer)
Public WooHooBuffer As DirectSoundBuffer
Public WooHooBufferDesc As DSBUFFERDESC

'The crapbuffer and its description (Krusty)
Public CrapBuffer As DirectSoundBuffer
Public CrapBufferDesc As DSBUFFERDESC

'Used in the creation of the sound buffers
Public WavFormat As WAVEFORMATEX

'This creates the sounds from a file
Sub CreateSoundBufFromFile(Buffer As DirectSoundBuffer, FileName As String, BufferDesc As DSBUFFERDESC, wFormat As WAVEFORMATEX)
    If ds Is Nothing Then Exit Sub

    Set Buffer = ds.CreateSoundBufferFromFile(FileName, BufferDesc, wFormat)
End Sub

'This is for easier transportation to the main initialization
Sub LoadSounds()
    Call CreateSoundBufFromFile(WooHooBuffer, App.Path & "\sounds\woohoo.wav", WooHooBufferDesc, WavFormat)
    Call CreateSoundBufFromFile(CrapBuffer, App.Path & "\sounds\crap.wav", CrapBufferDesc, WavFormat)
End Sub

'This plays the sound
Sub dsPlay(Buffer As DirectSoundBuffer, Looping As Boolean)
    Call Buffer.SetCurrentPosition(0)

    'If the sound is to loop
    If Looping = True Then
        'play it with a loop
        Call Buffer.Play(DSBPLAY_LOOPING)
    'If the sound isn't to loop
    Else
        'don't play it with a loop
        Call Buffer.Play(DSBPLAY_DEFAULT)
    End If
End Sub

'This stops the sound
Sub dsStop(Buffer As DirectSoundBuffer)
    Buffer.Stop
End Sub

'This stops all of the sounds
Sub StopSounds()
    Call dsStop(WooHooBuffer)
    Call dsStop(CrapBuffer)
End Sub

'This unloads all of the soundbuffers at the programs end
Sub UnloadSounds()
    Set WooHooBuffer = Nothing
    Set CrapBuffer = Nothing
End Sub
