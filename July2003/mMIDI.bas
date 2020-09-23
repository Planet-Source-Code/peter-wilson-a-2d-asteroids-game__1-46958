Attribute VB_Name = "mMIDI"
Option Explicit

' The MIDIOUTCAPS structure describes the capabilities of a MIDI output device.
Private Type MIDIOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * 32
    wTechnology As Integer
    wVoices As Integer
    wNotes As Integer
    wChannelMask As Integer
    dwSupport As Long
End Type

' Pointer to an HMIDIOUT handle. This location is filled with a handle identifying the opened MIDI output device. The handle is used to identify the device in calls to other MIDI output functions.
Private lphmo As Long

' Identifier of the MIDI output device that is to be opened.
Private uDeviceID As Long

' The midiOutOpen function opens a MIDI output device for playback.
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwCallbackInstance As Long, ByVal dwFlags As Long) As Long

' The midiOutClose function closes the specified MIDI output device.
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hmo As Long) As Long

' The midiOutGetNumDevs function retrieves the number of MIDI output devices present in the system.
Private Declare Function midiOutGetNumDevs Lib "winmm.dll" () As Integer

' The midiOutGetDevCaps function queries a specified MIDI output device to determine its capabilities.
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpMidiOutCaps As MIDIOUTCAPS, ByVal cbMidiOutCaps As Long) As Long

' The midiOutShortMsg function sends a short MIDI message to the specified MIDI output device.
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hmo As Long, ByVal dwMsg As Long) As Long

' The midiOutGetErrorText function retrieves a textual description for an error identified by the specified error code.
Private Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal mmrError As Long, ByVal lpText As String, ByVal cchText As Long) As Long

Public Enum mdrInstrument
    NoInstrument = 0
    AcousticGrandPiano = 1
    BrightAcousticPiano = 2
    ElectricGrandPiano = 3
    HonkyTonkPiano = 4
    RhodesPiano = 5
    ChorusedPiano = 6
    Harpsichord = 7
    Clavinet = 8
    Celesta = 9
    Glockenspiel = 10
    MusicBox = 11
    Vibraphone = 12
    Marimba = 13
    Xylophone = 14
    TubularBells = 15
    Dulcimer = 16
    HammondOrgan = 17
    PercussiveOrgan = 18
    RockOrgan = 19
    ChurchOrgan = 20
    ReedOrgan = 21
    Accordion = 22
    Harmonica = 23
    TangoAccordion = 24
    AcousticGuitar_Nylon = 25
    AcousticGuitar_Steel = 26
    ElectricGuitar_Jazz = 27
    ElectricGuitar_Clean = 28
    ElectricGuitar_Muted = 29
    OverdrivenGuitar = 30
    DistortionGuitar = 31
    GuitarHarmonics = 32
    AcousticBass = 33
    ElectricBass_Finger = 34
    ElectricBass_Pick = 35
    FretlessBass = 36
    SlapBass1 = 37
    SlapBass2 = 38
    SynthBass1 = 39
    SynthBass2 = 40
    Violin = 41
    Viola = 42
    Cello = 43
    Contrabass = 44
    TremoloStrings = 45
    PizzicatoStrings = 46
    OrchestralHarp = 47
    Timpani = 48
    StringEnsemble1 = 49
    StringEnsemble2 = 50
    SynthStrings1 = 51
    SynthStrings2 = 52
    ChoirAahs = 53
    VoiceOohs = 54
    SynthVoice = 55
    OrchestraHit = 56
    Trumpet = 57
    Trombone = 58
    Tuba = 59
    MutedTrumpet = 60
    FrenchHorn = 61
    BrassSection = 62
    SynthBrass1 = 63
    SynthBrass2 = 64
    SopranoSax = 65
    AltoSax = 66
    TenorSax = 67
    BaritoneSax = 68
    Oboe = 69
    EnglishHorn = 70
    Bassoon = 71
    Clarinet = 72
    Piccolo = 73
    Flute = 74
    Recorder = 75
    PanFlute = 76
    BottleBlow = 77
    Shakuhachi = 78
    Whistle = 79
    Ocarina = 80
    Lead1_Square = 81
    Lead2_Sawtooth = 82
    Lead3_CalliopeLead = 83
    Lead4_ChiffLead = 84
    Lead5_Charang = 85
    Lead6_Voice = 86
    Lead7_Fifths = 87
    Lead8_BrassAndLead = 88
    Pad1_NewAge = 89
    Pad2_Warm = 90
    Pad3_Polysynth = 91
    Pad4_Choir = 92
    Pad5_Bowed = 93
    Pad6_Metallic = 94
    Pad7_Halo = 95
    Pad8_Sweep = 96
    FX1_IceRain = 97
    FX2_SoundTrack = 98
    FX3_Crystal = 99
    FX4_Atmosphere = 100
    FX5_Brightness = 101
    FX6_Goblin = 102
    FX7_EchoDrops = 103
    FX8_StarTheme = 104
    Sitar = 105
    Banjo = 106
    Shamisen = 107
    Koto = 108
    Kalimba = 109
    BagPipe = 110
    Fiddle = 111
    Shanai = 112
    TinkleBell = 113
    Agogo = 114
    SteelDrums = 115
    Woodblock = 116
    Taiko = 117
    MelodicTom = 118
    SynthDrum = 119
    ReverseCymbal = 120
    GuitarFretNoise = 121
    BreathNoise = 122
    Seashore = 123
    BirdTweet = 124
    TelephoneRing = 125
    Helicopter = 126
    Applause = 127
    Gunshot = 128
End Enum

Private Sub RaiseMIDIErrorMessage(Source As String, ByVal intMIDIErrorNumber As Integer)
    
    ' Retrieves a textual description for an error identified by the specified error code,
    ' and then raises an error using the correct message.
    
    Dim strErrMsg As String
    Dim lngReturnValue As Integer

    strErrMsg = Space(128)
    lngReturnValue = midiOutGetErrorText(intMIDIErrorNumber, strErrMsg, 128)

    VBA.Err.Raise vbObjectError + 1001, App.Title & ".mMIDI." & Source, strErrMsg

End Sub

Private Function PackDWord(intA As Integer, intB As Integer, intC As Integer, intD As Integer) As Long
    
    ' This function helps us cram four integers into a single long data type.
    
    PackDWord = intB * &H10000 + intC * &H100 + intD

End Function

Public Sub ChangeInstrument(ByVal Channel As Integer, ByVal InstrumentIndex As mdrInstrument)
    
    ' I have assigned "InstrumentIndex" to "intInstrument" more for cosmetic reasons than anything else.
    Dim intInstrument As Integer
    intInstrument = InstrumentIndex
    
    ' Check for valid MIDI output device.
    If (lphmo = 0) Then Exit Sub
    
    ' Send a Control Command first (0,0)
    Call SendShortMIDIMessage(&HB0 + Channel, 0, 0)
    
    ' Send the data
    Call SendShortMIDIMessage(&HC0 + Channel, intInstrument - 1, 0)
    
End Sub

Public Sub SendShortMIDIMessage(StatusByte1 As Integer, DataByte2 As Integer, DataByte3 As Integer)
        
    ' The message is packed into a DWORD value with the first byte of the message
    ' in the low-order byte. The MIDI message is packed into the "dwMsg" parameter as follows:
    ' ----------------------------------------------------------------------------------------
    ' Word  |   Byte        |   Usage                                       |   VB
    ' ------+---------------+-----------------------------------------------+-----------------
    ' High  |   High-Order  |   Not used.                                   |   0 (zero)
    '       |   Low-Order   |   The second byte of MIDI data (when needed). |   DataByte3
    ' Low   |   High-Order  |   The first byte of MIDI data (when needed).  |   DataByte2
    '       |   Low-Order   |   The MIDI status                             |   StatusByte1
    ' ----------------------------------------------------------------------------------------
    
    Dim intError As Integer
    
    ' Check for valid MIDI output device.
    If (lphmo = 0) Then Exit Sub
    
    intError = midiOutShortMsg(lphmo, PackDWord(0, DataByte3, DataByte2, StatusByte1))
    If Not (intError = 0) Then Call RaiseMIDIErrorMessage("SendShortMIDIMessage", intError)
    
    
    ' =============
    ' MIDI Messages
    ' =============
    ' If you would like to know what the complete MIDI messages are (since it is a bit cryptic), then
    ' you will need to download or purchase the "Complete MIDI 1.0 Detailed Specification"
    '
    ' Also see:
    '   http://www.midi.org/about-midi/table1.shtml
    '   http://www.midi.org/about-midi/table2.shtml
    '   http://www.midi.org/about-midi/table3.shtml
    '
    ' The tables at the above URL's will tell you what value of
    ' StatusByte and DataBytes to use for a desired effect.
    
End Sub

Public Sub OpenMIDI()
    
    Dim intError As Integer
    
    ' Close before trying to open again.
    Call CloseMIDI
    
    ' Check if there are any MIDI devices we can open.
    If midiOutGetNumDevs = 0 Then Exit Sub
    
    ' Default device
    uDeviceID = 0
    
    ' Attempt to open a MIDI output device for playback.
    intError = midiOutOpen(lphmo, uDeviceID, 0, 0, 0)
    If Not (intError = 0) Then Call RaiseMIDIErrorMessage("OpenMIDI", intError)
    
End Sub

Public Sub CloseMIDI()
    
    Dim intError As Integer
    
    ' Check for valid MIDI output device.
    If (lphmo = 0) Then Exit Sub
    
    ' Attempt to close a MIDI output device.
    intError = midiOutClose(lphmo)
    If (intError = 0) Then
        lphmo = 0
    Else
        Call RaiseMIDIErrorMessage("OpenMIDI", intError)
    End If
    
End Sub


