Attribute VB_Name = "ModSound"
Option Explicit
'--------------------------------------------------
' Global variables, constants and declaration.
'--------------------------------------------------

' Functions and constants used to play sounds.
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_MEMORY = &H4
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10







Function NoiseGet(ByVal FileName) As String
'------------------------------------------------------------
' Load a sound file into a string variable.
'------------------------------------------------------------
Dim buffer As String
Dim f As Integer
Dim SoundBuffer As String

    On Error GoTo NoiseGet_Error

    buffer = Space$(1024)
    SoundBuffer = ""
    f = FreeFile
    Open FileName For Binary As f
    Do While Not EOF(f)
        Get #f, , buffer     ' Load in 1K chunks
        SoundBuffer = SoundBuffer & buffer
    Loop
    Close f
    NoiseGet = Trim$(SoundBuffer)
Exit Function

NoiseGet_Error:
    SoundBuffer = ""
    Exit Function
End Function

Sub NoisePlay(SoundBuffer As String, ByVal PlayMode As Integer)
'------------------------------------------------------------
' Plays a sound previously loaded into memory with function
' NoiseGet().
'------------------------------------------------------------
Dim retcode As Integer
    
    If SoundBuffer = "" Then Exit Sub

    ' Stop any sound that may currently be playing.
    retcode = sndStopSound(0, SND_ASYNC)

    ' PlayMode should be SND_SYNC or SND_ASYNC
    retcode = sndPlaySound(ByVal SoundBuffer, PlayMode Or SND_MEMORY)
End Sub


