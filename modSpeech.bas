Attribute VB_Name = "modSpeech"
' Constants "SPEAKFLAGS"
Const SPF_ASYNC = 1
Const SPF_DEFAULT = 0
Const SPF_IS_FILENAME = 4
Const SPF_IS_NOT_XML = 16
Const SPF_IS_XML = 8
Const SPF_NLP_MASK = 64
Const SPF_NLP_SPEAK_PUNC = 64
Const SPF_PERSIST_XML = 32
Const SPF_PURGEBEFORESPEAK = 2
Const SPF_UNUSED_FLAGS = -128
Const SPF_VOICE_MASK = 127

Public Sub ReadText(ByVal TextString As String)
    If application_speech_enable = True Then
        DoEvents
        Set ISpeechVoice = CreateObject("SAPI.SpVoice")
        Call ISpeechVoice.Speak(TextString, SPF_DEFAULT)
    End If
End Sub
