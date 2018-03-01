Attribute VB_Name = "Module1"
Private Declare Function PlaySound Lib "winmm.dll" Alias _
    "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, _
        ByVal dwFlags As Long) As Long
Public Function PlayWaveFile(strFileName As String, Optional blnAsync As Boolean) As Boolean
    Dim lngFlags As Long
    Const snd_sync = &H0
    Const snd_Async = &H1
    Const snd_Nodefault = &H2
    Const snd_Filename = &H20000
    lngFlags = snd_Nodefault Or snd_Filename Or snd_sync
    If blnAsync Then lngFlags = lngFlags Or snd_Async
    PlayWaveFile = PlaySound(strFileName, 0&, lngFlags)
End Function

