Attribute VB_Name = "mdlMci"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1 '�첽����
Private Const SND_SYNC = &H0  'ͬ������,ȱʡ

Public Sub PlaySound(strFileName)
    sndPlaySound strFileName, SND_ASYNC
End Sub
