Attribute VB_Name = "Speaker"
Option Explicit

'following lines are for Win9xMe platforms
'For these systems, the file WIN95IO.DLL must be copied
'to the Windows/System folder.
'WIN95IO.DLL is available from http://www.softcircuits.com
Declare Sub vbOut Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Function vbInp Lib "WIN95IO.DLL" (ByVal nPort As Integer) As Integer

'This line is for NT2000 platforms
Public Declare Function NtBeep Lib "kernel32" Alias "Beep" (ByVal FreqHz As Long, ByVal DurationMs As Long) As Long

'This is where the beep method depends on the operating system
Public Sub PcSpeakerBeep(ByVal FreqHz As Integer, ByVal LengthMs As Single)
    
    Select Case GetPlatform
        Case Win9xMe
            Call Win9xBeep(FreqHz, LengthMs)
        Case Nt2000
            Call NtBeep(CLng(FreqHz), CLng(LengthMs))
        Case OsUnknown
            Beep    'use the default beep routine, probably the sound card
    End Select
            
End Sub

'following routine largely by Jorge Loubet
Private Sub Win9xBeep(ByVal Freq As Integer, ByVal Length As Single)

    Dim LoByte As Integer
    Dim HiByte As Integer
    Dim Clicks As Integer
    Dim SpkrOn As Integer
    Dim SpkrOff As Integer
    Dim TimeEnd As Single
    
    TimeEnd = Timer + Length / 1000
    
    'Ports 66, 67, and 97 control timer and speaker
    '
    'Divide clock frequency by sound frequency
    'to get number of "clicks" clock must produce.
        Clicks = CInt(1193280 / Freq)
        LoByte = Clicks And &HFF
        HiByte = Clicks \ 256
    'Tell timer that data is coming
        vbOut 67, 182
    'Send count to timer
        vbOut 66, LoByte
        vbOut 66, HiByte
    'Turn speaker on by setting bits 0 and 1 of PPI chip.
        SpkrOn = vbInp(97) Or &H3
        vbOut 97, SpkrOn
    
    'Leave speaker on (while timer runs)
        Do While Timer < TimeEnd
            'Let processor do other tasks
            DoEvents
        Loop
    'Turn speaker off.
        SpkrOff = vbInp(97) And &HFC
        vbOut 97, SpkrOff
End Sub

Public Sub Warble(ByVal FreqHz As Integer, ByVal DurationMs As Single)
    Dim EndTime As Single
    EndTime = Timer + DurationMs / 1000
    
    If FreqHz < 100 Then FreqHz = 100
    Do While EndTime > Timer
        Call PcSpeakerBeep(FreqHz, 10)
        Call PcSpeakerBeep(FreqHz / 1.1, 10)
        Call PcSpeakerBeep(FreqHz / 1.2, 10)
        Call PcSpeakerBeep(FreqHz / 1.3, 10)
        Call PcSpeakerBeep(FreqHz / 1.2, 10)
        Call PcSpeakerBeep(FreqHz / 1.1, 10)
    Loop
End Sub
