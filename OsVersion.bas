Attribute VB_Name = "OsVersion"
Option Explicit
'adapted from Microsoft's Knowledgebase article (Q189249)

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Enum OsType
    Nt2000
    Win9xMe
    OsUnknown
End Enum

Public Function GetVersion() As String
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   With osinfo
   Select Case .dwPlatformId
      Case 1
         If .dwMinorVersion = 0 Then
            GetVersion = "Windows 95"
         ElseIf .dwMinorVersion = 10 Then
            GetVersion = "Windows 98"
         ElseIf .dwMinorVersion = 90 Then
            GetVersion = "Windows Me"
         End If
      Case 2
         If .dwMajorVersion = 3 Then
            GetVersion = "Windows NT 3.51"
         ElseIf .dwMajorVersion = 4 Then
            GetVersion = "Windows NT 4.0"
         ElseIf .dwMajorVersion = 5 Then
            GetVersion = "Windows 2000"
         End If
      Case Else
         GetVersion = "Failed"
   End Select
   End With
End Function

'for speaker beep function, only the platform type is relevant
Public Function GetPlatform() As OsType
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   
   Select Case osinfo.dwPlatformId
      Case 1
        GetPlatform = Win9xMe
      Case 2
        GetPlatform = Nt2000
      Case Else
        GetPlatform = OsUnknown
   End Select
   
End Function

