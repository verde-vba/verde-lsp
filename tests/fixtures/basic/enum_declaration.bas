Attribute VB_Name = "EnumDeclaration"

' Public Enum with hex values
Public Enum InterfaceColors
   icMistyRose = &HE1E4FF&
   icSlateGray = &H908070&
   icDodgerBlue = &HFF901E&
   icDeepSkyBlue = &HFFBF00&
   icSpringGreen = &H7FFF00&
   icForestGreen = &H228B22&
   icGoldenrod = &H20A5DA&
   icFirebrick = &H2222B2&
End Enum

' Enum with auto-increment values
Public Enum Days
    Monday
    Tuesday
    Wednesday
    Thursday
    Friday
    Saturday
    Sunday
End Enum

' Enum with explicit values
Public Enum eSoundAction
    Sound_Load
    Sound_Play
    Sound_Stop
End Enum

' Using Enum in Select Case
Sub UseEnum(m As Days)
    Select Case m
       Case Monday: Debug.Print "Monday"
       Case Tuesday: Debug.Print "Tuesday"
       Case Else: Debug.Print "Other"
    End Select
End Sub

' Enum for Excel operations
Public Enum Corner
    cnrTopLeft
    cnrTopRight
    cnrBottomLeft
    cnrBottomRight
End Enum

Public Enum OverwriteAction
    oaPrompt = 1
    oaOverwrite = 2
    oaSkip = 3
    oaError = 4
    oaCreateDirectory = 8
End Enum
