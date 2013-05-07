Attribute VB_Name = "ModINI"
Option Explicit
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public DefaultINIValue As Dictionary

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
    Dim StringBuffer As String
    Dim StringBufferSize As Long
    
    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)
    
    StringBufferSize = GetPrivateProfileString(INISection, INIKey, "", StringBuffer, StringBufferSize, INIFile)
    
    If StringBufferSize > 0 Then ReadINI = Left$(StringBuffer, StringBufferSize) Else ReadINI = vbNullString
    
    If ReadINI = vbNullString Then
        Call WriteINI(INISection, INIKey, CStr(DefaultINIValue(INIFile)(INISection)(INIKey)), INIFile)
        ReadINI = DefaultINIValue(INIFile)(INISection)(INIKey)
    End If
End Function

Public Sub InitAccountOpt()
On Error Resume Next
    AccOpt.InfName = ReadINI("INFO", "Account", ClientConfigurationFile)
    AccOpt.InfPass = ReadINI("INFO", "Password", ClientConfigurationFile)
    AccOpt.SpeechBubbles = CBool(ReadINI("CONFIG", "SpeechBubbles", ClientConfigurationFile))
    AccOpt.NpcBar = CBool(ReadINI("CONFIG", "NpcBar", ClientConfigurationFile))
    AccOpt.NpcName = CBool(ReadINI("CONFIG", "NPCName", ClientConfigurationFile))
    AccOpt.NpcDamage = CBool(ReadINI("CONFIG", "NPCDamage", ClientConfigurationFile))
    AccOpt.PlayBar = CBool(ReadINI("CONFIG", "PlayerBar", ClientConfigurationFile))
    AccOpt.PlayName = CBool(ReadINI("CONFIG", "PlayerName", ClientConfigurationFile))
    AccOpt.PlayDamage = CBool(ReadINI("CONFIG", "PlayerDamage", ClientConfigurationFile))
    AccOpt.MapGrid = CBool(ReadINI("CONFIG", "MapGrid", ClientConfigurationFile))
    AccOpt.Music = CBool(ReadINI("CONFIG", "Music", ClientConfigurationFile))
    AccOpt.Sound = CBool(ReadINI("CONFIG", "Sound", ClientConfigurationFile))
    AccOpt.Autoscroll = CBool(ReadINI("CONFIG", "AutoScroll", ClientConfigurationFile))
    AccOpt.NomObjet = CBool(ReadINI("CONFIG", "NomObjet", ClientConfigurationFile))
    AccOpt.LowEffect = CBool(ReadINI("CONFIG", "LowEffect", ClientConfigurationFile))
End Sub


Public Function ReadINISections(FileName As String) As String()
    Dim strBuffer As String, intLen As Integer
    
    
    Do While (intLen = Len(strBuffer) - 2) Or (intLen = 0)
        If strBuffer = vbNullString Then
            strBuffer = Space(256)
        Else
            strBuffer = String(Len(strBuffer) * 2, 0)
        End If
        
        intLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), FileName)
    Loop
    
    strBuffer = Left$(strBuffer, intLen - 1)
    
    ReadINISections = Split(strBuffer, vbNullChar)
End Function

Public Function ReadINIKeys(ByVal FileName As String, ByVal section As String) As String()
    Dim continue As Boolean
    Dim StringBuffer As String
    Dim StringBufferSize As Long
    Dim returnCode As Long

    continue = True
    StringBufferSize = 256

    Do While continue
        StringBuffer = String$(StringBufferSize, vbNullChar)
        returnCode = GetPrivateProfileString(section, vbNullString, "", StringBuffer, StringBufferSize, FileName)

        If returnCode = StringBufferSize - 2 Then
            StringBufferSize = StringBufferSize + 256
        Else
            If returnCode = 0 Then
                StringBuffer = ""
            Else
                StringBuffer = Left$(StringBuffer, returnCode - 1)
            End If
            continue = False
        End If
    Loop
    
    ReadINIKeys = Split(StringBuffer, vbNullChar)
End Function

Public Sub InitDefaultINIValue()
    Set DefaultINIValue = Nothing
    Set DefaultINIValue = New Dictionary
    
    Set DefaultINIValue(ClientConfigurationFile) = New Dictionary
    Set DefaultINIValue(ClientConfigurationFile)("INFO") = New Dictionary
    DefaultINIValue(ClientConfigurationFile)("INFO")("Account") = ""
    DefaultINIValue(ClientConfigurationFile)("INFO")("Password") = ""
    Set DefaultINIValue(ClientConfigurationFile)("CONFIG") = New Dictionary
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("SpeechBubbles") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("NpcBar") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("NpcName") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("NpcDamage") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("PlayerBar") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("PlayerName") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("PlayerDamage") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("MapGrid") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("Music") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("Sound") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("AutoScroll") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("NomObjet") = 1
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("LowEffect") = 0
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("WEBSITE") = "www.aride-online.com"
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("Music") = 0
    DefaultINIValue(ClientConfigurationFile)("CONFIG")("Port") = 0
    
    Set DefaultINIValue(OptionConfigurationFile) = New Dictionary
    Set DefaultINIValue(OptionConfigurationFile)("COMMAND") = New Dictionary
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("haut") = 90
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("bas") = 83
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("gauche") = 81
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("droite") = 68
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("attaque") = 69
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("courir") = 16
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("ramasser") = 32
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("action") = 65
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac1") = 49
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac2") = 50
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac3") = 51
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac4") = 52
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac5") = 53
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac6") = 54
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac7") = 55
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac8") = 56
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac9") = 57
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac10") = 48
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac11") = 112
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac12") = 113
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac13") = 114
    DefaultINIValue(OptionConfigurationFile)("COMMAND")("rac14") = 115
    
    Set DefaultINIValue(ThemeConfigurationFile) = New Dictionary
    Set DefaultINIValue(ThemeConfigurationFile)("THEMES") = New Dictionary
    DefaultINIValue(ThemeConfigurationFile)("THEMES")("Theme") = "\Themes\Aride Online"
    
    
    ' Writing default values
    Dim file As Variant
    Dim section As Variant
    Dim key As Variant
    
    For Each file In DefaultINIValue
        For Each section In DefaultINIValue(file)
            For Each key In DefaultINIValue(file)(section)
                Call ReadINI(CStr(section), CStr(key), CStr(file))
            Next key
        Next section
    Next file
End Sub
