Attribute VB_Name = "modDatabase"
Option Explicit
Public Const MAX_PATH As Integer = 260
Private Const ERROR_NO_MORE_FILES As Byte = 18
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpfilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then StripTerminator = Left$(strString, intZeroPos - 1) Else StripTerminator = strString
End Function

Function FileExiste(ByVal FileName As String) As Boolean
    On Error GoTo er:
    If dir$(App.Path & "\" & FileName) = vbNullString Then FileExiste = False Else FileExiste = True
    Exit Function
er:
FileExiste = False
End Function

Function FileExistes(ByVal AppPAndfile As String) As Boolean
    On Error GoTo er:
    If dir$(AppPAndfile) = vbNullString Then FileExistes = False Else FileExistes = True
Exit Function
er:
FileExistes = False
End Function

Sub SaveLocalMap(ByVal mapNum As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\maps\map" & mapNum & ".aoc"
                            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(mapNum)
    Close #f
End Sub

Sub LoadItem(ByVal ItemNum As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\items\item" & ItemNum & ".aoo"
        
    If Not FileExiste("items\item" & ItemNum & ".aoo") Then Exit Sub
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadMap(ByVal mapNum As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\maps\map" & mapNum & ".aoc"
        
'    If Not FileExiste("maps\map" & mapNum & ".aoc") Then Exit Sub
    If Not FileExiste("maps\map" & mapNum & ".aoc") Then
        'Call SaveMapVide(mapNum)
        Call ClearMap(mapNum)
        ReDim Map(mapNum).tile(0 To 30, 0 To 30) As TileRec
    Else
    
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(mapNum)
        Close #f
    End If
End Sub

Sub LoadCraft(ByVal CraftNum As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\Crafts\Craft" & CraftNum & ".aop"

    If Not FileExiste("\Crafts\Craft" & CraftNum & ".aop") Then Exit Sub
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Crafts(CraftNum)
    Close #f
End Sub

Sub LoadArea(ByVal AreaNum As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\Areas\Area" & AreaNum & ".aoz"

    If Not FileExiste("\Areas\Area" & AreaNum & ".aoz") Then Exit Sub
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Areas(AreaNum)
    Close #f
End Sub

Sub LoadDream(ByVal DreamNum As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\Dreams\Dream" & DreamNum & ".aod"

    If Not FileExiste("\Dreams\Dream" & DreamNum & ".aod") Then Exit Sub
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Dreams(DreamNum)
    Close #f
End Sub

Sub LoadQuete(ByVal QIndex As Long)
Dim FileName As String
Dim f As Long
    FileName = App.Path & "\quetes\quete" & QIndex & ".fcq"
        
    If Not FileExiste("quetes\quete" & QIndex & ".fcq") Then Exit Sub
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , quete(QIndex)
    Close #f
End Sub

Function ListMusic(ByVal sStartDir As String)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
   
    On Error Resume Next
   
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    frmMapProperties.lstMusic.Clear
   
    frmMapProperties.lstMusic.AddItem "Aucune", 0
   
    sStartDir = sStartDir & "*.*"
   
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
   
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                    If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                        sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                        If Right$(sTemp, 4) = ".mid" Then frmMapProperties.lstMusic.AddItem sTemp
                        If Right$(sTemp, 4) = ".mp3" Then frmMapProperties.lstMusic.AddItem sTemp
                        If Right$(sTemp, 4) = ".wma" Then frmMapProperties.lstMusic.AddItem sTemp
                        If Right$(sTemp, 4) = ".ogg" Then frmMapProperties.lstMusic.AddItem sTemp
                        If Right$(sTemp, 4) = ".wav" Then frmMapProperties.lstMusic.AddItem sTemp
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
            Sleep 1
        Loop
    End If
    
    lRet = FindClose(lFileHdl)
End Function

Function ListSounds(ByVal sStartDir As String, ByVal Form As Long)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
    
    On Error Resume Next
    
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    If Form = 1 Then
        frmSound.lstSound.Clear
    ElseIf Form = 2 Then
        frmNotice.lstSound.Clear
    End If
    
    sStartDir = sStartDir & "*.*"
    
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                    If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                        sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                        If Right$(sTemp, 4) = ".wav" Then
                            If Form = 1 Then
                                frmSound.lstSound.AddItem sTemp
                            ElseIf Form = 2 Then
                                frmNotice.lstSound.AddItem sTemp
                            End If
                        End If
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
            Sleep 1
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

Function ListPanorama(ByVal sStartDir As String)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
   
    On Error Resume Next
   
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    frmPanorama.lstPano.Clear
   
    frmPanorama.lstPano.AddItem "Aucune", 0
   
    sStartDir = sStartDir & "*.*"
   
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
   
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                    If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                        sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                        If Right$(sTemp, 4) = ".png" Then frmPanorama.lstPano.AddItem sTemp
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
            Sleep 1
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

Function ListFog(ByVal sStartDir As String)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
   
    On Error Resume Next
   
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    frmMapProperties.cmbFog.Clear
   
    frmMapProperties.cmbFog.AddItem "Aucune", 0
   
    sStartDir = sStartDir & "*.*"
   
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
   
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                    If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                        sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                        sTemp = Trim$(LCase(sTemp))
                        If Right$(sTemp, 4) = ".png" And Left$(sTemp, 3) = "fog" And IsNumeric(Mid$(sTemp, 4, 1)) Then frmMapProperties.cmbFog.AddItem Mid$(sTemp, 4, 1)
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
            Sleep 1
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

