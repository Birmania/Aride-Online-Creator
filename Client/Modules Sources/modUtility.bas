Attribute VB_Name = "modUtility"
'    Copyright (C) 2013  BRULTET Antoine
'
'    This file is part of Aride Online Creator.
'
'    Aride Online Creator is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Aride Online Creator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Aride Online Creator.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

'Put form on foreground
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

' API : Obtenir les coordonnées de la souris
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

''''
'begin Get System Directory

'Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetFolderPath Lib "shfolder" _
    Alias "SHGetFolderPathA" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, _
    ByVal hToken As Long, ByVal dwFlags As Long, _
    ByVal pszPath As String) As Long
    
Private Const S_OK = &H0
Private Const S_FALSE = &H1
Private Const E_INVALIDARG = &H80070057
    
Private Const CSIDL_LOCAL_APPDATA = &H1C& 'Not used for the moment
Private Const CSIDL_SYSTEMX86 = &H29&

Private Const SHGFP_TYPE_CURRENT = 0
    
''''

'' DLL Import
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal _
lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

' Pointeur de souris
Public Type POINTAPI
    X As Long
    Y As Long
End Type

' Debut de suppression de bordure de liste
' Declares
Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Enum enWindowStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_EX_TOOLWINDOW = &H80&
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    '\\ New from 95/NT4 onwards
    WS_EX_MDICHILD = &H40
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum
' Fin de suppression de bordure de liste


Sub SetQuestTime(ByVal time As Long)

    If time > 0 Then
        frmMirage.quetetimersec.Interval = 1000
        Seco = time - ((time \ 60) * 60)
        Minu = time \ 60
        frmMirage.tmpsquete.Visible = True
        If Len(Str$(Minu)) > 2 Then frmMirage.Minute.Caption = Minu & ":" Else frmMirage.Minute.Caption = "0" & Minu & ":"
        If Len(Str$(Seco)) > 2 Then frmMirage.seconde.Caption = Seco Else frmMirage.seconde.Caption = "0" & Seco
        frmMirage.quetetimersec.Enabled = True
        Exit Sub
    End If

End Sub

Public Function IsEmptyArray(ppsa As Long) As Boolean
    Dim psa As Long 'SAFEARRAY*
    
    ' Déférence le SAFEARRAY**
    CopyMemory psa, ByVal ppsa, LenB(psa)
    
    'Vérifie si une l'adresse est valide
    IsEmptyArray = psa = 0
    
End Function

Public Function GetMousePosition() As POINTAPI
    Dim pos As POINTAPI

    GetCursorPos pos
    ScreenToClient frmMirage.hwnd, pos
    GetMousePosition = pos
End Function

Sub ListDir(ByVal directory As String, ByRef directories() As String)
    Dim i As Integer
    Dim Fol, Fi
    Set Fol = CreateObject("Scripting.FileSystemObject").GetFolder(directory)

    If Fol.SubFolders.Count > 0 Then
        ReDim directories(0 To Fol.SubFolders.Count - 1) As String
        i = 0
        For Each Fi In Fol.SubFolders
            directories(i) = Fi.name
            i = i + 1
        Next Fi
    End If
    
    Set Fol = Nothing
    Set Fi = Nothing
End Sub

Sub ListFiles(ByVal directory As String, ByRef files() As String)
    Dim i As Integer
    Dim Fol, Fi
    Set Fol = CreateObject("Scripting.FileSystemObject").GetFolder(directory)

    If Fol.files.Count > 0 Then
        ReDim files(0 To Fol.files.Count - 1) As String
        i = 0
        For Each Fi In Fol.files
            files(i) = Fi.name
            i = i + 1
        Next Fi
    End If
    
    Set Fol = Nothing
    Set Fi = Nothing
End Sub

Public Function GetArraySize(ByRef MyArray As Variant)
    GetArraySize = 0
    On Error Resume Next
    GetArraySize = UBound(MyArray) + 1
End Function

Public Function GetPathOfFileIn(ByVal FileName As String, ByVal directory As String)
    Dim i As Integer, J As Integer
    Dim Path As String
    Path = vbNullString
    Dim allDirs() As String
    Dim allFiles() As String
    Call ListDir(directory, allDirs)
    For i = 0 To GetArraySize(allDirs) - 1
        Call ListFiles(directory & "\" & allDirs(i), allFiles)
        For J = 0 To GetArraySize(allFiles) - 1
            If allFiles(J) = FileName Then
                Path = directory & "\" & allDirs(i) & "\" & allFiles(J)
                Exit For
            End If
        Next J
        If Path <> vbNullString Then
            Exit For
        End If
    Next i
    
    GetPathOfFileIn = Path
End Function

'' Function

Public Function RemoveBorder(ByVal hwnd As Long)
    
    Dim lngRetVal As Long
    
    lngRetVal = GetWindowLong(hwnd, GWL_STYLE)
    
    lngRetVal = lngRetVal And (Not WS_BORDER) And (Not WS_DLGFRAME) And (Not WS_THICKFRAME)
    
    SetWindowLong hwnd, GWL_STYLE, lngRetVal
    
     lngRetVal = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    lngRetVal = lngRetVal And (Not WS_EX_CLIENTEDGE) And (Not WS_EX_STATICEDGE) And (Not WS_EX_WINDOWEDGE)
    
    
    SetWindowLong hwnd, GWL_EXSTYLE, lngRetVal
    
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function


'# Permet de tester l'existence d'un dossier...
Public Function FolderExists(ByRef vsPathFolder As String) As Boolean
    On Error Resume Next
    FolderExists = CBool(GetAttr(vsPathFolder) And vbDirectory)
End Function

Public Function MkDir(ByRef vsPathFolder As String) As Boolean
Dim i As Long
    '# Si le dossier n'existe pas...
    If Not FolderExists(vsPathFolder) Then
        '# On va procéder a un découpage : on récupère le dossier parent
        i = InStrRev(vsPathFolder, "\")
        If i = 0 Then
            '# On est surement arrivés au nom du lecteur, et il ne semble pas être présent
            MkDir = False
        ElseIf MkDir(Left$(vsPathFolder, i - 1)) Then
            '# Le repertoire parent existe ou a pu être créé...
            FileSystem.MkDir vsPathFolder
            MkDir = True
        Else
            '# Le repertoire parent n'a pu être créé
            MkDir = False
        End If
    Else
        '# Le repertoire existe bien
        MkDir = True
    End If
End Function

Public Sub SetIcon(ByRef Form As Form)
   If App.LogMode <> 0 Then
    Form.Icon = LoadResPicture("APPICON", vbResIcon)
   End If
End Sub

Public Function IsInArray(FindValue As Variant, arrSearch As Variant) As Boolean
    On Error GoTo LocalError
    If Not IsArray(arrSearch) Then Exit Function
    IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0
    
    Exit Function
LocalError:
    'Justin (just in case)
End Function


Public Function IsDirectoryEmpty(ByVal Path As String) As Boolean
    IsDirectoryEmpty = (dir(Path & "\*.*") = "")
End Function

Public Function KillFolder(ByVal FullPath As String) _
   As Boolean
   
'******************************************
'PURPOSE: DELETES A FOLDER, INCLUDING ALL SUB-
'         DIRECTORIES, FILES, REGARDLESS OF THEIR
'         ATTRIBUTES

'PARAMETER: FullPath = FullPath of Folder to Delete

'RETURNS:   True is successful, false otherwise

'REQUIRES:  'VB6
            'Reference to Microsoft Scripting Runtime
            'Caution in use for obvious reasons

'EXAMPLE:   'KillFolder("D:\MyOldFiles")

'******************************************
On Error Resume Next
Dim oFso As New Scripting.FileSystemObject

'deletefolder method does not like the "\"
'at end of fullpath

If Right(FullPath, 1) = "\" Then FullPath = _
    Left(FullPath, Len(FullPath) - 1)

If oFso.FolderExists(FullPath) Then
    
    'Setting the 2nd parameter to true
    'forces deletion of read-only files
    oFso.DeleteFolder FullPath, True
    
    KillFolder = Err.Number = 0 And _
      oFso.FolderExists(FullPath) = False
End If

End Function
  
Public Function IsDLLAvailable(ByVal DllFilename As String) As Boolean
''Dim X As New TypeLibInfo
'''MsgBox DllFilename
''X.ContainingFile = DllFilename
''Dim A As Object
'''MsgBox X.CoClasses(1).name
'''Set A = CreateObject(X.name & "." & X.CoClasses(1).name)
'
Dim hModule As Long

' attempt to load the module
hModule = LoadLibrary(DllFilename)
If hModule > 32 Then
FreeLibrary (hModule) ' decrement the DLL usage counter
IsDLLAvailable = True
End If
'
''Dim rc As Double
''rc = Shell("regsvr32 """ & DllFilename & """")
''MsgBox rc
''If rc = 0 Then
''    IsDLLAvailable = True
''Else
''    IsDLLAvailable = False
''End If
End Function

Public Function GetSystemDirectory() As String
    Dim systemDirectory As String
    Dim RetVal As Long

    'systemDirectory = Space$(255)
    systemDirectory = String(260, 0)
    
    'Call GetSystemDirectory(systemDirectory, Len(systemDirectory))
    'MsgBox systemDirectory
    
    RetVal = SHGetFolderPath(0, CSIDL_SYSTEMX86, 0, SHGFP_TYPE_CURRENT, systemDirectory)
    
    Select Case RetVal
        Case S_OK
            ' We retrieved the folder successfully
            ' All C Strings are null terminated
            ' So we need to return the string up to the first null character
            systemDirectory = Left(systemDirectory, InStr(1, systemDirectory, Chr(0)) - 1)
        Case S_FALSE
            ' The CSIDL in nFolder is valid, but the folder does not exist.
            ' Use CSIDL_FLAG_CREATE to have it created automatically
        Case E_INVALIDARG
            ' nFolder is invalid
    End Select
            
    GetSystemDirectory = systemDirectory
End Function
