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

' API : Obtenir les coordonnées de la souris
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' Pointeur de souris
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Function GetMousePosition() As POINTAPI
    Dim pos As POINTAPI

    GetCursorPos pos
    ScreenToClient frmMirage.hWnd, pos
    GetMousePosition = pos
End Function

Public Function IsEmptyArray(ppsa As Long) As Boolean
    Dim psa As Long 'SAFEARRAY*
    
    ' Déférence le SAFEARRAY**
    CopyMemory psa, ByVal ppsa, LenB(psa)
    
    'Vérifie si une l'adresse est valide
    IsEmptyArray = psa = 0
    
End Function

Sub ListDir(ByVal directory As String, ByRef directories() As String)
    Dim I As Integer
    Dim Fol, Fi
    Set Fol = CreateObject("Scripting.FileSystemObject").GetFolder(directory)

    If Fol.SubFolders.Count > 0 Then
        ReDim directories(0 To Fol.SubFolders.Count - 1) As String
        I = 0
        For Each Fi In Fol.SubFolders
            directories(I) = Fi.name
            I = I + 1
        Next Fi
    End If
    
    Set Fol = Nothing
    Set Fi = Nothing
End Sub

Sub ListFiles(ByVal directory As String, ByRef Files() As String)
    Dim Fol, Fi
    Dim I As Integer
    Set Fol = CreateObject("Scripting.FileSystemObject").GetFolder(directory)

    If Fol.Files.Count > 0 Then
        ReDim Files(0 To Fol.Files.Count - 1) As String
        I = 0
        For Each Fi In Fol.Files
            Files(I) = Fi.name
            I = I + 1
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
    Dim Path As String
    Dim I As Integer, j As Integer
    
    Path = vbNullString
    Dim allDirs() As String
    Dim allFiles() As String
    Call ListDir(directory, allDirs)
    For I = 0 To GetArraySize(allDirs) - 1
        Call ListFiles(directory & "\" & allDirs(I), allFiles)
        For j = 0 To GetArraySize(allFiles) - 1
            If allFiles(j) = FileName Then
                Path = directory & "\" & allDirs(I) & "\" & allFiles(j)
                Exit For
            End If
        Next j
        If Path <> vbNullString Then
            Exit For
        End If
    Next I
    
    GetPathOfFileIn = Path
End Function

'# Permet de tester l'existence d'un dossier...
Private Function FolderExists(ByRef vsPathFolder As String) As Boolean
    On Error Resume Next
    FolderExists = CBool(GetAttr(vsPathFolder) And vbDirectory)
End Function

Public Function MkDir(ByRef vsPathFolder As String) As Boolean
Dim I As Long
    '# Si le dossier n'existe pas...
    If Not FolderExists(vsPathFolder) Then
        '# On va procéder a un découpage : on récupère le dossier parent
        I = InStrRev(vsPathFolder, "\")
        If I = 0 Then
            '# On est surement arrivés au nom du lecteur, et il ne semble pas être présent
            MkDir = False
        ElseIf MkDir(Left$(vsPathFolder, I - 1)) Then
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

Public Sub CopyFolder(ByVal sourcePath As String, destPath As String)
    Dim fso, fld
    Call MkDir(destPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(sourcePath)
    fld.Copy destPath
End Sub

Public Function Minimum(ByVal number1 As Long, ByVal number2 As Long)
    If number1 <= number2 Then
        Minimum = number1
    Else
        Minimum = number2
    End If
End Function

Public Function ReadINISection(ByVal FileName As String, ByVal section As String) As String
    Dim continue As Boolean
    Dim StringBuffer As String
    Dim StringBufferSize As Long
    Dim returnCode As Long

    continue = True
    StringBufferSize = 256

    Do While continue

        'StringBuffer = Space$(255)
        StringBuffer = String$(StringBufferSize, vbNullChar)
        'StringBufferSize = Len(StringBuffer)
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
    
    ReadINISection = StringBuffer
End Function
