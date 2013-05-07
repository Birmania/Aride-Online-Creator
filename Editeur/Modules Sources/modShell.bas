Attribute VB_Name = "modShell"
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

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&

' For the compiling sub
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
' End of compiling sub

Public Sub ShellAndWait(ByVal program_name As String)
Dim process_id As Long
Dim process_handle As Long

    ' Start the program.
    On Error GoTo ShellError
    process_id = Shell(program_name, vbNormalFocus)
    On Error GoTo 0

    ' Hide.
    'Me.Visible = False
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    ' Reappear.
    'Me.Visible = True
    Exit Sub

ShellError:
    MsgBox "Error running '" & program_name & _
        "'" & vbCrLf & Err.description
End Sub

Public Sub CompileVB6Project(ByVal VB6VBP As String, ByVal destination As String)
    'Determiner le chemin vers VB6.exe
    Dim VB6Path As String

    VB6Path = String(660, 32)
    Call FindExecutable(VB6VBP, vbNullString, VB6Path)

    Call ShellAndWait("""" & Left$(VB6Path, InStr(VB6Path, Chr$(0)) - 1) & """ /MAKE """ & VB6VBP & """ /outdir """ & destination & """")
End Sub
