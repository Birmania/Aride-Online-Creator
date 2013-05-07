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

Public Declare Function IsUserAnAdmin Lib "shell32" () As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" ( _
     ByRef lpExecInfo As SHELLEXECUTEINFO) As Long

Private Const SW_NOSHOW = 0
Private Const SW_SHOWNORMAL = 1
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40
Private SEI As SHELLEXECUTEINFO

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
        
Private Const INFINITE = &HFFFFFFFF
        
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function ExecuteCommandAsAdmin(ByVal command As String) As Long
    ExecuteCommandAsAdmin = ExecuteApplicationAsAdmin("cmd", "/c " & command)
'    SEI.cbSize = Len(SEI)
'    SEI.lpVerb = "runas"
'    SEI.lpFile = "cmd"
'    SEI.lpParameters = "/c " & command
'    'SEI.lpParameters = "/c echo 'ls'"
'    SEI.nShow = SW_NOSHOW
'    SEI.fMask = SEE_MASK_NOCLOSEPROCESS
'
'    ShellExecuteEx SEI
'    'MsgBox SEI.hProcess
'    'Dim rc As Long
'    WaitForSingleObject SEI.hProcess, INFINITE
'    Call GetExitCodeProcess(SEI.hProcess, ExecuteCommandAsAdmin)
'    'Call GetExitCodeProcess(SEI.hProcess, rc)
'    'MsgBox rc
End Function

Public Function ExecuteApplicationAsAdmin(ByVal application As String, ByVal parameters As String) As Long
    SEI.hwnd = hwnd
    SEI.cbSize = Len(SEI)
    SEI.lpVerb = "runas"
    SEI.lpFile = application
    SEI.lpParameters = parameters
    'SEI.lpParameters = "/c echo 'ls'"
    SEI.nShow = SW_NOSHOW
    SEI.fMask = SEE_MASK_NOCLOSEPROCESS

    ShellExecuteEx SEI
    'MsgBox SEI.hProcess
    'Dim rc As Long
    WaitForSingleObject SEI.hProcess, INFINITE
    Call GetExitCodeProcess(SEI.hProcess, ExecuteApplicationAsAdmin)
    'Call GetExitCodeProcess(SEI.hProcess, rc)
    'MsgBox rc
    Call CloseHandle(SEI.hProcess)
End Function
