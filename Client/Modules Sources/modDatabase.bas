Attribute VB_Name = "modDatabase"
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

Public SOffsetX As Integer
Public SOffsetY As Integer

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then StripTerminator = Left$(strString, intZeroPos - 1) Else StripTerminator = strString
End Function

Function FileExist(ByVal FileName As String) As Boolean
On Error GoTo er:
    If dir$(FileName) = vbNullString Then FileExist = False Else FileExist = True
    Exit Function
er:
FileExist = False
End Function

Sub SaveLocalMap(ByVal mapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & mapNum & ".aoc"
                            
    f = FreeFile
    If dir$(FileName) <> vbNullString Then
        Kill FileName
    End If
    Open FileName For Binary As #f
        'Put #f, , Map(mapNum)
        Put #f, , Map
    Close #f
End Sub

Sub LoadItem(ByVal itemNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Items\item" & itemNum & ".aoo"

    If Not FileExist(FileName) Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , item(itemNum)
    Close #f
End Sub

Sub LoadCraft(ByVal craftNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Crafts\Craft" & craftNum & ".aop"

    If Not FileExist(FileName) Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Crafts(craftNum)
    Close #f
End Sub

Sub LoadSkill(ByVal skillNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Skills\skill" & skillNum & ".aos"

    If Not FileExist(FileName) Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , skill(skillNum)
    Close #f
End Sub

Sub LoadNpc(ByVal npcNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Npcs\npc" & npcNum & ".aon"

    If Not FileExist(FileName) Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Npc(npcNum)
    Close #f
End Sub

Sub LoadMap(ByVal mapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & mapNum & ".aoc"

    If Not FileExist(FileName) Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Map
    Close #f
    
End Sub

Sub MoveForm(f As Form, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GlobalX As Integer
Dim GlobalY As Integer

GlobalX = f.Left
GlobalY = f.Top

If Button = 1 Then f.Left = GlobalX + X - SOffsetX: f.Top = GlobalY + Y - SOffsetY
End Sub

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext, ByVal erLineNumber, ByVal stackTraceEnded As Boolean)
Dim FileName As String
Dim errorFile As Integer

    Call MkDir(App.Path & "\Logs")

    errorFile = FreeFile

    Open ErrorLogFile For Append As #errorFile
        Print #errorFile, "The following error occured at '" & procName & "' line '" & erLineNumber & "' in '" & contName & "'."
        Print #errorFile, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #errorFile, ""
    Close #errorFile
    
    'To Remove
    If stackTraceEnded Then
        If Not frmReport.wasShown Then
            Call frmReport.Show(1)
            Call GameDestroy
        End If
    Else
        Err.Raise erNumber, procName, "--" & erDesc
    End If
    
    Exit Sub
End Sub
