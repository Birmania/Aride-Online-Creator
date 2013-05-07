VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdater 
   Caption         =   "Updater"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLaunchGame 
      Caption         =   "Lancer le jeu !"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar downloadBar 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblFileDownload 
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -600
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "frmUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private activated As Boolean

Public Function DownloadFile(srcFileName As String, targetFileName As String)
Dim Size As Long, Remaining As Long, FFile As Integer, Chunk() As Byte
Dim fso As New FileSystemObject
Dim Extension As String

If Trim(targetFileName) <> "" Then

    If FileExist(App.Path & "\" & targetFileName) Then
        Kill App.Path & "\" & targetFileName
    End If

    If Extension <> "" Then
        If Mid(Extension, 1, 1) = "\" Or Mid(Extension, 1, 1) = "/" Then Extension = Mid(Extension, 2, Len(Extension))
        If Mid(Extension, Len(Extension), Len(Extension)) = "\" Or Mid(Extension, Len(Extension), Len(Extension)) = "/" Then Extension = Mid(Extension, 1, Len(Extension) - 1)
        
        If LCase(dir(App.Path & "\" & Extension, vbDirectory)) <> LCase(Extension) Then
            Call MkDir(App.Path & "\" & Extension)
        End If
    Else
        Call MkDir(App.Path & "\" & fso.GetParentFolderName(targetFileName))
        targetFileName = targetFileName
    End If
        
    frmUpdater.lblFileDownload.Caption = fso.GetBaseName(targetFileName)
    frmUpdater.Inet.Execute srcFileName, "GET"

    Do While frmUpdater.Inet.StillExecuting
        DoEvents
    Loop

    Size = Val(frmUpdater.Inet.GetHeader("Content-Length"))
    If Right$(Left$(frmUpdater.Inet.GetHeader, 12), 3) <> "200" Then
        frmUpdater.Inet.Cancel
        Call Err.Raise(10000, "frmUpdater", "Error : File " & fso.GetBaseName(targetFileName) & " do not exist on remote server. HTTP Header : " & frmUpdater.Inet.GetHeader)
    End If
'    If fso.GetBaseName(targetFileName) = "dx7vb.dll" Then
'        MsgBox Right$(Left$(frmUpdater.Inet.GetHeader, 12), 3)
'    End If
    Remaining = 0
    FFile = FreeFile
    
    If Extension <> "" Then
        Open App.Path & "\" & Extension & "\" & targetFileName For Binary Access Write As #FFile
    Else
        Open App.Path & "\" & targetFileName For Binary Access Write As #FFile
    End If
    Do Until Remaining >= Size
        If Size - Remaining > 1023 Then
            Chunk = frmUpdater.Inet.GetChunk(1024, icByteArray)
            Remaining = Remaining + 1024
        Else
            Chunk = frmUpdater.Inet.GetChunk(Size - Remaining, icByteArray)
            Remaining = Size
        End If
        Put #FFile, , Chunk

                ' - Affiche le statut du téléchargement
     If Size > 0 Then
     frmUpdater.lblPercent.Caption = Int((Remaining / Size) * 100) & " %"
     frmUpdater.downloadBar.Value = Int((Remaining / Size) * 100)
     End If
    Loop
    Close #FFile
    
    DoEvents
End If
End Function

Private Sub cmdLaunchGame_Click()
    Unload Me
    Call LaunchGame
End Sub

Private Sub Form_Activate()
    'Usefull because form is showed before changing is widget.
    'If not doing it, you wouldn't see the frame
    If Not activated Then
        Call CheckVersion
        activated = True
    End If
End Sub

Private Sub Form_Load()
    'frmUpdater.Icon = LoadResPicture("APPICON", vbResIcon)
    Call SetIcon(Me)
    
    'Call CheckVersion
    activated = False
End Sub

Private Sub RemoveFileAndParents(ByVal currentFile As String)
    Dim fso As New FileSystemObject
    Dim directoryToDelete As String

    If FileExist(App.Path & "\" & currentFile) Then
        Kill App.Path & CStr(currentFile)
    End If
    If FileExist(App.Path & "\" & currentFile & ".new") Then
        Kill App.Path & CStr(currentFile & ".new")
    End If
    
    directoryToDelete = fso.GetParentFolderName(App.Path & CStr(currentFile))
    Do While True
        If IsDirectoryEmpty(directoryToDelete) Then
        'fso.DeleteFolder(
    ' OppositeRmTree
            If Not KillFolder(directoryToDelete) Then
                Exit Do
            End If
            'Call fso.DeleteFolder(directoryToDelete, True)
            directoryToDelete = fso.GetParentFolderName(directoryToDelete)
        Else
            Exit Do
        End If
    Loop
    'fso.GetParentFolderName (App.Path & currentFile)
End Sub

Private Sub CheckVersion()
    Dim allPath As New Collection
    Dim Path As Variant
    Dim versionFile As String
    Dim webUpdater As String
    Dim fso As New FileSystemObject
    
    'Dim StringBuffer As String
    Dim oldFiles() As String
    Dim oldFilesNew() As String
    Dim currentFile As Variant
    'Me.Visible = True

    versionFile = App.Path & "\Version.ini"
    If dir$(versionFile) <> vbNullString Then
        oldFiles = ReadINIKeys(versionFile, "VERSION")
        'oldFiles = Split(StringBuffer, vbNullChar)
    End If

    If dir$(versionFile & ".new") <> vbNullString Then
        oldFilesNew = ReadINIKeys(versionFile & ".new", "VERSION")
        'oldFilesNew = Split(StringBuffer, vbNullChar)
        
        ' Go back to a stable state
        If GetArraySize(oldFilesNew) > 0 Then
            For Each currentFile In oldFilesNew
                If fso.GetFileName(currentFile) <> App.EXEName & ".exe" Then
                    If Not IsInArray(currentFile, oldFiles) Then
                        Call RemoveFileAndParents(currentFile)
                    End If
                End If
            Next
        End If
    End If

    'Call addDirectories(App.Path, allPath)

    'webUpdater = GetVar(App.path & "\Config\Updater.ini", "UPDATER", "WebUpdater")
    If Not FileExist(App.Path & "\Config\Updater.ini") Then
        webUpdater = "http://aride-online.com/client/"
    Else
        webUpdater = ReadINI("UPDATER", "WebUpdater", App.Path & "\Config\Updater.ini")
    End If

    DownloadFile (webUpdater & "Version.ini"), "Version.ini.new"

    'Debug.Print ReadINI("VERSION", 0, App.path & "\Version.ini")
    
    
'    Dim continue As Boolean
'    Dim StringBuffer As String
'    Dim StringBufferSize As Long
'    Dim returnCode As Long
'
'    continue = True
'    StringBufferSize = 256
'
'    Do While continue
'
'        'StringBuffer = Space$(255)
'        StringBuffer = String$(StringBufferSize, vbNullChar)
'        'StringBufferSize = Len(StringBuffer)
'        returnCode = GetPrivateProfileString("VERSION", vbNullString, "", StringBuffer, StringBufferSize, App.Path & "\Version.ini")
'
'        If returnCode = StringBufferSize - 2 Then
'            StringBufferSize = StringBufferSize + 256
'        Else
'            If returnCode = 0 Then
'                StringBuffer = ""
'            Else
'                StringBuffer = Left$(StringBuffer, returnCode - 1)
'            End If
'            continue = False
'        End If
'    Loop
    Dim files() As String
    Dim isClientChanged As Boolean
    isClientChanged = False
    files = ReadINIKeys(versionFile & ".new", "VERSION")
    'files = Split(StringBuffer, vbNullChar)
    
    ' Remove old files
    If GetArraySize(oldFiles) > 0 Then
        For Each currentFile In oldFiles
            If Not IsInArray(currentFile, files) Then
                Call RemoveFileAndParents(currentFile)
            End If
        Next
    End If
    
'    If GetArraySize(oldFilesNew) > 0 Then
'        For Each currentFile In oldFilesNew
'            If Not IsInArray(currentFile, files) Then
'                Call RemoveFileAndParents(currentFile)
'            End If
'        Next
'    End If
    
    Dim currentMD5 As String
    For Each currentFile In files
        currentMD5 = ReadINI("VERSION", CStr(currentFile), versionFile & ".new")
        If currentMD5 <> "variable" Then
            If currentMD5 <> MD5File(App.Path & CStr(currentFile)) Then
                DownloadFile (webUpdater & currentFile), Right$(CStr(currentFile) & ".new", Len(CStr(currentFile) & ".new") - 1)
    
                If fso.GetFileName(currentFile) = App.EXEName & ".exe" Then
                    isClientChanged = True
                Else
                    'fso.GetFile(App.path & CStr(currentFile)).Attributes = Normal
                    'fso.GetFolder(fso.GetParentFolderName(App.path & CStr(currentFile))).Attributes = Normal
    
                    Call fso.CopyFile(App.Path & CStr(currentFile) & ".new", App.Path & CStr(currentFile), True)
                    'Kill App.path & CStr(currentFile)
                    'Call fso.MoveFile(App.path & CStr(currentFile) & ".new", App.path & CStr(currentFile))
                    'Name App.path & CStr(currentFile) & ".new" As App.path & CStr(currentFile)
                    'FileCopy App.path & CStr(currentFile) & ".new", App.path & CStr(currentFile)
                    Kill App.Path & CStr(currentFile) & ".new"
                End If
            End If
        End If
    Next
    
'    If dir$(versionFile) <> vbNullString Then
'        Kill versionFile
'    End If
    'fso.GetFolder(fso.GetParentFolderName(versionFile)).Attributes = Normal
    'fso.GetFile(versionFile).Attributes = Normal
    Call fso.CopyFile(versionFile & ".new", versionFile, True)
    Kill versionFile & ".new"
    
    Call RegisterDLLs
    
    If isClientChanged Then
        'Call Shell(App.EXEName & ".exe")
        Call SelfDelete
    Else
        Me.cmdLaunchGame.Enabled = True
        Me.lblFileDownload.Caption = "Prêt à jouer !"
    End If
End Sub

Private Sub SelfDelete()
'    On Error GoTo ErrorHandler
    Open App.Path + "\" + App.EXEName + ".bat" For Output As #1
        'Print #1, "sleep 10"
        Print #1, ":loop"
        Print #1, "If Not exist " & Chr(34) & App.Path + "\" + App.EXEName + ".exe" & Chr(34); " goto :ends"
        Print #1, "del " & Chr(34) & App.Path + "\" + App.EXEName + ".exe" & Chr(34)
        Print #1, "goto loop"
        Print #1, ":ends"
        Print #1, "rename " & Chr(34) & App.Path + "\" + App.EXEName + ".exe.new" & Chr(34) & " " & Chr(34) & App.EXEName + ".exe" & Chr(34)
        Print #1, "start /d " & Chr(34) & App.Path & Chr(34) & " " & App.EXEName + ".exe"
        Print #1, "del " & Chr(34) & App.Path + "\" + App.EXEName + ".bat" & Chr(34)
        Close #1
 
    'Then call shell function.
       Call Shell(App.Path + "\" + App.EXEName + ".bat", vbHide)
       'Notice the End statement.  This makes it all work
       End
'ErrorHandler:
    'Catch any errors here if something goes wrong.
End Sub

Private Sub addDirectories(ByVal directory As String, allPath As Collection)
    Dim i As Integer
    Dim directories() As String

    Call addFilesOfDirectory(directory, allPath)
    
    Call ListDir(directory, directories)
    For i = 0 To GetArraySize(directories) - 1
        Call addDirectories(directory & "\" & directories(i), allPath)
    Next i
End Sub

Private Sub addFilesOfDirectory(ByVal directory As String, allPath As Collection)
    Dim files() As String
    Dim i As Integer
    
    Call ListFiles(directory, files)

    For i = 0 To GetArraySize(files) - 1
        allPath.Add (directory & "\" & files(i))
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Error Resume Next
    'frmUpdater.Inet.Execute "CLOSE"
    'frmUpdaterClosed = True
    ' Fermeture forcé de l'application si l'utilisateur quitte pendant un DL (Sinon, problème de "thread")
    If frmUpdater.Inet.StillExecuting Then
        frmUpdater.Inet.Cancel
        End
    End If
End Sub
