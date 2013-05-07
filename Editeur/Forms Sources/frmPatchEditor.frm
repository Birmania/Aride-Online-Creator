VERSION 5.00
Begin VB.Form frmPatchEditor 
   Caption         =   "Editeur de patch"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Générer une nouvelle version"
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdCreateNewVersion 
         Caption         =   "Créer"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton cmdBrowseDestinationDirectory 
         Caption         =   "Parcourir..."
         Height          =   375
         Left            =   7200
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdBrowseSourceDirectory 
         Caption         =   "Parcourir..."
         Height          =   375
         Left            =   7200
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDestinationDirectory 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   6975
      End
      Begin VB.TextBox txtSourceDirectory 
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Générer un Version.ini"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   8895
      Begin VB.CommandButton cmdBrowseDirectory 
         Caption         =   "Parcourir..."
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Générer"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtDirectory 
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmPatchEditor"
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

Private Sub cmdBrowseDestinationDirectory_Click()
    Dim dir As String
    
    dir = BrowseForFolder(Me, "Select a directory", txtDestinationDirectory.Text)
    If Len(dir) = 0 Then Exit Sub
    txtDestinationDirectory.Text = dir
End Sub

Private Sub cmdBrowseDirectory_Click()
    Dim dir As String
    
    dir = BrowseForFolder(Me, "Select a directory", txtDirectory.Text)
    If Len(dir) = 0 Then Exit Sub
    txtDirectory.Text = dir
End Sub

Private Sub cmdBrowseSourceDirectory_Click()
    Dim dir As String
    
    dir = BrowseForFolder(Me, "Select a directory", txtSourceDirectory.Text)
    If Len(dir) = 0 Then Exit Sub
    txtSourceDirectory.Text = dir
End Sub

Private Sub cmdCreateNewVersion_Click()
    Dim fso
    Dim productionDirectory As String
    Dim directoryToCrypt As String
    Dim directoryNotToCrypt As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    productionDirectory = txtDestinationDirectory.Text & "\production"
    directoryToCrypt = txtDestinationDirectory.Text & "\to_crypt"
    directoryNotToCrypt = txtDestinationDirectory.Text & "\not_to_crypt"

    fso.DeleteFolder productionDirectory, True
    Call MkDir(productionDirectory)

    'Getting all native sources
    Call CopyFolder(txtSourceDirectory.Text & "\Class Sources", App.Path & "\Tools\instrumentation\input\Class Sources")
    Call CopyFolder(txtSourceDirectory.Text & "\Forms Sources", App.Path & "\Tools\instrumentation\input\Forms Sources")
    Call CopyFolder(txtSourceDirectory.Text & "\Modules Sources", App.Path & "\Tools\instrumentation\input\Modules Sources")
    Call FileCopy(txtSourceDirectory.Text & "\Client.vbp", App.Path & "\Tools\instrumentation\input\Client.vbp")
    Call FileCopy(txtSourceDirectory.Text & "\Aride.RES", App.Path & "\Tools\instrumentation\input\Aride.RES")
    Call CopyFolder(App.Path & "\Tools\instrumentation\input", txtDestinationDirectory.Text & "\native")
    
    'Instrumentation
    Dim PythonPath As String
    PythonPath = String(660, 32)
    Call FindExecutable(App.Path & "\Tools\instrumentation\add_line_numbers.py", vbNullString, PythonPath)
    Call ShellAndWait("""" & Left$(PythonPath, InStr(PythonPath, Chr$(0)) - 1) & """ """ & App.Path & "\Tools\instrumentation\add_line_numbers.py""")
    Call CopyFolder(App.Path & "\Tools\instrumentation\output", txtDestinationDirectory.Text & "\instrumented")
    
    'Compilation of instrumented code
    On Error GoTo ErrorHandler
    Call CompileVB6Project(App.Path & "\Tools\instrumentation\output\Client.vbp", App.Path & "\Tools\instrumentation\output\")
    On Error GoTo 0
    
    Call FileCopy(App.Path & "\Tools\instrumentation\output\Client.exe", txtDestinationDirectory.Text & "\production\Client.exe")
    
    ' Copy and crypt files to crypt
    Dim allPath As New Collection
    Dim Path As Variant
    Dim currentDestination As String
    
    Call addDirectories(directoryToCrypt, allPath)
    
    For Each Path In allPath
        currentDestination = productionDirectory & Right$(Path, Len(Path) - Len(directoryToCrypt))

        Call MkDir(fso.GetParentFolderName(currentDestination))

        Call CryptAndSaveTo(Path, currentDestination)
    Next Path
    
    'Copy all non-crypted files
    Set allPath = Nothing
    Set allPath = New Collection
    
    Call addDirectories(directoryNotToCrypt, allPath)
    For Each Path In allPath
        currentDestination = productionDirectory & Right$(Path, Len(Path) - Len(directoryNotToCrypt))

        Call MkDir(fso.GetParentFolderName(currentDestination))
        
        Call FileCopy(Path, currentDestination)
    Next Path
    
    'Create initial INI version (without variables)
    Call CreateINIVersion(productionDirectory)
    
    'Fill version INI with variables
    Dim StringBuffer As String
    Dim Files() As String
    Dim File As Variant
    StringBuffer = ReadINISection(txtDestinationDirectory.Text & "\variable.txt", "VERSION")
    Files = Split(StringBuffer, vbNullChar)
    For Each File In Files
        WriteINI "VERSION", CStr(File), "variable", txtDestinationDirectory.Text & "\production\Version.ini"
    Next File
    
    MsgBox "Nouvelle version crée !"
    Exit Sub
ErrorHandler:
    MsgBox "Erreur dans la compilation !"
    Exit Sub
End Sub

Private Sub cmdGenerate_Click()

    
    'Call ListFiles(txtDirectory.Text, files)
    'Call ListDir(txtDirectory.Text, directories)
           
    Call CreateINIVersion(txtDirectory.Text)

   'Debug.Print directories(0)
End Sub

Private Sub CryptAndSaveTo(ByVal source As String, ByVal destination As String)
    Dim iNumFichier As Long
    Dim sBufferImage As String
    Dim oBufferImage As String
    
    iNumFichier = FreeFile
    
    Open source For Binary As #iNumFichier
        sBufferImage = Space(LOF(iNumFichier))
        Get #iNumFichier, , sBufferImage
    Close #iNumFichier
    
    oBufferImage = CryptData(sBufferImage, "test")
    
    Open destination For Binary As #iNumFichier
        Put #iNumFichier, , oBufferImage
    Close #iNumFichier
End Sub

Private Sub CreateINIVersion(ByVal directory As String)
    Dim allPath As New Collection
    Dim Path As Variant
    Dim versionFile As String

    versionFile = directory & "\Version.ini"
    If dir$(versionFile) <> vbNullString Then
        Kill versionFile
    End If
    
    Call addDirectories(directory, allPath)
    
    For Each Path In allPath
        WriteINI "VERSION", Right(CStr(Path), Len(CStr(Path)) - Len(directory)), MD5File(CStr(Path)), versionFile
    Next Path
End Sub

Private Sub addDirectories(ByVal directory As String, allPath As Collection)
    Dim I As Integer
    Dim directories() As String

    Call addFilesOfDirectory(directory, allPath)
    
    Call ListDir(directory, directories)
    For I = 0 To GetArraySize(directories) - 1
        Call addDirectories(directory & "\" & directories(I), allPath)
    Next I
    
'    Call addFilesOfDirectory(directory, allPath)
End Sub

Private Sub addFilesOfDirectory(ByVal directory As String, allPath As Collection)
    Dim Files() As String
    Dim I As Integer
    
    Call ListFiles(directory, Files)

    For I = 0 To GetArraySize(Files) - 1
        allPath.add (directory & "\" & Files(I))
    Next I
End Sub

Private Sub Form_Load()
    txtDirectory.Text = CurDir
End Sub
