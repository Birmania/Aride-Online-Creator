VERSION 5.00
Begin VB.Form frmDreamEditor 
   Caption         =   "Editer un rêve"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textDreamName 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdBeginningMap 
      Caption         =   "Définir comme carte de départ"
      Height          =   855
      Left            =   4080
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.ListBox lstMaps 
      Height          =   1425
      Left            =   5640
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.ListBox lstDreamMaps 
      Height          =   1425
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblBeginningPosition 
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Nom du rêve :"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblBeginningMap 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Cartes :"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Carte de départ :"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDreamEditor"
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

Private Sub cmdBeginningMap_Click()
    Dim beginningX, beginningY As String

    If lstDreamMaps.ListIndex >= 0 Then
        beginningX = "not numeric"
        Do While Not IsNumeric(beginningX) Or Val(beginningX) < 0
            beginningX = InputBox("Position X de départ ?")
            If StrPtr(beginningX) = 0 Then
                Exit Sub
            End If
        Loop
        beginningY = "not numeric"
        Do While Not IsNumeric(beginningY) Or Val(beginningY) < 0
            beginningY = InputBox("Position Y de départ ?")
            If StrPtr(beginningY) = 0 Then
                Exit Sub
            End If
        Loop
        lblBeginningMap.Caption = lstDreamMaps.ItemData(lstDreamMaps.ListIndex)
        lblBeginningPosition.Caption = beginningX & " : " & beginningY
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
    Call frmIndex.Load_Dreams
End Sub

Private Sub cmdOk_Click()
    Dim position() As String
'    Dreams(0).name = "Salut !"
'    Dreams(0).Beginning = 50
'    ReDim Dreams(0).maps(0 To 1)
'    Dreams(0).maps(0) = 50
'    Dreams(0).maps(1) = 51
'
'    Call SendSaveDream(0)
    position = Split(lblBeginningPosition.Caption, " : ")
    If lblBeginningMap.Caption = "-1" Or position(0) = "-1" Or position(1) = "-1" Then
        MsgBox "Erreur : Position de départ non correcte."
        Exit Sub
    End If

    Dreams(EditorIndex).name = textDreamName.Text
    Dreams(EditorIndex).beginning.Map = lblBeginningMap.Caption
    
    Dreams(EditorIndex).beginning.position.x = position(0)
    Dreams(EditorIndex).beginning.position.y = position(1)
    If lstDreamMaps.ListCount > 0 Then
        ReDim Dreams(EditorIndex).maps(0 To (lstDreamMaps.ListCount - 1))
        For I = 0 To lstDreamMaps.ListCount - 1
            Dreams(EditorIndex).maps(I) = lstDreamMaps.ItemData(I)
        Next I
    Else
        Erase Dreams(EditorIndex).maps
    End If

    Call SendSaveDream(EditorIndex)
    
    Unload Me
    Call frmIndex.Load_Dreams
End Sub

Private Sub Form_Load()
    Dim I, j As Integer
    Dim toAdd As Boolean

    ' Init the lists
    lstMaps.Clear
    lstDreamMaps.Clear
    textDreamName.Text = Trim$(Dreams(EditorIndex).name)
    lblBeginningMap.Caption = Dreams(EditorIndex).beginning.Map
    lblBeginningPosition.Caption = Dreams(EditorIndex).beginning.position.x & " : " & Dreams(EditorIndex).beginning.position.y
    For I = 0 To MAX_MAPS
        toAdd = True
        For j = 0 To GetDreamNbMaps(EditorIndex) - 1
            If Dreams(EditorIndex).maps(j) = I Then
                toAdd = False
                Exit For
            End If
        Next j
        
        If toAdd Then
            lstMaps.AddItem I & " : " & Map(I).name
            lstMaps.ItemData(lstMaps.ListCount - 1) = I
        Else
            lstDreamMaps.AddItem I & " : " & Map(I).name
            lstDreamMaps.ItemData(lstDreamMaps.ListCount - 1) = I
        End If
    Next I
    
End Sub

Private Sub lstDreamMaps_DblClick()
    Dim I As Integer

    For I = 0 To lstMaps.ListCount
        If lstMaps.ItemData(I) > lstDreamMaps.ItemData(lstDreamMaps.ListIndex) Then
            Exit For
        End If
    Next I
    
    If lstDreamMaps.ItemData(lstDreamMaps.ListIndex) = lblBeginningMap.Caption Then
        lblBeginningMap.Caption = -1
        lblBeginningPosition.Caption = "-1 : -1"
    End If
    
    Call lstMaps.AddItem(lstDreamMaps.List(lstDreamMaps.ListIndex), I)
    lstMaps.ItemData(I) = lstDreamMaps.ItemData(lstDreamMaps.ListIndex)
    lstDreamMaps.RemoveItem (lstDreamMaps.ListIndex)
    
End Sub

Private Sub lstMaps_DblClick()
    Dim I As Integer
    
    For I = 0 To lstDreamMaps.ListCount - 1
        If lstDreamMaps.ItemData(I) > lstMaps.ItemData(lstMaps.ListIndex) Then
            Exit For
        End If
    Next I
    Call lstDreamMaps.AddItem(lstMaps.List(lstMaps.ListIndex), I)
    lstDreamMaps.ItemData(I) = lstMaps.ItemData(lstMaps.ListIndex)
    lstMaps.RemoveItem (lstMaps.ListIndex)
End Sub
