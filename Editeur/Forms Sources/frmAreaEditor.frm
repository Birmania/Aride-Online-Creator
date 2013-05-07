VERSION 5.00
Begin VB.Form frmAreaEditor 
   Caption         =   "Edition de zones"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "Valider"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fréquence des tempêtes (0<fréquence<1)"
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3375
      Begin VB.TextBox txtThunderingFrequency 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtRainingFrequency 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSnowingFrequency 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSandStormFrequency 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Orage :"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Pluie :"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Neige :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Sable :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nom :"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAreaEditor"
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

Private Sub cmdCancel_Click()
    frmAreaEditor.Visible = False
End Sub

Private Sub cmdValidate_Click()
    Dim AreaIndex As Integer
    
    'Controle des frequences
    frmAreaEditor.txtSandStormFrequency.Text = Replace(frmAreaEditor.txtSandStormFrequency.Text, ".", ",")
    frmAreaEditor.txtSnowingFrequency.Text = Replace(frmAreaEditor.txtSnowingFrequency.Text, ".", ",")
    frmAreaEditor.txtRainingFrequency.Text = Replace(frmAreaEditor.txtRainingFrequency.Text, ".", ",")
    frmAreaEditor.txtThunderingFrequency.Text = Replace(frmAreaEditor.txtThunderingFrequency.Text, ".", ",")
    
    If (IsNumeric(frmAreaEditor.txtSandStormFrequency.Text) And IsNumeric(frmAreaEditor.txtSnowingFrequency.Text) _
    And IsNumeric(frmAreaEditor.txtRainingFrequency.Text) And IsNumeric(frmAreaEditor.txtThunderingFrequency.Text)) Then
        Dim total As Single
        total = CSng(frmAreaEditor.txtSandStormFrequency.Text) + CSng(frmAreaEditor.txtSnowingFrequency.Text) + _
        CSng(frmAreaEditor.txtRainingFrequency.Text) + CSng(frmAreaEditor.txtThunderingFrequency.Text)
        If Not (0 <= total And total <= 1) Then
            MsgBox "Erreur, la frequence doit être comrpise entre 0 et 1"
            GoTo WrongInput
        End If
    Else
        MsgBox "Erreur, la frequence doit être comrpise entre 0 et 1"
        GoTo WrongInput
    End If
    
    AreaIndex = frmAreaEditor.txtName.DataField

    Areas(AreaIndex).name = frmAreaEditor.txtName.Text

'    frequency = Format(frmAreaEditor.txtSandStormFrequency.Text, "##,##0.00"
    Areas(AreaIndex).SandStormFrequency = CSng(frmAreaEditor.txtSandStormFrequency.Text)
    Areas(AreaIndex).SnowingFrequency = CSng(frmAreaEditor.txtSnowingFrequency.Text)
    Areas(AreaIndex).RainingFrequency = CSng(frmAreaEditor.txtRainingFrequency.Text)
    Areas(AreaIndex).ThunderingFrequency = CSng(frmAreaEditor.txtThunderingFrequency.Text)

    Call SaveArea(AreaIndex)
    Call frmMirage.Editeurarea_Click
    frmAreaEditor.Visible = False

    Exit Sub
WrongInput:
End Sub
