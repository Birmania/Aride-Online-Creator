VERSION 5.00
Begin VB.Form frmSelectDirection 
   Caption         =   "SourceBorderDirection"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelectDown 
      Caption         =   "Bas"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelectLeft 
      Caption         =   "Gauche"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelectRight 
      Caption         =   "Droite"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelectUp 
      Caption         =   "Haut"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmSelectDirection"
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

Private direction As Byte

Private Sub cmdSelectDown_Click()
    direction = DIR_DOWN
    
    Unload Me
End Sub

Private Sub cmdSelectLeft_Click()
    direction = DIR_LEFT
    
    Unload Me
End Sub

Private Sub cmdSelectRight_Click()
    direction = DIR_RIGHT
    
    Unload Me
End Sub

Private Sub cmdSelectUp_Click()
    direction = DIR_UP
    
    Unload Me
End Sub

Public Function ShowWithResult()
    Me.Show vbModal
    
    ShowWithResult = direction
End Function

