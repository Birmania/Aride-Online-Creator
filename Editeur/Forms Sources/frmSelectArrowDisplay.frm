VERSION 5.00
Begin VB.Form frmSelectArrowDisplay 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   855
   End
   Begin VB.PictureBox picSelect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   2
      ToolTipText     =   "Image qui sera affiché dans l'inventaire "
      Top             =   3360
      Width           =   465
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3240
      LargeChange     =   10
      Left            =   2880
      Max             =   464
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3225
      Left            =   0
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   0
      ToolTipText     =   "Sélectionner une image pour l'objet"
      Top             =   0
      Width           =   2880
   End
End
Attribute VB_Name = "frmSelectArrowDisplay"
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

Private SelectArrowY As Integer

Private Sub Form_Load()
    Call AffSurfPic(DD_ArrowAnim, picSelect, 0, SelectArrowY * PIC_Y)
    Call AffSurfPic(DD_ArrowAnim, picPic, 0, VScroll1.value * PIC_X)
End Sub

Private Sub ok_Click()
    Call AffSurfPic(DD_ArrowAnim, frmItemEditor.picBow, 0, SelectArrowY * PIC_Y)
    frmItemEditor.picBow.DataField = SelectArrowY
    Me.Visible = False
End Sub

Private Sub picPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SelectArrowY = (y \ PIC_Y) + VScroll1.value
    Call AffSurfPic(DD_ArrowAnim, picSelect, 0, SelectArrowY * PIC_Y)
End Sub
