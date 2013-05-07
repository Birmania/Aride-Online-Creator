VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
    
Public num As Long
    
Public Target As Long
    
Public HP As Long
Public MaxHp As Long
Public MP As Long
Public SP As Long
    
Public Map As Long
Public X As Byte
Public Y As Byte
Public dir As Byte

' Client use only
Public XOffset As Integer
Public YOffset As Integer
Public Moving As Byte
Public MovingSave As Byte
Public Attacking As Byte
Public AttackTimer As Long

Public newNpcDir As New Collection
    
Public Destination As New clsPosition


Private Sub Class_Initialize()
    Call ClearMapNpc(Me)
End Sub
