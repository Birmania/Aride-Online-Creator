VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
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

Private Buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Flush
    
    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "Class_Initialize", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim Buffer(0 To BufferSize)
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "PreAllocate", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub Allocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve Buffer(0 To BufferSize)
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "Allocate", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub Flush()
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim Buffer(0)
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "Flush", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub Trim()
Dim C As Long

    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= Count Then
        Flush
    End If
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "Trim", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub WriteByte(ByVal nByte As Byte)
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead > BufferSize Then Allocate 1
 
    Buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "WriteByte", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
        
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory Buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "WriteBytes", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead + 1 > BufferSize Then Allocate 2
        
    CopyMemory Buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "WriteInteger", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub WriteLong(ByVal nLong As Long)
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory Buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "WriteLong", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteInteger sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory Buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
    
'    ' Error handler
'    Exit Sub
'ErrorHandler:
'    HandleError "WriteString", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Sub
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = Buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ReadByte", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ReadBoolean(Optional MoveReadHead As Boolean = True) As Byte
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead > BufferSize Then Exit Function
    
    If Buffer(ReadHead) = 1 Then
        ReadBoolean = True
    Else
        ReadBoolean = False
    End If
    
    If MoveReadHead Then ReadHead = ReadHead + 1
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ReadBoolean", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim data() As Byte
    
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
        
    ReDim data(nLength - 1)
        
    CopyMemory data(0), Buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = data
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ReadBytes", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead + 1 > BufferSize Then Exit Function
    
    CopyMemory ReadInteger, Buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ReadInteger", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, Buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ReadLong", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
    
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    sLength = ReadInteger(False)
    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), Buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ReadString", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function Count() As Long
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Count = (UBound(Buffer) - LBound(Buffer)) + 1
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "Count", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function Length() As Long
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Length = Count - ReadHead
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "Length", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ToArray() As Byte()
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ToArray = Buffer()
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ToArray", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Function ToString() As String
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ToString = StrConv(Buffer, vbUnicode)
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "ToString", "clsBuffer", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

