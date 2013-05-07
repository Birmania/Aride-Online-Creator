Attribute VB_Name = "modClientTCP"
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

Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long

Const IPPROTO_TCP = 6
Const TCP_NODELAY = &H1&

Public InGame As Boolean
Public TradePlayer As Long
Public nbPartyPlayer As Integer
Private MapNumS As Long

' Server adress variables
Public GAME_IP As String
Public GAME_PORT As Long

Public timerSend As Long

' Client adress variables
Public CLIENT_PORT As Long

Private PlayerBuffer As clsBuffer
Private PlayerBufferTCP As clsBuffer

Public OpenPort As Long

Sub TcpInit()
    ' If debug mode, handle error then exit out
  
    Set PlayerBuffer = New clsBuffer
    Set PlayerBufferTCP = New clsBuffer

End Sub

Sub TcpDestroy()
    frmMirage.SocketTCP.Close
    
    If frmMainMenu.fraPers.Visible Then frmMainMenu.fraPers.Visible = False
    If frmMainMenu.fraLogin.Visible Then frmMainMenu.fraLogin.Visible = False
End Sub

Function CheckServerStatus() As Boolean
    Dim PortFind As Boolean
    Dim i, J As Integer
    Dim IPs() As String
    Dim Packet As clsBuffer
    
' connect
    frmMirage.SocketUDPSend.RemoteHost = GAME_IP
    frmMirage.SocketUDPSend.RemotePort = GAME_PORT

    frmMirage.SocketTCP.Close
    frmMirage.SocketTCP.RemoteHost = GAME_IP
    frmMirage.SocketTCP.RemotePort = GAME_PORT + 1

    CheckServerStatus = False

    If ConnectToServer = True Then
        CheckServerStatus = True
    End If
End Function

Public Sub IncomingTCPData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    frmMirage.SocketTCP.GetData Buffer, vbUnicode, DataLength
    
    PlayerBufferTCP.WriteBytes Buffer()
    
    If PlayerBufferTCP.Length >= 4 Then pLength = PlayerBufferTCP.ReadLong(False)
    
    Do While pLength > 0 And pLength <= PlayerBufferTCP.Length - 4
        If pLength <= PlayerBufferTCP.Length - 4 Then
            PlayerBufferTCP.ReadLong
            HandleData PlayerBufferTCP.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBufferTCP.Length >= 4 Then pLength = PlayerBufferTCP.ReadLong(False)
    Loop
    PlayerBufferTCP.Trim
    DoEvents
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    frmMirage.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
End Sub

Sub HandleData(ByRef data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        GameDestroy
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        MsgBox "Erreur de packet : " + Trim(data())
        GameDestroy
        Exit Sub
    End If
    Debug.Print "message : " & MsgType
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub


Function ConnectToServer() As Boolean
Dim Wait As Long
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then ConnectToServer = True: Exit Function
    
    Wait = GetTickCount
    frmMirage.SocketTCP.Connect
    
    ' Wait until connected or 7 seconds have passed and report the server being down
    Do While (Not IsConnected) And GetTickCount <= Wait + 2000
        'Sleep 100
        'frmMirage.SocketTCP.Connect
        DoEvents
    Loop
    
    If IsConnected Then
        ConnectToServer = True
        Call setsockopt(frmMirage.SocketTCP.SocketHandle, IPPROTO_TCP, TCP_NODELAY, 1, 4)
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.SocketTCP.State = sckConnected Then IsConnected = True Else IsConnected = False
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If GetPlayerName(Index) <> vbNullString Then IsPlaying = True Else IsPlaying = False
End Function

Sub SendDataTCP(ByRef data() As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
            
    Buffer.WriteLong (UBound(data) - LBound(data)) + 1
    Buffer.WriteBytes data()
    On Error Resume Next ' SendData could throw an error if socket was closed in server side
    frmMirage.SocketTCP.SendData Buffer.ToArray()
End Sub

Sub SendData(ByRef data() As Byte)
Dim Buffer As clsBuffer
    
    Call SendDataTCP(data)
End Sub

Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Packet As clsBuffer
    Set Packet = New clsBuffer
    Packet.WriteLong CLogin
    Packet.WriteString name
    Packet.WriteString HashPassword(Password)
    Packet.WriteLong App.Major
    Packet.WriteLong App.Minor
    Packet.WriteLong App.Revision
    Packet.WriteString SEC_CODE1
    Packet.WriteString SEC_CODE2
    Packet.WriteString SEC_CODE3
    Packet.WriteString SEC_CODE4
    
    SendDataTCP Packet.ToArray()
    Set Packet = Nothing
End Sub

Public Function HashPassword(ByVal Password As String)
    ' STUB : Create your own method
    HashPassword = Password
End Function

Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal slot As Long)

End Sub

Sub SendDelChar(ByVal slot As Long)

End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CSayMsg
    Packet.WriteString Text
    SendData Packet.ToArray()
    
    Set Packet = Nothing
End Sub

Sub GuildeMsg(ByVal Text As String)

End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)

End Sub

Sub AdminMsg(ByVal Text As String)
End Sub

Sub SendPlayerMove()
Dim Packet As clsBuffer
    Debug.Print "Send player move : dir " & GetPlayerDir(MyIndex) & " controller : " & movementController & " timer " & Str(GetTickCount)
    Set Packet = New clsBuffer

    Packet.WriteLong CPlayerMove
    Packet.WriteByte movementController
    Packet.WriteByte GetPlayerX(MyIndex)
    Packet.WriteByte GetPlayerY(MyIndex)
    'Packet.WriteByte GetPlayerDir(MyIndex)
    Packet.WriteByte Player(MyIndex).Moving
    SendData Packet.ToArray()

    Set Packet = Nothing
    
    If movementController = 1 Then
        movementController = 0
    End If
End Sub

Sub SendPlayerStopMove()
Dim Packet As clsBuffer
    Debug.Print "Timer d'envoi stop move" & Str(GetTickCount)
    timerSend = GetTickCount
    Set Packet = New clsBuffer

    Packet.WriteLong CPlayerStopMove
    Packet.WriteByte GetPlayerX(MyIndex)
    Packet.WriteByte GetPlayerY(MyIndex)
    SendData Packet.ToArray()

    Set Packet = Nothing
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer

    If (Player(MyIndex).Moving > 0) Then
        Buffer.WriteLong CPlayerDirMove
        Buffer.WriteByte GetPlayerX(MyIndex)
        Buffer.WriteByte GetPlayerY(MyIndex)
    Else
        Buffer.WriteLong CPlayerDir
    End If
    Buffer.WriteByte GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerRequestNewMap()
    ' TODO
End Sub

Sub WarpMeTo(ByVal name As String)
    ' TODO
End Sub

Sub WarpToMe(ByVal name As String)
    ' TODO
End Sub

Sub WarpTo(ByVal mapNum As Long)
    ' TODO
End Sub

Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
    ' TODO
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
    ' TODO
End Sub

Sub SendPlayerInfoRequest(ByVal name As String)
    ' TODO
End Sub

Sub SendKick(ByVal name As String)
    ' TODO
End Sub

Sub SendBan(ByVal name As String)
    ' TODO
End Sub

Sub SendMapRespawn(ByVal numMap As Integer)
    ' TODO
End Sub

Sub UseSkill(ByVal Index As Integer)
    If Player(MyIndex).skill(Index) >= 0 Then
        If Player(MyIndex).AttackTimer < GetTickCount Then
            If Player(MyIndex).Moving = 0 Then
                If skill(Player(MyIndex).skill(Index)).TargetType = 2 Then
                    Call SendUseSkill(Player(MyIndex).skill(Index), -1, -1)
                Else
                    Call SetCursor(App.Path & Rep_Theme & "\skill.ani")
                    Player(MyIndex).castedSpell = Player(MyIndex).skill(Index)
                End If
            Else
                Call AddText("Vous ne pouvez lancer un sort en marchant!", BrightRed)
            End If
        End If
    Else
        Call AddText("Aucuns sort ici.", BrightRed)
    End If
End Sub

Public Sub SendUseSkill(ByVal skillNum As Integer, ByVal CurX As Integer, ByVal CurY As Integer)
    Dim Packet As clsBuffer
    Dim i As Integer

    Set Packet = New clsBuffer
    Packet.WriteLong CUseSkill
    Packet.WriteInteger skillNum
    If skill(skillNum).TargetType = 1 Then
        Packet.WriteByte CurX
        Packet.WriteByte CurY
        Call RestoreCursor
    ElseIf skill(skillNum).TargetType = 0 Then
        Dim Target() As Integer

        Target = FindIndexAtPos(GetPlayerMap(MyIndex), CurX, CurY)
        
        If Target(1) >= 0 Then
            Packet.WriteByte Target(1) 'Target Type
            Packet.WriteInteger Target(0) 'Target index
        Else
            Set Packet = Nothing
            Exit Sub
        End If
        Call RestoreCursor
    ElseIf skill(skillNum).TargetType = 2 Then
        ' Do nothing
    End If
    
    SendData Packet.ToArray()
    Set Packet = Nothing
    
    Player(MyIndex).Attacking = 1
    Player(MyIndex).castedSpell = -1
End Sub

Sub SendUseItem(ByVal InvNum As Long)
' Do not send if item is in cooldown
If IsItemInCooldown(Player(MyIndex).Inv(InvNum).num) Then
    Exit Sub
End If

Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CUseItem
    Packet.WriteByte InvNum
    
    SendData Packet.ToArray()
    
    Set Packet = Nothing
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Amount As Long)
Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CDropItem
    Packet.WriteByte InvNum
    Packet.WriteInteger Amount
    
    SendData Packet.ToArray()
    Set Packet = Nothing
End Sub

Sub SendDestroyItem(ByVal InvNum, ByVal Amount As Long)
    ' TODO
End Sub

Sub SendOnlineList()
    ' TODO
End Sub

Sub SendTradeRequest(ByVal playerIndex As Integer)
    ' TODO
End Sub

Sub SendAcceptTrade()
    ' TODO
End Sub

Sub SendDeclineTrade()
    ' TODO
End Sub

Sub SendRequestParty(ByVal name As String, Optional ByVal partyName As Variant)
Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CRequestParty
    Packet.WriteInteger FindPlayer(name)
    
    If Not IsMissing(partyName) Then
        Packet.WriteString (partyName)
    End If
    
    SendData Packet.ToArray()
    Set Packet = Nothing
End Sub

Sub SendJoinParty()
Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CJoinParty
    
    SendData Packet.ToArray()
    Set Packet = Nothing
End Sub

Sub SendLeaveParty()
Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CLeaveParty
    
    SendData Packet.ToArray()
    Set Packet = Nothing
End Sub

Sub SendSleep()
Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CPlayerSleep
    
    SendData Packet.ToArray()
    Set Packet = Nothing
End Sub


Sub SendRequestLocation()
    ' TODO
End Sub

Sub SendSetPlayerSprite(ByVal name As String, ByVal SpriteNum As Byte)
    ' TODO
End Sub

Sub SendSetPlayerName(ByVal name As String, ByVal NewName As String)
    ' TODO
End Sub

Sub SendSetPlayerstr(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerDef(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerVit(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerMagi(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerPk(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerNiveau(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerExp(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerPoint(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerMaxPv(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendSetPlayerMaxPm(ByVal name As String, ByVal num As Long)
    ' TODO
End Sub

Sub SendGetAdminHelp()
    ' TODO
End Sub
