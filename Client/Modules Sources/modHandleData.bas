Attribute VB_Name = "modHandleData"
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

Public timer As Long

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetAddress = FunAddr
    
'    ' Error handler
'    Exit Function
'ErrorHandler:
'    HandleError "GetAddress", "modHandleData", Err.Number, Err.description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
'    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleDataSub(SFindServer) = GetAddress(AddressOf HandleFindServer)
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SErrorLogin) = GetAddress(AddressOf HandleErrorLogin)
    'HandleDataSub(SAllChars) = GetAddress(AddressOf HandleAllChars)
    HandleDataSub(SEndWarp) = GetAddress(AddressOf HandleEndWarp)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SYourIndex) = GetAddress(AddressOf HandleYourIndex)
    HandleDataSub(SLife) = GetAddress(AddressOf HandleLife)
    HandleDataSub(SStamina) = GetAddress(AddressOf HandleStamina)
    HandleDataSub(SSleep) = GetAddress(AddressOf HandleSleep)
    HandleDataSub(SExperience) = GetAddress(AddressOf HandleExperience)
    HandleDataSub(SNextLevel) = GetAddress(AddressOf HandleNextLevel)
    HandleDataSub(SPartyBars) = GetAddress(AddressOf HandlePartyBars)
    HandleDataSub(SPlayerSkills) = GetAddress(AddressOf HandlePlayerSkills)
    HandleDataSub(SPlayerCrafts) = GetAddress(AddressOf HandlePlayerCrafts)
    'HandleDataSub(SPlayerSkill) = GetAddress(AddressOf HandlePlayerSkill)
    HandleDataSub(SInventory) = GetAddress(AddressOf HandleInventory)
    HandleDataSub(SInventorySlot) = GetAddress(AddressOf HandleInventorySlot)
    HandleDataSub(SWeaponSlot) = GetAddress(AddressOf HandleWeaponSlot)
    HandleDataSub(SArmorSlot) = GetAddress(AddressOf HandleArmorSlot)
    HandleDataSub(SHelmetSlot) = GetAddress(AddressOf HandleHelmetSlot)
    HandleDataSub(SShieldSlot) = GetAddress(AddressOf HandleShieldSlot)
    HandleDataSub(SJoinParty) = GetAddress(AddressOf HandleJoinParty)
    HandleDataSub(SLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    HandleDataSub(SPlayerStartMove) = GetAddress(AddressOf HandlePlayerStartMove)
    HandleDataSub(SPlayerStopMove) = GetAddress(AddressOf HandlePlayerStopMove)
    HandleDataSub(SPlayerDirMove) = GetAddress(AddressOf HandlePlayerDirMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcStartMove) = GetAddress(AddressOf HandleNpcStartMove)
    HandleDataSub(SNpcStopMove) = GetAddress(AddressOf HandleNpcStopMove)
    HandleDataSub(SNpcDirMove) = GetAddress(AddressOf HandleNpcDirMove)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPetStartMove) = GetAddress(AddressOf HandlePetStartMove)
    HandleDataSub(SPetStopMove) = GetAddress(AddressOf HandlePetStopMove)
    HandleDataSub(SPetDirMove) = GetAddress(AddressOf HandlePetDirMove)
    HandleDataSub(SPetDir) = GetAddress(AddressOf HandlePetDir)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SChatMsg) = GetAddress(AddressOf HandleChatMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SStatistics) = GetAddress(AddressOf HandleStatistics)
    HandleDataSub(SPetDead) = GetAddress(AddressOf HandlePetDead)
    HandleDataSub(SAreaWeather) = GetAddress(AddressOf HandleAreaWeather)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SMissileAppear) = GetAddress(AddressOf HandleMissileAppear)
    HandleDataSub(SMissileDisappear) = GetAddress(AddressOf HandleMissileDisappear)
    HandleDataSub(SSpawnMapItem) = GetAddress(AddressOf HandleSpawnMapItem)
    HandleDataSub(SDeleteMapItem) = GetAddress(AddressOf HandleDeleteMapItem)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SQuitMap) = GetAddress(AddressOf HandleQuitMap)
    HandleDataSub(SPlayerDead) = GetAddress(AddressOf HandlePlayerDead)
    HandleDataSub(STime) = GetAddress(AddressOf HandleTime)
    HandleDataSub(SConfirmUseItem) = GetAddress(AddressOf HandleConfirmUseItem)
    HandleDataSub(SCancelUseItem) = GetAddress(AddressOf HandleCancelUseItem)
    HandleDataSub(SPlayerStartInfos) = GetAddress(AddressOf HandlePlayerStartInfos)
    HandleDataSub(SPlayerPosition) = GetAddress(AddressOf HandlePlayerPosition)
    HandleDataSub(SDamageDisplay) = GetAddress(AddressOf HandleDamageDisplay)
    HandleDataSub(SGetItemDisplay) = GetAddress(AddressOf HandleGetItemDisplay)
    HandleDataSub(SRequestParty) = GetAddress(AddressOf HandleRequestParty)
End Sub

Sub HandleFindServer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Integer
Dim Buffer As clsBuffer
Dim numServer As Integer
Dim nbPlayer As Integer
Dim maxPlayer As Integer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    nbPlayer = Buffer.ReadInteger
    frmServerChooser.lstServers.List(frmServerChooser.lstServers.ListCount - 1) = frmServerChooser.lstServers.List(frmServerChooser.lstServers.ListCount - 1) & " - Ouvert (" & nbPlayer & "/" & MAX_PLAYERS & ")"

    CHECK_WAIT = False

    Set Buffer = Nothing
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    frmMirage.Visible = False
    frmsplash.Visible = False
    frmMainMenu.Visible = True
    
    Msg = Buffer.ReadString 'Parse(1)
 
    Set Buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Game_Name)
    Call GameDestroy
End Sub

Sub HandleErrorLogin(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    frmsplash.Visible = False
    frmMainMenu.fraLogin.Visible = True
    Call MsgBox(Buffer.ReadString, vbOKOnly, Game_Name)
    
    Set Buffer = Nothing
End Sub

Sub HandleYourIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    MyIndex = Buffer.ReadInteger
    
    Set Buffer = Nothing
End Sub

Sub HandleEndWarp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InitTiles = True

End Sub

Private Sub HandlePlayerPosition(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i, J As Integer
Dim X As Integer
Dim Y As Integer
Dim nbPlayer As Integer
Dim dir As Byte
Dim vitesse As Byte
Dim Buffer As clsBuffer
    timer = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    nbPlayer = Buffer.ReadInteger
    
    For J = 1 To nbPlayer
        i = Buffer.ReadInteger
        X = Buffer.ReadInteger
        Y = Buffer.ReadInteger
        dir = Buffer.ReadByte
        vitesse = Buffer.ReadByte
        
        Call SetPlayerMap(i, GetPlayerMap(MyIndex))
        Call SetPlayerX(i, X)
        Call SetPlayerY(i, Y)
        
        If vitesse > 0 Then
            Call StartMove(i, vitesse, dir)
        Else
            Call SetPlayerDir(i, dir)
        End If
        
        If Pets(i).num > -1 Then
            With Pets(i)
                .X = Buffer.ReadByte
                .Y = Buffer.ReadByte
                .dir = Buffer.ReadByte
                .Map = GetPlayerMap(i)
            End With
    
            vitesse = Buffer.ReadByte
            If vitesse > 0 Then
                Call StartNpcMove(Pets(i), vitesse)
            End If
        End If
    Next J
End Sub

Private Sub StartMove(ByVal Index As Long, ByVal movement As Integer, ByVal direction As Integer)
    If (Player(Index).Destination.X <> -1) Then
        Call SetPlayerX(Index, Player(Index).Destination.X)
        Call SetPlayerY(Index, Player(Index).Destination.Y)
        
        Call ClearPlayerMove(Index)
    End If
    
    Call beginPlayerMovement(Index, movement, direction)
End Sub

Private Sub HandlePlayerStartMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Integer
Dim dir As Byte
Dim n As Byte
Dim Buffer As clsBuffer

    timer = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadInteger
    dir = Buffer.ReadByte
    n = Buffer.ReadByte
    
    If Player(i).Moving > 0 Then
        If GetArraySize(Player(i).newDir) > 0 Then
            Call SetPlayerDir(i, Player(i).newDir(GetArraySize(Player(i).newDir) - 1).dir)
        End If
        
        If Player(i).Destination.X <> -1 Then
            Call SetPlayerX(i, Player(i).Destination.X)
            Call SetPlayerY(i, Player(i).Destination.Y)
        End If
        Call ClearPlayerMove(i)
    End If
    
    Call StartMove(i, n, dir)
End Sub

Private Sub HandleNpcStartMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim mapNpcNum As Integer
    Dim dir As Byte
    Dim movement As Byte
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    mapNpcNum = Buffer.ReadInteger
    dir = Buffer.ReadByte
    movement = Buffer.ReadByte
    
    Set Buffer = Nothing

    Call StartNpcMoveWithVerif(MapNpc(mapNpcNum), movement, dir)
End Sub

Private Sub HandlePetStartMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim petNum As Integer
    Dim dir As Byte
    Dim movement As Byte
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    petNum = Buffer.ReadInteger
    dir = Buffer.ReadByte
    movement = Buffer.ReadByte
    
    Set Buffer = Nothing

    Call StartNpcMoveWithVerif(Pets(petNum), movement, dir)
End Sub

Private Sub HandleNpcStopMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim mapNpcNum As Integer
    Dim newX, newY As Byte
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    mapNpcNum = Buffer.ReadInteger
    newX = Buffer.ReadByte
    newY = Buffer.ReadByte
    
    Call StopNpcMove(MapNpc(mapNpcNum), newX, newY)
End Sub

Private Sub HandlePetStopMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim petNpcNum As Integer
    Dim newX, newY As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    petNpcNum = Buffer.ReadInteger
    newX = Buffer.ReadByte
    newY = Buffer.ReadByte
    
    Call StopNpcMove(Pets(petNpcNum), newX, newY)

    Set Buffer = Nothing
End Sub

Private Sub StopNpcMove(ByRef MapNpc As clsMapNpc, ByVal newX As Byte, ByVal newY As Byte)
    Dim ancientX, ancientY As Byte
    
    ancientX = GetNpcX(MapNpc)
    ancientY = GetNpcY(MapNpc)
    
    Debug.Print "Npc stop move : " & ancientX & "/" & newX & " Y : " & ancientY & "/" & newY & " time : " & GetTickCount
    
    Select Case GetNpcDir(MapNpc)
        Case DIR_UP
            If newY <= ancientY And newX = ancientX Then
                MapNpc.Destination.X = newX
                MapNpc.Destination.Y = newY
            End If

        Case DIR_DOWN
            If newY >= ancientY And newX = ancientX Then
                MapNpc.Destination.X = newX
                MapNpc.Destination.Y = newY
            End If

        Case DIR_LEFT
            If newX <= ancientX And newY = ancientY Then
                MapNpc.Destination.X = newX
                MapNpc.Destination.Y = newY
            End If

        Case DIR_RIGHT
            If newX >= ancientX And newY = ancientY Then
                MapNpc.Destination.X = newX
                MapNpc.Destination.Y = newY
            End If
    End Select

    If MapNpc.Destination.X = -1 Then
        Call SetNpcX(MapNpc, newX)
        Call SetNpcY(MapNpc, newY)
        If MapNpc.newNpcDir.Count > 0 Then
            Call SetNpcDir(MapNpc, MapNpc.newNpcDir.item(1).dir)
        End If
        Call ClearMapNpcMove(MapNpc)
    End If
End Sub

Private Sub ChangeNpcDir(ByRef MapNpc As clsMapNpc, ByVal dir As Byte)
    If MapNpc.Moving > 0 Then
        Call SetNpcX(MapNpc, MapNpc.Destination.X)
        Call SetNpcY(MapNpc, MapNpc.Destination.Y)
        Call ClearMapNpcMove(MapNpc)
    End If
    
    Call SetNpcDir(MapNpc, dir)
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As String
Dim i As Long
Dim NeedMap As Byte
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    IsDead = False
    GettingMap = True
    
    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, -1)
        End If
        Call ClearPlayerMove(i)
        Call ClearMapNpcMove(Pets(i))
    Next

    ' Erase all temporary tile values and npcs if some are already presents (in game)
    If InGame Then
        Call ClearTempTile
        Call ClearMapNpcs
        Call ClearMapItems
        'Call ClearMap ' Attention : Va rendre indisponible MaxMapX et MaxMapY
    End If
    
    ' Get map num
    X = Buffer.ReadInteger
    ' Get revision
    Y = Buffer.ReadString
    Set Buffer = Nothing
    
    Call SetPlayerMap(MyIndex, X)

    If FileExist(App.Path & "\Maps\map" & X & ".aoc") Then
        ' Check to see if the revisions match
        If MD5File(App.Path & "\Maps\map" & X & ".aoc") = Y Then
            ' We do so we dont need the map

            ' Load the map
            Call LoadMap(X)

            Set Buffer = New clsBuffer
            Buffer.WriteLong CNeedMap
            Buffer.WriteByte 0
            SendData Buffer.ToArray()
            Set Buffer = Nothing
            
            Call InitMap
            Call InitPano
            
            GoTo LaunchGame
            'Exit Sub
        End If
    End If
    ' Either the revisions didn't match or we dont have the map, so we need it
    OldMusic = Map.Music
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteByte 1
    SendData Buffer.ToArray()
    Set Buffer = Nothing

LaunchGame:
    If Not InGame Then
        Call beginGame
    End If
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim nbDatas, nbStrings, nbNpcs As Long
Dim nbBorders As Long
Dim mapNum As Integer
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteBytes data()


    mapNum = Buffer.ReadInteger
    
    With Map
        .name = Buffer.ReadString
        .Moral = Buffer.ReadByte
        .Music = Buffer.ReadString
        .BootMap = Buffer.ReadLong
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .Indoors = Buffer.ReadBoolean
        
        ReDim Map.tile(0 To Buffer.ReadInteger - 1, 0 To Buffer.ReadInteger - 1) As TileRec

        For X = 0 To MaxMapX
            For Y = 0 To MaxMapY
                With Map.tile(X, Y)
                    .Ground = Buffer.ReadLong
                    .Mask = Buffer.ReadLong
                    .Anim = Buffer.ReadLong
                    .Mask2 = Buffer.ReadLong
                    .M2Anim = Buffer.ReadLong
                    .Mask3 = Buffer.ReadLong
                    .M3Anim = Buffer.ReadLong
                    .Fringe = Buffer.ReadLong
                    .FAnim = Buffer.ReadLong
                    .Fringe2 = Buffer.ReadLong
                    .F2Anim = Buffer.ReadLong
                    .Fringe3 = Buffer.ReadLong
                    .F3Anim = Buffer.ReadLong
                    .Type = Buffer.ReadByte
                    nbDatas = Buffer.ReadLong
                    If nbDatas > 0 Then
                        ReDim .Datas(0 To nbDatas - 1) As Long
        
                        For i = 0 To UBound(.Datas)
                            .Datas(i) = Buffer.ReadLong
                        Next i
                    End If
                    
                    nbStrings = Buffer.ReadLong
                    If nbStrings > 0 Then
                        ReDim .Strings(0 To nbStrings - 1) As String
                        
                        For i = 0 To UBound(.Strings)
                            .Strings(i) = Buffer.ReadString
                        Next i
                    End If
                    .Light = Buffer.ReadLong

                End With
            Next Y
        Next X
        
        nbNpcs = Buffer.ReadLong
        If nbNpcs > 0 Then
            ReDim .Npcs(0 To nbNpcs - 1) As NpcMapRec
            For X = 0 To GetMapNbNpcs() - 1
                With Map
                    .Npcs(X).id = Buffer.ReadInteger
                    
                    ReDim .Npcs(X).X(0 To Buffer.ReadLong - 1) As Byte
                    For i = 0 To UBound(.Npcs(X).X)
                        .Npcs(X).X(i) = Buffer.ReadByte
                    Next i
                    
                    ReDim .Npcs(X).Y(0 To Buffer.ReadLong - 1) As Byte
                    For i = 0 To UBound(.Npcs(X).Y)
                        .Npcs(X).Y(i) = Buffer.ReadByte
                    Next i
                    
                    .Npcs(X).dir = Buffer.ReadByte
                    .Npcs(X).Hasardp = Buffer.ReadByte
                    .Npcs(X).movementType = Buffer.ReadByte
                End With
            Next X
        End If
        
        .PanoInf = Buffer.ReadString
        .TranInf = Buffer.ReadByte
        .PanoSup = Buffer.ReadString
        .TranSup = Buffer.ReadByte
        .Fog = Buffer.ReadInteger
        .FogAlpha = Buffer.ReadByte
        
        nbBorders = Buffer.ReadLong
        If nbBorders > 0 Then
            ReDim .Borders(0 To nbBorders - 1) As BorderRec

            For X = 0 To nbBorders - 1
                .Borders(X).XSource = Buffer.ReadByte
                .Borders(X).YSource = Buffer.ReadByte
                .Borders(X).DirectionSource = Buffer.ReadByte
                .Borders(X).MapDestination = Buffer.ReadInteger
                .Borders(X).XDestination = Buffer.ReadByte
                .Borders(X).YDestination = Buffer.ReadByte
            Next X
        End If
        
        .Area = Buffer.ReadByte
    End With
    
    Call SaveLocalMap(mapNum)

    Set Buffer = Nothing
    
    Call InitMap
    Call InitPano

    ' Play music
    If OldMusic <> vbNullString Then
        'If Trim$(Map(GetPlayerMap(MyIndex)).Music) = Trim$(Map(OldMap).Music) Then
        If Trim$(Map.Music) = Trim$(OldMusic) Then
            ' Do Nothing
        ElseIf Trim$(Map.Music) <> "Aucune" Then
            Call PlayMidi(App.Path & "\" & Trim$(Map.Music))
        Else
            Call StopMidi
        End If
    Else
        If Trim$(Map.Music) <> "Aucune" Then Call PlayMidi(App.Path & "\" & Trim$(Map.Music)) Else Call StopMidi
    End If
    OldMusic = Map.Music
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i, J As Integer
    Dim Buffer As clsBuffer
    Dim vitesse As Byte
    Dim nbNpc As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    nbNpc = Buffer.ReadByte
    For J = 1 To nbNpc
        i = Buffer.ReadInteger

        Set MapNpc(i) = New clsMapNpc
        ' Mettre les valeurs par défaut dans le NPC (ex : destination)
        
        With MapNpc(i)
            .num = Buffer.ReadInteger
            .Map = Player(MyIndex).Map
            .X = Buffer.ReadByte
            .Y = Buffer.ReadByte
            .HP = Buffer.ReadLong
            .dir = Buffer.ReadByte
        End With
        
        vitesse = Buffer.ReadByte
        If vitesse > 0 Then
            Call StartNpcMove(MapNpc(i), vitesse)
        End If
    Next J
    
    Set Buffer = Nothing
End Sub

Private Sub HandleMapPetData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i, J As Integer
    Dim Buffer As clsBuffer
    Dim vitesse As Byte
    Dim nbPet As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    nbPet = Buffer.ReadInteger
    For J = 1 To nbPet
        i = Buffer.ReadInteger
        With Pets(i)
            .num = Buffer.ReadInteger
            .X = Buffer.ReadByte
            .Y = Buffer.ReadByte
            .HP = Buffer.ReadLong
            .dir = Buffer.ReadByte
            .Map = GetPlayerMap(MyIndex)
        End With

        vitesse = Buffer.ReadByte
        If vitesse > 0 Then
            Call StartNpcMove(Pets(i), vitesse)
        End If
    Next J
    
    Set Buffer = Nothing
End Sub

Private Sub StartNpcMoveWithVerif(ByRef MapNpc As clsMapNpc, ByVal movement As Byte, ByVal dir As Byte)
    With MapNpc
        If MapNpc.Moving > 0 Then 'If not arrived to his destination, teleport
            MapNpc.Destination.X = -1
            MapNpc.Destination.Y = -1
            
            Set MapNpc.newNpcDir = Nothing
            Set MapNpc.newNpcDir = New Collection
        End If
        .dir = dir
        Call StartNpcMove(MapNpc, movement)
    End With
End Sub

Private Sub StartNpcMove(ByRef MapNpc As clsMapNpc, ByVal movement As Byte)
With MapNpc
    .XOffset = 0
    .YOffset = 0
    .Moving = movement
    Debug.Print "Npc Start move : " & .dir & " time : " & GetTickCount

    Select Case .dir
        Case DIR_UP
            .YOffset = PIC_Y
            Call SetNpcY(MapNpc, GetNpcY(MapNpc) - 1)
        Case DIR_DOWN
            .YOffset = PIC_Y * -1
            Call SetNpcY(MapNpc, GetNpcY(MapNpc) + 1)
        Case DIR_LEFT
            .XOffset = PIC_X
            Call SetNpcX(MapNpc, GetNpcX(MapNpc) - 1)
        Case DIR_RIGHT
            .XOffset = PIC_X * -1
            Call SetNpcX(MapNpc, GetNpcX(MapNpc) + 1)
    End Select
    
End With
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleChatMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Long
Dim i As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadInteger
    Call AddText(Msg, Color)

End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim playerIndex As Integer
Dim Msg As String
Dim Color As Long
Dim i As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    playerIndex = Buffer.ReadInteger
    Msg = Buffer.ReadString

    If playerIndex > 0 Then
        Bubble(playerIndex).Text = Msg
        Bubble(playerIndex).Created = GetTickCount()
    
        Msg = Trim$(Player(playerIndex).name) & ":" & Msg
    End If
    Call AddText(Msg, 15)
End Sub

Private Sub HandleStatistics(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    Player(MyIndex).Str = Buffer.ReadInteger
    Player(MyIndex).Def = Buffer.ReadInteger
    Player(MyIndex).Dex = Buffer.ReadInteger
    Player(MyIndex).Sci = Buffer.ReadInteger
    Player(MyIndex).Lang = Buffer.ReadInteger
    Player(MyIndex).FreePoints = Buffer.ReadInteger
    
    frmMirage.lblSTR.Caption = Player(MyIndex).Str
    frmMirage.lblDEF.Caption = Player(MyIndex).Def
    frmMirage.lblDEX.Caption = Player(MyIndex).Dex
    frmMirage.lblSCI.Caption = Player(MyIndex).Sci
    frmMirage.lblLANG.Caption = Player(MyIndex).Lang
    frmMirage.lblPoints.Caption = Player(MyIndex).FreePoints
End Sub

Private Sub HandlePetDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim petIndex As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    petIndex = Buffer.ReadInteger
    Set Pets(petIndex) = New clsMapNpc

    If frmMirage.picRightClick.Visible Then
        If frmMirage.picRightClick.DataField = PET_TYPE Then
            If frmMirage.nom.DataField = petIndex Then
                frmMirage.picRightClick.Visible = False
            End If
        End If
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleAreaWeather(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim AreaCount As Byte
    Dim AreaIndex As Byte
    Dim AreaWeather As Byte
    Dim i As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    Set AreasWeather = Nothing
    Set AreasWeather = New Collection

    AreaCount = Buffer.ReadByte
    For i = 0 To AreaCount - 1
        AreaIndex = Buffer.ReadByte
        AreaWeather = Buffer.ReadByte
        Call AreasWeather.Add(AreaWeather, Str(AreaIndex))
    Next i

    If GameWeather <> AreasWeather.item(Str(Map.Area)) Then
    GameWeather = AreasWeather.item(Str(Map.Area))
        Debug.Print "game weather : " & GameWeather

        If IsEmptyArray(ArrPtr(DropRain)) Then ' Probleme ici car si on redim le tableau alors
        'qu'on le process ça peut bugger  en execution
            RainIntensity = 200

            MAX_RAINDROPS = RainIntensity

            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleSpawnMapItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim nbItem As Integer
    Dim i As Integer
    Dim X, Y As Byte
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    nbItem = Buffer.ReadInteger
    
    For i = 0 To nbItem - 1
        X = Buffer.ReadByte
        Y = Buffer.ReadByte
        
        ReDim Preserve MapItem(X, Y).items(0 To NbMapItems(X, Y))
        With MapItem(X, Y).items(NbMapItems(X, Y) - 1)
            .num = Buffer.ReadInteger
            .Value = Buffer.ReadInteger
            .dur = Buffer.ReadInteger
        End With
    Next i
    
    Set Buffer = Nothing
End Sub

Private Sub HandleDeleteMapItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X, Y As Byte
    Dim i As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    X = Buffer.ReadByte
    Y = Buffer.ReadByte
    
    With MapItem(X, Y)
        .items(0).Value = .items(0).Value - Buffer.ReadInteger
        If .items(0).Value <= 0 Then
            If NbMapItems(X, Y) = 1 Then
                Erase .items
            Else
                For i = 0 To NbMapItems(X, Y) - 2
                    .items(i) = .items(i + 1)
                Next i
                
                ReDim Preserve .items(0 To NbMapItems(X, Y) - 2)
            End If
        End If
    End With
    
    Set Buffer = Nothing
End Sub


Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Integer
Dim i As Integer
Dim Buffer As clsBuffer


    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadInteger
    
    Call MapNpc.Remove(n)
    
    Dim nbDamage As Integer
    nbDamage = NbDamageToDisplay()
    i = 0
    Do While i <= (nbDamage - 1)
        Debug.Print i
        If DamageDisplayer(i).TargetType = NPC_TYPE Then
            If DamageDisplayer(i).targetIndex = n Then
                Call RemoveDamageToDisplay(i)
                nbDamage = nbDamage - 1
            End If
        End If
        i = i + 1
    Loop

    If frmMirage.picRightClick.Visible Then
        If frmMirage.picRightClick.DataField = NPC_TYPE Then
            If frmMirage.nom.DataField = n Then
                frmMirage.picRightClick.Visible = False
            End If
        End If
    End If
End Sub

Private Sub HandleMissileAppear(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    Dim launcherIndex As Integer
    Dim id As Byte
    Dim missileType As Byte
    Dim direction, X, Y As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    launcherIndex = Buffer.ReadInteger
    Player(launcherIndex).MovementTimer = GetTickCount + (Player(launcherIndex).AttackSpeed / 2)
    Player(launcherIndex).AttackTimer = GetTickCount + Player(launcherIndex).AttackSpeed
    id = Buffer.ReadByte
    missileType = Buffer.ReadByte
    direction = Buffer.ReadByte
    X = Buffer.ReadByte
    Y = Buffer.ReadByte
    
    Set Buffer = Nothing

    ArrowsEffect.Add id, New clsArrowAnim

    With ArrowsEffect(id)
        .ArrowAnim = missileType
        .ArrowTime = GetTickCount
        .ArrowVarX = 0
        .ArrowVarY = 0
        .ArrowY = Y
        .ArrowX = X

        .ArrowPosition = direction
        If direction = DIR_DOWN Then
            .ArrowY = Y + 1
        End If
        If direction = DIR_UP Then
            .ArrowY = Y - 1
        End If
        If direction = DIR_RIGHT Then
            .ArrowX = X + 1
        End If
        If direction = DIR_LEFT Then
            .ArrowX = X - 1
        End If
    End With
End Sub

Private Sub HandleMissileDisappear(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim arrowIndex As Byte
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    arrowIndex = Buffer.ReadByte
    
    ArrowsEffect.Remove arrowIndex
    
    Set Buffer = Nothing
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
      
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    n = Buffer.ReadLong
    
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    n = Buffer.ReadByte
    TempTile(X, Y).DoorOpen = n
End Sub


Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim playerIndex As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    playerIndex = Buffer.ReadInteger
    
    Call ClearPlayer(playerIndex)
    
    Set Buffer = Nothing
    
    Call RefreshOnlineList
End Sub

Private Sub HandleQuitMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playerIndex As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    playerIndex = Buffer.ReadInteger
    
    Player(playerIndex).Map = -1

    ' Quand il part, il faut qu'il n'ai plus aucune destination, sinon en revenant, quand il fera un start move il peut s'y teleporter
    Call ClearPlayerMove(playerIndex)
    
    With Pets(playerIndex)
        .Map = -1
        Call ClearMapNpcMove(Pets(playerIndex))
    End With
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    IsDead = True
End Sub

Private Sub HandleLife(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

Dim Buffer As clsBuffer
Dim i As Integer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    i = Buffer.ReadInteger

    'Player HP
    Player(i).MaxHp = Buffer.ReadLong
    Call SetPlayerHP(i, Buffer.ReadLong)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleStamina(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
Dim Buffer As clsBuffer
Dim i As Integer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    i = Buffer.ReadInteger

    Player(i).MaxSTP = Buffer.ReadLong
    Call SetPlayerSTP(i, Buffer.ReadLong)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleSleep(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
 
    'Player SP
    Player(MyIndex).MaxSLP = Buffer.ReadLong
    Call SetPlayerSLP(MyIndex, Buffer.ReadLong)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleExperience(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Call SetPlayerExp(MyIndex, Buffer.ReadLong)
    
    Set Buffer = Nothing

End Sub

Private Sub HandleNextLevel(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
 
    nelvl = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePartyBars(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i, J As Integer
    Dim nbPlayer As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    nbPlayer = Buffer.ReadByte
    For J = 1 To nbPlayer
        i = Buffer.ReadInteger
        
        Player(i).partyIndex = Player(MyIndex).partyIndex
        nbPartyPlayer = nbPartyPlayer + 1
    
        Call SetArmorSlot(i, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
        Call SetWeaponSlot(i, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
        Call SetHelmetSlot(i, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
        Call SetShieldSlot(i, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    
        'Player HP
        Player(i).MaxHp = Buffer.ReadLong
        Call SetPlayerHP(i, Buffer.ReadLong)
    
        'Player SP
        Player(i).MaxSTP = Buffer.ReadLong
        Call SetPlayerSTP(i, Buffer.ReadLong)
    Next J
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerSkills(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim nbSkills As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    nbSkills = Buffer.ReadByte

    For i = 0 To nbSkills - 1
        Player(MyIndex).skill(i) = Buffer.ReadInteger
    Next i
    
    Set Buffer = Nothing
    
    Call Affspell
End Sub

Private Sub HandlePlayerSkill(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim slotNum As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    slotNum = Buffer.ReadInteger
    Player(MyIndex).skill(slotNum) = Buffer.ReadInteger
    
    Set Buffer = Nothing
End Sub

Private Sub HandleInventory(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    Call SetArmorSlot(MyIndex, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    Call SetWeaponSlot(MyIndex, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    Call SetHelmetSlot(MyIndex, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    Call SetShieldSlot(MyIndex, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    
    For i = 0 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadInteger)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadInteger)
        Call SetPlayerInvItemDur(MyIndex, i, Buffer.ReadInteger)
    Next i
    
    Call UpdateVisInv
    
    Set Buffer = Nothing
End Sub

Private Sub HandleInventorySlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim numSlot As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    numSlot = Buffer.ReadByte
    
    Call SetPlayerInvItemNum(MyIndex, numSlot, Buffer.ReadInteger)
    Call SetPlayerInvItemValue(MyIndex, numSlot, Buffer.ReadInteger)
    Call SetPlayerInvItemDur(MyIndex, numSlot, Buffer.ReadInteger)
    
    Call UpdateVisInv
    
    Set Buffer = Nothing
End Sub

Private Sub HandleWeaponSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Call SetWeaponSlot(Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleArmorSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Call SetArmorSlot(Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)

    Set Buffer = Nothing
End Sub

Private Sub HandleHelmetSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Call SetHelmetSlot(Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleShieldSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
  
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Call SetShieldSlot(Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger, Buffer.ReadInteger)
    
    Set Buffer = Nothing
    
End Sub

Sub HandleTime(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    GameTime = Buffer.ReadBoolean

    If GameTime = TIME_DAY Then Call AddText("Le jour se lève.", White) Else Call AddText("La nuit tombe.", White)
    
    Set Buffer = Nothing
End Sub

Sub HandleConfirmUseItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim itemNum As Integer
    Dim playerIndex As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    playerIndex = Buffer.ReadInteger()
    itemNum = Buffer.ReadInteger()
    
    Call ApplyItemEffects(playerIndex, itemNum)
    
    Set Buffer = Nothing
End Sub

Sub HandleCancelUseItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim itemNum As Integer
    Dim playerIndex As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    playerIndex = Buffer.ReadInteger()
    itemNum = Buffer.ReadInteger()
    
    Call RemoveItemEffects(playerIndex, itemNum)
    
    Set Buffer = Nothing
End Sub

Public Sub ApplyItemEffects(ByVal Index As Integer, ByVal itemNum As Integer)
    Dim i As Integer
    
    With item(itemNum)
        If .LifeEffect <> 0 Then
            Call SetPlayerHP(Index, GetPlayerHP(Index) + .LifeEffect)
        End If
        If .SleepEffect <> 0 Then
            Call SetPlayerSLP(Index, GetPlayerSLP(Index) + .SleepEffect)
        End If
        If .StaminaEffect <> 0 Then
            Call SetPlayerSTP(Index, GetPlayerSTP(Index) + .StaminaEffect)
        End If
        
        If .AddHP <> 0 Then
            Player(Index).MaxHp = GetPlayerMaxHP(Index) + .AddHP
            If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
                Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            End If
        End If
        
        If .AddSTP <> 0 Then
            Player(Index).MaxSTP = GetPlayerMaxSTP(Index) + .AddSTP
            If GetPlayerSTP(Index) > GetPlayerMaxSTP(Index) Then
                Call SetPlayerSTP(Index, GetPlayerMaxSTP(Index))
            End If
        End If
        
        If .AddSLP <> 0 Then
            Player(Index).MaxSLP = GetPlayerMaxSLP(Index) + .AddSLP
            If GetPlayerSLP(Index) > GetPlayerMaxSLP(Index) Then
                Call SetPlayerSLP(Index, GetPlayerMaxSLP(Index))
            End If
        End If
        
        Player(Index).StrBonus = Player(Index).StrBonus + .AddStr
        Player(Index).DefBonus = Player(Index).DefBonus + .AddDef
        Player(Index).DexBonus = Player(Index).DexBonus + .AddDex
        Player(Index).SciBonus = Player(Index).SciBonus + .AddSci
        Player(Index).LangBonus = Player(Index).LangBonus + .AddLang
        
        Player(Index).AttackSpeed = Player(Index).AttackSpeed + 1000 * (.AttackSpeed / 1000)
        
        If Index = MyIndex Then
            Call frmMirage.RefreshStatsBonus
        
            If .Type = ITEM_TYPE_POTION Then
                Dim itemCooldown As New clsItemCooldown
                itemCooldown.itemNum = itemNum
                itemCooldown.endCooldown = GetTickCount + item(itemNum).Datas(0)
                Call cooldownItem.Add(itemCooldown)
            End If
        End If
    End With
End Sub

Public Sub RemoveItemEffects(ByVal Index As Integer, ByVal itemNum As Integer)
    Dim i As Integer

    With item(itemNum)
        If .AddHP <> 0 Then
            Player(Index).MaxHp = GetPlayerMaxHP(Index) - .AddHP
            If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
                Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            End If
        End If
        
        If .AddSTP <> 0 Then
            Player(Index).MaxSTP = GetPlayerMaxSTP(Index) - .AddSTP
            If GetPlayerSTP(Index) > GetPlayerMaxSTP(Index) Then
                Call SetPlayerSTP(Index, GetPlayerMaxSTP(Index))
            End If
        End If
        
        If .AddSLP <> 0 Then
            Player(Index).MaxSLP = GetPlayerMaxSLP(Index) - .AddSLP
            If GetPlayerSLP(Index) > GetPlayerMaxSLP(Index) Then
                Call SetPlayerSLP(Index, GetPlayerMaxSLP(Index))
            End If
        End If
        
        Player(Index).StrBonus = Player(Index).StrBonus - .AddStr
        Player(Index).DefBonus = Player(Index).DefBonus - .AddDef
        Player(Index).DexBonus = Player(Index).DexBonus - .AddDex
        Player(Index).SciBonus = Player(Index).SciBonus - .AddSci
        Player(Index).LangBonus = Player(Index).LangBonus - .AddLang
    
        Player(Index).AttackSpeed = Player(Index).AttackSpeed - 1000 * (.AttackSpeed / 1000)
    
        If Index = MyIndex Then
            Call frmMirage.RefreshStatsBonus
            
            If .Type = ITEM_TYPE_POTION Then
                For i = 1 To cooldownItem.Count
                    If cooldownItem.item(i).itemNum = itemNum Then
                        Call cooldownItem.Remove(i)
                        Exit For
                    End If
                Next i
                'Call cooldownItem.Remove(itemNum)
                Call UpdateVisInv
            End If
        End If
    End With
End Sub

Sub HandlePlayerStartInfos(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim indexPlayer As Integer
    Dim sprite As Integer
    Dim partyIndex As Integer
    Dim i As Byte
    Dim nbPlayer As Byte
    Dim petTypeId As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    nbPlayer = Buffer.ReadByte

    For i = 1 To nbPlayer
        indexPlayer = Buffer.ReadInteger
    
        Call SetPlayerName(indexPlayer, Buffer.ReadString)

        If indexPlayer = MyIndex And Not InGame Then
            ' Depend du joueur
            Call initRac
        End If

        sprite = Buffer.ReadInteger
        Call SetPlayerSprite(indexPlayer, sprite)
        
        partyIndex = Buffer.ReadInteger
        Player(indexPlayer).partyIndex = partyIndex
        
        petTypeId = Buffer.ReadInteger
        If petTypeId > -1 Then
            With Pets(indexPlayer)
                .num = petTypeId
            End With
        End If
    Next i
    
    frmMirage.lblName.Caption = Player(MyIndex).name

    Set Buffer = Nothing
    Call RefreshOnlineList
End Sub

Sub HandlePlayerDirMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim dir As Integer
    Dim newX, newY, ancientX, ancientY As Byte
    Dim newDirection As clsDirection
    Set newDirection = New clsDirection

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    i = Buffer.ReadInteger
    dir = Buffer.ReadByte
    
    If dir < DIR_DOWN Or dir > DIR_UP Then Exit Sub
    If (Player(i).Moving = 0) Then
        Call SetPlayerDir(i, dir)
    Else
        newX = Buffer.ReadByte
        newY = Buffer.ReadByte
        
        ancientX = GetPlayerX(i)
        ancientY = GetPlayerY(i)
        
        With newDirection
            .dir = dir
            .X = newX
            .Y = newY
        End With
        
        Call addPlayerDirection(i, newDirection)
    
        If GetArraySize(Player(i).newDir) = 1 Then 'Seul la nouvelle direction est présente
            ' Si le joueur ne se déplace pas par la case indiqué comme fin de déplacement, on l'y téléporte pour faire une bonne synchronisation
            Select Case GetPlayerDir(i)
                Case DIR_UP
                    If newY > ancientY Or Not newX = ancientX Then
                        Call SetPlayerX(i, newX)
                        Call SetPlayerY(i, newY)
                        Player(i).XOffset = 0
                        Player(i).YOffset = 0
                    End If
        
                Case DIR_DOWN
                    If newY < ancientY Or Not newX = ancientX Then
                        Call SetPlayerX(i, newX)
                        Call SetPlayerY(i, newY)
                        Player(i).XOffset = 0
                        Player(i).YOffset = 0
                    End If
        
                Case DIR_LEFT
                    If newX > ancientX Or Not newY = ancientY Then
                        Call SetPlayerX(i, newX)
                        Call SetPlayerY(i, newY)
                        Player(i).XOffset = 0
                        Player(i).YOffset = 0
                    End If
        
                Case DIR_RIGHT
                    If newX < ancientX Or Not newY = ancientY Then
                        Call SetPlayerX(i, newX)
                        Call SetPlayerY(i, newY)
                        Player(i).XOffset = 0
                        Player(i).YOffset = 0
                    End If
            End Select
        End If

    End If

    Set Buffer = Nothing
End Sub

Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim dir As Integer
    Dim newX, newY, ancientX, ancientY As Byte
    Dim newDirection As clsDirection
    Set newDirection = New clsDirection

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    i = Buffer.ReadInteger
    dir = Buffer.ReadByte
    
    Debug.Print "dir without move !" & GetTickCount
    
    If dir < DIR_DOWN Or dir > DIR_UP Then Exit Sub

    If Player(i).Moving > 0 Then
        Call SetPlayerX(i, Player(i).Destination.X)
        Call SetPlayerY(i, Player(i).Destination.Y)
        Call ClearPlayerMove(i)
    End If
    
    Call SetPlayerDir(i, dir)

    Set Buffer = Nothing
End Sub

Sub HandleNpcDirMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim dir As Byte
    Dim newX, newY, ancientX, ancientY As Byte
    Dim newDirection As clsDirection
    Set newDirection = New clsDirection

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    i = Buffer.ReadInteger
    dir = Buffer.ReadByte
    newX = Buffer.ReadByte
    newY = Buffer.ReadByte
    Debug.Print "Npc change dir " & dir & " time : " & GetTickCount
    If MapNpc(i).Moving > 0 Then 'On ne sait jamais.
        ancientX = GetNpcX(MapNpc(i))
        ancientY = GetNpcY(MapNpc(i))
    
        Debug.Print "Npc change dir X : " & ancientX & "/" & newX & " Y : " & ancientY & "/" & newY
    
        If dir < DIR_DOWN Or dir > DIR_UP Then Exit Sub
    
        With newDirection
            .dir = dir
            .X = newX
            .Y = newY
        End With
        
        Call addNpcDirection(MapNpc(i), newDirection)
    
        ' Si le joueur ne se déplace pas par la case indiqué comme fin de déplacement, on l'y téléporte pour faire une bonne synchronisation
        Select Case GetNpcDir(MapNpc(i))
            Case DIR_UP
                If newY > ancientY Then
                    Call SetNpcX(MapNpc(i), newX)
                    Call SetNpcY(MapNpc(i), newY)
                    MapNpc(i).XOffset = 0
                    MapNpc(i).YOffset = 0
                End If
    
            Case DIR_DOWN
                If newY < ancientY Then
                    Call SetNpcX(MapNpc(i), newX)
                    Call SetNpcY(MapNpc(i), newY)
                    MapNpc(i).XOffset = 0
                    MapNpc(i).YOffset = 0
                End If
    
            Case DIR_LEFT
                If newX > ancientX Then
                    Call SetNpcX(MapNpc(i), newX)
                    Call SetNpcY(MapNpc(i), newY)
                    MapNpc(i).XOffset = 0
                    MapNpc(i).YOffset = 0
                End If
    
            Case DIR_RIGHT
                If newX < ancientX Then
                    Call SetNpcX(MapNpc(i), newX)
                    Call SetNpcY(MapNpc(i), newY)
                    MapNpc(i).XOffset = 0
                    MapNpc(i).YOffset = 0
                End If
        End Select
    End If

    Set Buffer = Nothing
End Sub

Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim npcIndex As Integer
    Dim dir As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    npcIndex = Buffer.ReadInteger
    dir = Buffer.ReadByte
    
    Call ChangeNpcDir(MapNpc(npcIndex), dir)
    
    Set Buffer = Nothing
End Sub

Sub HandlePetDirMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim dir As Byte
    Dim newX, newY As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    i = Buffer.ReadInteger
    dir = Buffer.ReadByte
    newX = Buffer.ReadByte
    newY = Buffer.ReadByte

    Call MapNpcDirMove(Pets(i), dir, newX, newY)

    Set Buffer = Nothing
End Sub

Sub HandlePetDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim petIndex As Integer
    Dim dir As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    petIndex = Buffer.ReadInteger
    dir = Buffer.ReadByte
    
    Call ChangeNpcDir(Pets(petIndex), dir)
    
    Set Buffer = Nothing
End Sub

Private Sub MapNpcDirMove(ByRef MapNpc As clsMapNpc, ByVal dir As Integer, ByVal newX As Byte, ByVal newY As Byte)
    Dim ancientX, ancientY As Byte
    Dim newDirection As clsDirection
    Set newDirection = New clsDirection
    
    If MapNpc.Moving > 0 Then 'On ne sait jamais.
        ancientX = GetNpcX(MapNpc)
        ancientY = GetNpcY(MapNpc)
    
        Debug.Print "Npc change dir X : " & ancientX & "/" & newX & " Y : " & ancientY & "/" & newY
    
        If dir < DIR_DOWN Or dir > DIR_UP Then Exit Sub
    
    
        With newDirection
            .dir = dir
            .X = newX
            .Y = newY
        End With
        
        Call addNpcDirection(MapNpc, newDirection)
    
        ' Si le joueur ne se déplace pas par la case indiqué comme fin de déplacement, on l'y téléporte pour faire une bonne synchronisation
        Select Case GetNpcDir(MapNpc)
            Case DIR_UP
                If newY > ancientY Then
                    Call SetNpcX(MapNpc, newX)
                    Call SetNpcY(MapNpc, newY)
                    MapNpc.XOffset = 0
                    MapNpc.YOffset = 0
                End If
    
            Case DIR_DOWN
                If newY < ancientY Then
                    Call SetNpcX(MapNpc, newX)
                    Call SetNpcY(MapNpc, newY)
                    MapNpc.XOffset = 0
                    MapNpc.YOffset = 0
                End If
    
            Case DIR_LEFT
                If newX > ancientX Then
                    Call SetNpcX(MapNpc, newX)
                    Call SetNpcY(MapNpc, newY)
                    MapNpc.XOffset = 0
                    MapNpc.YOffset = 0
                End If
    
            Case DIR_RIGHT
                If newX < ancientX Then
                    Call SetNpcX(MapNpc, newX)
                    Call SetNpcY(MapNpc, newY)
                    MapNpc.XOffset = 0
                    MapNpc.YOffset = 0
                End If
        End Select
    End If
End Sub

Sub HandleDamageDisplay(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
    On Error GoTo Err
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Integer

    Dim dommage As Long

    Dim attaquantType As Byte
    Dim attaquant As Integer
    Dim attaquantNom As String
    Dim cibleType As Byte
    Dim cible As Integer
    Dim cibleNom As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    attaquantType = Buffer.ReadByte
    attaquant = Buffer.ReadInteger
    cibleType = Buffer.ReadByte
    cible = Buffer.ReadInteger
    dommage = Buffer.ReadLong

    If cibleType = PLAYER_TYPE Then
        Call SetPlayerHP(cible, GetPlayerHP(cible) - dommage) ' Must always execute this for party members
        
        If GetPlayerMap(cible) <> GetPlayerMap(MyIndex) Then Exit Sub ' Exit if not on map
    End If

    If attaquantType = PLAYER_TYPE Then
        If GetPlayerMap(attaquant) <> GetPlayerMap(MyIndex) Then Exit Sub ' Exit if not on map
    
        Player(attaquant).MovementTimer = GetTickCount + (Player(attaquant).AttackSpeed / 2)
        Player(attaquant).AttackTimer = GetTickCount + Player(attaquant).AttackSpeed
        attaquantNom = Player(attaquant).name
        If attaquant = MyIndex Then
            Color = White
        End If
    ElseIf attaquantType = NPC_TYPE Then
        attaquantNom = Npc(MapNpc(attaquant).num).name
    End If

    If cibleType = PLAYER_TYPE Then
        cibleNom = Player(cible).name
        If cible = MyIndex Then
            Color = red
        End If
    ElseIf cibleType = NPC_TYPE Then
        cibleNom = Npc(MapNpc(cible).num).name
    End If

    ReDim Preserve DamageDisplayer(0 To NbDamageToDisplay())
    With DamageDisplayer(NbDamageToDisplay() - 1)
        .damage = dommage
        .TargetType = cibleType
        .targetIndex = cible
        .time = GetTickCount
    End With

    Msg = Trim$(attaquantNom) & " attaque " & Trim$(cibleNom) & " pour " & dommage & " de dommages."

    Set Buffer = Nothing

    Call DisplayReport(Msg, Color)
    
Err:
    'Peut arriver si on reçoit l'affichage des dégats d'un mob alors qu'il meurt
End Sub

Sub HandleGetItemDisplay(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Integer
    
    Dim itemId, itemVal As Integer
 
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
 
    itemId = Buffer.ReadInteger
    itemVal = Buffer.ReadInteger
 
    Msg = "Vous venez de ramasser " & itemVal & " " & Trim$(item(itemId).name) & "."
    Color = 14
 
    Set Buffer = Nothing
 
    Call DisplayReport(Msg, Color)
End Sub

Sub HandleRequestParty(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    frmMirage.messageInvitation(PARTY_MESSAGE).Text = GetPlayerName(Buffer.ReadInteger) + " vous a invitez à rejoindre son groupe. Voulez vous accepter ?"
    frmMirage.picInvitation(PARTY_MESSAGE).Visible = True
    
    Set Buffer = Nothing
End Sub

Sub HandleJoinParty(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim partyIndex As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    Player(MyIndex).partyIndex = Buffer.ReadInteger

    Set Buffer = Nothing
End Sub

Sub HandleLeaveParty(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playerIndex As Integer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    playerIndex = Buffer.ReadInteger
    Player(playerIndex).partyIndex = -1
    nbPartyPlayer = nbPartyPlayer - 1
    
    Set Buffer = Nothing

End Sub

Sub HandlePlayerCrafts(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim nbCraft As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    nbCraft = Buffer.ReadByte
    
    For i = 1 To nbCraft
        Call Player(MyIndex).Crafts.Add(Buffer.ReadInteger)
    Next i
    
    frmMirage.CraftsLoad
    
    Set Buffer = Nothing
End Sub

Sub HandleNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcIndex As Integer
    Dim spellNumber As Integer
    Dim i As Integer
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    npcIndex = Buffer.ReadInteger
    
    With Npc(npcIndex)
        .name = Buffer.ReadString
        .sprite = Buffer.ReadInteger
        .SpawnSecs = Buffer.ReadLong
        .Behavior = Buffer.ReadByte
        .Range = Buffer.ReadByte
        .Str = Buffer.ReadInteger
        .Def = Buffer.ReadInteger
        .Dex = Buffer.ReadInteger
        .Sci = Buffer.ReadInteger
        .Lang = Buffer.ReadInteger
        
        .MaxHp = Buffer.ReadLong
        .exp = Buffer.ReadLong
        
        .SpawnTime = Buffer.ReadByte
        
        For i = 1 To MAX_NPC_DROPS
            With .ItemNPC(i)
                .itemNum = Buffer.ReadLong
                .ItemValue = Buffer.ReadLong
                .Chance = Buffer.ReadLong
            End With
        Next i
        
        .Inv = Buffer.ReadBoolean
        .Vol = Buffer.ReadBoolean
        
        spellNumber = Buffer.ReadInteger
        
        If spellNumber > 0 Then
            ReDim .skill(0 To spellNumber - 1)
            
            For i = 0 To spellNumber - 1
                .skill(i) = Buffer.ReadInteger
            Next i
        End If
    End With

    Set Buffer = Nothing
End Sub


Sub HandlePlayerStopMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerIndex As Integer
    Dim newX, newY, ancientX, ancientY As Byte
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    playerIndex = Buffer.ReadInteger
    
    If playerIndex = MyIndex Then
        movementController = 1
    End If

    ' Ici on met trytomove a faux ce qui fait qu'en local, si on est bloqué par un autre joueur et qu'il bouge on va pas se redéplacer.
    ' Le problème est que si on met vrai en local, il risque d'y avoir des transpercement de joueur.
    newX = Buffer.ReadByte
    newY = Buffer.ReadByte
    ancientX = GetPlayerX(playerIndex)
    ancientY = GetPlayerY(playerIndex)
    
'    ' Si le joueur ne se déplace pas par la case indiqué comme fin de déplacement, on l'y téléporte pour faire une bonne synchronisation

    If GetArraySize(Player(playerIndex).newDir) > 0 Then
        Player(playerIndex).Destination.X = newX
        Player(playerIndex).Destination.Y = newY
    Else
        Select Case GetPlayerDir(playerIndex)
            Case DIR_UP
                If newY <= ancientY And newX = ancientX Then
                    Player(playerIndex).Destination.X = newX
                    Player(playerIndex).Destination.Y = newY
                End If
    
            Case DIR_DOWN
                If newY >= ancientY And newX = ancientX Then
                    Player(playerIndex).Destination.X = newX
                    Player(playerIndex).Destination.Y = newY
                End If
    
            Case DIR_LEFT
                If newX <= ancientX And newY = ancientY Then
                    Player(playerIndex).Destination.X = newX
                    Player(playerIndex).Destination.Y = newY
                End If
    
            Case DIR_RIGHT
                Debug.Print "OKKKKKKKKk : " & newX & " /" & ancientX
                If newX >= ancientX And newY = ancientY Then
                    Player(playerIndex).Destination.X = newX
                    Player(playerIndex).Destination.Y = newY
                End If
        End Select
    End If

    If Player(playerIndex).Destination.X = -1 Then
        Call ClearPlayerMove(playerIndex)
        Call SetPlayerX(playerIndex, newX)
        Call SetPlayerY(playerIndex, newY)
    Else
        If playerIndex = MyIndex Then
            Player(playerIndex).Moving = MOVING_RUNNING
        End If
    End If

    Set Buffer = Nothing
End Sub
