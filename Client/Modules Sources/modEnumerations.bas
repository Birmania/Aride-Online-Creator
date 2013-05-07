Attribute VB_Name = "modEnumerations"
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

' Packets sent by server to client
Public Enum ServerPackets
    SFindServer = 0
    SAlertMsg
    SErrorLogin
    SEndWarp
    SCheckForMap
    SYourIndex
    SLife
    SStamina
    SSleep
    SExperience
    SNextLevel
    SPartyBars
    SPlayerSkills
    SPlayerCrafts
    SInventory
    SInventorySlot
    SWeaponSlot
    SArmorSlot
    SHelmetSlot
    SShieldSlot
    SMapData
    SPlayerStartInfos
    SPlayerPosition
    SRequestParty
    SJoinParty
    SLeaveParty
    SChatMsg
    SPlayerStartMove
    SPlayerStopMove
    SPlayerDirMove
    SPlayerDir
    SNpcStartMove
    SNpcStopMove
    SNpcDirMove
    SNpcDir
    SPetStartMove
    SPetStopMove
    SPetDirMove
    SPetDir
    SLeft
    SQuitMap
    SPlayerDead
    SDamageDisplay
    SGetItemDisplay
    SMapNpcData
    SNpcDead
    SMissileAppear
    SMissileDisappear
    SSpawnMapItem
    SDeleteMapItem
    SPlayerMsg
    SStatistics
    SPetDead
    SAreaWeather
    STime
    SConfirmUseItem
    SCancelUseItem
    
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CFindServer = 0
    CLogin
    CNeedMap
    CRequestParty
    CJoinParty
    CLeaveParty
    CPlayerMove
    CPlayerStopMove
    CPlayerDirMove
    CPlayerDir
    CAttack
    CFire
    CMapGetItem
    CSayMsg
    CGoBorderMap
    CPlayerSleep
    CUseItem
    CDropItem
    CTakeOutWeapon
    CTakeOutArmor
    CTakeOutHelmet
    CTakeOutShield
    CMoveInventoryItem
    CAddStr
    CAddDef
    CAddDex
    CAddSci
    CAddLang
    CUseSkill
    CAdopt
    CAbandon
    CExecuteCraft

    CMSG_COUNT
End Enum

Public HandleDataSub(SMSG_COUNT) As Long
