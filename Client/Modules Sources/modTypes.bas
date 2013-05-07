Attribute VB_Name = "modTypes"
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


' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
'CallWindowProc called in modThreading
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long

' Allow transparent background color on option button
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Cursor
Public Declare Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function GetClassLongPtr Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
    
Public Const GCW_HCURSOR = (-12)
' Cursor end

' Configuration files
Public ClientConfigurationFile As String
Public OptionConfigurationFile As String
Public ThemeConfigurationFile As String
Public ColorConfigurationFile As String
Public ErrorLogFile As String

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0

' Map constants
'Public Const MaxMapX As Long = 30
'Public Const MaxMapY As Long = 30
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_NO_PENALTY As Byte = 2

' General constants
Public Const Game_Name As String = "Aride Online"
Public Const WEBSITE As String = "www.aride-online.fr"
Public Const MAX_PLAYERS As Long = 250
Public Const MAX_SKILLS As Long = 100
Public Const MAX_MAPS As Long = 150
Public Const MAX_SHOPS As Long = 50
Public Const MAX_ITEMS As Long = 100
Public Const MAX_NPCS As Long = 100
Public Const MAX_MAP_ITEMS As Long = 20
Public Const MAX_EMOTICONS As Long = 10
'Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public Const MAX_LEVEL As Long = 100
Public Const MAX_QUETES As Long = 100
Public MAX_DX_PETS As Long
Public Const MAX_CRAFTS As Long = 1000
Public Const MAX_MATERIALS As Long = 9

Public Const MAX_INV As Integer = 26
Public Const MAX_PARTY_MEMBERS As Byte = 20
Public Const MAX_PLAYER_SKILLS As Byte = 19
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10

Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Message constants
Public Enum Messages
    PARTY_MESSAGE = 0
    CHAT_MESSAGE
    TRADE_MESSAGE
    SLEEP_MESSAGE
    IMSG_COUNT
End Enum


' Account constants
Public Const NAME_LENGTH As Byte = 40
Public Const MAX_CHARS As Byte = 1
Public Const ACCOUNT_LENGTH As Byte = 12

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 As String = "kwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 As String = "lsisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 As String = "taiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 As String = "98978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Image constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32
Public Const PIC_PL As Byte = 64
Public Const PIC_NPC1 As Byte = 2
Public Const PIC_NPC2 As Byte = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_HEAL As Byte = 7
Public Const TILE_TYPE_KILL As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_CBLOCK As Byte = 10
Public Const TILE_TYPE_ARENA As Byte = 11
Public Const TILE_TYPE_SOUND As Byte = 12
Public Const TILE_TYPE_SPRITE_CHANGE As Byte = 13
Public Const TILE_TYPE_SIGN As Byte = 14
Public Const TILE_TYPE_DOOR As Byte = 15
Public Const TILE_TYPE_NOTICE As Byte = 16
Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_NPC_SPAWN As Byte = 20
Public Const TILE_TYPE_BANK As Byte = 21
Public Const TILE_TYPE_POISON As Byte = 26
Public Const TILE_TYPE_COFFRE As Byte = 22
Public Const TILE_TYPE_PORTE_CODE As Byte = 23
Public Const TILE_TYPE_BLOCK_MONTURE As Byte = 24
Public Const TILE_TYPE_BLOCK_NIVEAUX As Byte = 25
Public Const TILE_TYPE_TOIT As Byte = 26
Public Const TILE_TYPE_BLOCK_GUILDE As Byte = 27
Public Const TILE_TYPE_BLOCK_TOIT As Byte = 28
Public Const TILE_TYPE_BLOCK_DIR As Byte = 29

' quetes constant
Public Const QUETE_TYPE_AUCUN As Byte = 0
Public Const QUETE_TYPE_RECUP As Byte = 1
Public Const QUETE_TYPE_APORT As Byte = 2
Public Const QUETE_TYPE_PARLER As Byte = 3
Public Const QUETE_TYPE_TUER As Byte = 4
Public Const QUETE_TYPE_FINIR As Byte = 5
Public Const QUETE_TYPE_GAGNE_XP As Byte = 6
Public Const QUETE_TYPE_SCRIPT As Byte = 7
Public Const QUETE_TYPE_MINIQUETE As Byte = 8

' Messages constants
Public Const MSG_TYPE_INFO As Byte = 0
Public Const MSG_TYPE_QUEST As Byte = 1

' Item constants
Public Enum ItemTypes
    ITEM_TYPE_NONE = 0
    ITEM_TYPE_WEAPON
    ITEM_TYPE_THROWABLE
    ITEM_TYPE_MISSILE
    ITEM_TYPE_ARMOR
    ITEM_TYPE_HELMET
    ITEM_TYPE_SHIELD
    ITEM_TYPE_POTION
    ITEM_TYPE_KEY
    ITEM_TYPE_CURRENCY
    ITEM_TYPE_SPELL
    ITEM_TYPE_MONTURE
    ITEM_TYPE_SCRIPT
    ITEM_TYPE_PET
End Enum

' Bank constants
Public Const MOVE_TO_INV As Byte = 0
Public Const MOVE_TO_SAFE As Byte = 1

' Direction constants
Public Const DIR_UP As Byte = 3
Public Const DIR_DOWN As Byte = 0
Public Const DIR_LEFT As Byte = 1
Public Const DIR_RIGHT As Byte = 2

' Types d'individu
Public Const PLAYER_TYPE = 0
Public Const NPC_TYPE = 1
Public Const PET_TYPE = 2

' Constants for player movement
Public Const MOVING_WALKING As Byte = 4
Public Const MOVING_RUNNING As Byte = 8

' Weather constants
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_THUNDER As Byte = 3
Public Const WEATHER_SAND_STORMING As Byte = 4

' Time constants
Public Const TIME_DAY As Boolean = True
Public Const TIME_NIGHT As Boolean = False

' Admin constants
Public Const ADMIN_MONITER As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_GUARD As Byte = 4
Public Const NPC_BEHAVIOR_QUETEUR As Byte = 5

' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME As Integer = 4000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 16 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 20 ' In characters.
Public Const MAX_LINES As Byte = 5

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5
Public Const SPELL_TYPE_SCRIPT As Byte = 6
Public Const SPELL_TYPE_AMELIO As Byte = 7
Public Const SPELL_TYPE_DECONC As Byte = 8
Public Const SPELL_TYPE_PARALY As Byte = 9
Public Const SPELL_TYPE_DEFENC As Byte = 10

' options
Public Options As OptionsRec

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
End Type

Public Type DragAndDropRec
    Index As Integer
    Type As Byte
End Type

Public Loading As Boolean
Public deco As Boolean

Type ChatBubble
    Text As String
    Created As Long
End Type

Type ItemSlotRec
    num As Long
    Value As Long
    dur As Long
End Type

Type CoffreTempRec
    Numeros As Long
    valeur As Long
    Durabiliter As Long
End Type

Type ArrowRec
    'Arrow As Byte
    'arrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
End Type

Type IndRec
    Datas(0 To 2) As Long
    Strings(0) As String
End Type

Type PlayerQueteRec
    Temps As Long
    Datas(0 To 2) As Long
    Strings(0) As String
    indexe(1 To 15) As IndRec
End Type

Type PlayerRec

    ' General
    name As String
    Guild As String
    Guildaccess As Byte
    Class As Long
    sprite As Long
    level As Long
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    STP As Long
    SLP As Long 'Tired Point
    
    ' Stats
    Str As Integer
    Def As Integer
    Dex As Integer
    Sci As Integer
    Lang As Integer
    FreePoints As Integer
    'POINTS As Long
    
    ' Worn equipment
    ArmorSlot As ItemSlotRec
    WeaponSlot As ItemSlotRec
    HelmetSlot As ItemSlotRec
    ShieldSlot As ItemSlotRec
    PetSlot As ItemSlotRec
    
    ' Inventory
    Inv(0 To MAX_INV) As ItemSlotRec
    skill(0 To MAX_PLAYER_SKILLS) As Long
    Crafts As New Collection
    
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    dir As Byte
    newDir() As clsDirection

    ' Client use only
    StrBonus As Integer
    DefBonus As Integer
    DexBonus As Integer
    SciBonus As Integer
    LangBonus As Integer
    
    MaxHp As Long
    MaxSTP As Long
    MaxSLP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    MovingSave As Byte
'    Destination As PositionRec
    Destination As New clsPosition
    Attacking As Byte
    MovementTimer As Long
    AttackTimer As Long
    AttackSpeed As Long
    MapGetTimer As Long
    castedSpell As Integer
    partyIndex As Integer

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long
    
    LevelUp As Long
    LevelUpT As Long

    QueteEnCour As Long
    Quetep As PlayerQueteRec
    
    Anim As Byte
    'PAPERDOLL
    Casque As Long
    Armure As Long
    Arme As Long
    Bouclier As Long
    'FIN PAPERDOLL
End Type
    
Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Mask3 As Long
    M3Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Fringe3 As Long
    F3Anim As Long
    Type As Byte
    Datas() As Long
    Strings() As String
    Light As Long
End Type

Type NpcMapRec
    id As Integer
    X() As Byte
    Y() As Byte
    dir As Byte
    Hasardp As Boolean
    movementType As Byte
End Type

Type BorderRec
    XSource As Byte
    YSource As Byte
    DirectionSource As Byte
    
    MapDestination As Integer
    XDestination As Byte
    YDestination As Byte
End Type

Type MapRec
    name As String
    Moral As Byte
    Music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Indoors As Boolean
    tile() As TileRec
    Npcs() As NpcMapRec
    PanoInf As String
    TranInf As Byte
    PanoSup As String
    TranSup As Byte
    Fog As Integer
    FogAlpha As Byte
    Borders() As BorderRec
    Area As Byte
End Type

Type RecompRec
    exp As Long
    objn1 As Long
    objn2 As Long
    objn3 As Long
    objq1 As Long
    objq2 As Long
    objq3 As Long
End Type

Type QueteRec
    nom As String * 40
    Type As Long
    description As String
    reponse As String
    Temps As Long
    Datas(0 To 2) As Long
    Strings(0) As String
    Recompence As RecompRec
    indexe(1 To 15) As IndRec
    Case As Long
End Type

Type ClassRec
    name As String
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    Str As Long
    Def As Long
    Speed As Long
    Magi As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    name As String
    desc As String

    Pic As Long
    Type As Byte
    Datas() As Long
    StrReq As Integer
    DefReq As Integer
    DexReq As Integer
    SciReq As Integer
    LangReq As Integer
'    ClassReq As Long
'    AccessReq As Byte

    Empilable As Byte

    LifeEffect As Integer
    SleepEffect As Integer
    StaminaEffect As Integer
    AddHP As Integer
    AddSLP As Integer
    AddSTP As Integer
    AddStr As Integer
    AddDef As Integer
    AddSci As Integer
    AddDex As Integer
    AddLang As Integer
    AddEXP As Long
    AttackSpeed As Long

    NCoul As Long

    Sex As Byte
End Type

Type DisplayDamageRec
    damage As Integer
    TargetType As Byte
    targetIndex As Integer
    time As Long
    offset As Integer
End Type

Type MapItemRec
    items() As ItemSlotRec
End Type

Type NPCEditorRec
    itemNum As Long
    ItemValue As Long
    Chance As Long
End Type

Type NpcRec
    name As String
    'AttackSay As String
    
    sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    Str As Integer
    Def As Integer
    Dex As Integer
    Sci As Integer
    Lang As Integer
    MaxHp As Long
    exp As Long
    SpawnTime As Byte
    AttackSpeed As Integer
    
    ItemNPC() As NPCEditorRec
    Inv As Boolean
    Vol As Boolean
    
    skill() As Integer
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type TradeItemsRec
    Value(1 To MAX_TRADES) As TradeItemRec
End Type

Type ShopRec
    name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
    FixObjet As Long
End Type

Type SkillRec
    name As String
    LevelReq As Integer
    Sound As Integer
    MPCost As Integer
    Type As Byte
    Range As Byte
    
    Big As Byte
    
    SkillAnimId As Integer
    SkillAnimDuration As Integer
    SkillAnimOccurence As Integer
    
    SkillIco As Integer
    
    TargetType As Byte
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    Pic As Long
    command As String
End Type

Type DropRainRec
    X As Long
    Y As Long
    Randomized As Boolean
    Speed As Byte
End Type

Type MaterialRec
    itemNum As Integer
    Count As Integer
End Type

Type CraftRec
    name As String
    Materials() As MaterialRec
    Products() As MaterialRec
End Type

' Bubble thing
Public Bubble(1 To MAX_PLAYERS) As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public quete(0 To MAX_QUETES) As QueteRec
Public Map As MapRec
Public TempTile() As TempTileRec
Public Player(0 To MAX_PLAYERS) As PlayerRec
Public PlayerAnim(1 To MAX_PLAYERS, 0 To 4) As Long
Public Class() As ClassRec
Public item(0 To MAX_ITEMS) As ItemRec
Public cooldownItem As Collection
Public Npc(0 To MAX_NPCS) As NpcRec
Public DamageDisplayer() As DisplayDamageRec
Public MapItem() As MapItemRec
Public MapNpc As New Dictionary
Public Shop(0 To MAX_SHOPS) As ShopRec
Public skill(0 To MAX_SKILLS) As SkillRec
Public Emoticons(0 To MAX_EMOTICONS) As EmoRec
Public MapReport(1 To MAX_MAPS) As MapRec
Public CoffreTmp(1 To 30) As CoffreTempRec
Public Crafts(0 To MAX_CRAFTS) As CraftRec
' Passer Pets en collection
Public Pets(0 To MAX_PLAYERS) As New clsMapNpc

'User commands
Public UserCommand As New Dictionary
Public UserCommandLabel As New Collection

Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type
Type TradeRec
    items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type
Public Trade(1 To 6) As TradeRec

Public ArrowsEffect As Dictionary
Public SkillsEffect As Collection

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Long
    time As Long
    Done As Byte
    Y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    item As Long
    dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec


Public Minu As Long
Public Seco As Long

'Type pour stocker le contenu de Account.ini
Type TpAccOpt
    InfName As String
    InfPass As String
    SpeechBubbles As Boolean
    NpcBar As Boolean
    NpcName As Boolean
    NpcDamage As Boolean
    PlayBar As Boolean
    PlayName As Boolean
    PlayDamage As Boolean
    MapGrid As Boolean
    Music As Boolean
    Sound As Boolean
    Autoscroll As Boolean
    NomObjet As Boolean
    LowEffect As Boolean
End Type

Public rac(0 To 13, 0 To 1) As Integer
Public dragAndDrop As DragAndDropRec

Public AccOpt As TpAccOpt

' Configuration Menu Option des touches
Type optToucheRec
    nom As String
    'Value As Byte
End Type
Public nelvl As Long
Public Const TCHMAX = 51
Public optTouche As New Dictionary

Sub iniOptTouche()
    optTouche(vbKeyA) = "A"
    optTouche(vbKeyB) = "B"
    optTouche(vbKeyC) = "C"
    optTouche(vbKeyD) = "D"
    optTouche(vbKeyE) = "E"
    optTouche(vbKeyF) = "F"
    optTouche(vbKeyG) = "G"
    optTouche(vbKeyH) = "H"
    optTouche(vbKeyI) = "I"
    optTouche(vbKeyJ) = "J"
    optTouche(vbKeyK) = "K"
    optTouche(vbKeyL) = "L"
    optTouche(vbKeyM) = "M"
    optTouche(vbKeyN) = "N"
    optTouche(vbKeyO) = "O"
    optTouche(vbKeyP) = "P"
    optTouche(vbKeyQ) = "Q"
    optTouche(vbKeyR) = "R"
    optTouche(vbKeyS) = "S"
    optTouche(vbKeyT) = "T"
    optTouche(vbKeyU) = "U"
    optTouche(vbKeyV) = "V"
    optTouche(vbKeyW) = "W"
    optTouche(vbKeyX) = "X"
    optTouche(vbKeyY) = "Y"
    optTouche(vbKeyZ) = "Z"
    optTouche(vbKey0) = "0"
    optTouche(vbKey1) = "1"
    optTouche(vbKey2) = "2"
    optTouche(vbKey3) = "3"
    optTouche(vbKey4) = "4"
    optTouche(vbKey5) = "5"
    optTouche(vbKey6) = "6"
    optTouche(vbKey7) = "7"
    optTouche(vbKey8) = "8"
    optTouche(vbKey9) = "9"
    optTouche(vbKeyF1) = "F1"
    optTouche(vbKeyF2) = "F2"
    optTouche(vbKeyF3) = "F3"
    optTouche(vbKeyF4) = "F4"
    optTouche(vbKeyF5) = "F5"
    optTouche(vbKeyF6) = "F6"
    optTouche(vbKeyF7) = "F7"
    optTouche(vbKeyF8) = "F8"
    optTouche(vbKeyUp) = "Haut"
    optTouche(vbKeyDown) = "Bas"
    optTouche(vbKeyLeft) = "Gauche"
    optTouche(vbKeyRight) = "Droite"
    optTouche(vbKeyControl) = "Ctrl"
    optTouche(vbKeyMenu) = "Alt"
    optTouche(vbKeyShift) = "Shift"
    optTouche(vbKeySpace) = "Espace"
    optTouche(vbKeyReturn) = "Entrée"
End Sub

Sub ClearTempTile()
    Erase TempTile
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

With Player(Index)
    .name = vbNullString
    .Guild = vbNullString
    .Guildaccess = 0
    .Class = 0
    .level = 0
    .sprite = 0
    .exp = 0
    .Access = 0
    .PK = NO
        
    .HP = 0
    .STP = 0
    .SLP = 0
        
    .Dex = 0
    .Str = 0
    .Sci = 0
    .Lang = 0
    .Def = 0
    .FreePoints = 0
    
    .QueteEnCour = 0
    .Quetep.Datas(0) = 0
    .Quetep.Datas(1) = 0
    .Quetep.Datas(2) = 0
    .Quetep.Strings(0) = vbNullString
      
    For n = 1 To 15
    .Quetep.indexe(n).Datas(0) = 0
    .Quetep.indexe(n).Datas(1) = 0
    .Quetep.indexe(n).Datas(2) = 0
    .Quetep.indexe(n).Strings(0) = vbNullString
    Next n
        
    For n = 1 To MAX_INV
        .Inv(n).num = -1
        .Inv(n).Value = -1
        .Inv(n).dur = -1
    Next n
        
    .ArmorSlot.num = -1
    .ArmorSlot.Value = -1
    .ArmorSlot.dur = -1
    .WeaponSlot.num = -1
    .WeaponSlot.Value = -1
    .WeaponSlot.dur = -1
    .HelmetSlot.num = -1
    .HelmetSlot.Value = -1
    .HelmetSlot.dur = -1
    .ShieldSlot.num = -1
    .ShieldSlot.Value = -1
    .ShieldSlot.dur = -1
    .PetSlot.num = -1
    .PetSlot.Value = -1
    .PetSlot.dur = -1
    
    .Map = -1
    .X = 0
    .Y = 0
    .dir = 0
    
    ' Client use only
    .StrBonus = 0
    .DefBonus = 0
    .DexBonus = 0
    .SciBonus = 0
    .LangBonus = 0
    
    .MaxHp = 0
    .MaxSTP = 0
    .MaxSLP = 0
    Call ClearPlayerMove(Index)
    .Attacking = 0
    .MovementTimer = 0
    .AttackTimer = 0
    .AttackSpeed = 1000
    .MapGetTimer = 0
    .castedSpell = -1
    .EmoticonNum = -1
    .EmoticonTime = 0
    .EmoticonVar = 0
    
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).time = i
    Next i
    
    .QueteEnCour = 0
    
    For i = 0 To MAX_PLAYER_SKILLS
        .skill(i) = -1
    Next i
    
    Set .Crafts = Nothing
    Set .Crafts = New Collection

    .partyIndex = -1
End With

Call ClearMapNpc(Pets(Index))
End Sub

Sub ClearPlayerMove(ByVal Index As Integer)
    With Player(Index)
        .Moving = 0
        .XOffset = 0
        .YOffset = 0
        Erase .newDir
        .Destination.X = -1
        .Destination.Y = -1
    End With
End Sub

Sub ClearMapNpcMove(ByRef MapNpc As clsMapNpc)
    With MapNpc
        .Moving = 0
        .XOffset = 0
        .YOffset = 0
        Set .newNpcDir = Nothing
        Set .newNpcDir = New Collection
        .Destination.X = -1
        .Destination.Y = -1
    End With
End Sub

Sub ClearPlayerQuete(ByVal Index As Long)
Dim i As Long
With Player(MyIndex)
        .QueteEnCour = 0
        .Quetep.Datas(0) = 0
        .Quetep.Datas(1) = 0
        .Quetep.Datas(2) = 0
        .Quetep.Strings(0) = vbNullString
'        Accepter = False
        
        For i = 1 To 15
        .Quetep.indexe(i).Datas(0) = 0
        .Quetep.indexe(i).Datas(1) = 0
        .Quetep.indexe(i).Datas(2) = 0
        .Quetep.indexe(i).Strings(0) = 0
        Next i
End With
End Sub

Sub ClearItem(ByVal Index As Long)
With item(Index)
    .name = vbNullString
    .desc = vbNullString
    
    .Type = 0
    Erase .Datas
    .StrReq = 0
    .DefReq = 0
    .DexReq = 0
    .SciReq = 0
    .LangReq = 0
    
    .Empilable = 0
    
    .AddHP = 0
    .AddSLP = 0
    .AddSTP = 0
    .AddStr = 0
    .AddDef = 0
    .AddSci = 0
    .AddDex = 0
    .AddLang = 0
    .AddEXP = 0
    .AttackSpeed = 1000
    
    .NCoul = 0
End With
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal X As Long, ByVal Y As Long)
With MapItem(X, Y)
    Erase .items
End With
End Sub

Sub ClearMap()
    With Map
        .name = vbNullString
        .Moral = 0
        .Indoors = 0
            
         Erase Map.tile
            
        .PanoInf = vbNullString
        .TranInf = 0
        .PanoSup = vbNullString
        .TranSup = 0
        .Fog = 0
        .FogAlpha = 0
    End With
    
    Set TilesPic = Nothing
    Set TilesPic = New Collection
End Sub

Sub ClearMapItems()
Dim X, Y As Long

    For X = 0 To MaxMapX
        For Y = 0 To MaxMapY
            Call ClearMapItem(X, Y)
        Next Y
    Next X
End Sub

Sub ClearMapNpc(ByRef MapNpc As clsMapNpc)
With MapNpc
    .num = -1
    .Target = 0
    .HP = 0
    .SP = 0
    .Map = -1
    .X = 0
    .Y = 0
    .dir = 0

    ' Client use only
    Call ClearMapNpcMove(MapNpc)
    .Attacking = 0
    .AttackTimer = 0
End With
End Sub

Sub ClearMapNpcs()
Dim i As Long
    Call MapNpc.RemoveAll
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    Player(Index).name = name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)
    Player(Index).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal level As Long)
    Player(Index).level = level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal exp As Long)
    Player(Index).exp = exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
End Sub

Function GetPlayerSTP(ByVal Index As Long) As Long
    GetPlayerSTP = Player(Index).STP
End Function

Sub SetPlayerSTP(ByVal Index As Long, ByVal STP As Long)
    Player(Index).STP = STP

    If GetPlayerSTP(Index) > GetPlayerMaxSTP(Index) Then Player(Index).STP = GetPlayerMaxSTP(Index)
End Sub

Function GetPlayerSLP(ByVal Index As Long) As Long
    GetPlayerSLP = Player(Index).SLP
End Function

Sub SetPlayerSLP(ByVal Index As Long, ByVal SLP As Long)
    Player(Index).SLP = SLP

    If GetPlayerSLP(Index) > GetPlayerMaxSLP(Index) Then Player(Index).SLP = GetPlayerMaxSLP(Index)
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHp
End Function

Function GetPlayerMaxSTP(ByVal Index As Long) As Long
    GetPlayerMaxSTP = Player(Index).MaxSTP
End Function

Function GetPlayerMaxSLP(ByVal Index As Long) As Long
    GetPlayerMaxSLP = Player(Index).MaxSLP
End Function

Function GetPlayerstr(ByVal Index As Long) As Long
    GetPlayerstr = Player(Index).Str
End Function

Sub SetPlayerstr(ByVal Index As Long, ByVal Str As Long)
    Player(Index).Str = Str
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Def
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal Def As Long)
    Player(Index).Def = Def
End Sub

Function GetPlayerDEX(ByVal Index As Long) As Long
    GetPlayerDEX = Player(Index).Dex
End Function

Sub SetPlayerLANG(ByVal Index As Long, ByVal Lang As Long)
    Player(Index).Lang = Lang
End Sub

Function GetPlayerLANG(ByVal Index As Long) As Long
    GetPlayerLANG = Player(Index).Lang
End Function

Sub SetPlayerDEX(ByVal Index As Long, ByVal Dex As Long)
    Player(Index).Dex = Dex
End Sub

Function GetPlayerSCI(ByVal Index As Long) As Long
    GetPlayerSCI = Player(Index).Sci
End Function

Sub SetPlayerSCI(ByVal Index As Long, ByVal Sci As Long)
    Player(Index).Sci = Sci
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapNum As Long)
    Player(Index).Map = mapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Dim Packet As clsBuffer
    If X < 0 Then
        X = 0
    End If
    If X > MaxMapX Then
        X = MaxMapX
    End If
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    If Y < 0 Then
        Y = 0
    End If
    If Y > MaxMapY Then
        Y = MaxMapY
    End If
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Byte
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
    Player(Index).dir = dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal itemNum As Long)
    Player(Index).Inv(InvSlot).num = itemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Function GetPlayerInvItemTotalValue(ByVal Index As Long, ByVal itemNum As Integer) As Long
    Dim i As Integer
    
    For i = 0 To MAX_INV
        If GetPlayerInvItemNum(Index, i) = itemNum Then
            GetPlayerInvItemTotalValue = GetPlayerInvItemTotalValue + GetPlayerInvItemValue(Index, i)
        End If
    Next i
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).dur = ItemDur
End Sub

Sub ClearCraft(ByVal Index As Long)
    Dim i As Integer

    Crafts(Index).name = vbNullString
    
    Erase Crafts(Index).Materials
    Erase Crafts(Index).Products
End Sub

Sub ClearCrafts()
Dim i As Long

    For i = 0 To MAX_CRAFTS
        Call ClearCraft(i)
    Next i
End Sub

Public Function GetMapNbNpcs()
    If IsEmptyArray(ArrPtr(Map.Npcs)) Then
        GetMapNbNpcs = 0
    Else
        GetMapNbNpcs = UBound(Map.Npcs) + 1
    End If
End Function

Public Function RemovePlayerDirection(ByVal Index As Integer)
    Dim i As Byte
    
    If GetArraySize(Player(Index).newDir) > 1 Then
        i = 0
        Do While i < GetArraySize(Player(Index).newDir) - 1
            Player(Index).newDir(i) = Player(Index).newDir(i + 1)
            i = i + 1
        Loop
        ReDim Preserve Player(Index).newDir(0 To GetArraySize(Player(Index).newDir) - 2)
    Else
        Erase Player(Index).newDir
    End If
End Function

Public Sub addPlayerDirection(ByVal Index As Integer, ByRef direction As clsDirection)
    ReDim Preserve Player(Index).newDir(0 To GetArraySize(Player(Index).newDir))
    Set Player(Index).newDir(UBound(Player(Index).newDir)) = direction
End Sub

Public Sub addNpcDirection(ByRef MapNpc As clsMapNpc, ByRef direction As clsDirection)
    Call MapNpc.newNpcDir.Add(direction)
End Sub

Public Function RemoveNpcDirection(ByRef MapNpc As clsMapNpc)
    Call MapNpc.newNpcDir.Remove(1)
End Function

Public Sub beginPlayerMovement(ByVal Index As Integer, ByVal movement As Integer, ByVal direction As Integer)
    If direction < DIR_DOWN Or direction > DIR_UP Then Exit Sub

    Call SetPlayerDir(Index, direction)

    Player(Index).Moving = movement
    
    Call initMoveOffset(Index)
    
    If IsPlaying(Index) And Index < MAX_PLAYERS And Index > 0 Then If Player(Index).Anim = 0 Then Player(Index).Anim = 2 Else Player(Index).Anim = 0
End Sub

Public Sub initMoveOffset(ByVal Index As Long)
    Player(Index).Destination.X = -1
    Player(Index).Destination.Y = -1

    Call initOffset(Index)
End Sub

Public Sub initOffset(ByVal Index As Long)
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).YOffset = PIC_Y
            Call SetPlayerY(Index, GetPlayerY(Index) - 1)
        Case DIR_DOWN
            Player(Index).YOffset = (PIC_Y * -1)
            Call SetPlayerY(Index, GetPlayerY(Index) + 1)
        Case DIR_LEFT
            Player(Index).XOffset = PIC_X
            Call SetPlayerX(Index, GetPlayerX(Index) - 1)
        Case DIR_RIGHT
            Player(Index).XOffset = (PIC_X * -1)
            Call SetPlayerX(Index, GetPlayerX(Index) + 1)
    End Select
End Sub

Sub SetNpcX(ByRef MapNpc As clsMapNpc, ByVal X As Long)
    If X < 0 Then
        X = 0
    End If
    If X > MaxMapX Then
        X = MaxMapX
    End If
    MapNpc.X = X
End Sub

Function GetNpcX(ByRef MapNpc As clsMapNpc)
    GetNpcX = MapNpc.X
End Function

Sub SetNpcY(ByRef MapNpc As clsMapNpc, ByVal Y As Long)
    If Y < 0 Then
        Y = 0
    End If
    If Y > MaxMapY Then
        Y = MaxMapY
    End If
    MapNpc.Y = Y
End Sub

Function GetNpcY(ByRef MapNpc As clsMapNpc)
    GetNpcY = MapNpc.Y
End Function

Function SetNpcDir(ByRef MapNpc As clsMapNpc, ByVal direction As Long)
    MapNpc.dir = direction
End Function

Function GetNpcDir(ByRef MapNpc As clsMapNpc)
    GetNpcDir = MapNpc.dir
End Function

Public Function NbMapItems(ByVal X As Integer, ByVal Y As Integer)
    If IsEmptyArray(ArrPtr(MapItem(X, Y).items)) Then
        NbMapItems = 0
    Else
        NbMapItems = UBound(MapItem(X, Y).items) + 1
    End If
End Function

Public Function NbDamageToDisplay()
    If IsEmptyArray(ArrPtr(DamageDisplayer)) Then
        NbDamageToDisplay = 0
    Else
        NbDamageToDisplay = UBound(DamageDisplayer) + 1
    End If
End Function

Public Sub RemoveDamageToDisplay(ByVal damageNum As Integer)
    Dim J As Integer

    If NbDamageToDisplay() = 1 Then
        Erase DamageDisplayer
    Else
        For J = damageNum To (NbDamageToDisplay() - 2)
            DamageDisplayer(J) = DamageDisplayer(J + 1)
        Next J
        ReDim Preserve DamageDisplayer(0 To NbDamageToDisplay() - 2)
    End If
End Sub

Public Sub SetCursor(ByVal CursorFile As String)
    Dim i As Integer

    Call RestoreCursor ' retourne au curseur classique avant de partir sur un autre curseur

    'Load a cursor from a file
    frmMirage.Curs2Handle = LoadCursorFromFile(CursorFile)
    
    'Set the button's cursor
    Set frmMirage.SysCursHandle = New Collection
    For i = 1 To frmMirage.Ctrl2Handle.Count
        Call frmMirage.SysCursHandle.Add(SetClassLongPtr(frmMirage.Ctrl2Handle(i).hwnd, GCW_HCURSOR, frmMirage.Curs2Handle), frmMirage.Ctrl2Handle(i).name)
    Next i
    
    ' Change MousePointer argument to refresh the mouse icon
    Screen.MousePointer = Screen.MousePointer
End Sub

Public Sub RestoreCursor()
    If frmMirage.Curs2Handle Then
        Dim ctl As Control
        Dim i As Integer
        
        DestroyCursor frmMirage.Curs2Handle
    
        For i = 1 To frmMirage.Ctrl2Handle.Count
            Call SetClassLongPtr(frmMirage.Ctrl2Handle(i).hwnd, GCW_HCURSOR, frmMirage.SysCursHandle(frmMirage.Ctrl2Handle(i).name))
        Next i
        
    End If
End Sub

Public Sub InitMap()
    Debug.Print "oh ! "
    Call InitSurfacesAccordingMap
    Debug.Print "oh2 ! "

    ' Mettre ici toutes les commandes qui nécessitent MaxMapX et MaxMapY
    Dim i As Integer
    Dim X As Long, Y As Long

    ReDim TempTile(0 To MaxMapX, 0 To MaxMapY) As TempTileRec

    For Y = 0 To MaxMapY
        For X = 0 To MaxMapX
            TempTile(X, Y).DoorOpen = NO
        Next X
    Next Y
    
    ReDim MapItem(0 To MaxMapX, 0 To MaxMapY) As MapItemRec
End Sub

Public Function MinMapX()
    MinMapX = 0
End Function

Public Function MaxMapX()
    MaxMapX = UBound(Map.tile, 1)
End Function

Public Function MinMapY()
    MinMapY = 0
End Function

Public Function MaxMapY()
    MaxMapY = UBound(Map.tile, 2)
End Function

Public Function GetMapBordersCount()
    GetMapBordersCount = 0
    On Error Resume Next
    GetMapBordersCount = UBound(Map.Borders) + 1
End Function

Public Function GetNbMaterials(ByVal craftNum As Integer)
    GetNbMaterials = 0
    On Error Resume Next
    GetNbMaterials = UBound(Crafts(craftNum).Materials) + 1
End Function

Public Function GetNbProducts(ByVal craftNum As Integer)
    GetNbProducts = 0
    On Error Resume Next
    GetNbProducts = UBound(Crafts(craftNum).Products) + 1
End Function

Public Function IsItemInCooldown(ByVal itemNum As Integer)
    Dim i As Integer

    IsItemInCooldown = False
    
    For i = 1 To cooldownItem.Count
        If cooldownItem.item(i).itemNum = itemNum Then
            IsItemInCooldown = True
            Exit Function
        End If
    Next i
End Function

Public Sub SetWeaponSlot(ByVal Index As Integer, ByVal num As Integer, ByVal Value As Integer, ByVal dur As Integer)
    Call SetEquipement(Index, Player(Index).WeaponSlot, num, Value, dur)
End Sub

Public Sub SetArmorSlot(ByVal Index As Integer, ByVal num As Integer, ByVal Value As Integer, ByVal dur As Integer)
    Call SetEquipement(Index, Player(Index).ArmorSlot, num, Value, dur)
End Sub

Public Sub SetHelmetSlot(ByVal Index As Integer, ByVal num As Integer, ByVal Value As Integer, ByVal dur As Integer)
    Call SetEquipement(Index, Player(Index).HelmetSlot, num, Value, dur)
End Sub

Public Sub SetShieldSlot(ByVal Index As Integer, ByVal num As Integer, ByVal Value As Integer, ByVal dur As Integer)
    Call SetEquipement(Index, Player(Index).ShieldSlot, num, Value, dur)
End Sub

Private Sub SetEquipement(ByVal Index As Integer, ByRef slot As ItemSlotRec, ByVal num As Integer, ByVal Value As Integer, ByVal dur As Integer)
    With slot
        If .num >= 0 Then
            Call RemoveItemEffects(Index, .num)
        End If
        
        .num = num
        .Value = Value
        .dur = dur
        
        If .num >= 0 Then
            Call ApplyItemEffects(Index, .num)
        End If
        
        Call UpdateVisInv
    End With
End Sub

Public Sub ResetDragAndDrop()
    dragAndDrop.Index = -1
    dragAndDrop.Type = 0
    frmMirage.dragDropPicture.Visible = False
End Sub
