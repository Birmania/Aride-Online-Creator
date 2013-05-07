Attribute VB_Name = "modTypes"
'Option Explicit

'API
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)

' To compare user objects
'Public Declare Sub hmemcpy Lib "kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_CRAFTS As Long
Public MAX_AREAS As Long
Public MAX_MATERIALS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public HORS_LIGNE As Byte
Public MAX_LEVEL As Long
Public MAX_QUETES As Long
Public MAX_NPC_SPELLS As Long
Public MAX_DX_SPRITE As Long
Public MAX_DX_PAPERDOLL As Long
Public MAX_DX_SPELLS As Long
Public MAX_DX_BIGSPELLS As Long
Public MAX_DX_PETS As Long
Public MAX_PETS As Long

Public MAX_DREAMS As Long
Public Const MAX_DREAM_INSTANCES As Long = 100
Public Const MAX_DREAM_MAPS As Long = 50

Public Const MAX_ARROWS As Byte = 100
Public Const MAX_PLAYER_ARROWS As Byte = 100

Public MAX_INV As Integer
Public Const MAX_MAP_NPCS As Byte = 15
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_PLAYER_CRAFTS As Byte = 200
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10

Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 40
Public Const MAX_CHARS As Byte = 1

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 As String = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 As String = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 As String = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 As String = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
'Public Const MaxMapX = 30
'Public Const MaxMapY = 30
'Public MaxMapX As Long
'Public MaxMapY As Long
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_NO_PENALTY As Byte = 2

' Image constants/inconstants
Public Const PIC_X = 32
Public Const PIC_Y = 32
Public PIC_PL As Byte
Public PIC_NPC1 As Byte
Public PIC_NPC2 As Byte

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
'Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_NPC_SPAWN As Byte = 20
Public Const TILE_TYPE_BANK As Byte = 21
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
'    ITEM_TYPE_POTIONADDHP
'    ITEM_TYPE_POTIONADDMP
'    ITEM_TYPE_POTIONADDSP
'    ITEM_TYPE_POTIONSUBHP
'    ITEM_TYPE_POTIONSUBMP
'    ITEM_TYPE_POTIONSUBSP
    ITEM_TYPE_KEY
    ITEM_TYPE_CURRENCY
    ITEM_TYPE_SPELL
    ITEM_TYPE_MONTURE
    ITEM_TYPE_SCRIPT
    ITEM_TYPE_PET
End Enum
'Public Const ITEM_TYPE_NONE As Byte = 0
'Public Const ITEM_TYPE_WEAPON As Byte = 1
'Public Const ITEM_TYPE_MISSILE As Byte = 2
'Public Const ITEM_TYPE_ARMOR As Byte = 3
'Public Const ITEM_TYPE_HELMET As Byte = 4
'Public Const ITEM_TYPE_SHIELD As Byte = 5
'Public Const ITEM_TYPE_POTIONADDHP As Byte = 6
'Public Const ITEM_TYPE_POTIONADDTP As Byte = 7
'Public Const ITEM_TYPE_POTIONADDSP As Byte = 8
'Public Const ITEM_TYPE_POTIONSUBHP As Byte = 9
'Public Const ITEM_TYPE_POTIONSUBTP As Byte = 10
'Public Const ITEM_TYPE_POTIONSUBSP As Byte = 11
'Public Const ITEM_TYPE_KEY As Byte = 12
'Public Const ITEM_TYPE_CURRENCY As Byte = 13
'Public Const ITEM_TYPE_SPELL As Byte = 14
'Public Const ITEM_TYPE_MONTURE As Byte = 15
'Public Const ITEM_TYPE_SCRIPT As Byte = 16
'Public Const ITEM_TYPE_PET As Byte = 17

' Direction constants
Public Const DIR_UP As Byte = 3
Public Const DIR_DOWN As Byte = 0
Public Const DIR_LEFT As Byte = 1
Public Const DIR_RIGHT As Byte = 2

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Weather constants
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_THUNDER As Byte = 3

' Time constants
Public Const TIME_DAY As Byte = 0
Public Const TIME_NIGHT As Byte = 1

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
Public Const NPC_BEHAVIOR_SCRIPT As Byte = 6


' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME As Integer = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 23 ' In characters.
Public Const MAX_LINES As Byte = 3

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

'Public SourceBorder() As TileRec
'Public DestinationBorder() As TileRec

Public SourceBorder As New Collection
Public SourceBorderMap As Integer
Public SourceBorderDirection As Integer
Public DestinationBorder As New Collection

'To replace by clsPosition
Type PositionRec
    x As Integer
    y As Integer
End Type

Type MapPositionRec
    Map As Integer
    position As PositionRec
End Type

Type IndRec
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
End Type

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    num As Long
    value As Long
    dur As Long
End Type

Type CoffreTempRec 'coffre
    Numeros As Long
    Valeur As Long
    Durabiliter As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte
    
    SkillTime As Long
    SpellVar As Long
    SkillDone As Long
    
    Target As Long
    TargetType As Long
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
End Type

Type PlayerQueteRec
    Temps As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    indexe(1 To 15) As IndRec
End Type

Type PetPosRec
    x As Integer
    y As Integer
    dir As Byte
    XOffset As Integer
    YOffset As Integer
    Anim As Byte
End Type

Type PlayerRec

    ' General
    name As String
    guild As String
    Guildaccess As Byte
    Class As Long
    sprite As Long
    Level As Long
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    tp As Long
    
    ' Stats
    Str As Long
    def As Long
    speed As Long
    magi As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    PetSlot As Long
    
    ' Inventory
    inv() As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    pet As PetPosRec
    Crafts(0 To MAX_PLAYER_CRAFTS) As Integer
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    dir As Byte
    
    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxTP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SkillAnim() As SpellAnimRec
    BloodAnim As SpellAnimRec

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long
    
    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
    
    QueteEnCour As Long
    Quetep As PlayerQueteRec
    
    Anim As Byte
    
   'Paperdoll
   Casque As Long
   armure As Long
   arme As Long
   bouclier As Long
   'Fin paperdoll
End Type

Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Mask3 As Long '<--
    M3Anim As Long '<--
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Fringe3 As Long '<--
    F3Anim As Long '<--
    Type As Byte
    Datas() As Long
    Strings() As String
    Light As Long
    'GroundSet As Byte
    'MaskSet As Byte
    'AnimSet As Byte
    'Mask2Set As Byte
    'M2AnimSet As Byte
    'Mask3Set As Byte '<--
    'M3AnimSet As Byte '<--
    'FringeSet As Byte
    'FAnimSet As Byte
    'Fringe2Set As Byte
    'F2AnimSet As Byte
    'Fringe3Set As Byte '<--
    'F3AnimSet As Byte '<--
End Type

'Type TilePicRec
'    Mapping As Long
'    DD_Tile As DirectDrawSurface7
'End Type

Type NpcMapRec
    id As Integer
    x() As Byte
    y() As Byte
    dir As Byte
'    boucle As Byte
'    Hasardm As Byte
    Hasardp As Boolean
'    Imobile As Byte
    movementType As Byte
End Type

'Not use in editor environment for the moment
'Type InstanceRec
'    Players() As Integer
'End Type

Type AreaRec
    name As String

    SandStormFrequency As Single
    SnowingFrequency As Single
    RainingFrequency As Single
    ThunderingFrequency As Single
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
    'Up As Integer
    'Down As Integer
    'Left As Integer
    'Right As Integer
    Music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Indoors As Boolean
    tile() As TileRec
    'Npcs(1 To MAX_MAP_NPCS) As NpcMapRec
    Npcs() As NpcMapRec
    PanoInf As String
    TranInf As Byte
    PanoSup As String
    TranSup As Byte
    Fog As Integer
    FogAlpha As Byte
    borders() As BorderRec
    Area As Byte
End Type

Type DreamRec
    name As String
    'beginning As Integer
    beginning As MapPositionRec
    maps() As Integer
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
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
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
    def As Long
    speed As Long
    magi As Long
    
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
    
    'paperdoll As Byte
    'paperdollPic As Long
    
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

'Type ItemRec
'    name As String * NAME_LENGTH
'    desc As String * 150
'
'    Pic As Long
'    Type As Byte
'    Data1 As Long
'    Data2 As Long
'    Data3 As Long
'    StrReq As Long
'    DefReq As Long
'    SpeedReq As Long
'    ClassReq As Long
'    AccessReq As Byte
'
'    paperdoll As Byte
'    paperdollPic As Long
'
'    Empilable As Byte
'
'    AddHP As Long
'    AddMP As Long
'    AddSP As Long
'    AddStr As Long
'    AddDef As Long
'    AddMagi As Long
'    AddSpeed As Long
'    AddEXP As Long
'    attackSpeed As Long
'
'    NCoul As Long
'
'    Sex As Byte
'End Type

Type MapItemRec
    num As Long
    value As Long
    dur As Long
    
    x As Byte
    y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
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
    def As Integer
    Dex As Integer
    Sci As Integer
    Lang As Integer
    MaxHp As Long
    exp As Long
    SpawnTime As Byte
    AttackSpeed As Integer
    
    ItemNPC() As NPCEditorRec
    'quetenum As Long
    inv As Boolean
    vol As Boolean
    
    Spell() As Integer
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Long
    x As Byte
    y As Byte
    dir As Byte
    
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type TradeItemsRec
    value(1 To MAX_TRADES) As TradeItemRec
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
'    ClassReq As Long
    LevelReq As Integer
    Sound As Integer
    MPCost As Integer
    Type As Byte
'    Data1 As Long
'    Data2 As Long
'    Data3 As Long
    Range As Byte
    
    Big As Byte
    
    SkillAnim As Integer
    SkillTime As Integer
    SkillDone As Integer
    
    SkillIco As Integer
    
    Target As Byte
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

Type MaterialRec
    ItemNum As Integer
    Count As Integer
End Type

Type CraftRec
    name As String
    Materials() As MaterialRec
    Products() As MaterialRec
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

Type PetsRec
    nom As String * NAME_LENGTH
    sprite As Long
    addForce As Byte
    addDefence As Byte
End Type

'pour patchs
Type Fichiers
    nom As String
    version As String
    Chemins As String
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public quete() As QueteRec
Public Map() As MapRec
Public Dreams() As DreamRec
Public TempMap(0 To 5) As MapRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public PlayerAnim() As Long
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SkillRec
Public Crafts() As CraftRec
Public Areas() As AreaRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec
Public Experience() As Long
Public Pets() As PetsRec
Public CoffreTmp(1 To 30) As CoffreTempRec 'coffre

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
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type
Public Trade(1 To 6) As TradeRec

Type ArrowRec
    name As String
    Pic As Long
    Range As Byte
End Type
Public Arrows(0 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Long
    Time As Long
    Done As Byte
    y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

' Temporary variables
Public TempNpcsTab(0 To MAX_MAP_NPCS) As NpcMapRec ' Use to work on a temporaru variable when in map properties editor

Public Inventory As Long

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
    CPreVisu As Boolean
    LowEffect As Boolean
End Type
Public AccOpt As TpAccOpt

Sub ClearSpell(ByVal Index As Long)
    Spell(Index).name = vbNullString
'    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
'    Spell(Index).Data1 = 0
'    Spell(Index).Data2 = 0
'    Spell(Index).Data3 = 0
    Spell(Index).MPCost = 0
    Spell(Index).Sound = 0
    Spell(Index).Range = 0
    
    Spell(Index).Big = 0
    
    Spell(Index).SkillAnim = 0
    Spell(Index).SkillTime = 40
    Spell(Index).SkillDone = 1
    
    Spell(Index).SkillIco = 0
    
    Spell(Index).Target = 0
End Sub

Sub ClearShop(ByVal Index As Long)
Dim I As Long
Dim z As Long

    Shop(Index).name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString
    Shop(Index).FixesItems = 0
    Shop(Index).FixObjet = -1
    
    For z = 1 To 6
        For I = 1 To MAX_TRADES
            Shop(Index).TradeItem(z).value(I).GiveItem = 0
            Shop(Index).TradeItem(z).value(I).GiveValue = 0
            Shop(Index).TradeItem(z).value(I).GetItem = 0
            Shop(Index).TradeItem(z).value(I).GetValue = 0
        Next I
    Next z
End Sub

Sub ClearNpc(ByVal Index As Long)
Dim I As Long
With Npc(Index)
    .name = vbNullString
    '.AttackSay = vbNullString
    .sprite = 0
    .SpawnSecs = 0
    .Behavior = 0
    .Range = 0
'    .Str = 0
'    .def = 0
'    .speed = 0
'    .magi = 0
    .MaxHp = 0
    .exp = 0
    .SpawnTime = 0
'    .quetenum = 0
    .inv = 0
    .vol = 0
    Erase .ItemNPC
    Debug.Print "Clear NPc"
'    For i = 1 To MAX_NPC_DROPS
'        .ItemNPC(i).Chance = 0
'        .ItemNPC(i).ItemNum = -1
'        .ItemNPC(i).ItemValue = 0
'    Next i
End With
End Sub

Sub ClearDream(ByVal Index As Long)
    Dreams(Index).name = vbNullString
    With Dreams(Index).beginning
        .Map = -1
        .position.x = -1
        .position.y = -1
    End With
    'Dreams(Index).beginning = 0
    Erase Dreams(Index).maps
End Sub

Sub ClearCraft(ByVal Index As Long)
    Dim I As Integer

    Crafts(Index).name = vbNullString
    Erase Crafts(Index).Materials
    Erase Crafts(Index).Products
'    For I = 0 To MAX_MATERIALS
'        ' Init the need
'        Crafts(Index).Materials(I).ItemNum = -1
'        Crafts(Index).Materials(I).Count = 0
'
'        ' Init the products
'        Crafts(Index).Products(I).ItemNum = -1
'        Crafts(Index).Products(I).Count = 0
'    Next I
End Sub

Public Sub ClearArea(ByVal Index As Long)
    Areas(Index).name = vbNullString
    Areas(Index).SandStormFrequency = 0
    Areas(Index).RainingFrequency = 0
    Areas(Index).SnowingFrequency = 0
    Areas(Index).ThunderingFrequency = 0
End Sub

Sub ClearQuete(ByVal Index As Long)
    quete(Index).nom = vbNullString
    quete(Index).Data1 = 0
    quete(Index).Data2 = 0
    quete(Index).Data2 = 0
    quete(Index).description = vbNullString
    quete(Index).reponse = vbNullString
    quete(Index).String1 = vbNullString
    quete(Index).Temps = 0
    quete(Index).Type = 0
    Dim I As Long
    For I = 1 To 15
        quete(Index).indexe(I).Data1 = 1
        quete(Index).indexe(I).Data2 = 0
        quete(Index).indexe(I).Data3 = 0
        quete(Index).indexe(I).String1 = vbNullString
    Next I
    quete(Index).Recompence.exp = 0
    quete(Index).Recompence.objn1 = 1
    quete(Index).Recompence.objn2 = 1
    quete(Index).Recompence.objn3 = 1
    quete(Index).Recompence.objq1 = 0
    quete(Index).Recompence.objq2 = 0
    quete(Index).Recompence.objq3 = 0
    quete(Index).Case = 0
End Sub

Sub ClearTempTile()
Dim x As Long, y As Long

    ReDim TempTile(0 To MaxMapX, 0 To MaxMapY) As TempTileRec

    For y = 0 To MaxMapY
        For x = 0 To MaxMapX
            TempTile(x, y).DoorOpen = NO
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim n As Long

    Player(Index).name = vbNullString
    Player(Index).guild = vbNullString
    Player(Index).Guildaccess = 0
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).sprite = 0
    Player(Index).exp = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).tp = 0
        
    Player(Index).Str = 0
    Player(Index).def = 0
    Player(Index).speed = 0
    Player(Index).magi = 0
        
    For n = 1 To MAX_INV
        Player(Index).inv(n).num = 0
        Player(Index).inv(n).value = 0
        Player(Index).inv(n).dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
    Player(Index).PetSlot = 0
    
    Player(Index).pet.dir = 0
    Player(Index).pet.x = 0
    Player(Index).pet.y = 0
    
    Player(Index).Map = 0
    Player(Index).x = 0
    Player(Index).y = 0
    Player(Index).dir = 0
    
    ' Client use only
    Player(Index).MaxHp = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxTP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
    Player(Index).Moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0
    
    For I = 1 To MAX_SPELL_ANIM
        Player(Index).SkillAnim(I).CastedSpell = NO
        Player(Index).SkillAnim(I).SkillTime = 0
        Player(Index).SkillAnim(I).SpellVar = 0
        Player(Index).SkillAnim(I).SkillDone = 0
        
        Player(Index).SkillAnim(I).Target = 0
        Player(Index).SkillAnim(I).TargetType = 0
    Next I
    
    Player(Index).SpellNum = 0
    
    For I = 1 To MAX_BLT_LINE
        BattlePMsg(I).Index = 1
        BattlePMsg(I).Time = I
        BattleMMsg(I).Index = 1
        BattleMMsg(I).Time = I
    Next I
    
    Player(Index).QueteEnCour = 0
    Player(Index).Quetep.Data1 = 0
    Player(Index).Quetep.Data2 = 0
    Player(Index).Quetep.Data3 = 0
    Player(Index).Quetep.String1 = vbNullString
      
    For n = 1 To 15
        Player(Index).Quetep.indexe(n).Data1 = 0
        Player(Index).Quetep.indexe(n).Data2 = 0
        Player(Index).Quetep.indexe(n).Data3 = 0
        Player(Index).Quetep.indexe(n).String1 = vbNullString
    Next n
    
    Inventory = 1
End Sub

Sub ClearPlayerQuete(ByVal Index As Long)
Dim I As Long
        Player(MyIndex).QueteEnCour = 0
        Player(MyIndex).Quetep.Data1 = 0
        Player(MyIndex).Quetep.Data2 = 0
        Player(MyIndex).Quetep.Data3 = 0
        Player(MyIndex).Quetep.String1 = vbNullString
        Accepter = False
        
        For I = 1 To 15
            Player(MyIndex).Quetep.indexe(I).Data1 = 0
            Player(MyIndex).Quetep.indexe(I).Data2 = 0
            Player(MyIndex).Quetep.indexe(I).Data3 = 0
            Player(MyIndex).Quetep.indexe(I).String1 = 0
        Next I
End Sub

Sub ClearPet(ByVal Index As Long)
    Pets(Index).nom = ""
    Pets(Index).sprite = 0
    Pets(Index).addForce = 0
    Pets(Index).addDefence = 0
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).name = vbNullString
    Item(Index).desc = vbNullString
    
    Item(Index).Type = 0
    Erase Item(Index).Datas
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).DexReq = 0
    Item(Index).SciReq = 0
    Item(Index).LangReq = 0
'    Item(Index).ClassReq = -1
'    Item(Index).AccessReq = 0
    
    'Item(Index).paperdoll = 0
    'Item(Index).paperdollPic = 0
    
    Item(Index).Empilable = 0
    
    Item(Index).LifeEffect = 0
    Item(Index).AddHP = 0
    Item(Index).AddSLP = 0
    Item(Index).AddSTP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddSci = 0
    Item(Index).AddDex = 0
    Item(Index).AddLang = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
    
    Item(Index).NCoul = 0
End Sub

Sub ClearItems()
Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).num = 0
    MapItem(Index).value = 0
    MapItem(Index).dur = 0
    MapItem(Index).x = 0
    MapItem(Index).y = 0
End Sub

Sub ClearMap(ByVal mapNum As Long)
    Map(mapNum).name = vbNullString
    Map(mapNum).Moral = 0
    Map(mapNum).Indoors = False
    Map(mapNum).Music = "Aucune"
        
        
    Erase Map(mapNum).tile
    Erase Map(I).Npcs
    Map(I).PanoInf = vbNullString
    Map(I).TranInf = 0
    Map(I).PanoSup = vbNullString
    Map(I).TranSup = 0
    Map(I).Fog = 0
    Map(I).FogAlpha = 0
    Map(I).Area = 0
End Sub

Sub ClearMaps()
Dim I, j As Long
Dim x As Long
Dim y As Long

For I = 0 To MAX_MAPS
    Call ClearMap(I)
Next I

For I = 0 To 5
    TempMap(I).name = vbNullString
    TempMap(I).Moral = 0
    'TempMap(I).Up = 0
    'TempMap(I).Down = 0
    'TempMap(I).Left = 0
    'TempMap(I).Right = 0
    TempMap(I).Indoors = False
    TempMap(I).Music = "Aucune"
        
    For y = 0 To MaxMapY
        For x = 0 To MaxMapX
            TempMap(I).tile(x, y).Ground = 0
            TempMap(I).tile(x, y).Mask = 0
            TempMap(I).tile(x, y).Anim = 0
            TempMap(I).tile(x, y).Mask2 = 0
            TempMap(I).tile(x, y).M2Anim = 0
            TempMap(I).tile(x, y).Mask3 = 0 '<--
            TempMap(I).tile(x, y).M3Anim = 0 '<--
            TempMap(I).tile(x, y).Fringe = 0
            TempMap(I).tile(x, y).FAnim = 0
            TempMap(I).tile(x, y).Fringe2 = 0
            TempMap(I).tile(x, y).F2Anim = 0
            TempMap(I).tile(x, y).Fringe3 = 0 '<--
            TempMap(I).tile(x, y).F3Anim = 0 '<--
            TempMap(I).tile(x, y).Type = 0
            Erase TempMap(I).tile(x, y).Datas
            Erase TempMap(I).tile(x, y).Strings
            TempMap(I).tile(x, y).Light = 0
            'TempMap(i).tile(x, y).GroundSet = 0
            'TempMap(i).tile(x, y).MaskSet = 0
            'TempMap(i).tile(x, y).AnimSet = 0
            'TempMap(i).tile(x, y).Mask2Set = 0
            'TempMap(i).tile(x, y).M2AnimSet = 0
            'TempMap(i).tile(x, y).Mask3Set = 0 '<--
            'TempMap(i).tile(x, y).M3AnimSet = 0 '<--
            'TempMap(i).tile(x, y).FringeSet = 0
            'TempMap(i).tile(x, y).FAnimSet = 0
            'TempMap(i).tile(x, y).Fringe2Set = 0
            'TempMap(i).tile(x, y).F2AnimSet = 0
            'TempMap(i).tile(x, y).Fringe3Set = 0 '<--
            'TempMap(i).tile(x, y).F3AnimSet = 0 '<--
        Next x
    Next y
    Erase TempMap(I).Npcs
'    For x = 1 To MAX_MAP_NPCS
'        TempMap(i).Npcs(x).id = 0
''        TempMap(i).Npcs(x).boucle = 0
''        TempMap(i).Npcs(x).Hasardm = 1
'        TempMap(i).Npcs(x).Hasardp = 1
''        TempMap(i).Npcs(x).Imobile = 0
'        TempMap(i).Npcs(x).movementType = 0
'        Erase TempMap(i).Npcs(x).x
'        Erase TempMap(i).Npcs(x).y
'    Next x
    TempMap(I).PanoInf = vbNullString
    TempMap(I).TranInf = 0
    TempMap(I).PanoSup = vbNullString
    TempMap(I).TranSup = 0
    TempMap(I).Fog = 0
    TempMap(I).FogAlpha = 0
Next I
End Sub

Sub NetQueteType(ByVal Index As Integer)
    quete(Index).Data1 = 0
    quete(Index).Data2 = 0
    quete(Index).Data2 = 0
    quete(Index).String1 = vbNullString
    Dim I As Long
    For I = 1 To 15
        quete(Index).indexe(I).Data1 = 1
        quete(Index).indexe(I).Data2 = 0
        quete(Index).indexe(I).Data3 = 0
        quete(Index).indexe(I).String1 = vbNullString
    Next I
End Sub

Sub NetTempMap(ByVal Index As Byte)
Dim x As Long
Dim y As Long
    TempMap(Index).name = vbNullString
    TempMap(Index).Moral = 0
    TempMap(Index).Indoors = False


    Erase TempMap(Index).tile
    ReDim TempMap(Index).tile(0 To MaxMapX, 0 To MaxMapY) As TileRec
    TempMap(Index).PanoInf = vbNullString
    TempMap(Index).TranInf = 0
    TempMap(Index).PanoSup = vbNullString
    TempMap(Index).TranSup = 0
    TempMap(Index).Fog = 0
    TempMap(Index).FogAlpha = 0
End Sub

Sub ClearMapItems()
Dim x As Long

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).x = 0
    MapNpc(Index).y = 0
    MapNpc(Index).dir = 0
    
    ' Client use only
    MapNpc(Index).XOffset = 0
    MapNpc(Index).YOffset = 0
    MapNpc(Index).Moving = 0
    MapNpc(Index).Attacking = 0
    MapNpc(Index).AttackTimer = 0
    PNJAnim(Index) = 1
End Sub

Sub ClearMapNpcs()
Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(I)
    Next I
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    Player(Index).name = name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal guild As String)
    Player(Index).guild = guild
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
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
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
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Player(Index).HP = GetPlayerMaxHP(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then Player(Index).MP = GetPlayerMaxMP(Index)
End Sub

Function GetPlayerTP(ByVal Index As Long) As Long
    GetPlayerTP = Player(Index).tp
End Function

Sub SetPlayerTP(ByVal Index As Long, ByVal tp As Long)
    Player(Index).tp = tp

    If GetPlayerTP(Index) > GetPlayerMaxTP(Index) Then Player(Index).tp = GetPlayerMaxTP(Index)
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHp
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxTP(ByVal Index As Long) As Long
    GetPlayerMaxTP = Player(Index).MaxTP
End Function

Function Getplayerstr(ByVal Index As Long) As Long
    Getplayerstr = Player(Index).Str
End Function

Sub Setplayerstr(ByVal Index As Long, ByVal Str As Long)
    Player(Index).Str = Str
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).def
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal def As Long)
    Player(Index).def = def
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)
    Player(Index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).magi
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal magi As Long)
    Player(Index).magi = magi
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

'Function GetPlayerMap(ByVal Index As Long) As Long
'If Index <= 0 Then Exit Function
'    GetPlayerMap = Player(Index).Map
'End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapNum As Long)
    Player(Index).Map = mapNum
    WriteINI "CONFIG", "ERR", Val(mapNum), App.Path & "\Config.ini"
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
    Player(Index).dir = dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).inv(InvSlot).dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).inv(InvSlot).dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ShieldSlot = InvNum
End Sub

Function GetPlayerPetSlot(ByVal Index As Long) As Long
    GetPlayerPetSlot = Player(Index).PetSlot
End Function

Sub SetPlayerPetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).PetSlot = InvNum
End Sub

Public Sub SetTile(ByVal x As Integer, ByVal y As Integer)
Dim I As Integer
With Map(Player(MyIndex).Map).tile(x, y)
    If frmMirage.optBlocked.value = True Then .Type = TILE_TYPE_BLOCKED
    If frmMirage.optWarp.value = True Then
        .Type = TILE_TYPE_WARP
        ReDim .Datas(0 To 2) As Long
        .Datas(0) = EditorWarpMap
        .Datas(1) = EditorWarpX
        .Datas(2) = EditorWarpY
    ElseIf frmMirage.optMapBorder.value = True Then
        If SourceBorderMap = -1 Then
            For I = 1 To SourceBorder.Count
                If SourceBorder(I).x = x And SourceBorder(I).y = y Then
                    Exit Sub
                End If
            Next I
            
            If SourceBorder.Count = 0 Then
                Dim pos As New clsPosition
                pos.x = x
                pos.y = y
                SourceBorder.add pos
            Else
                Dim firstPos As clsPosition
                
                Set firstPos = SourceBorder(1)
                
                Set SourceBorder = Nothing
                Set SourceBorder = New Collection
                
                SourceBorder.add firstPos
                
                Dim beginLoop As Integer
                Dim endLoop As Integer
                Dim newPosition As clsPosition
                If Abs(firstPos.x - x) > Abs(firstPos.y - y) Then
                    ' On va faire une ligne horizontale
                
                    If firstPos.x > x Then
                        beginLoop = x
                        endLoop = firstPos.x
                        
                        For I = beginLoop To endLoop - 1
                            Set newPosition = New clsPosition
                            newPosition.x = endLoop - (I - beginLoop)
                            newPosition.y = firstPos.y
                            SourceBorder.add newPosition
                        Next I
                    Else
                        beginLoop = firstPos.x
                        endLoop = x
                        
                        For I = beginLoop + 1 To endLoop
                            Set newPosition = New clsPosition
                            newPosition.x = I
                            newPosition.y = firstPos.y
                            SourceBorder.add newPosition
                        Next I
'                        SourceBorder.add firstPos
                    End If
                    
'                    If firstPos.x > x Then
'                        SourceBorder.add firstPos
'                    End If
                Else
                    'Ligne verticale
                    If firstPos.y > y Then
                        beginLoop = y
                        endLoop = firstPos.y
                        
                        For I = beginLoop To endLoop - 1
                            Set newPosition = New clsPosition
                            newPosition.x = firstPos.x
                            newPosition.y = endLoop - (I - beginLoop)
                            SourceBorder.add newPosition
                        Next I
                    Else
                        beginLoop = firstPos.y
                        endLoop = y
                        
                        For I = beginLoop + 1 To endLoop
                            Set newPosition = New clsPosition
                            newPosition.x = firstPos.x
                            newPosition.y = I
                            SourceBorder.add newPosition
                        Next I
                        'SourceBorder.add firstPos
                    End If
                    

                    
'                    If firstPos.y > y Then
'                        SourceBorder.add firstPos
'                    End If
                End If
            End If
        
'            If SourceBorder.Count >= 2 Then
'                Dim firstPos As clsPosition
'                Dim lastPos As clsPosition
'
'                Set firstPos = SourceBorder(1)
'                Set lastPos = SourceBorder(SourceBorder.Count)
'
'                If firstPos.x = lastPos.x Then
'                    If lastPos.x <> x Then
'                        Set SourceBorder = Nothing
'                        Set SourceBorder = New Collection
'
'                        If lastPos.y = y Then
'                            SourceBorder.add lastPos
'                        End If
'                    End If
'                Else
'                    If lastPos.y <> y Then
'                        Set SourceBorder = Nothing
'                        Set SourceBorder = New Collection
'
'                        If lastPos.x = x Then
'                            SourceBorder.add lastPos
'                        End If
'                    End If
'                End If
'            End If
    
'            Dim pos As New clsPosition
'            pos.x = x
'            pos.y = y
'            SourceBorder.add pos
        End If
'        Else ' Source already selectionned
'            For i = 1 To SourceBorder.Count
'                With Map(Player(MyIndex).Map)
'                    ReDim Preserve .Borders(0 To GetMapBordersCount(Player(MyIndex).Map)) As BorderRec
'                    With .Borders(GetMapBordersCount(Player(MyIndex).Map) - 1)
'                        .XSource = DestinationBorder(i).x
'                        .YSource = DestinationBorder(i).y
'                        .MapDestination = SourceBorderMap
'                        .XDestination = SourceBorder(i).x
'                        .YDestination = SourceBorder(i).y
'                    End With
'                End With
'
'                With Map(SourceBorderMap)
'                    ReDim Preserve .Borders(0 To GetMapBordersCount(SourceBorderMap)) As BorderRec
'                    With .Borders(GetMapBordersCount(SourceBorderMap) - 1)
'                        .XSource = SourceBorder(i).x
'                        .YSource = SourceBorder(i).y
'                        .MapDestination = Player(MyIndex).Map
'                        .XDestination = DestinationBorder(i).x
'                        .YDestination = DestinationBorder(i).y
'                    End With
'                End With
'            Next i
'
'            Debug.Print "begin clear"
'            Set SourceBorder = Nothing
'            Set SourceBorder = New Collection
'            Set DestinationBorder = Nothing
'            Set DestinationBorder = New Collection
'            SourceBorderMap = -1
'            Debug.Print "end clear"
'        End If
    ElseIf frmMirage.optHeal.value = True Then
        .Type = TILE_TYPE_HEAL
    ElseIf frmMirage.optKill.value = True Then
        .Type = TILE_TYPE_KILL
    ElseIf frmMirage.optItem.value = True Then
        .Type = TILE_TYPE_ITEM
        ReDim .Datas(0 To 1) As Long
        .Datas(0) = ItemEditorNum
        .Datas(1) = ItemEditorValue
    ElseIf frmMirage.optNpcAvoid.value = True Then
        .Type = TILE_TYPE_NPCAVOID
    ElseIf frmMirage.optKey.value = True Then
        .Type = TILE_TYPE_KEY
        ReDim .Datas(0 To 1) As Long
        .Datas(0) = KeyEditorNum
        .Datas(1) = KeyEditorTake
    ElseIf frmMirage.optKeyOpen.value = True Then
        .Type = TILE_TYPE_KEYOPEN
        ReDim .Datas(0 To 1) As Long
        .Datas(0) = KeyOpenEditorX
        .Datas(1) = KeyOpenEditorY
        ReDim .Strings(0 To 0)
        .Strings(0) = KeyOpenEditorMsg
    ElseIf frmMirage.optShop.value = True Then
        .Type = TILE_TYPE_SHOP
        ReDim .Datas(0 To 0) As Long
        .Datas(0) = EditorShopNum
    ElseIf frmMirage.optCBlock.value = True Then
        .Type = TILE_TYPE_CBLOCK
        ReDim .Datas(0 To 2) As Long
        .Datas(0) = EditorItemNum1
        .Datas(1) = EditorItemNum2
        .Datas(2) = EditorItemNum3
    ElseIf frmMirage.optArena.value = True Then
        .Type = TILE_TYPE_ARENA
        ReDim .Datas(0 To 2) As Long
        .Datas(0) = Arena1
        .Datas(1) = Arena2
        .Datas(2) = Arena3
    ElseIf frmMirage.optSound.value = True Then
        .Type = TILE_TYPE_SOUND
        ReDim .Strings(0 To 0)
        .Strings(0) = SoundFileName
    ElseIf frmMirage.optSprite.value = True Then
        .Type = TILE_TYPE_SPRITE_CHANGE
        ReDim .Datas(0 To 2) As Long
        .Datas(0) = SpritePic
        .Datas(1) = SpriteItem
        .Datas(2) = SpritePrice
    ElseIf frmMirage.optSign.value = True Then
        .Type = TILE_TYPE_SIGN
        ReDim .Strings(0 To 0)
        .Strings(0) = SignLine1
    ElseIf frmMirage.optDoor.value = True Then
        .Type = TILE_TYPE_DOOR
    ElseIf frmMirage.optNotice.value = True Then
        .Type = TILE_TYPE_NOTICE
        ReDim .Strings(0 To 2)
        .Strings(0) = NoticeTitle
        .Strings(1) = NoticeText
        .Strings(2) = NoticeSound
    'elseif frmMirage.optChest.value = True Then
     '   .Type = TILE_TYPE_CHEST
      '  .Data1 = 0
       ' .Data2 = 0
        '.Data3 = 0
       ' .String1 = vbNullString
       ' .String2 = vbNullString
       ' .String3 = vbNullString                '
    ElseIf frmMirage.optClassChange.value = True Then
        .Type = TILE_TYPE_CLASS_CHANGE
        ReDim .Datas(0 To 1) As Long
        .Datas(0) = ClassChange
        .Datas(1) = ClassChangeReq
    ElseIf frmMirage.optScripted.value = True Then
        .Type = TILE_TYPE_SCRIPTED
        ReDim .Datas(0 To 0) As Long
        .Datas(0) = ScriptNum
    ElseIf frmMirage.OptBank.value = True Then
        .Type = TILE_TYPE_BANK
        ReDim .Strings(0 To 0)
        .Strings(0) = bankmsg
    ElseIf frmMirage.optcoffre.value = True Then
        .Type = TILE_TYPE_COFFRE
        ReDim .Datas(0 To 2) As Long
        .Datas(0) = CleCoffreNum
        .Datas(1) = CleCoffreSupr
        .Datas(2) = ObjCoffreNum
        ReDim .Strings(0 To 0)
        .Strings(0) = CodeCoffre
    ElseIf frmMirage.optportecode.value = True Then
        .Type = TILE_TYPE_PORTE_CODE
        ReDim .Strings(0 To 0)
        .Strings(0) = CodePorte
    ElseIf frmMirage.optBmont.value = True Then
        .Type = TILE_TYPE_BLOCK_MONTURE
    ElseIf frmMirage.optBniv.value Then
        .Type = TILE_TYPE_BLOCK_NIVEAUX
        ReDim .Datas(0 To 0) As Long
        .Datas(0) = NivMin
    ElseIf frmMirage.opttoit.value Then
        .Type = TILE_TYPE_TOIT
    ElseIf frmMirage.optBguilde.value Then
        .Type = TILE_TYPE_BLOCK_GUILDE
        ReDim .Strings(0 To 0)
        .Strings(0) = NomGuilde
    ElseIf frmMirage.optbtoit.value Then
        .Type = TILE_TYPE_BLOCK_TOIT
    ElseIf frmMirage.optBDir.value Then
        .Type = TILE_TYPE_BLOCK_DIR
        ReDim .Datas(0 To 2) As Long
        .Datas(0) = AccptDir1
        .Datas(1) = AccptDir2
        .Datas(2) = AccptDir3
    End If
End With
End Sub

Public Function GetMapNbNpcs(ByVal mapNum As Integer)
    If IsEmptyArray(ArrPtr(Map(mapNum).Npcs)) Then
        GetMapNbNpcs = 0
    Else
        GetMapNbNpcs = UBound(Map(mapNum).Npcs) + 1
    End If
End Function

Public Function GetDreamNbMaps(ByVal DreamNum As Integer)
    If IsEmptyArray(ArrPtr(Dreams(DreamNum).maps)) Then
        GetDreamNbMaps = 0
    Else
        GetDreamNbMaps = UBound(Dreams(DreamNum).maps) + 1
    End If
End Function

Public Function GetNpcNbDrops(ByVal npcNum As Integer)
    If IsEmptyArray(ArrPtr(Npc(npcNum).ItemNPC)) Then
        GetNpcNbDrops = 0
    Else
        GetNpcNbDrops = UBound(Npc(npcNum).ItemNPC) + 1
    End If
End Function

Public Function MaxMapX()
    MaxMapX = UBound(Map(0).tile, 1)
    On Error Resume Next
    MaxMapX = UBound(Map(Player(MyIndex).Map).tile, 1)
End Function

Public Function MaxMapY()
    MaxMapY = UBound(Map(0).tile, 2)
    On Error Resume Next
    MaxMapY = UBound(Map(Player(MyIndex).Map).tile, 2)
End Function

'Public Function GetBorderSize(ByRef MyArray() As TileRec)
'    GetBorderSize = 0
'    On Error Resume Next
'    GetBorderSize = UBound(MyArray) + 1
'End Function

'Public Sub AddToBorder(ByRef MyArray() As TileRec, ByRef tile As TileRec)
'    Dim i As Integer
'    For i = 0 To GetBorderSize(MyArray) - 1
'        If type2str(MyArray(i)) = type2str(tile) Then
'            Exit Sub
'        End If
'    Next i
'
'    ReDim MyArray(0 To GetBorderSize(MyArray)) As TileRec
'    MyArray(GetBorderSize(MyArray) - 1) = tile
'End Sub
'
'Public Function type2str(t As TileRec) As String
'    Dim s As String
'    s = Space$(Len(t))
'    'Call hmemcpy(ByVal s, t, Len(t))
'    Call CopyMemory(ByVal s, t, Len(t))
'    type2str = s
'End Function

Public Function GetMapBordersCount(ByVal mapNum As Integer)
    GetMapBordersCount = 0
    On Error Resume Next
    GetMapBordersCount = UBound(Map(mapNum).borders) + 1
End Function

Public Function GetNbMaterials(ByVal CraftNum As Integer)
    GetNbMaterials = 0
    On Error Resume Next
    GetNbMaterials = UBound(Crafts(CraftNum).Materials) + 1
End Function

Public Function GetNbProducts(ByVal CraftNum As Integer)
    GetNbProducts = 0
    On Error Resume Next
    GetNbProducts = UBound(Crafts(CraftNum).Products) + 1
End Function

Public Sub RemoveBorder(ByVal mapNum As Integer, ByVal x As Integer, ByVal y As Integer)
    Dim I As Integer
    Dim j As Integer

    For I = 0 To GetMapBordersCount(mapNum) - 1
        If Map(mapNum).borders(I).XSource = x And Map(mapNum).borders(I).YSource = y Then
            For j = I To GetMapBordersCount(mapNum) - 2
                Map(mapNum).borders(j) = Map(mapNum).borders(j + 1)
            Next j
            If GetMapBordersCount(mapNum) = 1 Then
                Erase Map(mapNum).borders
            Else
                ReDim Preserve Map(mapNum).borders(0 To GetMapBordersCount(mapNum) - 2) As BorderRec
            End If
            Exit For
        End If
    Next I
End Sub
