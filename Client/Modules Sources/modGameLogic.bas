Attribute VB_Name = "modGameLogic"
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


'***************************************************************************************************************************************************'
'ATTENTION : PENSER A NOTER LES MODIFICATIONS QUE VOUS APPORTER AU SOURCES POUR POUVOIR LES REFAIRE PLUS TARD SI VOUS DESIRER ACTUALISER LES SOURCES'
'***************************************************************************************************************************************************'

Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCAND As Long = &H8800C6
Public Const SRCCOPY As Long = &HCC0020
Public Const SRCPAINT As Long = &HEE0086

Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_ADDCHAR As Byte = 4
Public Const MENU_STATE_DELCHAR As Byte = 5
Public Const MENU_STATE_USECHAR As Byte = 6
Public Const MENU_STATE_INIT As Byte = 7

Public ShiftDown As Boolean
Public ControlDown As Boolean
Public directionOrder() As Byte
Public DirUpPressed As Boolean
Public DirDownPressed As Boolean
Public DirLeftPressed As Boolean
Public DirRightPressed As Boolean

Public TryToMove As Boolean

' Multi-Serveur
Public CHECK_WAIT As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Prediction of movement
Public movementController As Byte

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Boolean
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Utilise pour charger toutes les tiles avant affichage
Public InitTiles As Boolean

' Used to freeze controls when player is dead
Public IsDead As Boolean

' Used to check if in Toit or not
Public InToit As Boolean

' Game fps
Public GameFPS As Long

'Loc of pointer
Public CurX As Single '/case
Public CurY As Single '/case
Public PotX As Single 'réel
Public PotY As Single 'réel

' Used for atmosphere
Public AreasWeather As Collection
Public GameWeather As Long
Public GameTime As Boolean
Public RainIntensity As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public newX As Long
Public newY As Long
Public NewPlayerPicX As Long
Public NewPlayerPicY As Long
Public NewPlayerPOffsetX As Long
Public NewPlayerPOffsetY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long
Public DmgAddRem As Long
Public NPCDmgAddRem As Long

Public ii As Long, iii As Long
Public NpcIsDead As Boolean
Public sx As Long
Public sy As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public SoundFileName As String

Public Connucted As Boolean

'Pour la banque
Public bankmsg As String

'Pour les controlles
Public ConOff As Byte
Public OldMusic As String
Public Rep_Theme As String
Public numShop As Long

'Pour le mouvement des fenetre
Public drx As Long
Public dry As Long
Public dr As Boolean

'Pour les couleurs personalisables
Public AccModo As Long
Public AccMapeur As Long
Public AccDevelopeur As Long
Public AccAdmin As Long

'Variables pour FrmMirage
Public PicScWidth As Single
Public PicScHeight As Single
                    
Sub Main()
    If Len(command$()) > 0 Then
        If command$ = "reg_dll" Then
            Call RegisterDLLs
        End If
    Else
        ' Init error log very important constant
        ErrorLogFile = App.Path & "\Logs\errors.txt"
        
        If FileExist(ErrorLogFile) Then
            Kill ErrorLogFile
        End If
        
        'Init other constants
        ClientConfigurationFile = App.Path & "\Config\Client.ini"
        OptionConfigurationFile = App.Path & "\Config\Option.ini"
        ThemeConfigurationFile = App.Path & "\Themes.ini"
        Call InitDefaultINIValue
        Rep_Theme = ReadINI("THEMES", "Theme", ThemeConfigurationFile)
        ColorConfigurationFile = App.Path & Rep_Theme & "\Couleur.ini"
        
        Call RegisterDLLs
        
        If Not FileExist(App.Path & "\Config\Updater.ini") Or ReadINI("UPDATER", "actif", App.Path & "\Config\Updater.ini") = "1" Then
            'Launch Updater
            frmUpdater.Show
        Else
            Call LaunchGame
        End If
    End If
End Sub

Public Sub RegisterDLLs()
    ' Register DLLs
    Dim DLLDirectory As String
    Dim files() As String
    Dim file As Variant
    Dim fso As New FileSystemObject

    Dim systemDirectory As String

    systemDirectory = GetSystemDirectory
    
    'Check for DLL
    Dim commands As String
    commands = vbNullString
    DLLDirectory = App.Path & "\DLL"
    If FolderExists(DLLDirectory) Then
        Call ListFiles(DLLDirectory, files)
        If GetArraySize(files) > 0 Then
            For Each file In files
                If LCase(fso.GetExtensionName(CStr(file))) = "dll" Or LCase(fso.GetExtensionName(CStr(file))) = "ocx" Then
                    If Not IsDLLAvailable(CStr(file)) Then
                        commands = commands & " copy """ & DLLDirectory & "\" & CStr(file) & """ """ & systemDirectory & """ &"
                        commands = commands & " regsvr32 /s """ & systemDirectory & "\" & CStr(file) & """ &"
                    End If
                End If
            Next file
        End If
    End If
    
    If commands <> vbNullString Then
        If IsUserAnAdmin Then
            Call ExecuteCommandAsAdmin(commands)
        Else
            Call ExecuteApplicationAsAdmin(App.EXEName & ".exe", "reg_dll")
        End If
               
        'Check the success of the command
        For Each file In files
            If LCase(fso.GetExtensionName(CStr(file))) = "dll" Or LCase(fso.GetExtensionName(CStr(file))) = "ocx" Then
                If Not IsDLLAvailable(CStr(file)) Then
                    'Failure
                    MsgBox "You must accept UAC request in order to continue."
                    End
                End If
            End If
        Next file
    End If
End Sub

Sub LaunchGame()
Dim i As Long
Dim Ending As String
Dim t As Currency

    Call iniOptTouche
    
    dr = False
    frmsplash.Visible = True
    Call SetStatus("Vérification des dossiers...")
    DoEvents
                
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 400
    
    ReDim DD_SpriteSurf(0 To LoadMaxSprite()) As DirectDrawSurface7
    ReDim DDSD_Character(0 To LoadMaxSprite()) As DDSURFACEDESC2
        
    ReDim DD_PaperDollSurf(0 To LoadMaxPaperdolls()) As DirectDrawSurface7
    ReDim DDSD_PaperDoll(0 To LoadMaxPaperdolls()) As DDSURFACEDESC2
    
    ReDim DD_SpellAnim(0 To LoadMaxSpells()) As DirectDrawSurface7
    ReDim DDSD_SpellAnim(0 To LoadMaxSpells()) As DDSURFACEDESC2
    
    ReDim DD_BigSpellAnim(0 To LoadMaxBigSpells()) As DirectDrawSurface7
    ReDim DDSD_BigSpellAnim(0 To LoadMaxBigSpells()) As DDSURFACEDESC2
    
    ReDim DD_PetsSurf(0 To LoadMaxPet) As DirectDrawSurface7
    ReDim DDSD_Pets(0 To LoadMaxPet) As DDSURFACEDESC2
        
    With frmMirage
        .chkbubblebar.Value = ReadINI("CONFIG", "SpeechBubbles", ClientConfigurationFile)
        .chknpcbar.Value = ReadINI("CONFIG", "NpcBar", ClientConfigurationFile)
        .chknpcname.Value = ReadINI("CONFIG", "NPCName", ClientConfigurationFile)
        .chkplayerbar.Value = ReadINI("CONFIG", "PlayerBar", ClientConfigurationFile)
        .chkplayername.Value = ReadINI("CONFIG", "PlayerName", ClientConfigurationFile)
        .chkplayerdamage.Value = ReadINI("CONFIG", "NPCDamage", ClientConfigurationFile)
        .chknpcdamage.Value = ReadINI("CONFIG", "PlayerDamage", ClientConfigurationFile)
        .chkmusic.Value = ReadINI("CONFIG", "Music", ClientConfigurationFile)
        .chksound.Value = ReadINI("CONFIG", "Sound", ClientConfigurationFile)
        .chkAutoScroll.Value = ReadINI("CONFIG", "AutoScroll", ClientConfigurationFile)
        .chknobj.Value = Val(ReadINI("CONFIG", "NomObjet", ClientConfigurationFile))
        .chkLowEffect.Value = Val(ReadINI("CONFIG", "LowEffect", ClientConfigurationFile))
    End With
    
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 400
    
    If Not FileExist(ColorConfigurationFile) Then
        WriteINI "POLICE", "Police", "MS Sans Serif", ColorConfigurationFile
        WriteINI "POLICE", "PoliceSize", "8", ColorConfigurationFile
        WriteINI "POLICE", "PoliceChat", "MS Sans Serif", ColorConfigurationFile
        WriteINI "POLICE", "PoliceChatSize", "8", ColorConfigurationFile
    
        WriteINI "CHATBOX", "R", 152, ColorConfigurationFile
        WriteINI "CHATBOX", "G", 146, ColorConfigurationFile
        WriteINI "CHATBOX", "B", 120, ColorConfigurationFile
        
        WriteINI "CHATTEXTBOX", "R", 152, ColorConfigurationFile
        WriteINI "CHATTEXTBOX", "G", 146, ColorConfigurationFile
        WriteINI "CHATTEXTBOX", "B", 120, ColorConfigurationFile
        
        WriteINI "BACKGROUND", "R", 152, ColorConfigurationFile
        WriteINI "BACKGROUND", "G", 146, ColorConfigurationFile
        WriteINI "BACKGROUND", "B", 120, ColorConfigurationFile
        
        WriteINI "SPELLLIST", "R", 152, ColorConfigurationFile
        WriteINI "SPELLLIST", "G", 146, ColorConfigurationFile
        WriteINI "SPELLLIST", "B", 120, ColorConfigurationFile

        WriteINI "WHOLIST", "R", 152, ColorConfigurationFile
        WriteINI "WHOLIST", "G", 146, ColorConfigurationFile
        WriteINI "WHOLIST", "B", 120, ColorConfigurationFile
        
        WriteINI "NEWCHAR", "R", 152, ColorConfigurationFile
        WriteINI "NEWCHAR", "G", 146, ColorConfigurationFile
        WriteINI "NEWCHAR", "B", 120, ColorConfigurationFile
        
        WriteINI "BARE", "R", 128, ColorConfigurationFile
        WriteINI "BARE", "G", 128, ColorConfigurationFile
        WriteINI "BARE", "B", 255, ColorConfigurationFile
    End If
    
    Dim R1 As Long, G1 As Long, B1 As Long
    R1 = Val(ReadINI("CHATTEXTBOX", "R", ColorConfigurationFile))
    G1 = Val(ReadINI("CHATTEXTBOX", "G", ColorConfigurationFile))
    B1 = Val(ReadINI("CHATTEXTBOX", "B", ColorConfigurationFile))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)
       
    R1 = Val(ReadINI("FOND", "R", ColorConfigurationFile))
    G1 = Val(ReadINI("FOND", "V", ColorConfigurationFile))
    B1 = Val(ReadINI("FOND", "B", ColorConfigurationFile))
    With frmMirage
        .Picture9.BackColor = RGB(R1, G1, B1)
        .Picture8.BackColor = RGB(R1, G1, B1)
        .Picture11.BackColor = RGB(R1, G1, B1)
        .Picture13.BackColor = RGB(R1, G1, B1)
        .picInv3.BackColor = RGB(R1, G1, B1)
        .itmDesc.BackColor = RGB(R1, G1, B1)
        .picWhosOnline.BackColor = RGB(R1, G1, B1)
        .picGuildAdmin.BackColor = RGB(R1, G1, B1)
        .picGuild.BackColor = RGB(R1, G1, B1)
        .picEquip.BackColor = RGB(R1, G1, B1)
        .picPlayerSpells.BackColor = RGB(R1, G1, B1)
        .picOptions.BackColor = RGB(R1, G1, B1)
        .pictTouche.BackColor = RGB(R1, G1, B1)
        .chkbubblebar.BackColor = RGB(R1, G1, B1)
        .chknpcbar.BackColor = RGB(R1, G1, B1)
        .chknpcname.BackColor = RGB(R1, G1, B1)
        .chkplayerbar.BackColor = RGB(R1, G1, B1)
        .chkplayername.BackColor = RGB(R1, G1, B1)
        .chkplayerdamage.BackColor = RGB(R1, G1, B1)
        .chknpcdamage.BackColor = RGB(R1, G1, B1)
        .chkmusic.BackColor = RGB(R1, G1, B1)
        .chksound.BackColor = RGB(R1, G1, B1)
        .chkAutoScroll.BackColor = RGB(R1, G1, B1)
        .chknobj.BackColor = RGB(R1, G1, B1)
        .chkLowEffect.BackColor = RGB(R1, G1, B1)
    End With
    
    Debug.Print "test 3"
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 400
        
    R1 = Val(ReadINI("WHOLIST", "R", ColorConfigurationFile))
    G1 = Val(ReadINI("WHOLIST", "G", ColorConfigurationFile))
    B1 = Val(ReadINI("WHOLIST", "B", ColorConfigurationFile))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)
    
    'Init user commands
    Dim commands() As String, currentCommand As Variant
    commands = ReadINIKeys(OptionConfigurationFile, "COMMAND")

    If GetArraySize(commands) > 0 Then
        For Each currentCommand In commands
            Call UserCommand.Add(currentCommand, Val(ReadINI("COMMAND", CStr(currentCommand), OptionConfigurationFile)))
        Next
    End If
    
    Dim labels As New Collection
    Dim currentLabel As Variant
    
    With frmMirage
        UserCommandLabel.Add .lblCommandHaut
        UserCommandLabel.Add .lblCommandBas
        UserCommandLabel.Add .lblCommandGauche
        UserCommandLabel.Add .lblCommandDroite
        UserCommandLabel.Add .lblCommandAttaque
        UserCommandLabel.Add .lblCommandCourir
        UserCommandLabel.Add .lblCommandRamasser
        UserCommandLabel.Add .lblCommandAction
    
        For i = 0 To 13
            UserCommandLabel.Add .lblCommandRac(i)
        Next i
    End With
    
    Call SetStatus("Vérification du Statut...")
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 400
    
    frmsplash.Visible = True
    
    Call SetStatus("Initialisation des mises à jours...")
    Call InitAccountOpt
    Call InitMirageVars
    
    Call SetStatus("Initialisation du protocole TCP...")
        
    frmsplash.Shape1.Width = frmsplash.Shape1.Width + 400

    Call TcpInit
    Call InitMessages
    
    frmsplash.SetFocus
    frmServerChooser.Show
    
    ConOff = 0
    frmsplash.Visible = False
End Sub

Sub SetStatus(ByVal Caption As String)
    frmsplash.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    Connucted = True
    frmsplash.Visible = True
    frmsplash.Shape1.Width = 255
    Call SetStatus("Connection au Serveur...")
    Select Case State
        
        Case MENU_STATE_LOGIN
            frmMainMenu.fraLogin.Visible = False
            If CheckServerStatus Then Call SetStatus("Connecté, Envoie de la connexion au compte.."): Call SendLogin(frmMainMenu.txtName.Text, frmMainMenu.txtPassword.Text)
    End Select

    If Not IsConnected And Connucted = True Then
        frmMainMenu.Visible = True
        frmsplash.Visible = False
        Call MsgBox("Désoler, le serveur semble être indisponible, réessayer dans quelque minute ou visiter " & WEBSITE, vbOKOnly, Game_Name)
    End If
End Sub

Sub initRac()
Dim i As Integer
Dim PlayersConfigurationFile As String

    For i = 0 To 13
        rac(i, 0) = 0
        rac(i, 1) = 0
    Next i

    PlayersConfigurationFile = App.Path & "\Config\Players.ini"
    If Not FileExist(PlayersConfigurationFile) Then
        Call SaveRac
    End If
    
    For i = 0 To 13
        frmMirage.picRac(i).Picture = LoadPicture()
        rac(i, 0) = Val(ReadINI(GetPlayerName(MyIndex), "rac" & i, PlayersConfigurationFile))
        rac(i, 1) = Val(ReadINI(GetPlayerName(MyIndex), "type" & i, PlayersConfigurationFile))
    Next i
    
    Call AffRac
End Sub
Sub AffRac()
Dim i As Integer, indexNum As Integer
    For i = 0 To 13
        If Val(rac(i, 0)) >= 0 Then
            If Val(rac(i, 1)) = 1 Then
                indexNum = Player(MyIndex).skill(Val(rac(i, 0)))
            ElseIf Val(rac(i, 1)) = 2 Then
                indexNum = Player(MyIndex).Inv(Val(rac(i, 0))).num
            End If
            
            If indexNum < 0 Then
                frmMirage.picRac(i).Picture = LoadPicture()
            Else
                If Val(rac(i, 1)) = 1 Then
                    Call AffSurfPic(DD_ItemSurf, frmMirage.picRac(i), (skill(indexNum).SkillIco - (skill(indexNum).SkillIco \ 6) * 6) * PIC_X, (skill(indexNum).SkillIco \ 6) * PIC_Y)
                ElseIf Val(rac(i, 1)) = 2 Then
                    Call AffSurfPic(DD_ItemSurf, frmMirage.picRac(i), (item(indexNum).Pic - (item(indexNum).Pic \ 6) * 6) * PIC_X, (item(indexNum).Pic \ 6) * PIC_Y, False)
                    If IsItemInCooldown(indexNum) Then
                        Call ShadePictureBox(frmMirage.picRac(i))
                    End If
                    frmMirage.picRac(i).Refresh
                Else
                    frmMirage.picRac(i).Picture = LoadPicture()
                End If
            End If
        End If
    Next i
End Sub

Sub SaveRac()
Dim i As Integer
    For i = 0 To 13
        Call WriteINI(GetPlayerName(MyIndex), "rac" & i, CStr(rac(i, 0)), App.Path & "\Config\Players.ini")
        Call WriteINI(GetPlayerName(MyIndex), "type" & i, CStr(rac(i, 1)), App.Path & "\Config\Players.ini")
    Next i
End Sub

Sub useRac(Index As Integer)
Dim d As Byte
    If rac(Index, 0) >= 0 Then
        If rac(Index, 1) = 1 Then
            Call UseSkill(rac(Index, 0))
        End If
        
        If rac(Index, 1) = 2 Then
            If Player(MyIndex).Inv(rac(Index, 0)).num <= 0 Or Player(MyIndex).Inv(rac(Index, 0)).num > MAX_ITEMS Then Exit Sub
    
            Call SendUseItem(rac(Index, 0))
        End If
    Else
        Call AddText("Aucuns raccourci ici.", BrightRed)
    End If
End Sub

Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Byte
Dim FPS As Long
Dim TickMove As Long
Dim X As Long
Dim Y As Long
Dim i, J As Long
Dim rec_back As RECT
Dim Coulor As Long
Dim screen_xg As Integer 'Nb de cases a gauche du "milieu" de picscreen
Dim screen_xd As Integer 'Nb de cases a droite du "milieu" de picscreen
Dim screen_yh As Integer 'Nb de cases en haut du "milieu" de picscreen
Dim screen_yb As Integer 'Nb de cases en bas du "milieu" de picscreen
Dim MaxDrawMapX As Long 'Calcul du maximum a dessiner en X
Dim MinDrawMapX As Long 'Calcul du minimum a dessiner en X
Dim MaxDrawMapY As Long 'Calcul du maximum a dessiner en Y
Dim MinDrawMapY As Long 'Calcul du minimum a dessiner en Y
Dim PosXNpc As Integer 'Position X de référence du NPC
Dim PosYNpc As Integer 'Position Y de référence du NPC
Dim PosXNpcOffset As Long 'Décalage X de référence du NPC
Dim PosYNpcOffset As Long 'Décalage Y de référence du NPC
Dim PosXText As Integer 'Position X de référence du texte à afficher (Sauvegarder en cas de mort/respawn du npc)
Dim PosYText As Integer 'Position Y de référence du texte à afficher (Sauvegarder en cas de mort/respawn du npc)


    If Not InGame Then Exit Sub
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Modifier la police en jeu
    Call SetFont("Fixedsys", 20)
            
    ' Used for calculating fps
    TickFPS = 0
    TickMove = 0
    
    'Initialisation des variables pour les limites de la "vue" du joueur
    screen_xg = (frmMirage.picScreen.Width \ 64) - 1
    screen_xd = (frmMirage.picScreen.Width \ 32) - screen_xg - 1
    screen_yh = (frmMirage.picScreen.Height \ 64) - 1
    screen_yb = (frmMirage.picScreen.Height \ 32) - screen_yh - 1
    
    Do While InGame

        Tick = GetTickCount
        
        ' Check to make sure we are still connected
        InGame = IsConnected
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
rest:
            Do While NeedToRestoreSurfaces
                DoEvents
                Sleep 1
            Loop

            Err.Clear
            'On Error GoTo rest
            On Error Resume Next
            DD.RestoreAllSurfaces: Call DestroyDirectX: Call InitDirectX
            If Err.Number <> 0 Then
                Resume rest
            End If
        End If
        
        On Error GoTo 0

        ' On flush le frontbuffer
        rec.Top = 0
        rec.Left = 0
        rec.bottom = frmMirage.picScreen.Height + PIC_Y
        rec.Right = frmMirage.picScreen.Width + PIC_X
        Call DD_FrontBuffer.BltColorFill(rec, 0)
            
        If Not GettingMap And Not IsDead Then
            'Initialisation du RECT pour le backbuffer
            rec_back.Top = PIC_Y
            If frmMirage.picScreen.Height > (MaxMapY + 1) * PIC_Y Then
                rec_back.bottom = rec_back.Top + (MaxMapY + 1) * PIC_Y
            Else
                rec_back.bottom = rec_back.Top + frmMirage.picScreen.Height
            End If

            rec_back.Left = PIC_X
            If frmMirage.picScreen.Width > (MaxMapX + 1) * PIC_X Then
                rec_back.Right = rec_back.Left + (MaxMapX + 1) * PIC_X
            Else
                rec_back.Right = rec_back.Left + frmMirage.picScreen.Width
            End If

            sx = 32
            sy = 32
            
            'Calcul des variables pour l'affichage avec le scrolling
            If MaxMapX < screen_xg + screen_xd + 1 Then
                newX = Player(MyIndex).X * PIC_X + Player(MyIndex).XOffset
                NewXOffset = 0
                NewPlayerX = 0
                'sx = 0
            ElseIf Player(MyIndex).X <= screen_xg Then
                NewPlayerX = 0
                If Player(MyIndex).X = screen_xg And Player(MyIndex).dir = DIR_LEFT Then
                    newX = screen_xg * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                Else
                    newX = Player(MyIndex).X * PIC_X + Player(MyIndex).XOffset
                    NewXOffset = 0
                End If
            ElseIf MaxMapX - Player(MyIndex).X <= screen_xd Then
                NewPlayerX = MaxMapX - screen_xd - screen_xg
                If MaxMapX - Player(MyIndex).X = screen_xd And Player(MyIndex).dir = DIR_RIGHT Then
                    newX = screen_xg * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                Else
                    newX = (Player(MyIndex).X - MaxMapX + screen_xd + screen_xg) * PIC_X + Player(MyIndex).XOffset
                    NewXOffset = 0
                End If
            Else
                NewPlayerX = Player(MyIndex).X - screen_xg
                newX = screen_xg * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
            
            If MaxMapY < screen_yh + screen_yb + 1 Then
                newY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).YOffset
                NewYOffset = 0
                NewPlayerY = 0
                'sy = 0
            ElseIf Player(MyIndex).Y <= screen_yh Then
                NewPlayerY = 0
                If Player(MyIndex).Y = screen_yh And Player(MyIndex).dir = DIR_UP Then
                    newY = screen_yh * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                Else
                    newY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).YOffset
                    NewYOffset = 0
                End If
            ElseIf MaxMapY - Player(MyIndex).Y <= screen_yb Then
                NewPlayerY = MaxMapY - screen_yb - screen_yh
                If MaxMapY - Player(MyIndex).Y = screen_yb And Player(MyIndex).dir = DIR_DOWN Then
                    newY = screen_yh * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                Else
                    newY = (Player(MyIndex).Y - MaxMapY + screen_yb + screen_yh) * PIC_Y + Player(MyIndex).YOffset
                    NewYOffset = 0
                End If
            Else
                NewPlayerY = Player(MyIndex).Y - screen_yh
                newY = screen_yh * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
            
            'Calcul des variables de scrolling restante
            NewPlayerPicX = NewPlayerX * PIC_X
            NewPlayerPicY = NewPlayerY * PIC_Y
            NewPlayerPOffsetX = NewPlayerPicX + NewXOffset
            NewPlayerPOffsetY = NewPlayerPicY + NewYOffset
            
            MaxDrawMapX = NewPlayerX + screen_xg + screen_xd + 1
            MinDrawMapX = NewPlayerX - 1
            MaxDrawMapY = NewPlayerY + screen_yh + screen_yb + 1
            MinDrawMapY = NewPlayerY - 1
            If MaxDrawMapX > MaxMapX Then MaxDrawMapX = MaxMapX
            If MaxDrawMapY > MaxMapY Then MaxDrawMapY = MaxMapY
            If MinDrawMapX < 0 Then MinDrawMapX = 0
            If MinDrawMapY < 0 Then MinDrawMapY = 0

            ' Blit out tiles layers ground/anim1/anim2
            For Y = MinDrawMapY To MaxDrawMapY
                For X = MinDrawMapX To MaxDrawMapX
                    Call BltTile(X, Y)
                Next X
            Next Y
       
            For X = 0 To MaxMapX
                For Y = 0 To MaxMapY
                    If NbMapItems(X, Y) > 0 Then
                        Call BltItem(X, Y)
                    End If
                Next Y
            Next X

             For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) And Player(i).partyIndex = Player(MyIndex).partyIndex And Player(MyIndex).partyIndex > -1 Then
                    Call BltPlayerOmbre(i)
                    Call BltPlayerBar(i)
                End If
            Next i
             If AccOpt.PlayBar And Player(MyIndex).partyIndex > -1 Then
                 For i = 1 To MAX_PLAYERS
                     If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) And Player(i).partyIndex = Player(MyIndex).partyIndex Then
                         Call BltPlayerBar(i)
                     End If
                 Next i
             ElseIf AccOpt.PlayBar Then
                 Call BltPlayerBar(MyIndex)
             End If
             
             ' Blit out the sprite change attribute
             For Y = MinDrawMapY To MaxDrawMapY
                 For X = MinDrawMapX To MaxDrawMapX
                     If Map.tile(X, Y).Type = TILE_TYPE_SPRITE_CHANGE Then
                         Call BltSpriteChange(X, Y)
                         If PIC_PL > 1 Then Call BltSpriteChange2(X, Y)
                     End If
                 Next X
             Next Y
            
             ' Blit out the npcs
            For Each i In MapNpc.Keys
                If CLng(Npc(MapNpc(i).num).Vol) = 0 Then
                    'test = MapNpc(I)
                    Call BltNpc(MapNpc(i))
                    If AccOpt.NpcBar Then Call BltNpcBars(i)
                End If
            Next i
             
             ' Blit out players
             For i = 1 To MAX_PLAYERS
             
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call BltPlayer(i)
                End If

                If Pets(i).Map = GetPlayerMap(MyIndex) And Pets(i).Map <> -1 Then
                    Call BltNpc(Pets(i))
                End If
             Next i
             
             ' Blit arrows
             For Each i In ArrowsEffect.Keys
                Call BltArrowEffect(ArrowsEffect(i))
             Next i
             
             For i = 1 To SkillsEffect.Count
                Call BltSkillEffect(SkillsEffect(i))
             Next i

             ' Dessiner le haut des npc apres le bas des joueurs
             For Each i In MapNpc.Keys
                 If MapNpc(i).num > -1 And MapNpc(i).num < MAX_NPCS Then If CLng(Npc(MapNpc(i).num).Vol) = 0 Then If PIC_PL > 1 Then Call BltNpcTop(MapNpc(i))
             Next i
             
             For i = 1 To MAX_PLAYERS
                 If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                     'Ajout du haut du personnage pour le 32*64
                     If PIC_PL > 1 Then Call BltPlayerTop(i)
                     
                     If Player(i).LevelUpT + 3000 > Tick Then Call BltPlayerLevelUp(i) Else Player(i).LevelUpT = 0
                     Call BltEmoticons(i)
                     Call BltPlayerAnim(i)
                 End If
                 
                 If Pets(i).Map = GetPlayerMap(MyIndex) And Pets(i).Map <> -1 Then
                    If CLng(Npc(Pets(i).num).Vol) = 0 And PIC_PL > 1 Then
                        Call BltNpcTop(Pets(i))
                    End If
                 End If
             Next i
             
             'Blit blood
             For i = 1 To MAX_PLAYERS
                 If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                     Call BltBlood(i, PIC_X, PIC_Y, 40)
                     ' Call BltBlood(i) ferais aussi l'affaire car les autres paramètres peuvent être modifier selon le blood.png.
                     ' Le premier et le second paramètre sont la taille X et Y ce qui permet d'avoir des animations de sang 96X96 exemple.
                     ' Il se peux que le code demande à être modifié dans cette condition.
                     ' Le dernier paramètre est le temps de chaque image en ms (1000 ms = 1 seconde).
                 End If
             Next i
             
            'Verouiller le backbuffer pour pouvoir ecrire le nom des joueurs et de leur guildes
            TexthDC = DD_BackBuffer.GetDC
            If AccOpt.PlayName Then
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        Call BltPlayerGuildName(i)
                        Call BltPlayerName(i)
                    End If
                Next i
            End If
            
            'Draw NPC Names
            If AccOpt.NpcName Then
                For Each i In MapNpc.Keys
                    Call BltMapNPCName(i)
                Next i
            End If
        
            ' Draw damages
            If AccOpt.NpcDamage Then
                If NbDamageToDisplay() > 0 Then
                    i = 0
                    Do While i <= (NbDamageToDisplay() - 1)

                        If Tick < DamageDisplayer(i).time + 2000 Then
                            If DamageDisplayer(i).TargetType = NPC_TYPE Then
                                If MapNpc.Exists(DamageDisplayer(i).targetIndex) Then
                                    With MapNpc(DamageDisplayer(i).targetIndex)
                                        PosXText = .XOffset + (.X - NewPlayerX) * PIC_X + sx + PIC_X - ((1 + (Len(Str(DamageDisplayer(i).damage)) / 2)) * 10.6) - NewXOffset
                                        PosYText = .YOffset + (.Y - NewPlayerY) * PIC_Y + sx - 30 - NewYOffset - DamageDisplayer(i).offset
                                    End With
                                End If
                            ElseIf DamageDisplayer(i).TargetType = PLAYER_TYPE Then
                                With Player(DamageDisplayer(i).targetIndex)
                                    PosXText = .XOffset + (.X - NewPlayerX) * PIC_X + sx + PIC_X - ((1 + (Len(Str(DamageDisplayer(i).damage)) / 2)) * 10.6) - NewXOffset
                                    PosYText = .YOffset + (.Y - NewPlayerY) * PIC_Y + sx - 30 - NewYOffset - DamageDisplayer(i).offset
                                End With
                            ElseIf DamageDisplayer(i).TargetType = PET_TYPE Then
                                With Pets(DamageDisplayer(i).targetIndex)
                                    PosXText = .XOffset + (.X - NewPlayerX) * PIC_X + sx + PIC_X - ((1 + (Len(Str(DamageDisplayer(i).damage)) / 2)) * 10.6) - NewXOffset
                                    PosYText = .YOffset + (.Y - NewPlayerY) * PIC_Y + sx - 30 - NewYOffset - DamageDisplayer(i).offset
                                End With
                            End If

                            Call DrawText(TexthDC, PosXText, PosYText, DamageDisplayer(i).damage, QBColor(White))
                            
                            DamageDisplayer(i).offset = DamageDisplayer(i).offset + 1
                            i = i + 1
                        Else
                            Call RemoveDamageToDisplay(i)
                        End If
                    Loop
                End If
            End If
            
            Call DD_BackBuffer.ReleaseDC(TexthDC)
                                
            ' Blit out tile layer fringe
            For Y = MinDrawMapY To MaxDrawMapY
                For X = MinDrawMapX To MaxDrawMapX
                    Call BltFringeTile(X, Y)
                Next X
            Next Y

            'Dessiner les PNJs volant
            For Each i In MapNpc.Keys
                If MapNpc(i).num > -1 And MapNpc(i).num < MAX_NPCS Then
                    If CLng(Npc(MapNpc(i).num).Vol) <> 0 Then
                        Call BltNpc(MapNpc(i))
                        If AccOpt.NpcBar Then Call BltNpcBars(i)
                        If PIC_PL > 1 Then Call BltNpcTop(MapNpc(i))
                    End If
                End If
            Next i
        
            'If Not GettingMap And Not IsDead Then If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call BltWeather
            If Map.Indoors = 0 Then Call BltWeather
            
            'Dessin du brouillard
            If Map.Fog <> 0 And Not AccOpt.LowEffect And GameTime <> TIME_NIGHT Then Call BltFog(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
            
            If Player(MyIndex).SLP < (Player(MyIndex).MaxSLP / 2) Then
                Call BltSleep(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
            End If
    
            'Dessin de la nuit en "hight"
            If GameTime = TIME_NIGHT And Not AccOpt.LowEffect And Map.Indoors = 0 Then Call Night(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)
            
            TexthDC = DD_BackBuffer.GetDC
                If AccOpt.SpeechBubbles Then
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                            If Bubble(i).Text <> vbNullString Then Call BltPlayerText(i)
                            If Tick > Bubble(i).Created + DISPLAY_BUBBLE_TIME Then Bubble(i).Text = vbNullString
                        End If
                    Next i
                End If
            Call DD_BackBuffer.ReleaseDC(TexthDC)
            
            rec = rec_back
            rec.Left = rec.Left + XMapPadding
            rec.Right = rec.Right + XMapPadding
            rec.Top = rec.Top + YMapPadding
            rec.bottom = rec.bottom + YMapPadding
            
            If Not GettingMap And Not IsDead Then
                Call DD_FrontBuffer.Blt(rec, DD_BackBuffer, rec_back, DDBLT_WAIT)
            End If
        
            ' Dessiner les barres du joueur en haut à gauche de l'écran
            Call BltPlayerHotBars
        End If
        
        TexthDC = DD_FrontBuffer.GetDC
    

        ' Draw map name
        If Map.Moral = MAP_MORAL_NONE Then
            Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map.name)) / 2), 5 + sx, Trim$(Map.name), QBColor(White))
        ElseIf Map.Moral = MAP_MORAL_SAFE Then
            Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map.name)) / 2), 5 + sx, Trim$(Map.name), QBColor(White))
        ElseIf Map.Moral = MAP_MORAL_NO_PENALTY Then
            Call DrawText(TexthDC, (frmMirage.picScreen.Width / 2) - (Len(Trim$(Map.name)) / 2), 5 + sx, Trim$(Map.name), QBColor(Black))
        End If
        
        For i = 1 To MAX_BLT_LINE
            If BattlePMsg(i).Index > 0 Then
                If BattlePMsg(i).Color > 15 Then Coulor = BattlePMsg(i).Color Else Coulor = QBColor(BattlePMsg(i).Color)
                If BattlePMsg(i).time + 60000 > Tick Then Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).Y + PicScHeight - 80 + sx, Trim$(BattlePMsg(i).Msg), Coulor) Else BattlePMsg(i).Done = 0
            End If
            
            If BattleMMsg(i).Index > 0 Then
                If BattleMMsg(i).Color > 15 Then Coulor = BattleMMsg(i).Color Else Coulor = QBColor(BattleMMsg(i).Color)
                If BattleMMsg(i).time + 60000 > Tick Then Call DrawText(TexthDC, (PicScWidth - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).Y + PicScHeight - 80 + sx, Trim$(BattleMMsg(i).Msg), Coulor) Else BattleMMsg(i).Done = 0
            End If
        Next i

        'Dessin de la nuit en "low effect"
        If GameTime = TIME_NIGHT And AccOpt.LowEffect And Map.Indoors = 0 Then Call Night(MinDrawMapX, MaxDrawMapX, MinDrawMapY, MaxDrawMapY)

        ' Check if we are getting a map, and if we are tell them so
        If GettingMap Then
            Call DrawText(TexthDC, 36, 70, "Chargement de la Carte en cours...", QBColor(BrightCyan))
        ElseIf IsDead Then ' Check if we are dead, and if we are tell them so
            Call DrawText(TexthDC, 36, 70, "Vous êtes mort...", QBColor(BrightCyan))
        End If

        ' Release DC
        Call DD_FrontBuffer.ReleaseDC(TexthDC)
        
        
        rec.Top = PIC_Y
        rec.Left = PIC_X
        rec.bottom = rec.Top + frmMirage.picScreen.Height
        rec.Right = rec.Left + frmMirage.picScreen.Width
        
        Call dX.GetWindowRect(frmMirage.picScreen.hwnd, rec_pos)


        If Not NeedToRestoreSurfaces Then
            'Extraire la taille du picScreen depuis le FrontBuffer vers le PrimarySurf
            Call DD_PrimarySurf.Blt(rec_pos, DD_FrontBuffer, rec, DDBLT_WAIT)
        End If
        
        ' Refresh the party frame
        Call RefreshParty
        
        ' Refresh the skill enable
        Call RefreshSkills
        
        If TickMove <= Tick And Not GettingMap And Not IsDead Then
            ' Check if player is trying to move
            Call CheckMovement
            
            ' Check to see if player is trying to attack
            Call CheckAttack
            
            ' Process player movements (actually move them)
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                    
                    If Pets(i).Map = GetPlayerMap(MyIndex) Then
                        Call ProcessNpcMovement(Pets(i))
                    End If
                End If
            Next i
            
            ' Process npc movements (actually move them)
            ' Thanks to kryzalid who told me about this "kind of lag"
            For Each i In MapNpc.Keys
                If MapNpc(i).num > -1 Then Call ProcessNpcMovement(MapNpc(i))
            Next i
            
            ' Change map animation every 250 milliseconds
            If Tick > MapAnimTimer + 250 Then
                If Not MapAnim Then MapAnim = True Else MapAnim = False
                MapAnimTimer = Tick
            End If
            
            Call MakeMidiLoop
            TickMove = Tick + 30
            
            'Calcul des FPS
            TickFPS = TickFPS + 1
            If TickFPS >= 33 Then TickFPS = 0: GameFPS = FPS: FPS = 0
        End If


        If InitTiles Then
            ' On charge toutes les tiles pour ne pas avoir de ralentissement pendant que l'on marchera (accès disque)
            For X = 0 To MaxMapX
                For Y = 0 To MaxMapY
                    Call BltTile(X, Y)
                Next Y
            Next X
            
            If AreasWeather.Count > 0 Then
                GameWeather = AreasWeather.item(Str(Map.Area))
                    
                If IsEmptyArray(ArrPtr(DropRain)) Then
                    RainIntensity = 200
    
                    MAX_RAINDROPS = RainIntensity
                    ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
                    ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
                End If
            End If
            
            InitTiles = False
            GettingMap = False
        End If

        'Bloquer les FPS a 30 pour éviter de surcharger le processeur
        Do While GetTickCount < Tick + 30
            DoEvents
            Sleep 1
        Loop

        DoEvents
        FPS = FPS + 1
    Loop
    If Not deco Then
        frmMirage.Visible = False
        frmsplash.Visible = True
        Call SetStatus("Destroying game data...")
        
        ' Shutdown the game
        Call GameDestroy
    Else
        deco = False
    End If
End Sub

Sub GameDestroy()
    Dim i As Integer
    Dim ctl As Control

    On Error Resume Next
    
    Call RestoreCursor
    
    'Unhook mousewheel
    Set frmMirage.cSubclasserHooker = Nothing
    
    Call DestroyDirectX
    
    Call StopMidi
    
    ' DO NOT REMOVE THIS END
    End
End Sub

Sub BltTile(ByVal X As Long, ByVal Y As Long)
Dim Ground As Long
Dim Mask As Long
Dim Anim As Long
Dim Mask2 As Long
Dim M2Anim As Long
Dim Mask3 As Long
Dim M3Anim As Long
Dim GroundTileSet As Byte
Dim MaskTileSet As Byte
Dim AnimTileSet As Byte
Dim Mask2TileSet As Byte
Dim M2AnimTileSet As Byte
Dim Mask3TileSet As Byte
Dim M3AnimTileSet As Byte
Dim tx As Long
Dim ty As Long
    With Map.tile(X, Y)
        Ground = .Ground
        Mask = .Mask
        Anim = .Anim
        Mask2 = .Mask2
        M2Anim = .M2Anim
        Mask3 = .Mask3
        M3Anim = .M3Anim
    End With
    
    tx = (X - NewPlayerX) * PIC_X + sx - NewXOffset
    ty = (Y - NewPlayerY) * PIC_Y + sy - NewYOffset

    rec.Top = 0
    rec.bottom = 0
    rec.Left = 0
    rec.Right = 0

    Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Ground), rec, DDBLTFAST_WAIT)

    If (Not MapAnim) Or (Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask > 0 And TempTile(X, Y).DoorOpen = NO Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Mask), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Anim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (M2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Mask2), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M2Anim > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(M2Anim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (M3Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask3 > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Mask3), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M3Anim > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(M3Anim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'Utiliser pour dessiner le panorama
    With rec_pos
        .Top = (Y - NewPlayerY) * PIC_Y + sy - NewYOffset
        .bottom = .Top + PIC_Y
        .Left = (X - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    'Affichage du panorama inférieur si il y en à un
    If Trim$(Map.PanoInf) <> vbNullString Then
        rec.Top = Y * PIC_Y
        If rec.Top + PIC_Y > DDSD_PanoInf.lHeight Then rec.bottom = DDSD_PanoInf.lHeight: rec_pos.bottom = rec_pos.bottom - ((rec.Top + PIC_Y) - DDSD_PanoInf.lHeight) Else rec.bottom = rec.Top + PIC_Y
        rec.Left = X * PIC_X
        If rec.Left + PIC_Y > DDSD_PanoInf.lWidth Then rec.Right = DDSD_PanoInf.lWidth: rec_pos.Right = rec_pos.Right - ((rec.Left + PIC_X) - DDSD_PanoInf.lWidth) Else rec.Right = rec.Left + PIC_X
        If Map.TranInf = 1 And TypeName(DD_PanoInfSurf) <> "Nothing" Then Call DD_BackBuffer.Blt(rec_pos, DD_PanoInfSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC) Else If TypeName(DD_PanoInfSurf) <> "Nothing" Then Call DD_BackBuffer.Blt(rec_pos, DD_PanoInfSurf, rec, DDBLT_WAIT)
    End If
End Sub

Sub BltItem(ByVal X As Long, ByVal Y As Long)
    rec.Top = (item(MapItem(X, Y).items(0).num).Pic \ 6) * PIC_Y
    rec.bottom = rec.Top + PIC_Y
    rec.Left = (item(MapItem(X, Y).items(0).num).Pic - (item(MapItem(X, Y).items(0).num).Pic \ 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltSleep(ByVal MinX As Long, ByVal MaxX As Long, ByVal MinY As Long, ByVal MaxY As Long)
    Dim tx As Long, ty As Long
    Dim X As Long, Y As Long
    Dim opacity As Single
    
    tx = DDSD_Character(GetPlayerSprite(MyIndex)).lWidth / 4
    
    X = GetPlayerX(MyIndex) * PIC_X + sx + Player(MyIndex).XOffset + tx / 2
    Y = GetPlayerY(MyIndex) * PIC_Y + sy + Player(MyIndex).YOffset
    
    With rec
        .Top = DDSD_Sleep.lHeight / 2 - (Y - NewPlayerPOffsetY)
        ' +2 car on a une ligne qui peut se retrouver dans la partie haute non visible du backbuffer
        .bottom = .Top + (MaxY - MinY + 2) * PIC_Y
        .Left = DDSD_Sleep.lWidth / 2 - (X - NewPlayerPOffsetX)
        ' +2 car on a une ligne qui peut se retrouver dans la partie gauche non visible du backbuffer
        .Right = .Left + (MaxX - MinX + 2) * PIC_X
        If .Left < 0 Then .Left = 0
        If .Right > DDSD_Sleep.lWidth Then .Right = DDSD_Sleep.lWidth
        If .Top < 0 Then .Top = 0
        If .bottom > DDSD_Sleep.lHeight Then .bottom = DDSD_Sleep.lHeight
    End With

    With rec_pos
        .Top = Y - NewPlayerPOffsetY - (DDSD_Sleep.lHeight / 2 - rec.Top)
        .bottom = .Top + (rec.bottom - rec.Top)
        .Left = X - NewPlayerPOffsetX - (DDSD_Sleep.lWidth / 2 - rec.Left)
        .Right = .Left + (rec.Right - rec.Left)
    End With

    'Dessin du sommeil
    opacity = 1 - (Player(MyIndex).SLP / (Player(MyIndex).MaxSLP / 2))
    Call AlphaBlendDX(DD_SleepSurf, DDSD_Sleep, rec, rec_pos, opacity)
End Sub

Sub BltFog(ByVal MinX As Long, ByVal MaxX As Long, ByVal MinY As Long, ByVal MaxY As Long)
    'Initialisation du RECT source
    With rec_pos
        .Top = 0
        .bottom = (MaxY - MinY + 1) * PIC_Y
        .Left = 0
        .Right = .Left + (MaxX - MinX + 1) * PIC_X
    End With
    
    'Initialisation du RECT destination
    With rec
        .Top = -PIC_Y + (NewPlayerY * 32) + NewYOffset
        .bottom = .Top + rec_pos.bottom
        .Left = -PIC_X + (NewPlayerX * 32) + NewXOffset
        .Right = .Left + (MaxX - MinX + 1) * PIC_X
    End With
    
    'Dessin du brouillard
    ' TODO
    'Call AlphaBlendDX(rec_pos, rec, FogVerts)
End Sub

Sub BltFringeTile(ByVal X As Long, ByVal Y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long
Dim Fringe3 As Long
Dim F3Anim As Long
Dim FringeTileSet As Byte
Dim FAnimTileSet As Byte
Dim Fringe2TileSet As Byte
Dim F2AnimTileSet As Byte
Dim Fringe3TileSet As Byte
Dim F3AnimTileSet As Byte
Dim tx As Long
Dim ty As Long

    With Map.tile(X, Y)
        Fringe = .Fringe
        FAnim = .FAnim
        Fringe2 = .Fringe2
        F2Anim = .F2Anim
        Fringe3 = .Fringe3
        F3Anim = .F3Anim
    End With
    
    tx = (X - NewPlayerX) * PIC_X + sx - NewXOffset
    ty = (Y - NewPlayerY) * PIC_Y + sy - NewYOffset
    
    rec.Top = 0
    rec.bottom = 0
    rec.Left = 0
    rec.Right = 0
    
    If (Not MapAnim) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Fringe), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(FAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If

    If (Not MapAnim) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe2 > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Fringe2), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(F2Anim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (Not MapAnim) Or (F3Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe3 > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(Fringe3), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F3Anim > 0 Then
            Call DD_BackBuffer.BltFast(tx, ty, GetTileSurface(F3Anim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    'Affichage du panorama supérieur si il y en à un
    If Trim$(Map.PanoSup) <> vbNullString Then
        rec.Top = Y * PIC_Y
        If rec.Top + PIC_Y > DDSD_PanoSup.lHeight Then rec.bottom = DDSD_PanoSup.lHeight: rec_pos.bottom = rec_pos.bottom - ((rec.Top + PIC_Y) - DDSD_PanoSup.lHeight) Else rec.bottom = rec.Top + PIC_Y
        rec.Left = X * PIC_X
        If rec.Left + PIC_Y > DDSD_PanoSup.lWidth Then rec.Right = DDSD_PanoSup.lWidth: rec_pos.Right = rec_pos.Right - ((rec.Left + PIC_X) - DDSD_PanoSup.lWidth) Else rec.Right = rec.Left + PIC_X
        If Map.TranSup = 1 And TypeName(DD_PanoSupSurf) <> "Nothing" Then Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_PanoSupSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY) Else If TypeName(DD_PanoSupSurf) <> "Nothing" Then Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_PanoSupSurf, rec, DDBLTFAST_WAIT)
    End If
End Sub

Sub BltPlayerOmbre(ByVal Index As Long)
Dim X As Long, Y As Long

    If Index <= 0 And Index >= MAX_PLAYERS Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub

    X = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
    
    rec.Top = 5 * PIC_Y
    rec.bottom = rec.Top + PIC_Y
    rec.Left = 0 * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(X - NewPlayerPOffsetX, Y - NewPlayerPOffsetY, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub
Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim X As Long, Y As Long
Dim tx As Long, ty As Long

If Index <= 0 And Index >= MAX_PLAYERS Then Exit Sub
If Not IsPlaying(Index) Then Exit Sub
    ' Check for animation
    Anim = 1
    If Player(Index).Attacking = 0 Or Player(Index).Moving > 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
            Case DIR_LEFT
                If (Player(Index).XOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
        End Select
    Else
        If Player(Index).MovementTimer > GetTickCount Then Anim = 2
    End If

    ' Check to see if we want to stop making him attack
    If Player(Index).MovementTimer < GetTickCount Then Player(Index).Attacking = 0: Player(Index).MovementTimer = 0
    If Player(Index).AttackTimer < GetTickCount Then Player(Index).AttackTimer = 0
    
    ty = DDSD_Character(GetPlayerSprite(Index)).lHeight / 4
    tx = DDSD_Character(GetPlayerSprite(Index)).lWidth / 4
    
    rec.Top = GetPlayerDir(Index) * ty + (ty / 2)
    rec.bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx

    X = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16)
    Y = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset
    
    If X < 0 Then rec.Left = rec.Left - X: rec.Right = rec.Left + (tx + X): X = 0
    If Y < 0 Then rec.Top = rec.Top + (ty / 2): rec.bottom = rec.Top: Y = Player(Index).YOffset + sy
        
    Call DD_BackBuffer.BltFast(X - NewPlayerPOffsetX, Y - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(Index)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim X As Long, Y As Long
Dim tx As Long, ty As Long
    ' Check for animation
    Anim = 1
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
            Case DIR_LEFT
                If (Player(Index).XOffset > PIC_Y / 2) Then Anim = Player(Index).Anim
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = Player(Index).Anim
        End Select
    Else
        If Player(Index).MovementTimer > GetTickCount Then Anim = 2
    End If
   
    ' Check to see if we want to stop making him attack
    If Player(Index).MovementTimer < GetTickCount Then Player(Index).Attacking = 0: Player(Index).MovementTimer = 0
    If Player(Index).AttackTimer < GetTickCount Then Player(Index).AttackTimer = 0
                  
    ty = DDSD_Character(GetPlayerSprite(Index)).lHeight / 4
    tx = DDSD_Character(GetPlayerSprite(Index)).lWidth / 4
    
    rec.Top = GetPlayerDir(Index) * ty
    rec.bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx
    
    X = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    Y = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset - (ty / 2)
    
    If X < 0 Then rec.Left = rec.Left - X: rec.Right = rec.Left + (tx + X): X = 0
    If Y < 0 Then rec.Top = rec.Top + (ty / 2): rec.bottom = rec.Top: Y = Player(Index).YOffset + sy

     Call DD_BackBuffer.BltFast(X - NewPlayerPOffsetX, Y - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(Index)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Mid$(Trim$(Npc(MapNpc(Index).num).name), 1, 2) = "**" Then Exit Sub

With Npc(MapNpc(Index).num)
'Draw name
    TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
    If DDSD_Character(Npc(MapNpc(Index).num).sprite).lHeight = 128 And DDSD_Character(Npc(MapNpc(Index).num).sprite).lWidth = 128 Then
        TextY = MapNpc(Index).Y * PIC_Y - 14 + MapNpc(Index).YOffset - CLng(PIC_Y / 2) + 48
    Else
        TextY = MapNpc(Index).Y * PIC_Y - 14 + MapNpc(Index).YOffset - CLng(PIC_Y / 2) + 32
    End If
    If Npc(MapNpc(Index).num).Behavior = NPC_BEHAVIOR_QUETEUR Then
        DrawPlayerNameText TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY - (PIC_Y / 2), Trim$(.name), vbGreen
    Else
        DrawPlayerNameText TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY - (PIC_Y / 2), Trim$(.name), vbWhite
    End If
End With
End Sub

Sub BltNpc(ByRef MapNpc As clsMapNpc)
Dim Anim As Byte
Dim X As Long, Y As Long
Dim tx As Long, ty As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc.num < 0 Then Exit Sub
    
    ' Check for animation
    Anim = 1
    If MapNpc.Attacking = 0 Then
        Select Case MapNpc.dir
            Case DIR_UP
                If (MapNpc.YOffset > PIC_Y / 2) Then Anim = 0
            Case DIR_DOWN
                If (MapNpc.YOffset < PIC_Y / 2 * -1) Then Anim = 0
            Case DIR_LEFT
                If (MapNpc.XOffset > PIC_Y / 2) Then Anim = 0
            Case DIR_RIGHT
                If (MapNpc.XOffset < PIC_Y / 2 * -1) Then Anim = 0
        End Select
    Else
        If MapNpc.AttackTimer + 500 > GetTickCount Then Anim = 2
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc.AttackTimer + 1000 < GetTickCount Then MapNpc.Attacking = 0: MapNpc.AttackTimer = 0
    
    ty = DDSD_Character(Npc(MapNpc.num).sprite).lHeight / 4
    tx = DDSD_Character(Npc(MapNpc.num).sprite).lWidth / 4
    
    rec.Top = MapNpc.dir * ty + (ty / 2)
    rec.bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx

    X = MapNpc.X * PIC_X + sx + MapNpc.XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    Y = MapNpc.Y * PIC_Y + sx + MapNpc.YOffset
    
    If X < 0 Then rec.Left = rec.Left - X: rec.Right = rec.Left + (tx + X): X = 0
    If Y < 0 Then rec.Top = rec.Top + (ty / 2): rec.bottom = rec.Top: Y = MapNpc.YOffset + sy
    
    Call DD_BackBuffer.BltFast(X - NewPlayerPOffsetX, Y - NewPlayerPOffsetY, DD_SpriteSurf(Npc(MapNpc.num).sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltNpcTop(ByRef MapNpc As clsMapNpc)
Dim Anim As Byte
Dim X As Long, Y As Long
Dim tx As Long, ty As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc.num < 0 Then Exit Sub
    
    ' Check for animation
    Anim = 1
    If MapNpc.Attacking = 0 Then
        Select Case MapNpc.dir
            Case DIR_UP
                If (MapNpc.YOffset > PIC_Y / 2) Then Anim = 0
            Case DIR_DOWN
                If (MapNpc.YOffset < PIC_Y / 2 * -1) Then Anim = 0
            Case DIR_LEFT
                If (MapNpc.XOffset > PIC_Y / 2) Then Anim = 0
            Case DIR_RIGHT
                If (MapNpc.XOffset < PIC_Y / 2 * -1) Then Anim = 0
        End Select
    Else
        If MapNpc.AttackTimer + 500 > GetTickCount Then Anim = 2
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc.AttackTimer + 1000 < GetTickCount Then MapNpc.Attacking = 0: MapNpc.AttackTimer = 0
    

    ty = DDSD_Character(Npc(MapNpc.num).sprite).lHeight / 4
    tx = DDSD_Character(Npc(MapNpc.num).sprite).lWidth / 4
    
    rec.Top = MapNpc.dir * ty
    rec.bottom = rec.Top + (ty / 2)
    rec.Left = Anim * tx + tx
    rec.Right = rec.Left + tx
    
    If tx > 32 Then
        X = MapNpc.X * PIC_X + sx + MapNpc.XOffset - ((tx / 2) - 16) '(tx / 4) - ((tx / 4) / 2)
    Else
        X = MapNpc.X * PIC_X + sx + MapNpc.XOffset
    End If
    Y = MapNpc.Y * PIC_Y + sx + MapNpc.YOffset - (ty / 2)
    
    If X < 0 Then rec.Left = rec.Left - X: rec.Right = rec.Left + (tx + X): X = 0
    If Y < 0 Then rec.Top = rec.Top + (ty / 2): rec.bottom = rec.Top: Y = MapNpc.YOffset + sy

    Call DD_BackBuffer.BltFast(X - NewPlayerPOffsetX, Y - NewPlayerPOffsetY, DD_SpriteSurf(Npc(MapNpc.num).sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerLevelUp(ByVal Index As Long)
Dim X As Long
Dim Y As Long
    rec.Top = (32 \ TilesInSheets) * PIC_Y
    rec.bottom = rec.Top + PIC_Y
    rec.Left = (32 - (32 \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + 96
    
    If Index = MyIndex Then
        X = newX + sx
        Y = newY + sy
        Call DD_BackBuffer.BltFast(X - 32, Y - 10 - Player(Index).LevelUp, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        X = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
        Y = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset
        Call DD_BackBuffer.BltFast(X - NewPlayerPicX - 32 - NewXOffset, Y - NewPlayerPicY - 10 - Player(Index).LevelUp - NewYOffset, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    If Player(Index).LevelUp >= 3 Then Player(Index).LevelUp = Player(Index).LevelUp - 1 Else If Player(Index).LevelUp >= 1 Then Player(Index).LevelUp = Player(Index).LevelUp + 1
End Sub

Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0: Color = QBColor(Brown)
            Case 1: Color = AccModo
            Case 2: Color = AccMapeur
            Case 3: Color = AccDevelopeur
            Case 4: Color = AccAdmin
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
    ' Draw name
    TextX = Player(Index).X * PIC_X + sx + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    If DDSD_Character(GetPlayerSprite(Index)).lHeight = 128 And DDSD_Character(GetPlayerSprite(Index)).lWidth = 128 Then
        TextY = Player(Index).Y * PIC_Y + sx + Player(Index).YOffset - 40 - ((PIC_NPC1 - 1) * 10) + 16
    Else
        TextY = Player(Index).Y * PIC_Y + sx + Player(Index).YOffset - 40 - ((PIC_NPC1 - 1) * 10)
    End If
    Call DrawText(TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY, GetPlayerName(Index), Color)
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0: Color = QBColor(red)
            Case 1: Color = QBColor(BrightCyan)
            Case 2: Color = QBColor(Yellow)
            Case 3: Color = QBColor(BrightGreen)
            Case 4: Color = QBColor(Yellow)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
    ' Draw name
    TextX = Player(Index).X * PIC_X + sx + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)
    TextY = Player(Index).Y * PIC_Y + sx + Player(Index).YOffset - (PIC_Y \ 2) - 10 - ((PIC_NPC1 - 1) * 10)
    Call DrawText(TexthDC, TextX - NewPlayerPOffsetX, TextY - NewPlayerPOffsetY, GetPlayerGuild(Index), Color)
End Sub

Sub ProcessMovement(ByVal Index As Long)
' vérifier si le joueur(sprite) ne va pas trop loin

' De temps en temps les valeurs X, Y ne sont pas à 0 alors qu'on ne se déplace pas sur leurs axes. Il faut harmoniser.
' De plus, Il faut remettre les valeurs à 0 si on arrive en bout de case
If GetPlayerDir(Index) = DIR_UP Or GetPlayerDir(Index) = DIR_DOWN Then
    Player(Index).XOffset = 0
    If GetPlayerDir(Index) = DIR_UP And Player(Index).YOffset < 0 Then
        Player(Index).YOffset = 0
    End If
    If GetPlayerDir(Index) = DIR_DOWN And Player(Index).YOffset > 0 Then
        Player(Index).YOffset = 0
    End If
Else ' DIR_LEFT or DIR_RIGHT
    Player(Index).YOffset = 0
    If GetPlayerDir(Index) = DIR_LEFT And Player(Index).XOffset < 0 Then
        Player(Index).XOffset = 0
    End If
    If GetPlayerDir(Index) = DIR_RIGHT And Player(Index).XOffset > 0 Then
        Player(Index).XOffset = 0
    End If
End If

If Player(Index).XOffset = 0 And Player(Index).YOffset = 0 Then
    If Player(Index).Moving > 0 Then

        If Index = MyIndex Then
            If Player(Index).Destination.X <> -1 Then
                If (Player(Index).Destination.X = GetPlayerX(Index) And Player(Index).Destination.Y = GetPlayerY(Index)) Then
                    Call ClearPlayerMove(Index)
                    Exit Sub
                End If
            Else
                If (TryToMove And Not CanMove) Or (Not TryToMove) Then
                    Call ClearPlayerMove(Index)
                    Call SendPlayerStopMove
                    Exit Sub
                Else
                    If CheckTeleport Then
                        'On souhaite se TP
                        Exit Sub
                    End If
                End If
            End If
        Else
            If GetArraySize(Player(Index).newDir) > 0 Then
                If (Player(Index).newDir(0).X = GetPlayerX(Index) And Player(Index).newDir(0).Y = GetPlayerY(Index)) Then
                    Call SetPlayerDir(Index, Player(Index).newDir(0).dir)
                    Call RemovePlayerDirection(Index)
                    If Player(Index).MovingSave <> 0 Then
                        Player(Index).Moving = Player(Index).MovingSave
                        Player(Index).MovingSave = 0
                    End If
                Else
                    If Player(Index).MovingSave = 0 Then
                        Player(Index).MovingSave = Player(Index).Moving
                        Player(Index).Moving = MOVING_RUNNING 'Must be change to MOVING_CORRECTION
                    End If
                End If
            End If
            
            If GetArraySize(Player(Index).newDir) = 0 Then
            
                If Player(Index).Destination.X <> -1 Then
                    If (Player(Index).Destination.X = GetPlayerX(Index) And Player(Index).Destination.Y = GetPlayerY(Index)) Then
                        Call ClearPlayerMove(Index)
                        Exit Sub
                    Else
                        Player(Index).Moving = MOVING_RUNNING
                    End If
                End If
            End If
        End If

        Select Case GetPlayerDir(Index)
            Case DIR_UP
                Player(Index).YOffset = PIC_Y
                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
        
            Case DIR_DOWN
                Player(Index).YOffset = PIC_Y * -1
                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
        
            Case DIR_LEFT
                Player(Index).XOffset = PIC_X
                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
        
            Case DIR_RIGHT
                Player(Index).XOffset = PIC_X * -1
                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
        End Select
    End If
End If

' Verifier si le joueur à une monture
If Player(Index).ArmorSlot.num > 0 And Player(Index).ArmorSlot.num < MAX_ITEMS Then
If (Player(Index).Moving = MOVING_WALKING Or Player(Index).Moving = MOVING_RUNNING) And item(Player(Index).ArmorSlot.num).Type = ITEM_TYPE_MONTURE Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - (MOVING_WALKING * item(Player(Index).ArmorSlot.num).Datas(1))
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + (MOVING_WALKING * item(Player(Index).ArmorSlot.num).Datas(1))
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - (MOVING_WALKING * item(Player(Index).ArmorSlot.num).Datas(1))
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + (MOVING_WALKING * item(Player(Index).ArmorSlot.num).Datas(1))
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - (MOVING_WALKING + ((MOVING_WALKING / 100) * item(Player(Index).ArmorSlot.num).Datas(1)))
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + (MOVING_WALKING + ((MOVING_WALKING / 100) * item(Player(Index).ArmorSlot.num).Datas(1)))
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - (MOVING_WALKING + ((MOVING_WALKING / 100) * item(Player(Index).ArmorSlot.num).Datas(1)))
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + (MOVING_WALKING + ((MOVING_WALKING / 100) * item(Player(Index).ArmorSlot.num).Datas(1)))
            End Select
        End If

        ' Check if completed walking over to the next tile
        Exit Sub
    End If
End If

' Check if player is walking, and if so process moving them over
If Player(Index).Moving = MOVING_WALKING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - MOVING_WALKING
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + MOVING_WALKING
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - MOVING_WALKING
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + MOVING_WALKING
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - MOVING_WALKING
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + MOVING_WALKING
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - MOVING_WALKING
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + MOVING_WALKING
            End Select
        End If
        
    ' Check if completed walking over to the next tile
   ' Check if player is running, and if so process moving them over
    ElseIf Player(Index).Moving = MOVING_RUNNING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - MOVING_RUNNING
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + MOVING_RUNNING
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - MOVING_RUNNING
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + MOVING_RUNNING
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - MOVING_RUNNING
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + MOVING_RUNNING
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - MOVING_RUNNING
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + MOVING_RUNNING
            End Select

        End If
    End If
End Sub

Sub ProcessNpcMovement(ByRef MapNpc As clsMapNpc)
    ' Check if npc is walking, and if so process moving them over

    If MapNpc.Moving > 0 Then
        
        ' Check if completed walking over to the next tile
        If (MapNpc.XOffset = 0) And (MapNpc.YOffset = 0) Then
            If MapNpc.newNpcDir.Count > 0 Then
                If (MapNpc.newNpcDir.item(1).X = GetNpcX(MapNpc) And MapNpc.newNpcDir.item(1).Y = GetNpcY(MapNpc)) Then
                    Call SetNpcDir(MapNpc, MapNpc.newNpcDir.item(1).dir)
                    Call RemoveNpcDirection(MapNpc)
                    If MapNpc.MovingSave <> 0 Then
                        MapNpc.Moving = MapNpc.MovingSave
                        MapNpc.MovingSave = 0
                    End If
                Else
                    If MapNpc.MovingSave = 0 Then
                        MapNpc.MovingSave = MapNpc.Moving
                        MapNpc.Moving = MOVING_RUNNING 'Must be change to MOVING_CORRECTION
                    End If
                End If
            'Else
            End If
            ' Tout le temps tester la destination (et pas seulement quand il n'y a plus de direction) car sinon on peut la louper
            ' si par exemple on reçoit un changement de destination puis un stop move sur la case de changement de direction (bas d'un blocage)
            If MapNpc.newNpcDir.Count = 0 Then
                If MapNpc.Destination.X <> -1 Then
                    If (MapNpc.Destination.X = GetNpcX(MapNpc) And MapNpc.Destination.Y = GetNpcY(MapNpc)) Then
                        Call ClearMapNpcMove(MapNpc)
                        Exit Sub
                    Else
                        MapNpc.Moving = MOVING_RUNNING
                    End If
                End If
            End If
        
            Select Case MapNpc.dir
                Case DIR_UP
                    MapNpc.YOffset = PIC_Y
                    Call SetNpcY(MapNpc, GetNpcY(MapNpc) - 1)
            
                Case DIR_DOWN
                    MapNpc.YOffset = PIC_Y * -1
                    Call SetNpcY(MapNpc, GetNpcY(MapNpc) + 1)
            
                Case DIR_LEFT
                    MapNpc.XOffset = PIC_X
                    Call SetNpcX(MapNpc, GetNpcX(MapNpc) - 1)
            
                Case DIR_RIGHT
                    MapNpc.XOffset = PIC_X * -1
                    Call SetNpcX(MapNpc, GetNpcX(MapNpc) + 1)
            End Select
        End If
        
        Select Case MapNpc.dir
            Case DIR_UP
                MapNpc.YOffset = MapNpc.YOffset - MOVING_WALKING
            Case DIR_DOWN
                MapNpc.YOffset = MapNpc.YOffset + MOVING_WALKING
            Case DIR_LEFT
                MapNpc.XOffset = MapNpc.XOffset - MOVING_WALKING
            Case DIR_RIGHT
                MapNpc.XOffset = MapNpc.XOffset + MOVING_WALKING
        End Select
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim name As String
Dim i As Long
Dim n As Long

If Len(frmMirage.txtMyTextBox.Text) > 200 Then
    MyText = Left(frmMirage.txtMyTextBox.Text, 200)
Else
    MyText = frmMirage.txtMyTextBox.Text
End If
' Handle when the player presses the return key
    

    If (KeyAscii = vbKeyReturn) Then
        If frmMirage.txtMyTextBox.Locked = False Then
            frmMirage.txtMyTextBox.Locked = True
            frmMirage.txtMyTextBox.Visible = False
            frmMirage.Canal.Visible = False
        Else
            frmMirage.txtMyTextBox.Text = vbNullString
            frmMirage.txtMyTextBox.Locked = False
            frmMirage.txtMyTextBox.Visible = True
            frmMirage.txtMyTextBox.SetFocus
            frmMirage.Canal.Visible = True
            Exit Sub
        End If
    
        On Error Resume Next
        
        ' message de guilde
       If Mid$(MyText, 1, 1) = "*" Then
           ChatText = Mid$(MyText, 2, Len(MyText) - 1)
           If Len(Trim$(ChatText)) > 0 Then Call GuildeMsg(ChatText)
           MyText = vbNullString
           Exit Sub
       End If
       
        ' Player message
        If Mid$(MyText, 1, 1) = "!" Or Mid$(MyText, 1, 1) = "w" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            name = vbNullString
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 1) <> " " Then name = name & Mid$(ChatText, i, 1) Else Exit For
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Utiliser: !nomjoueur msgici", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        If (Mid$(MyText, 1, 2) = "/w" Or Mid$(MyText, 1, 2) = "/W") And Mid$(MyText, 1, 7) <> "/warpto" Then
            ChatText = Mid$(MyText, 3, Len(MyText) - 2)
            name = vbNullString
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 1) <> " " Then name = name & Mid$(ChatText, i, 1) Else Exit For
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Utiliser: /wnomjoueur msgici", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        ' // Commands //
        
        ' Verification User
        If LCase$(Mid$(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
                
                        
        ' Checking fps
        If LCase$(Mid$(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Show inventory
        If LCase$(Mid$(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInv3.Visible = True
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Request stats
        If LCase$(Mid$(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
         
        ' Refresh Player
        If LCase$(Mid$(MyText, 1, 8)) = "/refresh" Then
            ConOff = ConOff + 1
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase$(Mid$(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Or LCase$(Mid$(MyText, 1, 5)) = "/refu" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If
                
        ' Party request
        If LCase$(Mid$(MyText, 1, 6)) = "/party" Or LCase$(Mid$(MyText, 1, 7)) = "/groupe" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                If Player(MyIndex).partyIndex = -1 Then
                    Dim partyName As String
                    partyName = InputBox("Nom du groupe :", "Nom du groupe")
                    Call SendRequestParty(ChatText, partyName)
                Else
                    Call SendRequestParty(ChatText)
                End If
            Else
                Call AddText("Utiliser : /group nomdujoueur", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Dream request
        If LCase$(Mid$(MyText, 1, 6)) = "/sleep" Then
            If GetPlayerSLP(MyIndex) <= 0.2 * GetPlayerMaxSLP(MyIndex) Then
                Call SendSleep
            Else
                Call AddText("Tu n'est pas assez fatigué pour dormir.", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' day night command
            If LCase$(Mid$(MyText, 1, 9)) = "/daynight" Or LCase$(Mid$(MyText, 1, 9)) = "/journuit" Then
                Call SendGameTime
                MyText = vbNullString
                Exit Sub
            End If
            
            ' weather command
            If LCase$(Mid$(MyText, 1, 8)) = "/weather" Or LCase$(Mid$(MyText, 1, 6)) = "/temps" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & Val(MyText) & SEP_CHAR & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Or Trim$(LCase$(MyText)) = "rien" Then i = 0
                        If Trim$(LCase$(MyText)) = "rain" Or Trim$(LCase$(MyText)) = "pluie" Then i = 1
                        If Trim$(LCase$(MyText)) = "snow" Or Trim$(LCase$(MyText)) = "neige" Then i = 2
                        If Trim$(LCase$(MyText)) = "thunder" Or Trim$(LCase$(MyText)) = "orage" Then i = 3
                        Call SendData("weather" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Admin Message
            If Mid$(MyText, 1, 1) = "=" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then Call AdminMsg(ChatText)
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase$(Mid$(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Map report
            If LCase$(Mid$(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase$(Mid$(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid$(MyText, 17, 1))
                
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Changement de nom de joueur
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayername" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid$(MyText, 17, 1))
                
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerName(i, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Respawn request
            If Mid$(MyText, 1, 8) = "/respawn" Then
            
                MyText = Mid$(MyText, 10, Len(MyText) - 9)
                Call SendMapRespawn(Val(MyText))
                MyText = vbNullString
                Exit Sub
            End If
        End If
                
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid$(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid$(MyText, 12, 1))
                
                MyText = Mid$(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(i).command) = Trim$(MyText) And Trim$(Emoticons(i).command) <> "/" Then
                    Call SendData("checkemoticons" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    MyText = vbNullString
                Exit Sub
                End If
            Next i
            Call SendData("checkcommands" & SEP_CHAR & MyText & SEP_CHAR & END_CHAR)
            MyText = vbNullString
        Exit Sub
        End If
            
        ' Say message
        If Len(Trim$(MyText)) > 0 Then
            '//Début du code de canaux
'            If frmMirage.Canal.ListIndex = 1 Then
                'Call BroadcastMsg(MyText)
'                MyText = vbNullString
'                Exit Sub
            If frmMirage.Canal.ListIndex = 2 Then
                Call GuildeMsg(MyText)
                MyText = vbNullString
                Exit Sub
            ElseIf frmMirage.Canal.ListIndex = 3 Then
                name = vbNullString
                   
                For i = 1 To Len(MyText)
                    If Mid$(MyText, i, 1) <> " " Then name = name & Mid$(MyText, i, 1) Else Exit For
                Next i
                   
                If Len(MyText) - i > 0 Then
                    MyText = Mid$(MyText, i + 1, Len(MyText) - i)
                   
                    Call PlayerMsg(MyText, name)
                Else
                    Call AddText("Vous avez oublié le nom du joueur", AlertColor)
                End If
                    MyText = vbNullString
                    Exit Sub
            ElseIf frmMirage.Canal.ListIndex = 0 Then
                Call SayMsg(MyText)
            Else
                Call SayMsg(MyText)
            End If
            
            If (KeyAscii = vbKeyReturn) And frmMirage.txtMyTextBox.Locked = False Then
                frmMirage.txtMyTextBox.Locked = True
                frmMirage.txtMyTextBox.Visible = False
                frmMirage.Canal.Visible = False
            End If
        End If
    MyText = vbNullString
    Exit Sub
    End If
End Sub

Sub CheckMapGetItem()
    Dim Packet As clsBuffer
    Dim MapObjectNum As Integer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Not frmMirage.txtMyTextBox.Visible Then
        MapObjectNum = ObjetNumPos(Player(MyIndex).X, Player(MyIndex).Y)

        If MapObjectNum > -1 Then
            Player(MyIndex).MapGetTimer = GetTickCount
            Set Packet = New clsBuffer
            Packet.WriteLong CMapGetItem
            Packet.WriteByte GetPlayerX(MyIndex)
            Packet.WriteByte GetPlayerY(MyIndex)
            SendData Packet.ToArray()
            Set Packet = Nothing
        End If
    End If
End Sub

Sub CheckAttack()
Dim Packet As clsBuffer
Dim Target() As Integer
    If ControlDown = True And Player(MyIndex).AttackTimer < GetTickCount And Player(MyIndex).Attacking = 0 Then
        Player(MyIndex).Attacking = 1

        If Player(MyIndex).WeaponSlot.num > 0 Then
            If item(Player(MyIndex).WeaponSlot.num).Type = ITEM_TYPE_MISSILE Or item(Player(MyIndex).WeaponSlot.num).Type = ITEM_TYPE_THROWABLE Then
                If item(Player(MyIndex).WeaponSlot.num).Type = ITEM_TYPE_MISSILE Then
                    If GetPlayerInvItemTotalValue(MyIndex, item(Player(MyIndex).WeaponSlot.num).Datas(3)) = 0 Then
                        Exit Sub
                    End If
                End If ' Pas de test pour le throwable car disparaitra des équipements à 0
                Set Packet = New clsBuffer
                Packet.WriteLong CFire

                SendData Packet.ToArray()
                
                Set Packet = Nothing
                Exit Sub
            End If
        End If
        
        Select Case Player(MyIndex).dir
            Case DIR_UP
                Target = FindIndexAtPos(Player(MyIndex).Map, Player(MyIndex).X, Player(MyIndex).Y - 1)
            Case DIR_DOWN
                Target = FindIndexAtPos(Player(MyIndex).Map, Player(MyIndex).X, Player(MyIndex).Y + 1)
            Case DIR_LEFT
                Target = FindIndexAtPos(Player(MyIndex).Map, Player(MyIndex).X - 1, Player(MyIndex).Y)
            Case DIR_RIGHT
                Target = FindIndexAtPos(Player(MyIndex).Map, Player(MyIndex).X + 1, Player(MyIndex).Y)
        End Select

        If Target(1) >= 0 Then
            ' Cible trouvé, attaque
            Set Packet = New clsBuffer
            Packet.WriteLong CAttack
        
            Packet.WriteByte Target(1) 'Target Type
            Packet.WriteInteger Target(0) 'Target index
            
            SendData Packet.ToArray()
            
            Set Packet = Nothing
        End If
    End If
End Sub

Public Function GetNbDirections()
    If IsEmptyArray(ArrPtr(directionOrder)) Then
        GetNbDirections = 0
    Else
        GetNbDirections = UBound(directionOrder) + 1
    End If
End Function

Public Function RemoveDirection(ByVal direction As Byte)
    Dim i As Byte
    
    If GetNbDirections() > 1 Then
        Do While directionOrder(i) <> direction
            ' Find up index
            i = i + 1
        Loop
    
        Do While i < GetNbDirections() - 1
            directionOrder(i) = directionOrder(i + 1)
            i = i + 1
        Loop
        ReDim Preserve directionOrder(0 To GetNbDirections() - 2)
    Else
        Erase directionOrder
        TryToMove = False
    End If
End Function

Public Sub addDirection(ByVal direction As Byte)
    If Not directionEnable(direction) Then
        ReDim Preserve directionOrder(0 To GetNbDirections()) As Byte
        directionOrder(UBound(directionOrder)) = direction
        
        TryToMove = True
    End If
End Sub

Public Function directionEnable(ByVal direction As Byte) As Boolean
    If (direction = DIR_UP) Then
        directionEnable = DirUpPressed
    ElseIf (direction = DIR_DOWN) Then
        directionEnable = DirDownPressed
    ElseIf (direction = DIR_LEFT) Then
        directionEnable = DirLeftPressed
    ElseIf (direction = DIR_RIGHT) Then
        directionEnable = DirRightPressed
    End If
End Function

Sub CheckInput(ByVal KeyState As Byte, ByVal keyCode As Integer, ByVal Shift As Integer)

    If KeyState = 1 Then
        If Not GettingMap And Not IsDead Then
            If keyCode = UserCommand.item("ramasser") Then If frmMirage.txtQ.Visible Then frmMirage.txtQ.Visible = False Else Call CheckMapGetItem
            If keyCode = UserCommand.item("attaque") Then ControlDown = True
            If keyCode = UserCommand.item("haut") Then
                Call addDirection(DIR_UP)
                DirUpPressed = True
            End If
            If keyCode = UserCommand.item("bas") Then
                Call addDirection(DIR_DOWN)
                DirDownPressed = True
            End If
            If keyCode = UserCommand.item("gauche") Then
                Call addDirection(DIR_LEFT)
                DirLeftPressed = True
            End If
            If keyCode = UserCommand.item("droite") Then
                Call addDirection(DIR_RIGHT)
                DirRightPressed = True
            End If
            If keyCode = UserCommand.item("courir") Then ShiftDown = True
        End If
    Else
            If keyCode = UserCommand.item("haut") Then
                Call RemoveDirection(DIR_UP)
                DirUpPressed = False
            End If
            If keyCode = UserCommand.item("bas") Then
                Call RemoveDirection(DIR_DOWN)
                DirDownPressed = False
            End If
            If keyCode = UserCommand.item("gauche") Then
                Call RemoveDirection(DIR_LEFT)
                DirLeftPressed = False
            End If
            If keyCode = UserCommand.item("droite") Then
                Call RemoveDirection(DIR_RIGHT)
                DirRightPressed = False
            End If
            If keyCode = UserCommand.item("courir") Then ShiftDown = False
            If keyCode = UserCommand.item("attaque") Then ControlDown = False
    End If
End Sub

Public Sub InitMirageVars()
    PicScWidth = frmMirage.picScreen.Width
    PicScHeight = frmMirage.picScreen.Height
End Sub

Sub CaseChange(ByVal cx, ByVal cy)
Dim ONum As Long

If Val(ReadINI("CONFIG", "NomObjet", ClientConfigurationFile)) = 0 Then frmMirage.ObjNm.Visible = False: Exit Sub

ONum = ObjetNumPos(cx, cy)

If ONum > -1 Then
    frmMirage.OName.Caption = Trim$(item(ONum).name) & "(" & ObjetValPos(cx, cy) & ")"
    frmMirage.OName.ForeColor = item(ONum).NCoul
    frmMirage.ObjNm.Left = PotX + 10
    frmMirage.ObjNm.Top = PotY - 30
    frmMirage.ObjNm.Width = frmMirage.OName.Width / Screen.TwipsPerPixelY + 240 / Screen.TwipsPerPixelY
    frmMirage.OName.Left = 120
    frmMirage.ObjNm.Visible = True
Else
    frmMirage.ObjNm.Visible = False
End If

End Sub

Function CanMove() As Boolean
Dim i, d As Long ' I doit être variant
Dim X As Long, Y As Long
Dim PX As Long, PY As Long
Dim Dire As Long
Dim Packet As clsBuffer

    CanMove = True

    d = GetPlayerDir(MyIndex)
    PX = 0
    PY = 0
    If GetNbDirections() > 0 Then
        Call SetPlayerDir(MyIndex, directionOrder(GetNbDirections() - 1))
        Dire = directionOrder(GetNbDirections() - 1)
        
        ' Set the new direction if they weren't facing that direction
        If d <> Dire Then Call SendPlayerDir

        If ControlDown Then CanMove = False: Exit Function

        If Dire = DIR_UP Then
            If GetPlayerY(MyIndex) > 0 Then
                PX = 0
                PY = -1
            Else
                ' Check if they can warp to a new map
                If CheckIsOnBorders Then
                    CanMove = True
                    Exit Function
                Else
                    CanMove = False
                    Exit Function
                End If
            End If
        ElseIf Dire = DIR_DOWN Then
            If GetPlayerY(MyIndex) < MaxMapY Then
                PX = 0
                PY = 1
            Else
                ' Check if they can warp to a new map
                If CheckIsOnBorders Then
                    CanMove = True
                    Exit Function
                Else
                    CanMove = False
                    Exit Function
                End If
            End If
        ElseIf Dire = DIR_LEFT Then
            If GetPlayerX(MyIndex) > 0 Then
                PX = -1
                PY = 0
            Else
                ' Check if they can warp to a new map
                If CheckIsOnBorders Then
                    CanMove = True
                    Exit Function
                Else
                    CanMove = False
                    Exit Function
                End If
            End If
        ElseIf Dire = DIR_RIGHT Then
            If GetPlayerX(MyIndex) < MaxMapX Then
                PX = 1
                PY = 0
            Else
                ' Check if they can warp to a new map
                If CheckIsOnBorders Then
                    CanMove = True
                    Exit Function
                Else
                    CanMove = False
                    Exit Function
                End If
            End If
        End If
    End If
    If PX = 0 And PY = 0 Then CanMove = False: Exit Function
        ' Check to see if the map tile is blocked or not
            If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCKED Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_SIGN Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_NIVEAUX Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_MONTURE Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_GUILDE Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_TOIT Then
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_MONTURE Then
                    If Player(MyIndex).ArmorSlot.num > 0 Then
                        If item(Player(MyIndex).ArmorSlot.num).Type = ITEM_TYPE_MONTURE Then CanMove = False Else CanMove = True
                    Else
                        CanMove = True
                    End If
                    
                ElseIf Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_NIVEAUX Then
                    If Player(MyIndex).level < Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(0) Then CanMove = False Else CanMove = True
                ElseIf Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_GUILDE Then
                    If Trim$(Player(MyIndex).Guild) = Trim$(Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Strings(0)) Then CanMove = True Else CanMove = False
                Else
                    CanMove = False
                End If
            End If
            
            If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_CBLOCK Then
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(0) = Player(MyIndex).Class Then Exit Function
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(1) = Player(MyIndex).Class Then Exit Function
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(2) = Player(MyIndex).Class Then Exit Function
                CanMove = False
            End If
            
            If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_BLOCK_DIR Then
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(0) = Player(MyIndex).dir Then CanMove = True: Exit Function
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(1) = Player(MyIndex).dir Then CanMove = True: Exit Function
                If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Datas(2) = Player(MyIndex).dir Then CanMove = True: Exit Function
                CanMove = False
                
            End If
        ' verif atribut toit
        Call SuprTileToit(PY, PX)
                                                    
            ' Check to see if the key door is open or not
            If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_KEY Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_DOOR Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_COFFRE Or Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_PORTE_CODE Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                    CanMove = False
                Else
                    If Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).Type = TILE_TYPE_COFFRE Then CanMove = False
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(i) = GetPlayerX(MyIndex) + PX) And (GetPlayerY(i) = GetPlayerY(MyIndex) + PY) Then
                            CanMove = False
                            Exit Function
                        End If
                    End If
                    
                    If Pets(i).Map = GetPlayerMap(MyIndex) And Pets(i).Map <> -1 Then
                        If (Pets(i).X = GetPlayerX(MyIndex) + PX) And (Pets(i).Y = GetPlayerY(MyIndex) + PY) Then
                            CanMove = False
                            Exit Function
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For Each i In MapNpc.Keys
                If MapNpc(i).num > -1 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex) + PX) And (MapNpc(i).Y = GetPlayerY(MyIndex) + PY) And Npc(MapNpc(i).num).Vol = 0 Then
                        CanMove = False
                        Exit Function
                    End If
                End If
            Next i
End Function

Sub SuprTileToit(ByVal dy As Long, ByVal dX As Long)
' verif atribut toit
On Error Resume Next
With Map
    If .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type <> TILE_TYPE_WALKABLE And .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type <> TILE_TYPE_WALKABLE And .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
    If .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Fringe > 0 Or .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).Fringe2 > 0 Or .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).F2Anim > 0 Or .tile(GetPlayerX(MyIndex) + dX, GetPlayerY(MyIndex) + dy).FAnim > 0 Then
        Dim MX As Long
        Dim MY As Long
        Dim er As Long
        Dim i As Long
    
        
        If InToit = False Then
        
        For er = Player(MyIndex).Y To MaxMapY
        If er < MaxMapY Then
        If .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_TOIT Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
            For i = Player(MyIndex).X To MaxMapX
            If i < MaxMapX Then
                If .tile(i + dX, er + dy).Type = TILE_TYPE_TOIT Or .tile(i + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(i + dX, er + dy).Fringe = 0
                    .tile(i + dX, er + dy).Fringe2 = 0
                    .tile(i + dX, er + dy).Fringe3 = 0
                    .tile(i + dX, er + dy).FAnim = 0
                    .tile(i + dX, er + dy).F2Anim = 0
                    .tile(i + dX, er + dy).F3Anim = 0
                Else
                If .tile(i + dX, er + dy).Type = TILE_TYPE_DOOR Or .tile(i + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(i + dX, er + dy).Type = TILE_TYPE_WARP Or .tile(i + dX, er + dy).Type = TILE_TYPE_KEY Then
                    .tile(i + dX, er + dy).Fringe = 0
                    .tile(i + dX, er + dy).Fringe2 = 0
                    .tile(i + dX, er + dy).Fringe3 = 0
                    .tile(i + dX, er + dy).FAnim = 0
                    .tile(i + dX, er + dy).F2Anim = 0
                    .tile(i + dX, er + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(i + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(i + dX, er + dy).Fringe = 0
                    .tile(i + dX, er + dy).Fringe2 = 0
                    .tile(i + dX, er + dy).Fringe3 = 0
                    .tile(i + dX, er + dy).FAnim = 0
                    .tile(i + dX, er + dy).F2Anim = 0
                    .tile(i + dX, er + dy).F3Anim = 0
                End If
            Else
                If .tile(i, er + dy).Type = TILE_TYPE_TOIT Or .tile(i, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(i, er + dy).Fringe = 0
                    .tile(i, er + dy).Fringe2 = 0
                    .tile(i, er + dy).Fringe3 = 0
                    .tile(i, er + dy).FAnim = 0
                    .tile(i, er + dy).F2Anim = 0
                    .tile(i, er + dy).F3Anim = 0
                Else
                If .tile(i, er + dy).Type = TILE_TYPE_DOOR Or .tile(i, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(i, er + dy).Type = TILE_TYPE_WARP Or .tile(i, er + dy).Type = TILE_TYPE_KEY Then
                    .tile(i, er + dy).Fringe = 0
                    .tile(i, er + dy).Fringe2 = 0
                    .tile(i, er + dy).Fringe3 = 0
                    .tile(i, er + dy).FAnim = 0
                    .tile(i, er + dy).F2Anim = 0
                    .tile(i, er + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(i, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(i, er + dy).Fringe = 0
                    .tile(i, er + dy).Fringe2 = 0
                    .tile(i, er + dy).Fringe3 = 0
                    .tile(i, er + dy).FAnim = 0
                    .tile(i, er + dy).F2Anim = 0
                    .tile(i, er + dy).F3Anim = 0
                End If
            End If
            Next i
                MX = Player(MyIndex).X
            For i = 0 To Player(MyIndex).X
                If .tile(MX + dX, er + dy).Type = TILE_TYPE_TOIT Or .tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(MX + dX, er + dy).Fringe = 0
                    .tile(MX + dX, er + dy).Fringe2 = 0
                    .tile(MX + dX, er + dy).Fringe3 = 0
                    .tile(MX + dX, er + dy).FAnim = 0
                    .tile(MX + dX, er + dy).F2Anim = 0
                    .tile(MX + dX, er + dy).F3Anim = 0
                Else
                If .tile(MX + dX, er + dy).Type = TILE_TYPE_DOOR Or .tile(MX + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(MX + dX, er + dy).Type = TILE_TYPE_WARP Or .tile(MX + dX, er + dy).Type = TILE_TYPE_KEY Then
                    .tile(MX + dX, er + dy).Fringe = 0
                    .tile(MX + dX, er + dy).Fringe2 = 0
                    .tile(MX + dX, er + dy).Fringe3 = 0
                    .tile(MX + dX, er + dy).FAnim = 0
                    .tile(MX + dX, er + dy).F2Anim = 0
                    .tile(MX + dX, er + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For
                    .tile(MX + dX, er + dy).Fringe = 0
                    .tile(MX + dX, er + dy).Fringe2 = 0
                    .tile(MX + dX, er + dy).Fringe3 = 0
                    .tile(MX + dX, er + dy).FAnim = 0
                    .tile(MX + dX, er + dy).F2Anim = 0
                    .tile(MX + dX, er + dy).F3Anim = 0
                End If
                MX = MX - 1
            Next i
        Else
        If .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_DOOR Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_WARP Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_KEY Then
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
                Exit For
            End If
            If .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
            .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
        End If
        Else
        If .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_TOIT Or .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_BLOCK_TOIT Then
            For i = Player(MyIndex).X To MaxMapX
            If i < MaxMapX Then
                If .tile(i + dX, er).Type = TILE_TYPE_TOIT Or .tile(i + dX, er).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(i + dX, er).Fringe = 0
                    .tile(i + dX, er).Fringe2 = 0
                    .tile(i + dX, er).Fringe3 = 0
                    .tile(i + dX, er).FAnim = 0
                    .tile(i + dX, er).F2Anim = 0
                    .tile(i + dX, er).F3Anim = 0
                Else
                If .tile(i + dX, er).Type = TILE_TYPE_DOOR Or .tile(i + dX, er).Type = TILE_TYPE_PORTE_CODE Or .tile(i + dX, er).Type = TILE_TYPE_WARP Or .tile(i + dX, er).Type = TILE_TYPE_KEY Then
                    .tile(i + dX, er).Fringe = 0
                    .tile(i + dX, er).Fringe2 = 0
                    .tile(i + dX, er).Fringe3 = 0
                    .tile(i + dX, er).FAnim = 0
                    .tile(i + dX, er).F2Anim = 0
                    .tile(i + dX, er).F3Anim = 0
                    Exit For
                End If
                    If .tile(i + dX, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(i + dX, er).Fringe = 0
                    .tile(i + dX, er).Fringe2 = 0
                    .tile(i + dX, er).Fringe3 = 0
                    .tile(i + dX, er).FAnim = 0
                    .tile(i + dX, er).F2Anim = 0
                    .tile(i + dX, er).F3Anim = 0
                End If
            Else
                If .tile(i, er).Type = TILE_TYPE_TOIT Or .tile(i, er).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(i, er).Fringe = 0
                    .tile(i, er).Fringe2 = 0
                    .tile(i, er).Fringe3 = 0
                    .tile(i, er).FAnim = 0
                    .tile(i, er).F2Anim = 0
                    .tile(i, er).F3Anim = 0
                Else
                If .tile(i, er).Type = TILE_TYPE_DOOR Or .tile(i, er).Type = TILE_TYPE_PORTE_CODE Or .tile(i, er).Type = TILE_TYPE_WARP Or .tile(i, er).Type = TILE_TYPE_KEY Then
                    .tile(i, er).Fringe = 0
                    .tile(i, er).Fringe2 = 0
                    .tile(i, er).Fringe3 = 0
                    .tile(i, er).FAnim = 0
                    .tile(i, er).F2Anim = 0
                    .tile(i, er).F3Anim = 0
                    Exit For
                End If
                    If .tile(i, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(i, er).Fringe = 0
                    .tile(i, er).Fringe2 = 0
                    .tile(i, er).Fringe3 = 0
                    .tile(i, er).FAnim = 0
                    .tile(i, er).F2Anim = 0
                    .tile(i, er).F3Anim = 0
                End If
            End If
            Next i
                MX = Player(MyIndex).X
            For i = 0 To Player(MyIndex).X
                If .tile(MX + dX, er).Type = TILE_TYPE_TOIT Or .tile(MX + dX, er).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(MX + dX, er).Fringe = 0
                    .tile(MX + dX, er).Fringe2 = 0
                    .tile(MX + dX, er).Fringe3 = 0
                    .tile(MX + dX, er).FAnim = 0
                    .tile(MX + dX, er).F2Anim = 0
                    .tile(MX + dX, er).F3Anim = 0
                Else
                If .tile(MX + dX, er).Type = TILE_TYPE_DOOR Or .tile(MX + dX, er).Type = TILE_TYPE_PORTE_CODE Or .tile(MX + dX, er).Type = TILE_TYPE_WARP Or .tile(MX + dX, er).Type = TILE_TYPE_KEY Then
                    .tile(MX + dX, er).Fringe = 0
                    .tile(MX + dX, er).Fringe2 = 0
                    .tile(MX + dX, er).Fringe3 = 0
                    .tile(MX + dX, er).FAnim = 0
                    .tile(MX + dX, er).F2Anim = 0
                    .tile(MX + dX, er).F3Anim = 0
                    Exit For
                End If
                    If .tile(MX + dX, er).Type = TILE_TYPE_BLOCKED Then Exit For
                    .tile(MX + dX, er).Fringe = 0
                    .tile(MX + dX, er).Fringe2 = 0
                    .tile(MX + dX, er).Fringe3 = 0
                    .tile(MX + dX, er).FAnim = 0
                    .tile(MX + dX, er).F2Anim = 0
                    .tile(MX + dX, er).F3Anim = 0
                End If
                MX = MX - 1
            Next i
        Else
        If .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_DOOR Or .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_PORTE_CODE Or .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_WARP Or .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_KEY Then
                .tile(GetPlayerX(MyIndex) + dX, er).Fringe = 0
                .tile(GetPlayerX(MyIndex) + dX, er).Fringe2 = 0
                .tile(GetPlayerX(MyIndex) + dX, er).Fringe3 = 0
                .tile(GetPlayerX(MyIndex) + dX, er).FAnim = 0
                .tile(GetPlayerX(MyIndex) + dX, er).F2Anim = 0
                .tile(GetPlayerX(MyIndex) + dX, er).F3Anim = 0
                Exit For
            End If
            If .tile(GetPlayerX(MyIndex) + dX, er).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
            .tile(GetPlayerX(MyIndex) + dX, er).Fringe = 0
            .tile(GetPlayerX(MyIndex) + dX, er).Fringe2 = 0
            .tile(GetPlayerX(MyIndex) + dX, er).Fringe3 = 0
            .tile(GetPlayerX(MyIndex) + dX, er).FAnim = 0
            .tile(GetPlayerX(MyIndex) + dX, er).F2Anim = 0
            .tile(GetPlayerX(MyIndex) + dX, er).F3Anim = 0
        End If
        End If
        Next er
        
        er = Player(MyIndex).Y
        For MY = 0 To Player(MyIndex).Y
        If .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_TOIT Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
            For i = Player(MyIndex).X To MaxMapX
            If i < MaxMapX Then
                If .tile(i + dX, er + dy).Type = TILE_TYPE_TOIT Or .tile(i + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(i + dX, er + dy).Fringe = 0
                    .tile(i + dX, er + dy).Fringe2 = 0
                    .tile(i + dX, er + dy).Fringe3 = 0
                    .tile(i + dX, er + dy).FAnim = 0
                    .tile(i + dX, er + dy).F2Anim = 0
                    .tile(i + dX, er + dy).F3Anim = 0
                Else
                If .tile(i + dX, er + dy).Type = TILE_TYPE_DOOR Or .tile(i + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(i + dX, er + dy).Type = TILE_TYPE_WARP Or .tile(i + dX, er + dy).Type = TILE_TYPE_KEY Then
                    .tile(i + dX, er + dy).Fringe = 0
                    .tile(i + dX, er + dy).Fringe2 = 0
                    .tile(i + dX, er + dy).Fringe3 = 0
                    .tile(i + dX, er + dy).FAnim = 0
                    .tile(i + dX, er + dy).F2Anim = 0
                    .tile(i + dX, er + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(i + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(i + dX, er + dy).Fringe = 0
                    .tile(i + dX, er + dy).Fringe2 = 0
                    .tile(i + dX, er + dy).Fringe3 = 0
                    .tile(i + dX, er + dy).FAnim = 0
                    .tile(i + dX, er + dy).F2Anim = 0
                    .tile(i + dX, er + dy).F3Anim = 0
                End If
                Else
                If .tile(i, er + dy).Type = TILE_TYPE_TOIT Or .tile(i, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(i, er + dy).Fringe = 0
                    .tile(i, er + dy).Fringe2 = 0
                    .tile(i, er + dy).Fringe3 = 0
                    .tile(i, er + dy).FAnim = 0
                    .tile(i, er + dy).F2Anim = 0
                    .tile(i, er + dy).F3Anim = 0
                Else
                If .tile(i, er + dy).Type = TILE_TYPE_DOOR Or .tile(i, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(i, er + dy).Type = TILE_TYPE_WARP Or .tile(i, er + dy).Type = TILE_TYPE_KEY Then
                    .tile(i, er + dy).Fringe = 0
                    .tile(i, er + dy).Fringe2 = 0
                    .tile(i, er + dy).Fringe3 = 0
                    .tile(i, er + dy).FAnim = 0
                    .tile(i, er + dy).F2Anim = 0
                    .tile(i, er + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(i, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(i, er + dy).Fringe = 0
                    .tile(i, er + dy).Fringe2 = 0
                    .tile(i, er + dy).Fringe3 = 0
                    .tile(i, er + dy).FAnim = 0
                    .tile(i, er + dy).F2Anim = 0
                    .tile(i, er + dy).F3Anim = 0
                End If
            End If
            Next i
                MX = Player(MyIndex).X
            For i = 0 To Player(MyIndex).X
                If .tile(MX + dX, er + dy).Type = TILE_TYPE_TOIT Or .tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(MX + dX, er + dy).Fringe = 0
                    .tile(MX + dX, er + dy).Fringe2 = 0
                    .tile(MX + dX, er + dy).Fringe3 = 0
                    .tile(MX + dX, er + dy).FAnim = 0
                    .tile(MX + dX, er + dy).F2Anim = 0
                    .tile(MX + dX, er + dy).F3Anim = 0
                Else
                If .tile(MX + dX, er + dy).Type = TILE_TYPE_DOOR Or .tile(MX + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(MX + dX, er + dy).Type = TILE_TYPE_WARP Or .tile(MX + dX, er + dy).Type = TILE_TYPE_KEY Then
                    .tile(MX + dX, er + dy).Fringe = 0
                    .tile(MX + dX, er + dy).Fringe2 = 0
                    .tile(MX + dX, er + dy).Fringe3 = 0
                    .tile(MX + dX, er + dy).FAnim = 0
                    .tile(MX + dX, er + dy).F2Anim = 0
                    .tile(MX + dX, er + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(MX + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(MX + dX, er + dy).Fringe = 0
                    .tile(MX + dX, er + dy).Fringe2 = 0
                    .tile(MX + dX, er + dy).Fringe3 = 0
                    .tile(MX + dX, er + dy).FAnim = 0
                    .tile(MX + dX, er + dy).F2Anim = 0
                    .tile(MX + dX, er + dy).F3Anim = 0
                End If
                MX = MX - 1
            Next i
        Else
            If .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_DOOR Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_WARP Or .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_KEY Then
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
                .tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
                Exit For
            End If
            If .tile(GetPlayerX(MyIndex) + dX, er + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
            .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe = 0
            .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe2 = 0
            .tile(GetPlayerX(MyIndex) + dX, er + dy).Fringe3 = 0
            .tile(GetPlayerX(MyIndex) + dX, er + dy).FAnim = 0
            .tile(GetPlayerX(MyIndex) + dX, er + dy).F2Anim = 0
            .tile(GetPlayerX(MyIndex) + dX, er + dy).F3Anim = 0
        End If
        er = er - 1
        Next MY
        
        For er = Player(MyIndex).X To MaxMapX
        If er < MaxMapX Then
        If .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
            For i = Player(MyIndex).Y To MaxMapY
            If i < MaxMapY Then
                If .tile(er + dX, i + dy).Type = TILE_TYPE_TOIT Or .tile(er + dX, i + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er + dX, i + dy).Fringe = 0
                    .tile(er + dX, i + dy).Fringe2 = 0
                    .tile(er + dX, i + dy).Fringe3 = 0
                    .tile(er + dX, i + dy).FAnim = 0
                    .tile(er + dX, i + dy).F2Anim = 0
                    .tile(er + dX, i + dy).F3Anim = 0
                Else
                If .tile(er + dX, i + dy).Type = TILE_TYPE_DOOR Or .tile(er + dX, i + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, i + dy).Type = TILE_TYPE_WARP Or .tile(er + dX, i + dy).Type = TILE_TYPE_KEY Then
                    .tile(er + dX, i + dy).Fringe = 0
                    .tile(er + dX, i + dy).Fringe2 = 0
                    .tile(er + dX, i + dy).Fringe3 = 0
                    .tile(er + dX, i + dy).FAnim = 0
                    .tile(er + dX, i + dy).F2Anim = 0
                    .tile(er + dX, i + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(er + dX, i + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er + dX, i + dy).Fringe = 0
                    .tile(er + dX, i + dy).Fringe2 = 0
                    .tile(er + dX, i + dy).Fringe3 = 0
                    .tile(er + dX, i + dy).FAnim = 0
                    .tile(er + dX, i + dy).F2Anim = 0
                    .tile(er + dX, i + dy).F3Anim = 0
                End If
            Else
            If .tile(er + dX, i).Type = TILE_TYPE_TOIT Or .tile(er + dX, i).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er + dX, i).Fringe = 0
                    .tile(er + dX, i).Fringe2 = 0
                    .tile(er + dX, i).Fringe3 = 0
                    .tile(er + dX, i).FAnim = 0
                    .tile(er + dX, i).F2Anim = 0
                    .tile(er + dX, i).F3Anim = 0
                Else
                If .tile(er + dX, i).Type = TILE_TYPE_DOOR Or .tile(er + dX, i).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, i).Type = TILE_TYPE_WARP Or .tile(er + dX, i).Type = TILE_TYPE_KEY Then
                    .tile(er + dX, i).Fringe = 0
                    .tile(er + dX, i).Fringe2 = 0
                    .tile(er + dX, i).Fringe3 = 0
                    .tile(er + dX, i).FAnim = 0
                    .tile(er + dX, i).F2Anim = 0
                    .tile(er + dX, i).F3Anim = 0
                    Exit For
                End If
                    If .tile(er + dX, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er + dX, i).Fringe = 0
                    .tile(er + dX, i).Fringe2 = 0
                    .tile(er + dX, i).Fringe3 = 0
                    .tile(er + dX, i).FAnim = 0
                    .tile(er + dX, i).F2Anim = 0
                    .tile(er + dX, i).F3Anim = 0
                End If
            End If
            Next i
                MY = Player(MyIndex).Y
            For i = 0 To Player(MyIndex).Y
                If .tile(er + dX, MY + dy).Type = TILE_TYPE_TOIT Or .tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er + dX, MY + dy).Fringe = 0
                    .tile(er + dX, MY + dy).Fringe2 = 0
                    .tile(er + dX, MY + dy).Fringe3 = 0
                    .tile(er + dX, MY + dy).FAnim = 0
                    .tile(er + dX, MY + dy).F2Anim = 0
                    .tile(er + dX, MY + dy).F3Anim = 0
                Else
                If .tile(er + dX, MY + dy).Type = TILE_TYPE_DOOR Or .tile(er + dX, MY + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, MY + dy).Type = TILE_TYPE_WARP Or .tile(er + dX, MY + dy).Type = TILE_TYPE_KEY Then
                    .tile(er + dX, MY + dy).Fringe = 0
                    .tile(er + dX, MY + dy).Fringe2 = 0
                    .tile(er + dX, MY + dy).Fringe3 = 0
                    .tile(er + dX, MY + dy).FAnim = 0
                    .tile(er + dX, MY + dy).F2Anim = 0
                    .tile(er + dX, MY + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er + dX, MY + dy).Fringe = 0
                    .tile(er + dX, MY + dy).Fringe2 = 0
                    .tile(er + dX, MY + dy).Fringe3 = 0
                    .tile(er + dX, MY + dy).FAnim = 0
                    .tile(er + dX, MY + dy).F2Anim = 0
                    .tile(er + dX, MY + dy).F3Anim = 0
                End If
                MY = MY - 1
            Next i
        Else
            If .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_DOOR Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_WARP Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_KEY Then
                .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
                Exit For
            End If
            If .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
            .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
        End If
        Else
        If .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
            For i = Player(MyIndex).Y To MaxMapY
            If i < MaxMapY Then
                If .tile(er, i + dy).Type = TILE_TYPE_TOIT Or .tile(er, i + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er, i + dy).Fringe = 0
                    .tile(er, i + dy).Fringe2 = 0
                    .tile(er, i + dy).Fringe3 = 0
                    .tile(er, i + dy).FAnim = 0
                    .tile(er, i + dy).F2Anim = 0
                    .tile(er, i + dy).F3Anim = 0
                Else
                If .tile(er, i + dy).Type = TILE_TYPE_DOOR Or .tile(er, i + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er, i + dy).Type = TILE_TYPE_WARP Or .tile(er, i + dy).Type = TILE_TYPE_KEY Then
                    .tile(er, i + dy).Fringe = 0
                    .tile(er, i + dy).Fringe2 = 0
                    .tile(er, i + dy).Fringe3 = 0
                    .tile(er, i + dy).FAnim = 0
                    .tile(er, i + dy).F2Anim = 0
                    .tile(er, i + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(er, i + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er, i + dy).Fringe = 0
                    .tile(er, i + dy).Fringe2 = 0
                    .tile(er, i + dy).Fringe3 = 0
                    .tile(er, i + dy).FAnim = 0
                    .tile(er, i + dy).F2Anim = 0
                    .tile(er, i + dy).F3Anim = 0
                End If
            Else
            If .tile(er, i).Type = TILE_TYPE_TOIT Or .tile(er, i).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er, i).Fringe = 0
                    .tile(er, i).Fringe2 = 0
                    .tile(er, i).Fringe3 = 0
                    .tile(er, i).FAnim = 0
                    .tile(er, i).F2Anim = 0
                    .tile(er, i).F3Anim = 0
                Else
                If .tile(er, i).Type = TILE_TYPE_DOOR Or .tile(er, i).Type = TILE_TYPE_PORTE_CODE Or .tile(er, i).Type = TILE_TYPE_WARP Or .tile(er, i).Type = TILE_TYPE_KEY Then
                    .tile(er, i).Fringe = 0
                    .tile(er, i).Fringe2 = 0
                    .tile(er, i).Fringe3 = 0
                    .tile(er, i).FAnim = 0
                    .tile(er, i).F2Anim = 0
                    .tile(er, i).F3Anim = 0
                    Exit For
                End If
                    If .tile(er, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er, i).Fringe = 0
                    .tile(er, i).Fringe2 = 0
                    .tile(er, i).Fringe3 = 0
                    .tile(er, i).FAnim = 0
                    .tile(er, i).F2Anim = 0
                    .tile(er, i).F3Anim = 0
                End If
            End If
            Next i
                MY = Player(MyIndex).Y
            For i = 0 To Player(MyIndex).Y
                If .tile(er, MY + dy).Type = TILE_TYPE_TOIT Or .tile(er, MY + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er, MY + dy).Fringe = 0
                    .tile(er, MY + dy).Fringe2 = 0
                    .tile(er, MY + dy).Fringe3 = 0
                    .tile(er, MY + dy).FAnim = 0
                    .tile(er, MY + dy).F2Anim = 0
                    .tile(er, MY + dy).F3Anim = 0
                Else
                If .tile(er, MY + dy).Type = TILE_TYPE_DOOR Or .tile(er, MY + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er, MY + dy).Type = TILE_TYPE_WARP Or .tile(er, MY + dy).Type = TILE_TYPE_KEY Then
                    .tile(er, MY + dy).Fringe = 0
                    .tile(er, MY + dy).Fringe2 = 0
                    .tile(er, MY + dy).Fringe3 = 0
                    .tile(er, MY + dy).FAnim = 0
                    .tile(er, MY + dy).F2Anim = 0
                    .tile(er, MY + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(er, MY + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er, MY + dy).Fringe = 0
                    .tile(er, MY + dy).Fringe2 = 0
                    .tile(er, MY + dy).Fringe3 = 0
                    .tile(er, MY + dy).FAnim = 0
                    .tile(er, MY + dy).F2Anim = 0
                    .tile(er, MY + dy).F3Anim = 0
                End If
                MY = MY - 1
            Next i
        Else
            If .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_DOOR Or .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_WARP Or .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_KEY Then
                .tile(er, GetPlayerY(MyIndex) + dy).Fringe = 0
                .tile(er, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                .tile(er, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                .tile(er, GetPlayerY(MyIndex) + dy).FAnim = 0
                .tile(er, GetPlayerY(MyIndex) + dy).F2Anim = 0
                .tile(er, GetPlayerY(MyIndex) + dy).F3Anim = 0
                Exit For
            End If
            If .tile(er, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
            .tile(er, GetPlayerY(MyIndex) + dy).Fringe = 0
            .tile(er, GetPlayerY(MyIndex) + dy).Fringe2 = 0
            .tile(er, GetPlayerY(MyIndex) + dy).Fringe3 = 0
            .tile(er, GetPlayerY(MyIndex) + dy).FAnim = 0
            .tile(er, GetPlayerY(MyIndex) + dy).F2Anim = 0
            .tile(er, GetPlayerY(MyIndex) + dy).F3Anim = 0
        End If
        End If
        Next er
        
        er = Player(MyIndex).X
        For MX = 0 To Player(MyIndex).X
        If .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_TOIT Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCK_TOIT Then
            For i = Player(MyIndex).Y To MaxMapY
            If i < MaxMapY Then
                If .tile(er + dX, i + dy).Type = TILE_TYPE_TOIT Or .tile(er + dX, i + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er + dX, i + dy).Fringe = 0
                    .tile(er + dX, i + dy).Fringe2 = 0
                    .tile(er + dX, i + dy).Fringe3 = 0
                    .tile(er + dX, i + dy).FAnim = 0
                    .tile(er + dX, i + dy).F2Anim = 0
                    .tile(er + dX, i + dy).F3Anim = 0
                Else
                If .tile(er + dX, i + dy).Type = TILE_TYPE_DOOR Or .tile(er + dX, i + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, i + dy).Type = TILE_TYPE_WARP Or .tile(er + dX, i + dy).Type = TILE_TYPE_KEY Then
                    .tile(er + dX, i + dy).Fringe = 0
                    .tile(er + dX, i + dy).Fringe2 = 0
                    .tile(er + dX, i + dy).Fringe3 = 0
                    .tile(er + dX, i + dy).FAnim = 0
                    .tile(er + dX, i + dy).F2Anim = 0
                    .tile(er + dX, i + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(er + dX, i + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er + dX, i + dy).Fringe = 0
                    .tile(er + dX, i + dy).Fringe2 = 0
                    .tile(er + dX, i + dy).Fringe3 = 0
                    .tile(er + dX, i + dy).FAnim = 0
                    .tile(er + dX, i + dy).F2Anim = 0
                    .tile(er + dX, i + dy).F3Anim = 0
                End If
            Else
                If .tile(er + dX, i).Type = TILE_TYPE_TOIT Or .tile(er + dX, i).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er + dX, i).Fringe = 0
                    .tile(er + dX, i).Fringe2 = 0
                    .tile(er + dX, i).Fringe3 = 0
                    .tile(er + dX, i).FAnim = 0
                    .tile(er + dX, i).F2Anim = 0
                    .tile(er + dX, i).F3Anim = 0
                Else
                If .tile(er + dX, i).Type = TILE_TYPE_DOOR Or .tile(er + dX, i).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, i).Type = TILE_TYPE_WARP Or .tile(er + dX, i).Type = TILE_TYPE_KEY Then
                    .tile(er + dX, i).Fringe = 0
                    .tile(er + dX, i).Fringe2 = 0
                    .tile(er + dX, i).Fringe3 = 0
                    .tile(er + dX, i).FAnim = 0
                    .tile(er + dX, i).F2Anim = 0
                    .tile(er + dX, i).F3Anim = 0
                    Exit For
                End If
                    If .tile(er + dX, i).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er + dX, i).Fringe = 0
                    .tile(er + dX, i).Fringe2 = 0
                    .tile(er + dX, i).Fringe3 = 0
                    .tile(er + dX, i).FAnim = 0
                    .tile(er + dX, i).F2Anim = 0
                    .tile(er + dX, i).F3Anim = 0
                End If
            End If
            Next i
                MY = Player(MyIndex).Y
            For i = 0 To Player(MyIndex).Y
                If .tile(er + dX, MY + dy).Type = TILE_TYPE_TOIT Or .tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCK_TOIT Then
                    .tile(er + dX, MY + dy).Fringe = 0
                    .tile(er + dX, MY + dy).Fringe2 = 0
                    .tile(er + dX, MY + dy).Fringe3 = 0
                    .tile(er + dX, MY + dy).FAnim = 0
                    .tile(er + dX, MY + dy).F2Anim = 0
                    .tile(er + dX, MY + dy).F3Anim = 0
                Else
                If .tile(er + dX, MY + dy).Type = TILE_TYPE_DOOR Or .tile(er + dX, MY + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, MY + dy).Type = TILE_TYPE_WARP Or .tile(er + dX, MY + dy).Type = TILE_TYPE_KEY Then
                    .tile(er + dX, MY + dy).Fringe = 0
                    .tile(er + dX, MY + dy).Fringe2 = 0
                    .tile(er + dX, MY + dy).Fringe3 = 0
                    .tile(er + dX, MY + dy).FAnim = 0
                    .tile(er + dX, MY + dy).F2Anim = 0
                    .tile(er + dX, MY + dy).F3Anim = 0
                    Exit For
                End If
                    If .tile(er + dX, MY + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
                    .tile(er + dX, MY + dy).Fringe = 0
                    .tile(er + dX, MY + dy).Fringe2 = 0
                    .tile(er + dX, MY + dy).Fringe3 = 0
                    .tile(er + dX, MY + dy).FAnim = 0
                    .tile(er + dX, MY + dy).F2Anim = 0
                    .tile(er + dX, MY + dy).F3Anim = 0
                End If
                MY = MY - 1
            Next i
        Else
            If .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_DOOR Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_PORTE_CODE Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_WARP Or .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_KEY Then
                .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
                .tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
                Exit For
            End If
            If .tile(er + dX, GetPlayerY(MyIndex) + dy).Type = TILE_TYPE_BLOCKED Then Exit For 'avoir
            .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe2 = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).Fringe3 = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).FAnim = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).F2Anim = 0
            .tile(er + dX, GetPlayerY(MyIndex) + dy).F3Anim = 0
        End If
        er = er - 1
        Next MX
        InToit = True
        Else
        If InToit = True Then
        Call LoadMap(GetPlayerMap(MyIndex))
        End If
        InToit = False
        End If
    End If
    Else
        If InToit = True Then
        Call LoadMap(GetPlayerMap(MyIndex))
        End If
        InToit = False
    End If
End With

End Sub

Sub CheckMovement()
    If Player(MyIndex).Moving = 0 Then
        If Not GettingMap And Not IsDead And TryToMove And Player(MyIndex).Attacking = 0 Then
        
            ' On doit tester CanMove après avoir setter la direction car sinon CanMove va setter la direction et envoyer un changement de direction au serveur
            If CanMove Then
                If CheckTeleport Then
                    Exit Sub
                End If
          
                If ShiftDown Then Player(MyIndex).Moving = MOVING_RUNNING Else Player(MyIndex).Moving = MOVING_WALKING
                Call SendPlayerMove
                Call initMoveOffset(MyIndex)
            End If
        End If
    End If
End Sub

Function CheckTeleport() As Boolean
    ' Check for border map
    Dim valueToSend As Byte

    CheckTeleport = False
    If GetPlayerDir(MyIndex) = DIR_UP Then
        If GetPlayerY(MyIndex) <= 0 Then
            ' Check if they can warp to a new map
            If CheckIsOnBorders Then
                CheckTeleport = True
                valueToSend = GetPlayerX(MyIndex)
            End If
        End If
    ElseIf GetPlayerDir(MyIndex) = DIR_DOWN Then
        If GetPlayerY(MyIndex) >= MaxMapY Then
            ' Check if they can warp to a new map
            If CheckIsOnBorders Then
                CheckTeleport = True
                valueToSend = GetPlayerX(MyIndex)
            End If
        End If
    ElseIf GetPlayerDir(MyIndex) = DIR_LEFT Then
        If GetPlayerX(MyIndex) <= 0 Then
            ' Check if they can warp to a new map
            If CheckIsOnBorders Then
                CheckTeleport = True
                valueToSend = GetPlayerY(MyIndex)
            End If
        End If
    ElseIf GetPlayerDir(MyIndex) = DIR_RIGHT Then
        If GetPlayerX(MyIndex) >= MaxMapX Then
            ' Check if they can warp to a new map
            If CheckIsOnBorders Then
                CheckTeleport = True
                valueToSend = GetPlayerY(MyIndex)
            End If
        End If
    End If
    
    If CheckTeleport Then
        GettingMap = True
    
        If Player(MyIndex).Moving Then
            Call ClearPlayerMove(MyIndex)
            Call SendPlayerStopMove
        End If
    
        Dim Packet As clsBuffer
        Set Packet = New clsBuffer

        Packet.WriteLong CGoBorderMap
        
        Packet.WriteByte valueToSend

        SendData Packet.ToArray()
        Set Packet = Nothing
        
        Exit Function
    End If
    'Exit Sub
End Function

Function CheckIsOnBorders()
    Dim i As Integer
    
    CheckIsOnBorders = False
    For i = 0 To GetMapBordersCount() - 1
        If Map.Borders(i).XSource = GetPlayerX(MyIndex) And Map.Borders(i).YSource = GetPlayerY(MyIndex) And Map.Borders(i).DirectionSource = GetPlayerDir(MyIndex) Then
            CheckIsOnBorders = True
            Exit Function
        End If
    Next i
End Function

Function FindIndexAtPos(ByVal mapNum As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer()
    Dim Target(0 To 1) As Integer
    Dim i As Variant ' I doit être variant
    
    Target(0) = 0
    Target(1) = -1
    i = 1
    
    Do While Target(0) = 0 And i <= MAX_PLAYERS
        If Player(i).Map = mapNum And Player(i).X = X And Player(i).Y = Y Then
            Target(0) = i
            'Type de cible
            Target(1) = PLAYER_TYPE
            FindIndexAtPos = Target()
            Exit Function
        End If
        
        If Pets(i).Map = mapNum And Pets(i).X = X And Pets(i).Y = Y Then
            Target(0) = i
            Target(1) = PET_TYPE
            FindIndexAtPos = Target()
            Exit Function
        End If
        i = i + 1
    Loop
    
    For Each i In MapNpc.Keys
        If MapNpc(i).Map = mapNum And MapNpc(i).X = X And MapNpc(i).Y = Y Then
            Target(0) = i
            ' Type de cible
            Target(1) = NPC_TYPE
            FindIndexAtPos = Target()
            Exit Function
        End If
    Next i
    
    FindIndexAtPos = Target()
End Function

Function FindPlayer(ByVal name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(name)))) = UCase$(Trim$(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Function FindOpenInvSlot(ByVal itemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(MyIndex) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then Exit Function
    
    If item(itemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 0 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) = itemNum Then FindOpenInvSlot = i: Exit Function
        Next i
    End If
    
    For i = 0 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(MyIndex, i) <= -1 Then FindOpenInvSlot = i: Exit Function
    Next i
End Function

Public Sub UpdateTradeInventory()
Dim i As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For i = 0 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).name)
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Aucun>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Function ObjetNumPos(ByVal X As Long, ByVal Y As Long) As Long
ObjetNumPos = -1

If NbMapItems(X, Y) > 0 Then
    ObjetNumPos = MapItem(X, Y).items(0).num
End If

End Function

Function ObjetValPos(ByVal X As Long, ByVal Y As Long) As Long
    ObjetValPos = MapItem(X, Y).items(0).Value
End Function

Sub PlayerSearch(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Long, y1 As Long

    x1 = (X \ PIC_X)
    y1 = (Y \ PIC_Y)
    
    If (x1 >= 0) And (x1 <= MaxMapX) And (y1 >= 0) And (y1 <= MaxMapY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub

Sub BltTile2(ByVal X As Long, ByVal Y As Long, ByVal tile As Long)
    rec.Top = (tile \ TilesInSheets) * PIC_Y
    rec.bottom = rec.Top + PIC_Y
    rec.Left = (tile - (tile \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(X - NewPlayerPicX + sx - NewXOffset, Y - NewPlayerPicY + sx - NewYOffset, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerText(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String

    strWords() = Split(Bubble(Index).Text, " ")
    
    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)
        
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y) + 85
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1
            
        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1
            
            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If
            
            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - NewPlayerPicX + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY - NewPlayerPicY + sx - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub
Sub BltPlayerBar(ByVal Index As Integer)
Dim X As Long, Y As Long, ty As Long
    
    If Player(Index).HP <> 0 Then
        ty = (DDSD_Character(GetPlayerSprite(Index)).lHeight / 4) / 2
        X = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - NewPlayerPOffsetX
        Y = (GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset) - NewPlayerPOffsetY + ty + 3
        'draws the back bars
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 2, X + 32, Y - 2)
    
        'draws HP
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 2, X + (Player(Index).HP / Player(Index).MaxHp * 32), Y - 2)
    End If
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim X As Long, Y As Long, ty As Long

If MapNpc(Index).HP = 0 Or MapNpc(Index).MaxHp <= 0 Or MapNpc(Index).num < 1 Then Exit Sub

    ty = (DDSD_Character(Npc(MapNpc(Index).num).sprite).lHeight / 4) / 2
    X = (MapNpc(Index).X * PIC_X + sx + MapNpc(Index).XOffset) - NewPlayerPOffsetX
    Y = (MapNpc(Index).Y * PIC_Y + sy + MapNpc(Index).YOffset) - NewPlayerPOffsetY + ty + 3
    
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(X, Y, X + 32, Y + 4)
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(X, Y, X + (MapNpc(Index).HP / MapNpc(Index).MaxHp * 32), Y + 4)
    If MapNpc(Index).MaxSp > 0 Then
       Call DD_BackBuffer.SetFillColor(RGB(122, 10, 122))
       Call DD_BackBuffer.DrawBox(X, Y + 4, X + 32, Y + 4 + 4)
       Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
       Call DD_BackBuffer.DrawBox(X, Y + 4, X + (MapNpc(Index).SP / MapNpc(Index).MaxSp * 32), Y + 4 + 4)
    End If
End Sub

Public Sub AffInv()
Dim i As Integer
Dim invSlotNum As Long
Dim Qq As Long
    For invSlotNum = 0 To MAX_INV
        Qq = Player(MyIndex).Inv(invSlotNum).num

        If Qq = -1 Then
            frmMirage.picInv(invSlotNum).Picture = LoadPicture()
        Else
            Call AffSurfPic(DD_ItemSurf, frmMirage.picInv(invSlotNum), (item(Qq).Pic - (item(Qq).Pic \ 6) * 6) * PIC_X, (item(Qq).Pic \ 6) * PIC_Y, False)
            
            If IsItemInCooldown(Qq) Then
                Call ShadePictureBox(frmMirage.picInv(invSlotNum))
            End If
            
            frmMirage.picInv(invSlotNum).Refresh
        End If
    Next invSlotNum
End Sub

Public Sub Affspell()
Dim Q As Long
Dim Qq As Long
    For Q = 0 To MAX_PLAYER_SKILLS
        Qq = Player(MyIndex).skill(Q)
        If Qq = -1 Then
            frmMirage.picspell(Q).Picture = LoadPicture()
            frmMirage.picspell(Q).Enabled = False
        Else
            Call AffSurfPic(DD_ItemSurf, frmMirage.picspell(Q), (skill(Qq).SkillIco - (skill(Qq).SkillIco \ 6) * 6) * PIC_X, (skill(Qq).SkillIco \ 6) * PIC_Y)
            frmMirage.picspell(Q).Enabled = True
        End If
    Next Q
End Sub

Public Sub EnableSkill(ByVal skillNum As Integer)
    Dim skillId As Integer
    
    skillId = Player(MyIndex).skill(skillNum)
    
    frmMirage.picspell(skillNum).Enabled = True
    Call AffSurfPic(DD_ItemSurf, frmMirage.picspell(skillNum), (skill(skillId).SkillIco - (skill(skillId).SkillIco \ 6) * 6) * PIC_X, (skill(skillId).SkillIco \ 6) * PIC_Y)
End Sub

Public Sub DisableSkill(ByVal skillNum As Integer)
    frmMirage.picspell(skillNum).Enabled = False
    Call ShadePictureBox(frmMirage.picspell(skillNum))
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long

frmMirage.ShieldImage.Picture = LoadPicture()
frmMirage.WeaponImage.Picture = LoadPicture()
frmMirage.HelmetImage.Picture = LoadPicture()
frmMirage.ArmorImage.Picture = LoadPicture()
frmMirage.PetImage.Picture = LoadPicture()

    With Player(MyIndex).ShieldSlot
        If .num > 0 Then
            Call AffSurfPic(DD_ItemSurf, frmMirage.ShieldImage, (item(.num).Pic - (item(.num).Pic \ 6) * 6) * PIC_X, (item(.num).Pic \ 6) * PIC_Y)
        End If
    End With
    With Player(MyIndex).WeaponSlot
        If .num > 0 Then
            Call AffSurfPic(DD_ItemSurf, frmMirage.WeaponImage, (item(.num).Pic - (item(.num).Pic \ 6) * 6) * PIC_X, (item(.num).Pic \ 6) * PIC_Y)
        End If
    End With
    With Player(MyIndex).HelmetSlot
        If .num > 0 Then
            Call AffSurfPic(DD_ItemSurf, frmMirage.HelmetImage, (item(.num).Pic - (item(.num).Pic \ 6) * 6) * PIC_X, (item(.num).Pic \ 6) * PIC_Y)
        End If
    End With
    With Player(MyIndex).ArmorSlot
        If .num > 0 Then
            Call AffSurfPic(DD_ItemSurf, frmMirage.ArmorImage, (item(.num).Pic - (item(.num).Pic \ 6) * 6) * PIC_X, (item(.num).Pic \ 6) * PIC_Y)
        End If
    End With
    With Player(MyIndex).PetSlot
        If .num > 0 Then
            Call AffSurfPic(DD_ItemSurf, frmMirage.PetImage, (item(.num).Pic - (item(.num).Pic \ 6) * 6) * PIC_X, (item(.num).Pic \ 6) * PIC_Y)
        End If
    End With

    Call AffInv
    Call AffRac
End Sub
Public Sub QueteMsg(ByVal Index As Long, ByVal Msg As String)
frmMirage.txtQ.Visible = True
frmMirage.TxtQ2.Text = Msg
End Sub

Sub BltSpriteChange(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, y2 As Long
    rec.Top = Map.tile(X, Y).Datas(0) * (PIC_NPC1 * 32) + PIC_NPC2
    rec.bottom = rec.Top + PIC_Y
    rec.Left = 128
    rec.Right = rec.Left + PIC_X
    
    x2 = X * PIC_X + sx
    y2 = Y * PIC_Y + sy
                                       
    Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(MyIndex)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltSpriteChange2(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, y2 As Long
    rec.Top = Map.tile(X, Y).Datas(0) * 64
    rec.bottom = rec.Top + PIC_Y
    rec.Left = 128
    rec.Right = rec.Left + PIC_X
    
    x2 = X * PIC_X + sx
    y2 = Y * PIC_Y + sy - 32
    If x2 < 0 Then x2 = 0
    If y2 < 0 Then y2 = 0
                        
    Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_SpriteSurf(GetPlayerSprite(MyIndex)), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub SendGameTime()
'Dim Packet As clsBuffer
'
'    Set Packet = New clsBuffer
'
'    Packet.WriteLong CSetGameTime
'    Packet.WriteLong GameTime
'
'    SendData Packet.ToArray()
'    Set Packet = Nothing
End Sub

'Sub SendGameTime()
'Dim Packet As String
'
'Packet = "GmTime" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
'Call SendData(Packet)
'End Sub

Sub ItemSelected(ByVal Index As Long, ByVal Selected As Long)
Dim index2 As Long
index2 = Trade(Selected).items(Index).ItemGetNum

    frmTrade.shpSelect.Top = frmTrade.picItem(Index - 1).Top - 1
    frmTrade.shpSelect.Left = frmTrade.picItem(Index - 1).Left - 1

    If index2 <= 0 Then Call clearItemSelected: Exit Sub

    frmTrade.descName.Caption = Trim$(item(index2).name)
    frmTrade.descName.ForeColor = item(index2).NCoul
    frmTrade.descQuantity.Caption = Trade(Selected).items(Index).ItemGetVal
    
    frmTrade.descStr.Caption = item(index2).StrReq
    frmTrade.descDef.Caption = item(index2).DefReq
    frmTrade.descDex.Caption = item(index2).DexReq
    frmTrade.descSci.Caption = item(index2).SciReq
    frmTrade.descLang.Caption = item(index2).LangReq
    
    frmTrade.descAStr.Caption = item(index2).AddStr
    frmTrade.descADef.Caption = item(index2).AddDef
    frmTrade.descASci.Caption = item(index2).AddSci
    frmTrade.descADex.Caption = item(index2).AddDex
    frmTrade.descALang.Caption = item(index2).AddLang
    
    frmTrade.descHp.Caption = item(index2).AddHP
    frmTrade.descMp.Caption = item(index2).AddSLP
    frmTrade.descSp.Caption = item(index2).AddSTP

    frmTrade.descAExp.Caption = item(index2).AddEXP
    frmTrade.desc.Caption = Trim$(item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = Trim$(item(Trade(Selected).items(Index).ItemGiveNum).name)
    frmTrade.lblTradeFor.ForeColor = item(Trade(Selected).items(Index).ItemGiveNum).NCoul
    frmTrade.lblQuantity.Caption = Trade(Selected).items(Index).ItemGiveVal
End Sub

Sub clearItemSelected()
    frmTrade.lblTradeFor.Caption = vbNullString
    frmTrade.lblQuantity.Caption = vbNullString
    
    frmTrade.descName.Caption = vbNullString
    frmTrade.descQuantity.Caption = vbNullString
    
    frmTrade.descStr.Caption = 0
    frmTrade.descDef.Caption = 0
    frmTrade.descDex.Caption = 0
    frmTrade.descSci.Caption = 0
    frmTrade.descLang.Caption = 0
    
    frmTrade.descAStr.Caption = 0
    frmTrade.descADef.Caption = 0
    frmTrade.descASci.Caption = 0
    frmTrade.descADex.Caption = 0
    frmTrade.descALang.Caption = 0
    
    frmTrade.descHp.Caption = 0
    frmTrade.descMp.Caption = 0
    frmTrade.descSp.Caption = 0

    frmTrade.descAExp.Caption = 0
    frmTrade.desc.Caption = vbNullString
End Sub

Sub AffSurfPic(DD_Surf As DirectDrawSurface7, ByVal picBox As PictureBox, ByVal X As Long, ByVal Y As Long, Optional mustRefresh As Boolean = True)
Dim sRECT As RECT
Dim dRECT As RECT

    If Not (DD_Surf Is Nothing) Then
    If DD_Surf Is Nothing Then Exit Sub
    picBox.Picture = LoadPicture()
    With dRECT
        .Top = 0
        .bottom = picBox.Height
        .Left = 0
        .Right = picBox.Width
    End With
    With sRECT
        .Top = Y
        .bottom = .Top + picBox.Height
        .Left = X
        .Right = .Left + picBox.Width
    End With
    Call DD_Surf.BltToDC(picBox.hdc, sRECT, dRECT)
    If mustRefresh Then
        picBox.Refresh
    End If
    End If
End Sub

Sub RefreshParty()
    Dim indexCourant As Integer
    Dim nombrePage As Integer
    Dim pageDatas() As String
    Dim C As Byte

    If Player(MyIndex).partyIndex > -1 Then
        If (Not frmMirage.picParty.Visible) Then frmMirage.picParty.Visible = True
        nombrePage = ((nbPartyPlayer - 1) \ 3) + 1
        pageDatas = Split(frmMirage.lbl_party_numPage.Caption, "/")
        If pageDatas(0) > nombrePage Then 'Si des joueurs sont parties et qu'on était sur une page qui n'existe plus, il faut changer le numéro de page
            pageDatas(0) = nombrePage
        End If
        frmMirage.lbl_party_numPage.Caption = pageDatas(0) + "/" + Str(nombrePage)
        
        C = 0
        For indexCourant = 1 To MAX_PLAYERS
            If IsPlaying(indexCourant) And Player(indexCourant).partyIndex = Player(MyIndex).partyIndex And C < 3 And indexCourant <> MyIndex Then
                C = C + 1
                frmMirage.lblPPName(C - 1).Tag = indexCourant
            End If
        Next
    
        For indexCourant = 0 To 2
            frmMirage.lblPPName(indexCourant).Visible = (indexCourant < C)
            frmMirage.backPPLife(indexCourant).Visible = frmMirage.lblPPName(indexCourant).Visible
            frmMirage.backPPMana(indexCourant).Visible = frmMirage.lblPPName(indexCourant).Visible
            If frmMirage.lblPPName(indexCourant).Visible Then
                frmMirage.lblPPName(indexCourant).Caption = Trim$(Player(Val(frmMirage.lblPPName(indexCourant).Tag)).name) & " - " & Player(Val(frmMirage.lblPPName(indexCourant).Tag)).level
                Debug.Print "valeur d'index : " & Player(Val(frmMirage.lblPPName(indexCourant).Tag)).HP
                frmMirage.shpPPLife(indexCourant).Width = Player(Val(frmMirage.lblPPName(indexCourant).Tag)).HP / Player(Val(frmMirage.lblPPName(indexCourant).Tag)).MaxHp * frmMirage.backPPLife(indexCourant).Width
                frmMirage.shpPPMana(indexCourant).Width = Player(Val(frmMirage.lblPPName(indexCourant).Tag)).STP / Player(Val(frmMirage.lblPPName(indexCourant).Tag)).MaxSTP * frmMirage.backPPMana(indexCourant).Width
                frmMirage.lblPPLife(indexCourant).Caption = "PV : " & Player(Val(frmMirage.lblPPName(indexCourant).Tag)).HP & "/" & Player(Val(frmMirage.lblPPName(indexCourant).Tag)).MaxHp
                frmMirage.lblPPMana(indexCourant).Caption = "PM : " & Player(Val(frmMirage.lblPPName(indexCourant).Tag)).STP & "/" & Player(Val(frmMirage.lblPPName(indexCourant).Tag)).MaxSTP
            End If
        Next
    End If
End Sub

Sub RefreshSkills()
    Dim i As Integer
    
    For i = 0 To MAX_PLAYER_SKILLS
        If Player(MyIndex).skill(i) > -1 Then
            If skill(Player(MyIndex).skill(i)).Type = 0 Or skill(Player(MyIndex).skill(i)).Type = 1 Then
                If Pets(MyIndex).num > -1 Then
                    If frmMirage.picspell(i).Enabled = False Then
                        Call EnableSkill(i)
                    End If
                Else
                    If frmMirage.picspell(i).Enabled = True Then
                        Call DisableSkill(i)
                    End If
                End If
            End If
        End If
    Next i
End Sub

Sub RefreshOnlineList()
    Dim i As Integer
    Dim selectedPlayerName As String
    
    selectedPlayerName = frmMirage.lstOnline.List(frmMirage.lstOnline.ListIndex)
    frmMirage.lstOnline.Clear
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If MyIndex <> i Then
                frmMirage.lstOnline.AddItem Player(i).name
                frmMirage.lstOnline.ItemData(frmMirage.lstOnline.ListCount - 1) = i
            End If
            If selectedPlayerName = Player(i).name Then
                frmMirage.lstOnline.ListIndex = frmMirage.lstOnline.ListCount - 1
            End If
        End If
    Next i
      
    Dim sngOldHeight
    Dim previousHeight

    frmMirage.m_sngLBRowHeight = 1
    If frmMirage.lstOnline.ListCount > 0 Then
        With frmMirage.lstOnline
            .TopIndex = .ListCount - 1
            sngOldHeight = .Height
            Debug.Print "bug serveur"
            Do While .TopIndex = 0
                previousHeight = .Height
                .Height = .Height - 10
                .TopIndex = .ListCount - 1

                If previousHeight = .Height Then ' Minimum Height reach
                    Exit Do
                End If
            Loop
            frmMirage.m_sngLBRowHeight = .Height / (.ListCount - .TopIndex + 1)
            .Height = sngOldHeight
            .TopIndex = 0
        End With
    End If
End Sub

Sub GameInit()
    Dim i As Integer
    
    Call ResetDragAndDrop
    
    movementController = 0
    
    GettingMap = True
    IsDead = False
    
    Set cooldownItem = Nothing
    Set cooldownItem = New Collection
    
    Set ArrowsEffect = Nothing
    Set ArrowsEffect = New Dictionary
    
    Set SkillsEffect = Nothing
    Set SkillsEffect = New Collection
    
    Set AreasWeather = Nothing
    Set AreasWeather = New Collection

    ' Initialiser les variables qui l'étaient avant avec le paquet maxinfo
    MAX_BLT_LINE = 10
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    
    Call ClearCrafts
    
    For i = 1 To MAX_PLAYERS
        PlayerAnim(i, 0) = 0
    Next i
    
    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).command = vbNullString
    Next i

    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next i
    
    ' Load the items
    For i = 0 To MAX_ITEMS
        DoEvents
        Call LoadItem(i)
    Next i
    
    For i = 0 To MAX_CRAFTS
        DoEvents
        Call LoadCraft(i)
    Next i
    
    For i = 0 To MAX_SKILLS
        DoEvents
        Call LoadSkill(i)
    Next i
    
    For i = 0 To MAX_NPCS
        DoEvents
        Call LoadNpc(i)
    Next i
    
    ' Display
    frmMirage.lblSTRBonus.Caption = ""
    frmMirage.lblDEFBonus.Caption = ""
    frmMirage.lblDEXBonus.Caption = ""
    frmMirage.lblSCIBonus.Caption = ""
    frmMirage.lblLANGBonus.Caption = ""
End Sub

Sub beginGame()
    InGame = True
    
    frmMirage.Show
    frmMainMenu.Visible = False
    frmsplash.Visible = False
    frmMirage.picScreen.Visible = True
    
    Call GameLoop
End Sub

Sub DisplayDescription(itemSlot As ItemSlotRec)

    If item(itemSlot.num).Type = ITEM_TYPE_CURRENCY And Trim$(item(itemSlot.num).desc) = vbNullString Then
        frmMirage.itmDesc.Height = 17
    ElseIf Trim$(item(itemSlot.num).desc) = vbNullString Then
        frmMirage.itmDesc.Height = 161
    ElseIf Trim$(item(itemSlot.num).desc) > vbNullString Then
        frmMirage.itmDesc.Height = 249
    End If

    Dim pos As POINTAPI

    pos = GetMousePosition

    frmMirage.itmDesc.Top = pos.Y + 15
    frmMirage.itmDesc.Left = pos.X + 15

    If item(itemSlot.num).Type = ITEM_TYPE_CURRENCY Or item(itemSlot.num).Empilable <> 0 Then
        frmMirage.descName.Caption = Trim$(item(itemSlot.num).name) & " (" & itemSlot.Value & ")"
    Else
        frmMirage.descName.Caption = Trim$(item(itemSlot.num).name) & " (" & itemSlot.Value & ")"
    End If
    If item(itemSlot.num).Type = ITEM_TYPE_PET Then
        frmMirage.descStr.Caption = Npc(Pets(item(itemSlot.num).Datas(0)).num).Str & " Force"
        frmMirage.descDef.Caption = Npc(Pets(item(itemSlot.num).Datas(0)).num).Def & " Défense"
    Else
        frmMirage.descStr.Caption = item(itemSlot.num).StrReq & " Force"
        frmMirage.descDef.Caption = item(itemSlot.num).DefReq & " Défense"
    End If
    frmMirage.descDex.Caption = item(itemSlot.num).DexReq & " Dexterité"
    frmMirage.descHpMp.Caption = "PV: " & item(itemSlot.num).AddHP & " PM: " & item(itemSlot.num).AddSLP & " End: " & item(itemSlot.num).AddSTP
    frmMirage.descSD.Caption = "FOR: " & item(itemSlot.num).AddStr & " Def: " & item(itemSlot.num).AddDef
    frmMirage.descMS.Caption = "Science: " & item(itemSlot.num).AddSci & " Vitesse: " & item(itemSlot.num).AddDex
    If (item(itemSlot.num).Type >= ITEM_TYPE_WEAPON) And (item(itemSlot.num).Type <= ITEM_TYPE_SHIELD) Then
        If item(itemSlot.num).Datas(0) <= 0 Then frmMirage.Usure.Caption = "Usure : Ind." Else frmMirage.Usure.Caption = "Usure : " & itemSlot.dur & "/" & item(itemSlot.num).Datas(0)
    End If
    frmMirage.desc.Caption = Trim$(item(itemSlot.num).desc)
    frmMirage.descName.ForeColor = item(itemSlot.num).NCoul
    frmMirage.itmDesc.Visible = True
End Sub
