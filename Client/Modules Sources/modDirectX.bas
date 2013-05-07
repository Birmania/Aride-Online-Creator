Attribute VB_Name = "modDirectX"
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

'Le code utilise pour l'alpha blending est modifie a partir
'du code de Matt Hafermann(http://www.rookscape.com/vbgaming/tutBT.php)

Public Const TilesInSheets As Byte = 14

Public dX As New DirectX7
Public DD As DirectDraw7

Public D3D As Direct3D7
Public Dev As Direct3DDevice7

Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_FrontClip As DirectDrawClipper
Public DD_FrontBuffer As DirectDrawSurface7
Public DDSD_FrontBuffer As DDSURFACEDESC2

Public DD_SpriteSurf() As DirectDrawSurface7
Public DDSD_Character() As DDSURFACEDESC2

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_SpellAnim() As DirectDrawSurface7
Public DDSD_SpellAnim() As DDSURFACEDESC2

Public DD_BigSpellAnim() As DirectDrawSurface7
Public DDSD_BigSpellAnim() As DDSURFACEDESC2

Public TilesPic As New Collection

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public DDSD_Outil As DDSURFACEDESC2
Public DD_OutilSurf As DirectDrawSurface7

Public DDSD_Snow As DDSURFACEDESC2
Public DD_SnowSurf As DirectDrawSurface7

Public DDSD_Sand As DDSURFACEDESC2
Public DD_SandSurf As DirectDrawSurface7

'PAPERDOLL

Public DD_PaperDollSurf() As DirectDrawSurface7
Public DDSD_PaperDoll() As DDSURFACEDESC2

'FIN PAPERDOLL

Public DD_PetsSurf() As DirectDrawSurface7
Public DDSD_Pets() As DDSURFACEDESC2

Public DDSD_Blood As DDSURFACEDESC2
Public DD_Blood As DirectDrawSurface7

Public DDSD_PanoInf As DDSURFACEDESC2
Public DD_PanoInfSurf As DirectDrawSurface7

Public DDSD_PanoSup As DDSURFACEDESC2
Public DD_PanoSupSurf As DirectDrawSurface7

Public DDSD_Sleep As DDSURFACEDESC2
Public DD_SleepSurf As DirectDrawSurface7

Public DDSD_Night As DDSURFACEDESC2
Public DD_NightSurf As DirectDrawSurface7

Public DDSD_Fog As DDSURFACEDESC2
Public DD_FogSurf As DirectDrawSurface7

Public DDSD_Temp As DDSURFACEDESC2
Public DD_TmpSurf As DirectDrawSurface7

Public rec As RECT
Public rec_pos As RECT

Public AlphaBlendDXIsInit As Boolean
Public DirectXIsInit As Boolean
Public ABDXWidth As Integer
Public ABDXHeight As Integer
Public ABDXAlpha As Single

' Image gestion
Public GestionImage As New clsGestionImage

Sub InitDirectX()

    'On Error GoTo Propagate
    Set DD = dX.DirectDrawCreate(vbNullString)
    
    Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.lBackBufferCount = 1
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hwnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
    
    'Initisalisation de D3D
    Set D3D = DD.GetDirect3D
    
    DirectXIsInit = True
End Sub

Function LoadMaxSprite() As Long
Dim i As Long
    For i = 0 To 9999
        If Not FileExist(App.Path & "\GFX\Sprites\Sprites" & i & ".png") Then
            If i < 1 Then
                MsgBox "Erreur : Sprite introuvable !"
                End
            Else
                LoadMaxSprite = i - 1
            End If
            Exit Function
        Else
            LoadMaxSprite = i + 1
        End If
    Next i
End Function

Function LoadMaxPaperdolls() As Long
Dim i As Long
    For i = 0 To 9999
        If Not FileExist(App.Path & "\GFX\Paperdolls\Paperdolls" & i & ".png") Then
            If i < 1 Then
                MsgBox "Erreur : Paperdolls introuvable !"
                End
            Else
                LoadMaxPaperdolls = i - 1
            End If
            Exit Function
        Else
            LoadMaxPaperdolls = i + 1
        End If
    Next i
End Function

Function LoadMaxSpells() As Long
Dim i As Long
    For i = 0 To 9999
        If Not FileExist(App.Path & "\GFX\Spells\Spells" & i & ".png") Then
            If i < 1 Then
                MsgBox "Erreur : Spells introuvable !"
                End
            Else
                LoadMaxSpells = i - 1
            End If
            Exit Function
        Else
            LoadMaxSpells = i + 1
        End If
    Next i
End Function

Function LoadMaxBigSpells() As Long
Dim i As Long
    For i = 0 To 9999
        If Not FileExist(App.Path & "\GFX\BigSpells\BigSpells" & i & ".png") Then
            If i < 1 Then
                MsgBox "Erreur : BigSpells introuvable !"
                End
            Else
                LoadMaxBigSpells = i - 1
            End If
            Exit Function
        Else
            LoadMaxBigSpells = i + 1
        End If
    Next i
End Function

Function LoadMaxPet() As Long
Dim i As Long
    For i = 0 To 9999
        If Not FileExist(App.Path & "\GFX\Pets\Pet" & i & ".png") Then
            If i < 1 Then
                MsgBox "Erreur : Famillier introuvable !"
                End
            Else
                LoadMaxPet = i - 1
            End If
            Exit Function
        Else
             LoadMaxPet = i + 1
        End If
    Next i
    
End Function

Sub InitSurfaces()
Dim key As DDCOLORKEY
Dim i As Long

    'On Error GoTo Propagate
    ' Check for files existing
    If Not FileExist(App.Path & "\GFX\items.png") Or Not FileExist(App.Path & "\GFX\emoticons.png") Or Not FileExist(App.Path & "\GFX\Outils.png") Or Not FileExist(App.Path & "\GFX\arrows.png") Then Call MsgBox("Plusieur fichier manquants", vbOKOnly, Game_Name): Call GameDestroy
    
    ' Set the key for masks
    key.low = 0
    key.high = 0
    
    Set DD_FrontBuffer = Nothing
    
    Dim FrontBufferWidth As Integer
    Dim FrontBufferHeight As Integer
    
    DDSD_FrontBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_FrontBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    ' On aurait pu imaginer ne prendre que le width et height mais on prend un
    ' peu plus pour respecter tout les blits déjà existant qui se faisait dans le backbuffer
    DDSD_FrontBuffer.lWidth = frmMirage.picScreen.Width + PIC_X
    DDSD_FrontBuffer.lHeight = frmMirage.picScreen.Height + PIC_Y
    Set DD_FrontBuffer = DD.CreateSurface(DDSD_FrontBuffer)
            
            
    Set DD_OutilSurf = Nothing
    
    DDSD_Outil.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Outil.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_OutilSurf = LoadImage(App.Path & "\GFX\Outils.png", DD, DDSD_Outil)
    SetMaskColorFromPixel DD_OutilSurf, 0, 0

    DDSD_Sleep.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
    DDSD_Sleep.ddsCaps.lCaps = DDSCAPS_TEXTURE
    DDSD_Sleep.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
    Set DD_SleepSurf = LoadImage(App.Path & "\GFX\sleep_vision.png", DD, DDSD_Sleep)
    Call SetMaskColor(DD_SleepSurf, vbWhite)
        
    DDSD_Snow.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Snow.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SnowSurf = LoadImage(App.Path & "\GFX\snow.png", DD, DDSD_Snow)
    SetMaskColorFromPixel DD_SnowSurf, 0, 0
    
    DDSD_Sand.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Sand.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SandSurf = LoadImage(App.Path & "\GFX\sand.png", DD, DDSD_Sand)
    SetMaskColorFromPixel DD_SandSurf, 0, 0
    
    For i = 0 To LoadMaxSprite()
        Set DD_SpriteSurf(i) = Nothing
    
        DDSD_Character(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Character(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_SpriteSurf(i) = LoadImage(App.Path & "\GFX\Sprites\Sprites" & i & ".png", DD, DDSD_Character(i))
        SetMaskColorFromPixel DD_SpriteSurf(i), 0, 0
    Next i
    
    Set DD_ItemSurf = Nothing
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = LoadImage(App.Path & "\GFX\items.png", DD, DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0
        
    Set DD_EmoticonSurf = Nothing
        
    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = LoadImage(App.Path & "\GFX\emoticons.png", DD, DDSD_Emoticon)
    SetMaskColorFromPixel DD_EmoticonSurf, 0, 0
    
    ' Init spells ddsd type and load the bitmap
    For i = 0 To LoadMaxSpells()
        Set DD_SpellAnim(i) = Nothing
        'Call ZeroMemory(ByVal VarPtr(DDSD_SpellAnim(i)), LenB(DDSD_SpellAnim(i)))
    
        DDSD_SpellAnim(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_SpellAnim(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_SpellAnim(i) = LoadImage(App.Path & "\GFX\Spells\Spells" & i & ".png", DD, DDSD_SpellAnim(i))
        SetMaskColorFromPixel DD_SpellAnim(i), 0, 0
    Next i
        
    For i = 0 To LoadMaxBigSpells()
        Set DD_BigSpellAnim(i) = Nothing
    
        DDSD_BigSpellAnim(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_BigSpellAnim(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_BigSpellAnim(i) = LoadImage(App.Path & "\GFX\BigSpells\BigSpells" & i & ".png", DD, DDSD_BigSpellAnim(i))
        SetMaskColorFromPixel DD_BigSpellAnim(i), 0, 0
    Next i
        
    Set DD_ArrowAnim = Nothing
        
    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = LoadImage(App.Path & "\GFX\arrows.png", DD, DDSD_ArrowAnim)
    SetMaskColorFromPixel DD_ArrowAnim, 0, 0
    
    Set DD_Blood = Nothing
    
    'Chargement de la planche de sang
    DDSD_Blood.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Blood.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_Blood = LoadImage(App.Path & "\GFX\blood.png", DD, DDSD_Blood)
    SetMaskColorFromPixel DD_Blood, 0, 0
    
    'PAPERDOLL
    For i = 0 To LoadMaxPaperdolls()
        Set DD_PaperDollSurf(i) = Nothing
        'Call ZeroMemory(ByVal VarPtr(DDSD_PaperDoll(i)), LenB(DDSD_PaperDoll(i)))
    
        DDSD_PaperDoll(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_PaperDoll(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        Set DD_PaperDollSurf(i) = LoadImage(App.Path & "\GFX\Paperdolls\Paperdolls" & i & ".png", DD, DDSD_PaperDoll(i))
        SetMaskColorFromPixel DD_PaperDollSurf(i), 0, 0
    Next i
    'FIN PAPERDOLL
    
    For i = 0 To LoadMaxPet
        Set DD_PetsSurf(i) = Nothing
        'Call ZeroMemory(ByVal VarPtr(DDSD_Pets(i)), LenB(DDSD_Pets(i)))
    
        DDSD_Pets(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Pets(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_PetsSurf(i) = LoadImage(App.Path & "\GFX\Pets\Pet" & i & ".png", DD, DDSD_Pets(i))
        SetMaskColorFromPixel DD_PetsSurf(i), 0, 0
    Next i
    
    If InGame Then
        Call InitSurfacesAccordingMap
    End If
End Sub

Sub InitSurfacesAccordingMap()
    ' On reinitialise le backbuffer car sa taille dépend de la taille de la map
    Set DD_BackBuffer = Nothing
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MaxMapX + 3) * PIC_X
    DDSD_BackBuffer.lHeight = (MaxMapY + 3) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    Call DestroyAlphaBlendDX
    
    On Error GoTo EndAlphaBlending
    
    'Initisialisation de la surface temporaire
    DDSD_Temp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    'DDSCAPS_3DDEVICE est utilise pour pouvoir utiliser D3D sur cette surface
    DDSD_Temp.lWidth = (MaxMapX + 3) * PIC_X
    DDSD_Temp.lHeight = (MaxMapY + 3) * PIC_Y

    Set DD_TmpSurf = DD.CreateSurface(DDSD_Temp)
    Call DD_TmpSurf.SetForeColor(0)
    Call DD_TmpSurf.setDrawStyle(5)
    
    Call InitNightSurface
    Call InitFogSurface
    
    Call InitAlphaBlendDX
    
EndAlphaBlending:
    
End Sub

Sub InitFogSurface()
    'Initialisation du brouillard
    If Map.Fog <> 0 Then
        'Initialisation de la texture pour le brouillard si il y en a un
        DDSD_Fog.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        DDSD_Fog.ddsCaps.lCaps = DDSCAPS_TEXTURE
        DDSD_Fog.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        Set DD_FogSurf = LoadImage(App.Path & "\GFX\fog" & Map.Fog & ".png", DD, DDSD_Fog)
    Else
        Set DD_FogSurf = Nothing
    End If
End Sub

Sub InitNightSurface()
    Dim X As Long
    Dim Y As Long
    Dim tile As Long
            
    'Initialisation de la texture pour la nuit
    DDSD_Night.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
    DDSD_Night.ddsCaps.lCaps = DDSCAPS_TEXTURE
    DDSD_Night.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
    DDSD_Night.lWidth = (MaxMapX + 1) * PIC_X
    DDSD_Night.lHeight = (MaxMapY + 1) * PIC_Y
    Set DD_NightSurf = DD.CreateSurface(DDSD_Night)

    With rec
        .Top = 0
        .bottom = (MaxMapX + 1) * PIC_X
        .Left = 0
        .Right = (MaxMapY + 1) * PIC_Y
    End With
    Call DD_NightSurf.BltColorFill(rec, 0)
    Call SetMaskColor(DD_NightSurf, vbWhite)
End Sub

Sub DestroyDirectX()
Dim i As Long

    If DirectXIsInit Then
        DD.RestoreDisplayMode
        
        Set dX = Nothing
        Set DD = Nothing
        Set DD_PrimarySurf = Nothing
        Set DD_BackBuffer = Nothing
        Set DD_FrontBuffer = Nothing
        For i = 0 To LoadMaxSprite()
            Set DD_SpriteSurf(i) = Nothing
        Next i
        
        Set TilesPic = Nothing
        Set DD_ItemSurf = Nothing
        Set DD_EmoticonSurf = Nothing
        For i = 0 To LoadMaxSpells()
            Set DD_SpellAnim(i) = Nothing
        Next i
        For i = 0 To LoadMaxBigSpells()
            Set DD_BigSpellAnim(i) = Nothing
        Next i
        Set DD_ArrowAnim = Nothing
        Set DD_PanoInfSurf = Nothing
        Set DD_PanoSupSurf = Nothing
        Set DD_FogSurf = Nothing
        Set DD_NightSurf = Nothing
    
        DirectXIsInit = False
    End If
    
    If AlphaBlendDXIsInit Then
        Call DestroyAlphaBlendDX
    End If
End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = DD.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then NeedToRestoreSurfaces = False Else NeedToRestoreSurfaces = True
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

With TmpR
.Left = X
.Top = Y
.Right = X
.bottom = Y
End With

TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

With TmpColorKey
.low = TheSurface.GetLockedPixel(X, Y)
.high = .low
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

TheSurface.Unlock TmpR
End Sub

Public Sub SetDefaultMaskColor(ByRef TheSurface As DirectDrawSurface7)
Dim TmpColorKey As DDCOLORKEY

With TmpColorKey
    .low = 0
    .high = 0
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
End Sub

Public Sub SetMaskColor(ByRef TheSurface As DirectDrawSurface7, ByVal Color As Long)
    Dim TmpColorKey As DDCOLORKEY
    
    With TmpColorKey
        .low = Color
        .high = Color
    End With
    
    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
End Sub

Sub DisplayFx(intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, tile As Long, ByRef lngDestDC As Long, ByRef lngSrcDC As Long)
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (tile - (tile \ TilesInSheets) * TilesInSheets) * PIC_X, (tile \ TilesInSheets) * PIC_Y, lngROP
End Sub

Sub Night(ByVal MinX As Long, ByVal MaxX As Long, ByVal MinY As Long, ByVal MaxY As Long)
Dim X As Long, Y As Long
    If AccOpt.LowEffect Then
        Dim lngSrcDC As Long
        
        lngSrcDC = DD_OutilSurf.GetDC
        For Y = MinY To MaxY
            For X = MinX To MaxX
                If Map.tile(X, Y).Light <= 0 Then
                    DisplayFx (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31, TexthDC, lngSrcDC
                Else
                    DisplayFx (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sy - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map.tile(X, Y).Light, TexthDC, lngSrcDC
                End If
            Next X
        Next Y
        
        DD_OutilSurf.ReleaseDC lngSrcDC
    Else
        'Initialisation du RECT source
        'Initialisation du RECT source
        With rec_pos
            .Top = 0
            .bottom = (MaxY - MinY + 2) * PIC_Y
            .Left = 0
            .Right = .Left + (MaxX - MinX + 2) * PIC_X
        End With
        
        'Initialisation du RECT destination
        With rec
            .Top = -PIC_Y + (NewPlayerY * 32) + NewYOffset
            .bottom = .Top + rec_pos.bottom
            .Left = -PIC_X + (NewPlayerX * 32) + NewXOffset
            .Right = .Left + (MaxX - MinX + 2) * PIC_X
        End With
        
        'Dessin de la nuit
        Call AlphaBlendDX(DD_NightSurf, DDSD_Night, rec, rec_pos, 0.5)
    End If
End Sub

Sub BltWeather()
Dim i As Long
    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then frmMirage.tmrRainDrop.Interval = 200: frmMirage.tmrRainDrop.Tag = "123"
                End If
            End If
        Next i
    ElseIf GameWeather = WEATHER_SNOWING Or GameWeather = WEATHER_SAND_STORMING Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then frmMirage.tmrSnowDrop.Interval = 200: frmMirage.tmrSnowDrop.Tag = "123"
                End If
            End If
        Next i
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then Call ClearRainDrop(BLT_RAIN_DROPS)
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
    
    If GameWeather = WEATHER_NONE Then Exit Sub
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If Not ((DropRain(i).X = 0) Or (DropRain(i).Y = 0)) Then
                DropRain(i).X = DropRain(i).X + DropRain(i).Speed
                DropRain(i).Y = DropRain(i).Y + DropRain(i).Speed
                Call DD_BackBuffer.DrawLine(DropRain(i).X, DropRain(i).Y, DropRain(i).X + DropRain(i).Speed, DropRain(i).Y + DropRain(i).Speed)
                If (DropRain(i).X > (MaxMapX + 1) * PIC_X) Or (DropRain(i).Y > (MaxMapY + 1) * PIC_Y) Then DropRain(i).Randomized = False
            End If
        Next i
    ElseIf GameWeather = WEATHER_SNOWING Or GameWeather = WEATHER_SAND_STORMING Then
    
        rec.Top = 0
        rec.Left = 0
        rec.bottom = PIC_Y
        rec.Right = PIC_X
        
        Dim DD_Surf As DirectDrawSurface7
        If GameWeather = WEATHER_SAND_STORMING Then
            Set DD_Surf = DD_SandSurf
        ElseIf GameWeather = WEATHER_SNOWING Then
            Set DD_Surf = DD_SnowSurf
        End If
            
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).X = 0) Or (DropSnow(i).Y = 0)) Then
                DropSnow(i).X = DropSnow(i).X + DropSnow(i).Speed
                DropSnow(i).Y = DropSnow(i).Y + DropSnow(i).Speed
                Call DD_BackBuffer.BltFast(DropSnow(i).X + DropSnow(i).Speed, DropSnow(i).Y + DropSnow(i).Speed, DD_Surf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).X > (MaxMapX + 1) * PIC_X) Or (DropSnow(i).Y > (MaxMapY + 1) * PIC_Y) Then DropSnow(i).Randomized = False
            End If
        Next i
        
    End If

    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then DD_BackBuffer.SetFillColor RGB(255, 255, 255): Call PlaySound("Thunder.wav"): Call DD_BackBuffer.DrawBox(PIC_X, PIC_Y, (MaxMapX + 2) * PIC_X, (MaxMapY + 2) * PIC_Y)
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)

Start:
    DropRain(RDNumber).X = Int((((MaxMapX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).Y = Int((((MaxMapY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).Y > (MaxMapY + 1) * PIC_Y / 4) And (DropRain(RDNumber).X > (MaxMapX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).Speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).X = 0
    DropRain(RDNumber).Y = 0
    DropRain(RDNumber).Speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    With DropSnow(RDNumber)
        .X = Int((((MaxMapX + 1) * PIC_X) * Rnd) + 1)
        .Y = Int((((MaxMapY + 1) * PIC_Y) * Rnd) + 1)
        If (.Y > (MaxMapY + 1) * PIC_Y / 4) And (.X > (MaxMapX + 1) * PIC_X / 4) Then GoTo Start
        .Speed = Int((10 * Rnd) + 6)
        .Randomized = True
    End With
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    With DropSnow(RDNumber)
        .X = 0
        .Y = 0
        .Speed = 0
        .Randomized = False
    End With
End Sub

Sub BltPlayerAnim(ByVal Index As Long)
Dim X As Long, Y As Long
    If PlayerAnim(Index, 0) = 0 Then Exit Sub
    
    X = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
    
    If GetTickCount > PlayerAnim(Index, 1) + 500 Then
        PlayerAnim(Index, 2) = PlayerAnim(Index, 2) + 1
        If PlayerAnim(Index, 2) > 11 Then PlayerAnim(Index, 2) = 0
        PlayerAnim(Index, 1) = GetTickCount
    End If
    
    rec.Top = 0 * PIC_Y
    rec.bottom = rec.Top + PIC_Y
    rec.Left = PlayerAnim(Index, 2) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(X - NewPlayerPOffsetX, Y - NewPlayerPOffsetY, DD_SpellAnim(PlayerAnim(Index, 0) - 1), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    If PlayerAnim(Index, 3) <> 0 Then
         If PlayerAnim(Index, 1) > PlayerAnim(Index, 3) Then
            PlayerAnim(Index, 0) = 0
            PlayerAnim(Index, 1) = GetTickCount
            PlayerAnim(Index, 2) = 0
            PlayerAnim(Index, 3) = 0
            If PlayerAnim(Index, 4) > 0 Then
                Call SendData("exscript" & SEP_CHAR & PlayerAnim(Index, 4) - 1 & SEP_CHAR & END_CHAR)
            End If
            PlayerAnim(Index, 4) = 0
         End If
    End If
End Sub

Sub BltPlayerHotBars()
rec.Top = 6 * PIC_Y + (PIC_Y / 2)
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_FrontBuffer.BltFast(35, 40, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

rec.Top = 6 * PIC_Y
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerHP(MyIndex)) / GetPlayerMaxHP(MyIndex))
Call DD_FrontBuffer.BltFast(35, 40, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
TexthDC = DD_FrontBuffer.GetDC

DrawTextInter TexthDC, 38, 39, Trim$(GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex))
Call DD_FrontBuffer.ReleaseDC(TexthDC)

rec.Top = 7 * PIC_Y + (PIC_Y / 2)
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_FrontBuffer.BltFast(35, 60, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

rec.Top = 7 * PIC_Y
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerSTP(MyIndex)) / GetPlayerMaxSTP(MyIndex))
Call DD_FrontBuffer.BltFast(35, 60, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
TexthDC = DD_FrontBuffer.GetDC
DrawTextInter TexthDC, 38, 59, Trim$(GetPlayerSTP(MyIndex) & " / " & GetPlayerMaxSTP(MyIndex))
Call DD_FrontBuffer.ReleaseDC(TexthDC)

rec.Top = 8 * PIC_Y + (PIC_Y / 2)
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_FrontBuffer.BltFast(35, 80, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

rec.Top = 8 * PIC_Y
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerExp(MyIndex)) / nelvl)
Call DD_FrontBuffer.BltFast(35, 80, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
TexthDC = DD_FrontBuffer.GetDC
DrawTextInter TexthDC, 38, 79, Trim$(GetPlayerExp(MyIndex) & " / " & nelvl)
Call DD_FrontBuffer.ReleaseDC(TexthDC)

rec.Top = 9 * PIC_Y + (PIC_Y / 2)
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (PIC_X * 4)
Call DD_FrontBuffer.BltFast(35, 100, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

rec.Top = 9 * PIC_Y
rec.bottom = rec.Top + (PIC_Y / 2)
rec.Left = 0 * PIC_X
rec.Right = rec.Left + (((PIC_X * 4) * GetPlayerSLP(MyIndex)) / GetPlayerMaxSLP(MyIndex))
Call DD_FrontBuffer.BltFast(35, 100, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
TexthDC = DD_FrontBuffer.GetDC
DrawTextInter TexthDC, 38, 99, Trim$(GetPlayerSLP(MyIndex) & " / " & GetPlayerMaxSLP(MyIndex))
Call DD_FrontBuffer.ReleaseDC(TexthDC)

End Sub

Sub BltSkillEffect(ByRef SkillAnim As clsSkillAnim)
Dim X As Long, Y As Long, i As Long

If skill(SkillAnim.castedSpell).SkillAnimId < 0 Then Exit Sub

    With SkillAnim
        If Not .SkillDone Then
            'TODO
        End If
    End With
End Sub
' Effet de sang
Sub BltBlood(ByVal Index As Long, Optional TailleX As Byte, Optional TailleY As Byte, Optional ImgTime As Byte)
    'TODO
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
    If Player(Index).EmoticonNum < 0 Then Exit Sub

    With Player(Index)
        If .EmoticonTime + 1300 > GetTickCount Then
            If GetTickCount >= .EmoticonTime + (108 * (.EmoticonVar + 1)) Then .EmoticonVar = .EmoticonVar + 1
                
            rec.Top = .EmoticonNum * PIC_Y
            rec.bottom = rec.Top + PIC_Y
            rec.Left = .EmoticonVar * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If Index = MyIndex Then
                x2 = newX + sx + 16
                y2 = newY + sx - 32
                
                If y2 < 0 Then Exit Sub
                
                Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                x2 = GetPlayerX(Index) * PIC_X + sx + .XOffset + 16
                y2 = GetPlayerY(Index) * PIC_Y + sx + .YOffset - (PIC_Y * 2)
                
                If y2 < 0 Then Exit Sub
                
                Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    End With
End Sub

'Sub BltArrow(ByVal Index As Long)
Sub BltArrowEffect(ByRef arrow As clsArrowAnim)
Dim X As Long, Y As Long, i As Long, z As Long
Dim BX As Long, by As Long

With arrow
If .ArrowAnim > 0 Then

    rec.Top = .ArrowAnim * PIC_Y
    rec.bottom = rec.Top + PIC_Y
    rec.Left = .ArrowPosition * PIC_X
    rec.Right = rec.Left + PIC_X
    
    If GetTickCount > .ArrowTime + 30 Then .ArrowTime = GetTickCount: .ArrowVarX = .ArrowVarX + 10: .ArrowVarY = .ArrowVarY + 10
    
    If .ArrowPosition = 0 Then
        X = .ArrowX
        Y = .ArrowY + (.ArrowVarY \ 32)
        
        If Y <= MaxMapY Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + .ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If .ArrowPosition = 1 Then
        X = .ArrowX - (.ArrowVarX \ 32)
        Y = .ArrowY
                    
        If X >= 0 Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - .ArrowVarX, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If .ArrowPosition = 2 Then
        X = .ArrowX + (.ArrowVarX \ 32)
        Y = .ArrowY
        
        If X <= MaxMapX Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + .ArrowVarX, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If .ArrowPosition = 3 Then
        X = .ArrowX
        Y = .ArrowY - (.ArrowVarY \ 32)
                    
        If Y >= 0 Then Call DD_BackBuffer.BltFast((.ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (.ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - .ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End If
End With
End Sub

Sub ChrgSpriteSurf()
Dim i As Long
Set DD = dX.DirectDrawCreate("")
Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
' Init sprite ddsd type and load the bitmap
For i = 0 To LoadMaxSprite()
    DDSD_Character(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Character(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf(i) = LoadImage(App.Path & "\GFX\Sprites\Sprites" & i & ".png", DD, DDSD_Character(i))
    SetMaskColorFromPixel DD_SpriteSurf(i), 0, 0
Next i

End Sub

'Public Sub InitPano(ByVal mapNum As Long)
Public Sub InitPano()
    If Trim$(Map.PanoInf) <> vbNullString And InStr(1, Trim$(Map.PanoInf), ".png") > 0 Then
        If Not FileExist(App.Path & "\GFX\" & Trim$(Map.PanoInf)) Then
            Map.PanoInf = vbNullString
        Else
            'Initiialisation de la surface PanoInfSurf
            DDSD_PanoInf.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_PanoInf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_PanoInfSurf = LoadImage(App.Path & "\GFX\" & Map.PanoInf, DD, DDSD_PanoInf)
            If Map.TranInf = 1 Then SetMaskColorFromPixel DD_PanoInfSurf, 0, 0
        End If
    End If
    
     If Trim$(Map.PanoSup) <> vbNullString And InStr(1, Trim$(Map.PanoSup), ".png") > 0 Then
        If Not FileExist(App.Path & "\GFX\" & Trim$(Map.PanoSup)) Then
            Map.PanoSup = vbNullString
        Else
            'Initiialisation de la surface PanoSupSurf
            DDSD_PanoSup.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_PanoSup.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_PanoSupSurf = LoadImage(App.Path & "\GFX\" & Map.PanoSup, DD, DDSD_PanoSup)
            If Map.TranSup = 1 Then SetMaskColorFromPixel DD_PanoSupSurf, 0, 0
        End If
    End If
End Sub

Public Sub DestroyAlphaBlendDX()
    If Not Dev Is Nothing Then
        'Desactivation de l'Alpha Blending
        Dev.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
    End If
    
    'Destruction de D3D et de sont Device
    Set Dev = Nothing
    Set D3D = Nothing
    
    AlphaBlendDXIsInit = False
End Sub

Sub InitAlphaBlendDX()
    'Verification de l'initialisation de DirectX
    If DD Is Nothing Then Exit Sub
    
    'Initisalisation de D3D si il ne l'est pas
    If D3D Is Nothing Then Set D3D = DD.GetDirect3D

    'Initialisation du Device avec la surface de destination
    Set Dev = D3D.CreateDevice("IID_IDirect3DHALDevice", DD_TmpSurf)
    
    AlphaBlendDXIsInit = True
End Sub

Sub AlphaBlendDX(SSurf As DirectDrawSurface7, SSurfDesc As DDSURFACEDESC2, Srec As RECT, Drec As RECT, ByVal Alpha As Single)
    'Verification de l'initialisation de l'Alpha Blending
    If Not AlphaBlendDXIsInit Then Exit Sub

    Dev.BeginScene
        
    'Activation de l'Alpha Blending
    Dev.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
    
    ' Test
    Dev.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
    Dev.SetRenderState D3DRENDERSTATE_COLORKEYBLENDENABLE, True

    'Initialisation des parametres du Device
    Dev.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
    Dev.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
    Dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTA_TFACTOR

    'Definition de la texture utiliser par le Device
    Dev.SetTexture 0, SSurf
    Dev.SetTextureStageState 0, D3DTSS_MIPFILTER, 3

    Dim Verts(3) As D3DTLVERTEX

    'Initialisation des variable pour la taille de la texture et l'alpha
    ABDXHeight = SSurfDesc.lHeight
    ABDXWidth = SSurfDesc.lWidth
    ABDXAlpha = Alpha

    'Call SetMaskColor(SSurf, 0)

    'Initialisation du vecteur qui permet de dessiner la texture ou il faut sur la surface
    Call SetUpVerts(Verts(), Srec, Drec, Alpha)
    
    Call DD_TmpSurf.BltFast(Drec.Left, Drec.Top, DD_BackBuffer, Drec, DDBLTFAST_WAIT)

    'Dessin de la texture
    Dev.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, Verts(0), 4, D3DDP_DEFAULT
    
    Dev.EndScene
    
    Call DD_BackBuffer.BltFast(Drec.Left, Drec.Top, DD_TmpSurf, Drec, DDBLTFAST_WAIT)
End Sub

Public Sub SetUpVerts(Verts() As D3DTLVERTEX, Src As RECT, Dest As RECT, ByVal Alpha As Single)
On Error Resume Next
    'Ce sub permet d'initialiser 4 points qui forme 3 vecteurs
    'pour une texture rectangulaire
    'Les 3 vecteurs forme 2 triangles(polygones) qui forme le rectangle
    'pour la surface, ces points sont définis par plusieurs parametres,
    'leur coordones,taille et leur couleur au format RGBA si R,G ou B sont a 1
    'la surface serat dessine avec un filtre de la/les couleurs correspondants
    'A permet de regler la valeur de l'opaciter de 0 a 1(0=invisible,1=opaque)
    'par exemple si A est a 0.5 la transparence est de 50%
    
    Dim SurfW As Single
    Dim SurfH As Single
    Dim XCor As Single
    Dim YCor As Single
    Dim AlphaColor As Long
    
    'Récupération de la couleur avec l'alpha
    AlphaColor = GetColorPlusAlpha(Alpha)
        
    'Largeur de la surface
    SurfW = ABDXWidth 'Src.Right - Src.Left
    'Hauteur de la surface
    SurfH = ABDXHeight 'Src.Bottom - Src.Top

   'Coordonees du 1er point sur la surface de destination
    XCor = Dest.Left
    YCor = Dest.bottom
    
    '1er point - en bas a gauche
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    Src.Left / SurfW, _
    (Src.bottom) / SurfH, _
    Verts(0)
    
    'Coordonees du 2eme point sur la surface de destination
    XCor = Dest.Left
    YCor = Dest.Top
    
    '2eme point - en haut a gauche
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    Src.Left / SurfW, _
    Src.Top / SurfH, _
    Verts(1)
    
    'Coordonees du 3eme point sur la surface de destination
    XCor = Dest.Right
    YCor = Dest.bottom
    
    '3eme point - en bas a droite
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    (Src.Right) / SurfW, _
    (Src.bottom) / SurfH, _
    Verts(2)
    
    'Coordonees du 4eme point sur la surface de destination
    XCor = Dest.Right
    YCor = Dest.Top
    
    '4eme point - en haut a droite
    dX.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    (Src.Right) / SurfW, _
    Src.Top / SurfH, _
    Verts(3)
End Sub

Sub DD_SetD3DSprite(lX As Long, lY As Long, lW As Long, lH As Long, coloring As Long, _
                    tx As Single, ty As Single, ByVal tW As Single, ByVal tH As Single, D3DSprite() As D3DTLVERTEX)
dX.CreateD3DTLVertex lX, lY + lH, 0, 1, coloring, 0, tx, ty + tH, D3DSprite(0)
dX.CreateD3DTLVertex lX, lY, 0, 1, coloring, 0, tx, ty, D3DSprite(1)
dX.CreateD3DTLVertex lX + lW, lY + lH, 0, 1, coloring, 0, tx + tW, ty + tH, D3DSprite(2)
dX.CreateD3DTLVertex lX + lW, lY, 0, 1, coloring, 0, tx + tW, ty, D3DSprite(3)
End Sub

Public Function GetColorPlusAlpha(ByVal Alpha As Single) As Long
    'Simulation de dX.CreateColorRGBA pour R=G=B=1 et pour une valeur d'alpha
    'La fonction doit etre simuler pour la compatibiliter avec vista
    
    'Valeur par defaut
    GetColorPlusAlpha = -1
    
    'Verification du parametre
    If Alpha < 0 Or Alpha > 1 Then Exit Function
    
    If Alpha = 0 Then
        'Valeur definit pour 0
        GetColorPlusAlpha = 16777215
    ElseIf Alpha = 1 Then
        'Valeur definit pour 1
        GetColorPlusAlpha = -1
    ElseIf Alpha <= 0.5 Then
        Dim b As Currency
        Dim Tmp As Currency
        Dim n As Byte
        'On decompose le calcul pour eviter les "depassement de capaciter"
        If (Alpha * 100 - (Fix(Alpha * 10) * 10)) Mod 2 = 0 Then
            Tmp = 16777215
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) - (Fix(Alpha * 10) \ 2)
            Tmp = Tmp + (n * 33554432)
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) + (Fix(Alpha * 10) \ 2)
            b = n * 25165824
            Tmp = Tmp + (b * 2)
        Else
            Tmp = 16777215
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) - ((Fix(Alpha * 10) + 1) \ 2)
            Tmp = Tmp + (n * 33554432)
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) + ((Fix(Alpha * 10) - 1) \ 2)
            If Fix(Alpha * 10) = 0 Then n = n - 1
            b = n * 25165824
            Tmp = Tmp + (b * 2)
        End If
        GetColorPlusAlpha = CLng(Tmp)
    Else
        GetColorPlusAlpha = (GetColorPlusAlpha(0.5 - (Alpha - 0.5)) + 2) * -1
    End If
End Function

Public Function GetTileSurface(ByVal tileNum As Long)
    Set GetTileSurface = Nothing
    
    On Error GoTo AddTile
    
    Set GetTileSurface = TilesPic.item(Str(tileNum))
    
    Exit Function
AddTile:
    
    If GetTileSurface Is Nothing Then
        Set GetTileSurface = AddTile(tileNum)
    End If
End Function

Private Function AddTile(ByVal tileNum As Long)
    Dim Path As String
    Path = App.Path & "\GFX\Tiles\Tile_" & tileNum & ".png"

    If Path <> vbNullString Then

        DDSD_Temp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        TilesPic.Add LoadImage(Path, DD, DDSD_Temp), Str(tileNum)
        SetDefaultMaskColor TilesPic.item(Str(tileNum))
        Set AddTile = TilesPic.item(Str(tileNum))
    Else
        MsgBox "Erreur : Le fichier Tile_" & tileNum & ".png" & " n'existe pas..."
        End
    End If
End Function

Sub ShadePictureBox(picBox As PictureBox)
    Dim lngFor1 As Long, lngFor2 As Long
    Dim red As Byte, green As Byte, blue As Byte
    Dim moy As Integer

    Set GestionImage.PictureBox = picBox

    For lngFor1 = 0 To picBox.ScaleWidth - 1
        For lngFor2 = 0 To picBox.ScaleHeight - 1
            Call GestionImage.GetPixelRGB(lngFor1, lngFor2, red, green, blue)
            moy = 0
            moy = moy + red
            moy = moy + green
            moy = moy + blue
            moy = moy / 3
            
            red = moy
            green = moy
            blue = moy
            Call GestionImage.SetPixelRGB(lngFor1, lngFor2, red, green, blue)
        Next lngFor2
    Next lngFor1
    
    Call GestionImage.Refresh
End Sub

Public Function XMapPadding() As Long
    XMapPadding = 0
    If (MaxMapX + 1) * PIC_X < frmMirage.picScreen.Width Then
        XMapPadding = (frmMirage.picScreen.Width - (MaxMapX + 1) * PIC_X) / 2
    End If
End Function

Public Function YMapPadding() As Long
    YMapPadding = 0
    If (MaxMapY + 1) * PIC_Y < frmMirage.picScreen.Height Then
        YMapPadding = (frmMirage.picScreen.Height - (MaxMapY + 1) * PIC_Y) / 2
    End If
End Function
