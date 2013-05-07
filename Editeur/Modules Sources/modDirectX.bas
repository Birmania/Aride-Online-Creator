Attribute VB_Name = "modDirectX"
Option Explicit

'Le code utilise pour l'alpha blending est modifie a partir
'du code de Matt Hafermann(http://www.rookscape.com/vbgaming/tutBT.php)

Public Const TilesInSheets As Byte = 14
'Public ExtraSheets As Byte
Public TileFolders() As String

Public Dx As New DirectX7
Public DD As DirectDraw7

Public D3D As Direct3D7
Public Dev As Direct3DDevice7

Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DDSD_Outil As DDSURFACEDESC2
Public DD_OutilSurf As DirectDrawSurface7

Public DD_SpriteSurf() As DirectDrawSurface7
Public DDSD_Character() As DDSURFACEDESC2

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

'Public DD_FrontBuffer As DirectDrawSurface7
'Public DDSD_FrontBuffer As DDSURFACEDESC2

Public DD_SpellAnim() As DirectDrawSurface7
Public DDSD_SpellAnim() As DDSURFACEDESC2

Public DD_BigSpellAnim() As DirectDrawSurface7
Public DDSD_BigSpellAnim() As DDSURFACEDESC2

'Public DD_TileSurf() As DirectDrawSurface7
'Public DDSD_Tile() As DDSURFACEDESC2
'Public TileFile() As Boolean
'Public TilesPic() As clsTilePic
Public TilesPic As New Collection

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

'PAPERDOLL
Public DD_PaperDollSurf() As DirectDrawSurface7
Public DDSD_PaperDoll() As DDSURFACEDESC2
'FIN PAPERDOLL

Public DD_PetsSurf() As DirectDrawSurface7
Public DDSD_Pets() As DDSURFACEDESC2


'EFFET DE SANG
Public DDSD_Blood As DDSURFACEDESC2
Public DD_Blood As DirectDrawSurface7
'FIN EFFET DE SANG

Public DDSD_Temp As DDSURFACEDESC2
Public DDSD_TilesTemp As DDSURFACEDESC2
Public DD_Temp As DirectDrawSurface7

Public DDSD_PanoInf As DDSURFACEDESC2
Public DD_PanoInfSurf As DirectDrawSurface7

Public DDSD_PanoSup As DDSURFACEDESC2
Public DD_PanoSupSurf As DirectDrawSurface7

Public DDSD_Night As DDSURFACEDESC2
Public DD_NightSurf As DirectDrawSurface7
Public NightVerts(3) As D3DTLVERTEX

Public DDSD_Fog As DDSURFACEDESC2
Public DD_FogSurf As DirectDrawSurface7
Public FogVerts(3) As D3DTLVERTEX

Public DDSD_Tmp As DDSURFACEDESC2
Public DD_TmpSurf As DirectDrawSurface7

Public rec As RECT
Public rec_pos As RECT

Public AlphaBlendDXIsInit As Boolean
Public ABDXWidth As Integer
Public ABDXHeight As Integer
Public ABDXAlpha As Single

Sub InitDirectX()
    Set DD = Dx.DirectDrawCreate(vbNullString)
        
    AlphaBlendDXIsInit = False
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    '
    DDSD_Primary.lBackBufferCount = 1
    '
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hWnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
    
    'Initisalisation de D3D
    Set D3D = DD.GetDirect3D
    
'    frmMirage.Show
End Sub

Sub LoadMaxSprite()
Dim I As Long
    For I = 0 To 9999
        If Not FileExiste("\GFX\Sprites\Sprites" & I & ".png") Then
            If I < 1 Then
                MsgBox "Erreur : Sprite introuvable !"
                End
            Else
                MAX_DX_SPRITE = I - 1
            End If
            Exit Sub
        Else
            MAX_DX_SPRITE = I + 1
        End If
    Next I
End Sub

Sub LoadMaxPet()
Dim I As Long
    For I = 0 To 9999
        If Not FileExiste("\GFX\Pets\Pet" & I & ".png") Then
            If I < 1 Then
                MsgBox "Erreur : Famillier introuvable !"
                End
            Else
                MAX_DX_PETS = I - 1
            End If
            Exit Sub
        Else
            MAX_DX_PETS = I + 1
        End If
    Next I
    
End Sub

Sub LoadMaxPaperdolls()
Dim I As Long
    For I = 0 To 9999
        If Not FileExiste("\GFX\Paperdolls\Paperdolls" & I & ".png") Then
            If I < 1 Then
                MsgBox "Erreur : Paperdoll introuvable !"
                End
            Else
                MAX_DX_PAPERDOLL = I - 1
            End If
            Exit Sub
        Else
            MAX_DX_PAPERDOLL = I + 1
        End If
    Next I
End Sub

'Sub LoadMaxSpells()
'Dim i As Long
'    For i = 0 To 9999
'        If Not FileExiste("\GFX\Spells\Spells" & i & ".png") Then
'            If i < 1 Then
'                MsgBox "Erreur : Spells introuvable !"
'                End
'            Else
'                MAX_DX_SPELLS = i - 1
'            End If
'            Exit Sub
'        Else
'            MAX_DX_SPELLS = i + 1
'        End If
'    Next i
'End Sub

Sub LoadMaxBigSpells()
Dim I As Long
    For I = 0 To 9999
        If Not FileExiste("\GFX\BigSpells\BigSpells" & I & ".png") Then
            If I < 1 Then
                MsgBox "Erreur : BigSpells introuvable !"
                End
            Else
                MAX_DX_BIGSPELLS = I - 1
            End If
            Exit Sub
        Else
            MAX_DX_BIGSPELLS = I + 1
        End If
    Next I
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY
Dim I As Long

    ' Check for files existing
    If FileExiste("\GFX\items.png") = False Or FileExiste("\GFX\emoticons.png") = False Or FileExiste("\GFX\Outils.png") = False Or FileExiste("\GFX\arrows.png") = False Then Call MsgBox("Plusieurs fichiers manquants", vbOKOnly, GAME_NAME): Call GameDestroy
    
    ' Set the key for masks
    key.low = 0
    key.high = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MaxMapX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MaxMapY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
        
'    Set DD_FrontBuffer = Nothing
'
'    DDSD_FrontBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
'    DDSD_FrontBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
'    DDSD_FrontBuffer.lWidth = (MaxMapX + 1) * PIC_X
'    DDSD_FrontBuffer.lHeight = (MaxMapY + 1) * PIC_Y
'    Set DD_FrontBuffer = DD.CreateSurface(DDSD_FrontBuffer)
        
    DDSD_Outil.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Outil.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_OutilSurf = LoadImage(App.Path & "\GFX\Outils.png", DD, DDSD_Outil)
    SetMaskColorFromPixel DD_OutilSurf, 0, 0
        
    ' Init sprite ddsd type and load the bitmap
    For I = 0 To MAX_DX_SPRITE
        DDSD_Character(I).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Character(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_SpriteSurf(I) = LoadImage(App.Path & "\GFX\Sprites\Sprites" & I & ".png", DD, DDSD_Character(I))
        SetMaskColorFromPixel DD_SpriteSurf(I), 0, 0
    Next I
    
    ' Init tiles ddsd type and load the bitmap
'    For i = 0 To ExtraSheets
'        If dir$(App.Path & "\GFX\tiles" & i & ".png") <> vbNullString Then
'            DDSD_Tile(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
'            DDSD_Tile(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
'            Set DD_TileSurf(i) = LoadImage(App.Path & "\GFX\tiles" & i & ".png", DD, DDSD_Tile(i))
'            SetMaskColorFromPixel DD_TileSurf(i), 0, 0
'            TileFile(i) = True
'        Else
'            TileFile(i) = False
'        End If
'    Next i
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = LoadImage(App.Path & "\GFX\items.png", DD, DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0
       
    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = LoadImage(App.Path & "\GFX\emoticons.png", DD, DDSD_Emoticon)
    SetMaskColorFromPixel DD_EmoticonSurf, 0, 0
    
    ' Init spells ddsd type and load the bitmap
    For I = 0 To MAX_DX_SPELLS
        DDSD_SpellAnim(I).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_SpellAnim(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_SpellAnim(I) = LoadImage(App.Path & "\GFX\spells\spells" & I & ".png", DD, DDSD_SpellAnim(I))
        SetMaskColorFromPixel DD_SpellAnim(I), 0, 0
    Next I
    
    For I = 0 To MAX_DX_BIGSPELLS
        DDSD_BigSpellAnim(I).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_BigSpellAnim(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_BigSpellAnim(I) = LoadImage(App.Path & "\GFX\BigSpells\BigSpells" & I & ".png", DD, DDSD_BigSpellAnim(I))
        SetMaskColorFromPixel DD_BigSpellAnim(I), 0, 0
    Next I
    
    For I = 0 To MAX_DX_PETS
        DDSD_Pets(I).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Pets(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_PetsSurf(I) = LoadImage(App.Path & "\GFX\Pets\Pet" & I & ".png", DD, DDSD_Pets(I))
        SetMaskColorFromPixel DD_PetsSurf(I), 0, 0
    Next I
    
    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = LoadImage(App.Path & "\GFX\arrows.png", DD, DDSD_ArrowAnim)
    SetMaskColorFromPixel DD_ArrowAnim, 0, 0
    
    ' Prends la planche de sang
    DDSD_Blood.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Blood.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_Blood = LoadImage(App.Path & "\GFX\blood.png", DD, DDSD_Blood)
    SetMaskColorFromPixel DD_Blood, 0, 0

    ' Init temp ddsd type and load the bitmap
    DDSD_Temp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_Temp = LoadImage(App.Path & "\GFX\tiles0.png", DD, DDSD_Temp)
    SetMaskColorFromPixel DD_Temp, 0, 0

    'Initisialisation de la surface temporaire
    DDSD_Tmp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_Tmp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE  'Or DDSCAPS_OPTIMIZED 'DDSCAPS_TEXTURE 'DDSCAPS_OPTIMIZED
    'DDSCAPS_3DDEVICE est utilise pour pouvoir utiliser D3D sur cette surface
    DDSD_Tmp.lWidth = (MaxMapX + 1) * PIC_X
    DDSD_Tmp.lHeight = (MaxMapY + 1) * PIC_Y
    Set DD_TmpSurf = DD.CreateSurface(DDSD_Tmp)
    Call DD_TmpSurf.SetForeColor(0)
    Call DD_TmpSurf.setDrawStyle(5)
    
    'Paperdoll
    For I = 0 To MAX_DX_PAPERDOLL
        DDSD_PaperDoll(I).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_PaperDoll(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        Set DD_PaperDollSurf(I) = LoadImage(App.Path & "\GFX\Paperdolls\Paperdolls" & I & ".png", DD, DDSD_PaperDoll(I))
        SetMaskColorFromPixel DD_PaperDollSurf(I), 0, 0
    Next I
    
End Sub

Sub DestroyDirectX()
Dim I As Long

    Set Dx = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    '
    Set DD_BackBuffer = Nothing
    '
    For I = 0 To MAX_DX_SPRITE
        Set DD_SpriteSurf(I) = Nothing
    Next I
    
    Set TilesPic = Nothing
'    Erase TilesPic
'    For i = 0 To ExtraSheets
'        If TileFile(i) Then Set DD_TileSurf(i) = Nothing
'    Next i
    Set DD_ItemSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    For I = 0 To MAX_DX_SPELLS
        Set DD_SpellAnim(I) = Nothing
    Next I
    For I = 0 To MAX_DX_BIGSPELLS
        Set DD_BigSpellAnim(I) = Nothing
    Next I
    Set DD_ArrowAnim = Nothing
    Set DD_PanoInfSurf = Nothing
    Set DD_PanoSupSurf = Nothing
    Set DD_FogSurf = Nothing
    Set DD_NightSurf = Nothing
End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = DD.TestCooperativeLevel
    If (TestCoopRes = DD_OK) Then NeedToRestoreSurfaces = False Else NeedToRestoreSurfaces = True
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal x As Long, ByVal y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

With TmpR
    .Left = x
    .Top = y
    .Right = x
    .Bottom = y
End With

TheSurface.lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

With TmpColorKey
    .low = TheSurface.GetLockedPixel(x, y)
    .high = .low
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

TheSurface.Unlock TmpR
End Sub

Public Sub SetDefaultMaskColor(ByRef TheSurface As DirectDrawSurface7)
'Dim TmpR As RECT
'Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

'With TmpR
'    .Left = x
'    .Top = y
'    .Right = x
'    .Bottom = y
'End With

'TheSurface.lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

With TmpColorKey
    .low = 0
    .high = 0
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

'TheSurface.Unlock TmpR
End Sub

Sub DisplayFx(intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, tile As Long, ByRef lngDestDC As Long, ByRef lngSrcDC As Long)
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (tile - Int(tile / TilesInSheets) * TilesInSheets) * PIC_X, Int(tile / TilesInSheets) * PIC_Y, lngROP
End Sub

Sub Night(ByVal MinX As Long, ByVal MaxX As Long, ByVal MinY As Long, ByVal MaxY As Long)
Dim x As Long, y As Long
    If AccOpt.LowEffect Then
        Dim lngSrcDC As Long
        
        lngSrcDC = DD_OutilSurf.GetDC
        For y = MinY To MaxY
            For x = MinX To MaxX
                If Map(Player(MyIndex).Map).tile(x, y).Light <= 0 Then
                    DisplayFx (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31, TexthDC, lngSrcDC
                Else
                    DisplayFx (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sy - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(Player(MyIndex).Map).tile(x, y).Light, TexthDC, lngSrcDC
                End If
            Next x
        Next y
        DD_OutilSurf.ReleaseDC lngSrcDC
    Else
        'Initialisation du RECT source
        With rec_pos
            .Top = 0
            .Bottom = (MaxY - MinY + 1) * PIC_Y
            .Left = 0
            .Right = .Left + (MaxX - MinX + 1) * PIC_X
        End With
        
        If VZoom > 3 Then
            'Initialisation du RECT destination
            With rec
                .Top = (NewPlayerY * 32) + NewYOffset
                .Bottom = .Top + rec_pos.Bottom
                .Left = (NewPlayerX * 32) + NewXOffset
                .Right = .Left + (MaxX - MinX + 1) * PIC_X
            End With
        Else
            'Initialisation du RECT destination
            With rec
                .Top = -PIC_Y + (NewPlayerY * 32) + NewYOffset
                .Bottom = .Top + rec_pos.Bottom
                .Left = -PIC_X + (NewPlayerX * 32) + NewXOffset
                .Right = .Left + (MaxX - MinX + 1) * PIC_X
            End With
        End If
        
        'Dessin de la nuit
        Call AlphaBlendDX(rec_pos, rec, NightVerts)
    End If
End Sub

Sub BltWeather()
Dim I As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For I = 1 To MAX_RAINDROPS
            If DropRain(I).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then frmMirage.tmrRainDrop.Interval = 200: frmMirage.tmrRainDrop.Tag = "123"
                End If
            End If
        Next I
    ElseIf GameWeather = WEATHER_SNOWING Then
        For I = 1 To MAX_RAINDROPS
            If DropSnow(I).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then frmMirage.tmrSnowDrop.Interval = 200: frmMirage.tmrSnowDrop.Tag = "123"
                End If
            End If
        Next I
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then Call ClearRainDrop(BLT_RAIN_DROPS)
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
    
    If GameWeather = WEATHER_NONE Then Exit Sub
    
    For I = 1 To MAX_RAINDROPS
        If Not ((DropRain(I).x = 0) Or (DropRain(I).y = 0)) Then
            DropRain(I).x = DropRain(I).x + DropRain(I).speed
            DropRain(I).y = DropRain(I).y + DropRain(I).speed
            Call DD_BackBuffer.DrawLine(DropRain(I).x, DropRain(I).y, DropRain(I).x + DropRain(I).speed, DropRain(I).y + DropRain(I).speed)
            If (DropRain(I).x > (MaxMapX + 1) * PIC_X) Or (DropRain(I).y > (MaxMapY + 1) * PIC_Y) Then DropRain(I).Randomized = False
        End If
    Next I
    
'    If TileFile(ExtraSheets) Then
    rec.Top = (14 \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (14 - (14 \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
        
    For I = 1 To MAX_RAINDROPS
        If Not ((DropSnow(I).x = 0) Or (DropSnow(I).y = 0)) Then
            DropSnow(I).x = DropSnow(I).x + DropSnow(I).speed
            DropSnow(I).y = DropSnow(I).y + DropSnow(I).speed
            Call DD_BackBuffer.BltFast(DropSnow(I).x + DropSnow(I).speed, DropSnow(I).y + DropSnow(I).speed, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If (DropSnow(I).x > (MaxMapX + 1) * PIC_X) Or (DropSnow(I).y > (MaxMapY + 1) * PIC_Y) Then DropSnow(I).Randomized = False
        End If
    Next I
'    End If
        
    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)
            Call PlaySound("Thunder.wav")
            Call DD_BackBuffer.DrawBox(0, 0, (MaxMapX + 1) * PIC_X, (MaxMapY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).x = Int((((MaxMapX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MaxMapY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MaxMapY + 1) * PIC_Y / 4) And (DropRain(RDNumber).x > (MaxMapX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).x = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).x = Int((((MaxMapX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).y = Int((((MaxMapY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).y > (MaxMapY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).x > (MaxMapX + 1) * PIC_X / 4) Then GoTo Start
    DropSnow(RDNumber).speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropSnow(RDNumber).x = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltPlayerAnim(ByVal Index As Long)
Dim x As Long, y As Long
    If PlayerAnim(Index, 0) = 0 Then Exit Sub
    
    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
    
    If GetTickCount > PlayerAnim(Index, 1) + 500 Then
        PlayerAnim(Index, 2) = PlayerAnim(Index, 2) + 1
        If PlayerAnim(Index, 2) > 11 Then PlayerAnim(Index, 2) = 0
        PlayerAnim(Index, 1) = GetTickCount
    End If
    
    rec.Top = 0 * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = PlayerAnim(Index, 2) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpellAnim(PlayerAnim(Index, 0) - 1), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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

Sub BltSpell(ByVal Index As Long)
Dim x As Long, y As Long, I As Long

If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then Exit Sub
If Spell(Player(Index).SpellNum).SkillAnim < 0 Then Exit Sub

For I = 1 To MAX_SPELL_ANIM
    With Player(Index).SkillAnim(I)
    If .CastedSpell = YES Then
        If .SkillDone < Spell(Player(Index).SpellNum).SkillDone Then
            If .SpellVar > 10 Then .SkillDone = .SkillDone + 1: .SpellVar = 0
            If GetTickCount > .SkillTime + Spell(Player(Index).SpellNum).SkillTime Then .SkillTime = GetTickCount: .SpellVar = .SpellVar + 1
            
            If Spell(Player(Index).SpellNum).Big > 0 Then
                rec.Top = 0 * (PIC_Y * 2)
                rec.Bottom = rec.Top + (PIC_Y * 2)
                rec.Left = .SpellVar * (PIC_X * 2)
                rec.Right = rec.Left + (PIC_X * 2)

            Else
                rec.Top = 0 * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .SpellVar * PIC_X
                rec.Right = rec.Left + PIC_X
            End If
            
            If .TargetType = 0 Then
                If .Target > 0 Then
                    If .Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sy
                        If Spell(Player(Index).SpellNum).Big > 0 Then
                            Call DD_BackBuffer.BltFast(x - 16, y - 32, DD_BigSpellAnim(Spell(Player(Index).SpellNum).SkillAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        Else
                            Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim(Spell(Player(Index).SpellNum).SkillAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    Else
                        x = GetPlayerX(.Target) * PIC_X + sx + Player(.Target).XOffset
                        y = GetPlayerY(.Target) * PIC_Y + sy + Player(.Target).YOffset
                        
                        If Spell(Player(Index).SpellNum).Big > 0 Then
                            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX - 16, y - NewPlayerPOffsetY - 32, DD_BigSpellAnim(Spell(Player(Index).SpellNum).SkillAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        Else
                            Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpellAnim(Spell(Player(Index).SpellNum).SkillAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End If
                    End If
                End If
            Else
                x = MapNpc(.Target).x * PIC_X + sx + MapNpc(.Target).XOffset
                y = MapNpc(.Target).y * PIC_Y + sy + MapNpc(.Target).YOffset
                If Spell(Player(Index).SpellNum).Big > 0 Then
                    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX - 16, y - NewPlayerPOffsetY - 16, DD_BigSpellAnim(Spell(Player(Index).SpellNum).SkillAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    Call DD_BackBuffer.BltFast(x - NewPlayerPOffsetX, y - NewPlayerPOffsetY, DD_SpellAnim(Spell(Player(Index).SpellNum).SkillAnim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
        Else
            .CastedSpell = NO
        End If
    End If
    End With
Next I
End Sub

Sub BltBlood(ByVal Index As Long, Optional TailleX As Byte, Optional TailleY As Byte, Optional ImgTime As Byte)
Dim x As Integer, y As Integer, I As Integer, BloodNum As Byte


If IsMissing(TailleX) Or TailleX <= 0 Then TailleX = PIC_X
If IsMissing(TailleY) Or TailleY <= 0 Then TailleY = PIC_Y
If IsMissing(ImgTime) Or ImgTime <= 0 Then ImgTime = 40

BloodNum = CByte(Player(Index).BloodAnim.SkillDone)

If Player(Index).BloodAnim.CastedSpell = YES Then
    With Player(Index).BloodAnim
            If GetTickCount > .SkillTime + ImgTime Then .SkillTime = GetTickCount: .SpellVar = .SpellVar + 1
            
            rec.Top = BloodNum * TailleY
            rec.Bottom = rec.Top + TailleY
            rec.Left = .SpellVar * TailleX
            rec.Right = rec.Left + TailleX
            
            If .TargetType = 0 Then
                If .Target > 0 Then
                    If .Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sy
                        Call DD_BackBuffer.BltFast(x, y, DD_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        x = GetPlayerX(.Target) * PIC_X + sx + Player(.Target).XOffset
                        y = GetPlayerY(.Target) * PIC_Y + sy + Player(.Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset - 16, DD_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                If .Target > 0 And MapNpc(.Target).num > 0 And MapNpc(.Target).HP > 0 Then
                    If Npc(MapNpc(.Target).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(.Target).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(.Target).num).Behavior <> NPC_BEHAVIOR_QUETEUR Then
                        x = MapNpc(.Target).x * PIC_X + sx + MapNpc(.Target).XOffset
                        y = MapNpc(.Target).y * PIC_Y + sy + MapNpc(.Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset - 16, DD_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End If
            If .SpellVar = 11 Then .CastedSpell = NO: .SpellVar = 0
    End With
End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
    If Player(Index).EmoticonNum < 0 Then Exit Sub
    
    If Player(Index).EmoticonTime + 1300 > GetTickCount Then
        If GetTickCount >= Player(Index).EmoticonTime + (108 * (Player(Index).EmoticonVar + 1)) Then Player(Index).EmoticonVar = Player(Index).EmoticonVar + 1
        
        rec.Top = Player(Index).EmoticonNum * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).EmoticonVar * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If Index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sy - 32
            
            If y2 < 0 Then Exit Sub
                        
            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sy + Player(Index).YOffset - (PIC_Y * 2)
            
            If y2 < 0 Then Exit Sub
                        
            Call DD_BackBuffer.BltFast(x2 - NewPlayerPOffsetX, y2 - NewPlayerPOffsetY, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltArrow(ByVal Index As Long)
Dim x As Long, y As Long, I As Long, z As Long
Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS
    If Player(Index).Arrow(z).Arrow > 0 Then
    
        rec.Top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If GetTickCount > Player(Index).Arrow(z).ArrowTime + 30 Then
            Player(Index).Arrow(z).ArrowTime = GetTickCount
            Player(Index).Arrow(z).ArrowVarX = Player(Index).Arrow(z).ArrowVarX + 10
            Player(Index).Arrow(z).ArrowVarY = Player(Index).Arrow(z).ArrowVarY + 10
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 0 Then
            x = Player(Index).Arrow(z).ArrowX
            y = Player(Index).Arrow(z).ArrowY + Int(Player(Index).Arrow(z).ArrowVarY / 32)
            If y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then Player(Index).Arrow(z).Arrow = 0
            If y <= MaxMapY Then Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sy - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        ElseIf Player(Index).Arrow(z).ArrowPosition = 1 Then
            x = Player(Index).Arrow(z).ArrowX
            y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)
            If y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then Player(Index).Arrow(z).Arrow = 0
            If y >= 0 Then Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sy - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        ElseIf Player(Index).Arrow(z).ArrowPosition = 2 Then
            x = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
            y = Player(Index).Arrow(z).ArrowY
            If x > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then Player(Index).Arrow(z).Arrow = 0
            If x <= MaxMapX Then Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sy - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        ElseIf Player(Index).Arrow(z).ArrowPosition = 3 Then
            x = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
            y = Player(Index).Arrow(z).ArrowY
            If x < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then Player(Index).Arrow(z).Arrow = 0
            If x >= 0 Then Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sy - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If x >= 0 And x <= MaxMapX Then If y >= 0 And y <= MaxMapY Then If Map(Player(MyIndex).Map).tile(x, y).Type = TILE_TYPE_BLOCKED Or Map(Player(MyIndex).Map).tile(x, y).Type = TILE_TYPE_BLOCK_TOIT Or Map(Player(MyIndex).Map).tile(x, y).Type = TILE_TYPE_BLOCK_DIR Then Player(Index).Arrow(z).Arrow = 0
        
        For I = 1 To MAX_PLAYERS
           If IsPlaying(I) And Player(I).Map = Player(MyIndex).Map Then
                If GetPlayerX(I) = x And GetPlayerY(I) = y Then
                    If Index = MyIndex Then Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & I & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    If Index <> I Then Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next I
        
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(I).num > 0 Then
                If MapNpc(I).x = x And MapNpc(I).y = y Then
                    If Index = MyIndex Then Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & I & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next I
    End If
Next z
End Sub

Public Sub InitPano(ByVal mapNum As Long)
    If mapNum <= 0 Or mapNum > MAX_MAPS Or DD Is Nothing Then Exit Sub
    If Trim$(Map(mapNum).PanoInf) <> vbNullString And InStr(1, Trim$(Map(mapNum).PanoInf), ".png") > 0 Then
        If Not FileExiste("GFX\" & Trim$(Map(mapNum).PanoInf)) Then
            Map(mapNum).PanoInf = vbNullString
        Else
            'Initialisation de la surface PanoInfSurf
            DDSD_PanoInf.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_PanoInf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_PanoInfSurf = LoadImage(App.Path & "\GFX\" & Map(mapNum).PanoInf, DD, DDSD_PanoInf)
            If Map(mapNum).TranInf = 1 Then SetMaskColorFromPixel DD_PanoInfSurf, 0, 0
        End If
    End If
    
    If Trim$(Map(mapNum).PanoSup) <> vbNullString And InStr(1, Trim$(Map(mapNum).PanoSup), ".png") > 0 Then
        If Not FileExiste("GFX\" & Trim$(Map(mapNum).PanoSup)) Then
            Map(mapNum).PanoSup = vbNullString
        Else
            'Initialisation de la surface PanoSupSurf
            DDSD_PanoSup.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_PanoSup.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_PanoSupSurf = LoadImage(App.Path & "\GFX\" & Map(mapNum).PanoSup, DD, DDSD_PanoSup)
            If Map(mapNum).TranSup = 1 Then SetMaskColorFromPixel DD_PanoSupSurf, 0, 0
        End If
    End If
End Sub

Public Sub InitNightAndFog(ByVal mapNum As Long)
    If mapNum <= 0 Or mapNum > MAX_MAPS Or DD Is Nothing Then Exit Sub
    
    'Initialisation du brouillard
    If Map(mapNum).Fog <> 0 Then
        'Initialisation de la texture pour le brouillard si il y en a un
        DDSD_Fog.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        DDSD_Fog.ddsCaps.lCaps = DDSCAPS_TEXTURE
        DDSD_Fog.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        Set DD_FogSurf = LoadImage(App.Path & "\GFX\fog" & Map(mapNum).Fog & ".png", DD, DDSD_Fog)
        
        'Initialisation du RECT de la texture
        With rec
            .Top = 0
            .Bottom = DDSD_Fog.lHeight
            .Left = 0
            .Right = DDSD_Fog.lWidth
        End With
        
        'Initialisation de l'Alpha Blending
        Call InitAlphaBlendDX(DD_TmpSurf, DD_FogSurf, rec, FogVerts(), Map(mapNum).FogAlpha / 100)
    Else
        Set DD_FogSurf = Nothing
    End If
    
    'Initialisation de la nuit
    If GameTime = TIME_NIGHT Or (frmMirage.nuitjour.Checked And InEditor) Then
        Dim x As Long
        Dim y As Long
        Dim tile As Long
                
        'Initialisation de la texture pour la nuit
        DDSD_Night.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        DDSD_Night.ddsCaps.lCaps = DDSCAPS_TEXTURE
        DDSD_Night.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        DDSD_Night.lWidth = (MaxMapX + 1) * PIC_X
        DDSD_Night.lHeight = (MaxMapY + 1) * PIC_Y
        Set DD_NightSurf = DD.CreateSurface(DDSD_Night)
        
        'Dessin de l'effet de la nuit en "low effect"
        tile = 31
        With rec
            .Top = (tile \ TilesInSheets) * PIC_Y
            .Bottom = .Top + PIC_X
            .Left = (tile - (tile \ TilesInSheets) * TilesInSheets) * PIC_X
            .Right = .Left + PIC_Y
        End With
        DD_NightSurf.BltFast 0, 0, DD_OutilSurf, rec, DDBLTFAST_WAIT
        
        'Récupération de la couleur de la nuit en "low effect" en changeant le colokey du Tiles ou sont les effets
        SetMaskColorFromPixel DD_OutilSurf, (tile - (tile \ TilesInSheets) * TilesInSheets) * PIC_X, (tile \ TilesInSheets) * PIC_Y
        
        'Remplissage de la surface en noir
        With rec
            .Top = 0
            .Bottom = (MaxMapX + 1) * PIC_X
            .Left = 0
            .Right = (MaxMapY + 1) * PIC_Y
        End With
        Call DD_NightSurf.BltColorFill(rec, 0)
        
        'Dessin des lumiéres
        For y = 0 To MaxMapY
            For x = 0 To MaxMapX
                If Map(Player(MyIndex).Map).tile(x, y).Light > 0 Then
                    tile = Map(Player(MyIndex).Map).tile(x, y).Light
                    With rec
                        .Top = (tile \ TilesInSheets) * PIC_Y
                        .Bottom = .Top + PIC_X
                        .Left = (tile - (tile \ TilesInSheets) * TilesInSheets) * PIC_X
                        .Right = .Left + PIC_Y
                    End With
                    DD_NightSurf.BltFast x * PIC_X, y * PIC_Y, DD_OutilSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            Next x
        Next y
        
        'Remettre la colorkey de base
        SetMaskColorFromPixel DD_OutilSurf, 0, 0
                
        'Initialisation du RECT de la texture
        With rec
            .Top = 0
            .Bottom = DDSD_Night.lHeight
            .Left = 0
            .Right = DDSD_Night.lWidth
        End With
        
        'Initialisation de l'Alpha Blending
        Call InitAlphaBlendDX(DD_TmpSurf, DD_NightSurf, rec, NightVerts(), 0.5)
    Else
        Set DD_NightSurf = Nothing
    End If
End Sub

Public Sub DestroyAlphaBlendDX()
    'Desactivation de l'Alpha Blending
    Dev.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
    
    'Destruction de D3D et de sont Device
    Set Dev = Nothing
    Set D3D = Nothing
    
    AlphaBlendDXIsInit = False
End Sub

Public Sub InitAlphaBlendDX(DSurf As DirectDrawSurface7, TSurf As DirectDrawSurface7, Vrec As RECT, DVerts() As D3DTLVERTEX, ByVal Alpha As Single)
    'Verification de l'initialisation de DirectX
    If DD Is Nothing Then Exit Sub
    
    'Verification de la surface
    If DSurf Is Nothing Then Exit Sub
    
    'Verification de l'alpha
    If Alpha < 0 Or Alpha > 1 Then Alpha = 0.5
    
    'Initisalisation de D3D si il ne l'est pas
    If D3D Is Nothing Then Set D3D = DD.GetDirect3D
    
    'Initialisation du Device avec la surface de destination
    Set Dev = D3D.CreateDevice("IID_IDirect3DHALDevice", DSurf)
    
    'Activation de l'Alpha Blending
    Dev.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
    
    'Initialisation des parametres du Device
    Dev.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
    Dev.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
    Dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTA_TFACTOR
    
    'Definition de la texture utiliser par le Device
    Dev.SetTexture 0, TSurf
    Dev.SetTextureStageState 0, D3DTSS_MIPFILTER, 3
    
    'Initialisation des variable pour la taille de la texture et l'alpha
    ABDXHeight = Vrec.Bottom
    ABDXWidth = Vrec.Right
    ABDXAlpha = Alpha
    
    'Initialisation du vecteur qui permet de dessiner la texture ou il faut sur la surface
    Call SetUpVerts(DVerts(), Vrec, Vrec, Alpha)
    
    AlphaBlendDXIsInit = True
End Sub

Public Sub AlphaBlendDX(Srec As RECT, Drec As RECT, DVerts() As D3DTLVERTEX)
    'Verification de l'initialisation de l'Alpha Blending
    If Not AlphaBlendDXIsInit Then Exit Sub

    'Initialisation du vecteur qui permet de dessiner la texture ou il faut sur la surface
    Call SetUpVerts(DVerts(), Drec, Srec, ABDXAlpha)
    
    'Debute la scene D3D (obligatoire avant tous dessin)
    Dev.BeginScene

    'Récupération du contenu du BackBuffer
    Call DD_TmpSurf.BltFast(Srec.Left, Srec.Top, DD_BackBuffer, Srec, DDBLTFAST_WAIT)
    
    'Dessin de la texture
    Dev.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, DVerts(0), 4, D3DDP_DEFAULT
    
    'Fin de la scene D3D (obligatoire apres tous dessin)
    Dev.EndScene
    
    Sleep 15
    
    'Dessin du résultat sur le BackBuffer
    Call DD_BackBuffer.BltFast(Srec.Left, Srec.Top, DD_TmpSurf, Srec, DDBLTFAST_WAIT)
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
    
    'Dest.Left = Dest.Left
    'Dest.Top = Dest.Top
    'Dest.Bottom = Dest.Bottom
    'Dest.Right = Dest.Right
   'Coordonees du 1er point sur la surface de destination
    XCor = Dest.Left
    YCor = Dest.Bottom
    
    '1er point - en bas a gauche
    Dx.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    Src.Left / SurfW, _
    (Src.Bottom) / SurfH, _
    Verts(0)
    
    'Coordonees du 2eme point sur la surface de destination
    XCor = Dest.Left
    YCor = Dest.Top
    
    '2eme point - en haut a gauche
    Dx.CreateD3DTLVertex _
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
    YCor = Dest.Bottom
    
    '3eme point - en bas a droite
    Dx.CreateD3DTLVertex _
    XCor, _
    YCor, _
    0, _
    1, _
    AlphaColor, _
    0, _
    (Src.Right) / SurfW, _
    (Src.Bottom) / SurfH, _
    Verts(2)
    
    'Coordonees du 4eme point sur la surface de destination
    XCor = Dest.Right
    YCor = Dest.Top
    
    '4eme point - en haut a droite
    Dx.CreateD3DTLVertex _
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
                    tx As Single, ty As Single, ByVal tW As Single, ByVal th As Single, D3DSprite() As D3DTLVERTEX)
Dx.CreateD3DTLVertex lX, lY + lH, 0, 1, coloring, 0, tx, ty + th, D3DSprite(0)
Dx.CreateD3DTLVertex lX, lY, 0, 1, coloring, 0, tx, ty, D3DSprite(1)
Dx.CreateD3DTLVertex lX + lW, lY + lH, 0, 1, coloring, 0, tx + tW, ty + th, D3DSprite(2)
Dx.CreateD3DTLVertex lX + lW, lY, 0, 1, coloring, 0, tx + tW, ty, D3DSprite(3)
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
        Dim tmp As Currency
        Dim n As Byte
        'On decompose le calcul pour eviter les "depassement de capaciter"
        If (Alpha * 100 - (Fix(Alpha * 10) * 10)) Mod 2 = 0 Then
            tmp = 16777215
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) - (Fix(Alpha * 10) \ 2)
            tmp = tmp + (n * 33554432)
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) + (Fix(Alpha * 10) \ 2)
            b = n * 25165824
            tmp = tmp + (b * 2)
        Else
            tmp = 16777215
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) - ((Fix(Alpha * 10) + 1) \ 2)
            tmp = tmp + (n * 33554432)
            
            n = (Alpha * 100) - ((Alpha * 100) \ 2) + ((Fix(Alpha * 10) - 1) \ 2)
            If Fix(Alpha * 10) = 0 Then n = n - 1
            b = n * 25165824
            tmp = tmp + (b * 2)
        End If
        GetColorPlusAlpha = CLng(tmp)
    Else
        GetColorPlusAlpha = (GetColorPlusAlpha(0.5 - (Alpha - 0.5)) + 2) * -1
    End If
End Function

Public Function GetTileSurface(ByVal tileNum As Long)
    Set GetTileSurface = Nothing

'    Dim i As Integer
'
'    For i = 0 To GetArraySize(TilesPic) - 1
'        If TilesPic(i).Mapping = tileNum Then
'            Set GetTileSurface = TilesPic(i).DD_Tile
'            Exit For
'        End If
'    Next i
'
'    If GetTileSurface Is Nothing Then
'        Set GetTileSurface = AddTile(tileNum)
'    End If
    On Error GoTo AddTile
    
    Set GetTileSurface = TilesPic.Item(Str(tileNum))
    
    Exit Function
AddTile:
    
    If GetTileSurface Is Nothing Then
        Set GetTileSurface = AddTile(tileNum)
    End If
End Function

Private Function AddTile(ByVal tileNum As Long)
'If GetTileSurface(tileNum) Is Nothing Then
    Dim Path As String
    Path = GetPathOfFileIn("Tile_" & tileNum & ".png", App.Path & "\GFX\Tiles")

    If Path <> vbNullString Then
'        ReDim Preserve TilesPic(0 To GetArraySize(TilesPic))
'        Set TilesPic(GetArraySize(TilesPic) - 1) = New clsTilePic
'        TilesPic(GetArraySize(TilesPic) - 1).Mapping = tileNum
'
'        DDSD_TilesTemp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
'        DDSD_TilesTemp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
'        Set TilesPic(GetArraySize(TilesPic) - 1).DD_Tile = LoadImage(Path, DD, DDSD_TilesTemp)
'        'SetMaskColorFromPixel TilesPic(GetArraySize(TilesPic) - 1).DD_Tile, 0, 0
'        SetDefaultMaskColor TilesPic(GetArraySize(TilesPic) - 1).DD_Tile
'
'        Set AddTile = TilesPic(GetArraySize(TilesPic) - 1).DD_Tile
        'ReDim Preserve TilesPic(0 To GetArraySize(TilesPic))
        'Set TilesPic(GetArraySize(TilesPic) - 1) = New clsTilePic
        'TilesPic(GetArraySize(TilesPic) - 1).Mapping = tileNum
    
        DDSD_Temp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
'        Set TilesPic(GetArraySize(TilesPic) - 1).DD_Tile = LoadImage(Path, DD, DDSD_Temp)
        TilesPic.add LoadImage(Path, DD, DDSD_Temp), Str(tileNum)
        'SetMaskColorFromPixel TilesPic(GetArraySize(TilesPic) - 1).DD_Tile, 0, 0
'        SetDefaultMaskColor TilesPic(GetArraySize(TilesPic) - 1).DD_Tile
        SetDefaultMaskColor TilesPic.Item(Str(tileNum))
        
'        Set AddTile = TilesPic(GetArraySize(TilesPic) - 1).DD_Tile
        Set AddTile = TilesPic.Item(Str(tileNum))
    Else
        MsgBox "Erreur : Le fichier Tile_" & tileNum & ".png" & " n'existe pas..."
    End If
'End If
End Function
