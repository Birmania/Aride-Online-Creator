VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditer..."
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Editer..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   2415
   End
   Begin VB.ListBox lstIndex 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      ItemData        =   "frmIndex.frx":0000
      Left            =   240
      List            =   "frmIndex.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Choix d'Édition"
      TabPicture(0)   =   "frmIndex.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rechercher (numéros ou nom) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   2670
         Width           =   2280
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edition"
      Visible         =   0   'False
      Begin VB.Menu couper 
         Caption         =   "Couper"
         Shortcut        =   ^X
      End
      Begin VB.Menu copier 
         Caption         =   "Copier"
         Shortcut        =   ^C
      End
      Begin VB.Menu coller 
         Caption         =   "Coller"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTRLD As Boolean

Private Sub cmdOk_Click()
    EditorIndex = lstIndex.ListIndex
    
    If EditorIndex < 0 Then Exit Sub

    If InQuetesEditor = True Then
        If HORS_LIGNE = 1 Then Call QuetesEditorInit Else Call SendData("EDITQUETES" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InItemsEditor = True Then
        If HORS_LIGNE = 1 Then Call ItemEditorInit Else Call SendData("EDITITEM" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InNpcEditor = True Then
        If HORS_LIGNE = 1 Then Call NpcEditorInit Else Call SendData("EDITNPC" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InShopEditor = True Then
        If HORS_LIGNE = 1 Then Call ShopEditorInit Else Call SendData("EDITSHOP" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InSpellEditor = True Then
        If HORS_LIGNE = 1 Then Call SpellEditorInit Else Call SendData("EDITSPELL" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InCraftEditor = True Then
        If HORS_LIGNE = 1 Then Call CraftEditorInit(EditorIndex) Else Call SendData("EDITCRAFT" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InAreaEditor = True Then
        If HORS_LIGNE = 1 Then Call AreaEditorInit(EditorIndex) Else Call SendData("EDITAREA" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InDreamEditor = True Then
        If HORS_LIGNE = 1 Then Call DreamEditorInit(EditorIndex) Else Call SendData("EDITDREAM" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InEmoticonEditor = True Then
        If HORS_LIGNE = 1 Then Call EmoticonEditorInit Else Call SendData("EDITEMOTICON" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InArrowEditor = True Then
        If HORS_LIGNE = 1 Then Call ArrowEditorInit Else Call SendData("EDITARROW" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    ElseIf InPetsEditor = True Then
        If HORS_LIGNE = 1 Then Call PetEditorInit Else Call SendData("EDITPET" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If
End Sub

Private Sub cmdCancel_Click()
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InSpellEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InQuetesEditor = False
    DonID = 0
    Unload frmIndex
    frmMirage.SetFocus
End Sub

Private Sub coller_Click()
Dim FileName As String
Dim f As Long
If DonID = lstIndex.ListIndex Then Exit Sub
    If InQuetesEditor Then
        If FileExiste("quetes\quete" & DonID & ".fcq") Then Call FileCopy(App.Path & "\quetes\quete" & DonID & ".fcq", App.Path & "\quetes\quete" & lstIndex.ListIndex & ".fcq") Else Call SendSaveQuete(DonID): Call FileCopy(App.Path & "\quetes\quete" & DonID & ".fcq", App.Path & "\quetes\quete" & lstIndex.ListIndex & ".fcq")
        Call ClearQuete(lstIndex.ListIndex)
        If FileExiste("quetes\quete" & lstIndex.ListIndex & ".fcq") Then
            FileName = App.Path & "\quetes\quete" & lstIndex.ListIndex & ".fcq"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , quete(lstIndex.ListIndex)
            Close #f
        End If
        Call SendSaveQuete(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & quete(lstIndex.ListIndex).nom
    ElseIf InItemsEditor Then
        If FileExiste("items\item" & DonID & ".fco") Then Call FileCopy(App.Path & "\items\item" & DonID & ".aoo", App.Path & "\items\item" & lstIndex.ListIndex & ".aoo") Else Call SendSaveItem(DonID): Call FileCopy(App.Path & "\items\item" & DonID & ".aoo", App.Path & "\items\item" & lstIndex.ListIndex & ".aoo")
        Call ClearItem(lstIndex.ListIndex)
        If FileExiste("items\item" & lstIndex.ListIndex & ".fco") Then
            FileName = App.Path & "\items\item" & lstIndex.ListIndex & ".aoo"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Item(lstIndex.ListIndex)
            Close #f
        End If
        Call SendSaveItem(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Item(lstIndex.ListIndex).name
    ElseIf InNpcEditor Then
        If FileExiste("pnjs\npc" & DonID & ".aon") Then Call FileCopy(App.Path & "\pnjs\npc" & DonID & ".aon", App.Path & "\pnjs\npc" & lstIndex.ListIndex & ".aon") Else Call SendSaveNpc(DonID): Call FileCopy(App.Path & "\pnjs\npc" & DonID & ".aon", App.Path & "\pnjs\npc" & lstIndex.ListIndex & ".aon")
        Call ClearNpc(lstIndex.ListIndex)
        If FileExiste("pnjs\npc" & lstIndex.ListIndex & ".aon") Then
            FileName = App.Path & "\pnjs\npc" & lstIndex.ListIndex & ".aon"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Npc(lstIndex.ListIndex)
            Close #f
        End If
        Call SendSaveNpc(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Npc(lstIndex.ListIndex).name
    ElseIf InShopEditor Then
        If FileExiste("shops\shop" & DonID & ".fcm") Then Call FileCopy(App.Path & "\shops\shop" & DonID & ".fcm", App.Path & "\shops\shop" & lstIndex.ListIndex & ".fcm") Else Call SendSaveShop(DonID): Call FileCopy(App.Path & "\shops\shop" & DonID & ".fcm", App.Path & "\shops\shop" & lstIndex.ListIndex & ".fcm")
        Call ClearShop(lstIndex.ListIndex)
        If FileExiste("shops\shop" & lstIndex.ListIndex & ".fcm") Then
            FileName = App.Path & "\shops\shop" & lstIndex.ListIndex & ".fcm"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Shop(lstIndex.ListIndex)
            Close #f
        End If
        Call SendSaveShop(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Shop(lstIndex.ListIndex).name
    ElseIf InSpellEditor Then
        If FileExiste("skills\skill" & DonID & ".aos") Then Call FileCopy(App.Path & "\skills\skill" & DonID & ".aos", App.Path & "\skills\skill" & lstIndex.ListIndex & ".aos") Else Call SendSaveSpell(DonID): Call FileCopy(App.Path & "\skills\skill" & DonID & ".aos", App.Path & "\skills\skill" & lstIndex.ListIndex & ".aos")
        Call ClearSpell(lstIndex.ListIndex)
        If FileExiste("spells\spell" & lstIndex.ListIndex & ".aos") Then
            FileName = App.Path & "\skills\skill" & lstIndex.ListIndex & ".aos"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Spell(lstIndex.ListIndex)
            Close #f
        End If
        Call SendSaveSpell(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Spell(lstIndex.ListIndex).name
    ElseIf InEmoticonEditor Then
        Emoticons(lstIndex.ListIndex).Command = Trim$(Emoticons(DonID - 1).Command)
        Emoticons(lstIndex.ListIndex).Pic = Val(Emoticons(DonID - 1).Pic)
        Call SendSaveEmoticon(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Emoticons(lstIndex.ListIndex).Command
    ElseIf InArrowEditor Then
        Arrows(lstIndex.ListIndex).name = Trim$(Arrows(DonID).name)
        Arrows(lstIndex.ListIndex).Pic = Val(Arrows(DonID).Pic)
        Arrows(lstIndex.ListIndex).Range = Arrows(DonID).Range
        Call SendSaveArrow(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Arrows(lstIndex.ListIndex).name
    End If
    
    If DonTP = 1 Then
        If InQuetesEditor Then
            If FileExiste("quetes\quete" & DonID & ".fcq") Then Call Kill(App.Path & "\quetes\quete" & DonID & ".fcq")
            Call ClearQuete(DonID)
            Call SendSaveQuete(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InItemsEditor Then
            If FileExiste("items\item" & DonID & ".fco") Then Call Kill(App.Path & "\items\item" & DonID & ".aoo")
            Call ClearItem(DonID)
            Call SendSaveItem(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InNpcEditor Then
            If FileExiste("pnjs\npc" & DonID & ".aon") Then Call Kill(App.Path & "\pnjs\npc" & DonID & ".aon")
            Call ClearNpc(DonID)
            Call SendSaveNpc(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InShopEditor Then
            If FileExiste("shops\shop" & DonID & ".fcm") Then Call Kill(App.Path & "\shops\shop" & DonID & ".fcm")
            Call ClearShop(DonID)
            Call SendSaveShop(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InSpellEditor Then
            If FileExiste("skills\skill" & DonID & ".aos") Then Call Kill(App.Path & "\skills\skill" & DonID & ".aos")
            Call ClearSpell(DonID)
            Call SendSaveSpell(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InEmoticonEditor Then
            Emoticons(DonID - 1).Command = vbNullString
            Emoticons(DonID - 1).Pic = 0
            Call SendSaveEmoticon(DonID)
            lstIndex.List(DonID - 1) = DonID - 1 & " : "
        ElseIf InArrowEditor Then
            Arrows(DonID).name = vbNullString
            Arrows(DonID).Pic = 0
            Arrows(DonID).Range = 0
            Call SendSaveArrow(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InPetsEditor Then
            If FileExiste("pets\pet" & DonID & ".fcf") Then Call Kill(App.Path & "\pets\pet" & DonID & ".fcf")
            Call ClearPet(DonID)
            Call SendSavePet(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        End If
    End If
End Sub

Private Sub copier_Click()
    DonID = lstIndex.ListIndex
    DonTP = 2
End Sub

Private Sub couper_Click()
    DonID = lstIndex.ListIndex
    DonTP = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdOk_Click
If KeyCode = vbKeyControl Then CTRLD = True
If CTRLD And KeyCode = vbKeyC Then Call copier_Click
If CTRLD And KeyCode = vbKeyV Then Call coller_Click
If CTRLD And KeyCode = vbKeyX Then Call couper_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then CTRLD = False
End Sub

Private Sub lstIndex_DblClick()
Call cmdOk_Click
End Sub

Private Sub lstIndex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdOk_Click
End Sub

Private Sub lstIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If DonID > 0 Then coller.Enabled = True Else coller.Enabled = False
    Call PopupMenu(edit)
End If
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdOk_Click
End Sub

Private Sub Text1_Change()
On Error GoTo er:
If Trim$(Text1.Text) = vbNullString Then lstIndex.ListIndex = 0: Exit Sub
If IsNumeric(Text1.Text) Then
    lstIndex.ListIndex = Val(Text1.Text) - 1
Else
    Dim I As Long
    For I = 0 To lstIndex.ListCount
        If InStr(1, lstIndex.List(I), Trim$(Text1.Text)) Then lstIndex.ListIndex = I
    Next I
End If
Exit Sub
er:
MsgBox "Numéros ou Nom introuvable!", vbCritical
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Public Sub Load_Dreams()
    Dim I As Integer
    
    frmIndex.lstIndex.Clear
    ' Add the names
    For I = 0 To MAX_DREAMS
        frmIndex.lstIndex.AddItem I & " : " & Trim$(Dreams(I).name)
    Next I
    frmIndex.lstIndex.ListIndex = 0
End Sub
