VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainMenu 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Principal"
   ClientHeight    =   7200
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ForeColor       =   &H000000FF&
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Palette         =   "frmMainMenu.frx":000C
   Picture         =   "frmMainMenu.frx":4ACD
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgl 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   65280
      _Version        =   393216
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Frame1"
      Height          =   2265
      Left            =   6120
      TabIndex        =   11
      Top             =   480
      Width           =   3390
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Musique"
         Height          =   195
         Left            =   2160
         TabIndex        =   18
         Top             =   1680
         Width           =   180
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   720
         Width           =   3075
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   3075
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Save Password"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   195
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Musique"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label picConnect 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "                             "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memoriser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Image imgLogin 
         Height          =   2265
         Left            =   0
         MousePointer    =   5  'Size
         Picture         =   "frmMainMenu.frx":130B11
         Top             =   0
         Width           =   3390
      End
   End
   Begin VB.Timer splash 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7200
      Top             =   0
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   0
   End
   Begin VB.CheckBox chk_fullscreen 
      BackColor       =   &H80000009&
      Caption         =   "Plein ecran"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Frame fraPers 
      Caption         =   "Frame1"
      Height          =   2550
      Left            =   3000
      TabIndex        =   13
      Top             =   1320
      Width           =   3385
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   930
         ItemData        =   "frmMainMenu.frx":15208D
         Left            =   120
         List            =   "frmMainMenu.frx":15208F
         TabIndex        =   5
         Top             =   800
         Width           =   3105
      End
      Begin VB.Label picUseChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   435
         Left            =   1440
         TabIndex        =   7
         Top             =   2160
         Width           =   1845
      End
      Begin VB.Label picDelChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label picNewChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1800
         Width           =   3285
      End
      Begin VB.Label picCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   0
         Width           =   420
      End
      Begin VB.Image imgPers 
         Height          =   2550
         Left            =   0
         MousePointer    =   5  'Size
         Picture         =   "frmMainMenu.frx":152091
         Top             =   0
         Width           =   3390
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Plein ecran"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label versionlbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   690
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   6840
      Width           =   1305
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Public animi As Long
Public DragImg As Long
Public DragX As Long
Public DragY As Long
Private twippx As Long
Private twippy As Long

Public Function getreselotionX()
    getreselotionX = Screen.Width \ Screen.TwipsPerPixelX
End Function

Public Function getreselotionY()
    getreselotionY = Screen.Height \ Screen.TwipsPerPixelY
End Function

Private Sub Check1_Click()
If Check1.Value = "0" Then StopMidi Else If FileExist(App.Path & "\Music\mainmenu.mid") Then Call PlayMidi(App.Path & "\Music\mainmenu.mid") Else Call PlayMidi(App.Path & "\Music\mainmenu.mp3")

Call WriteINI("CONFIG", "Music", Str$(Check1.Value), ClientConfigurationFile)
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    Call SetIcon(Me)

    Call InitDirectX
    
    Check1.Value = Val(ReadINI("CONFIG", "Music", ClientConfigurationFile))
    
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
        If FileExist(App.Path & Rep_Theme & "\Login\connexion" & Ending) Then imgLogin.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\connexion" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Login\personnage" & Ending) Then imgPers.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\personnage" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Login\fond" & Ending) Then Me.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\fond" & Ending)
    Next i
        
    If Check1.Value = 1 Then If FileExist(App.Path & "\Music\mainmenu.mid") Then Call PlayMidi(App.Path & "\Music\mainmenu.mid") Else Call PlayMidi(App.Path & "\Music\mainmenu.mp3")
            
    'Picsprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png", True)
    
    fraPers.Visible = False
    txtName.Text = Trim$(ReadINI("INFO", "Account", ClientConfigurationFile))
    txtPassword.Text = Trim$(ReadINI("INFO", "Password", ClientConfigurationFile))
    
    If Trim$(txtPassword.Text) <> vbNullString Then Check2.Value = Checked Else Check2.Value = Unchecked
    
    fraLogin.Visible = True
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
        
    twippy = Screen.TwipsPerPixelY
    twippx = Screen.TwipsPerPixelX
    
    versionlbl.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision

    fraLogin.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 1
DragX = X
DragY = Y
End Sub

Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragImg = 1 Then fraLogin.Top = fraLogin.Top + ((Y / twippy) - (DragY / twippy)): fraLogin.Left = fraLogin.Left + ((X / twippx) - (DragX / twippx))
End Sub

Private Sub imgLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub imgPers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 3
    DragX = X
    DragY = Y
End Sub

Private Sub imgPers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragImg = 3 Then fraPers.Top = fraPers.Top + ((Y / twippy) - (DragY / twippy)): fraPers.Left = fraPers.Left + ((X / twippx) - (DragX / twippx))
End Sub

Private Sub imgPers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub lstChars_DblClick()
Call picUseChar_Click
End Sub

Private Sub lstChars_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call picUseChar_Click: KeyAscii = 0
End Sub

Private Sub picCancel_Click()
    Dim i As Integer
    For i = 1 To MAX_INV - 1
        Unload frmMirage.picInv(i)
    Next
    Call TcpDestroy
    fraLogin.Visible = True
    fraPers.Visible = False
End Sub

Private Sub picConnect_Click()
    Call GameInit

    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        If Len(Trim$(txtName.Text)) < 3 Or Len(Trim$(txtPassword.Text)) < 3 Then MsgBox "Votre nom et votre mot de passe doivent contenir plus de 3 caractéres": Exit Sub
        Call MenuState(MENU_STATE_LOGIN)
        Call WriteINI("INFO", "Account", txtName.Text, (ClientConfigurationFile))
        If Check2.Value = Checked Then Call WriteINI("INFO", "Password", txtPassword.Text, (ClientConfigurationFile)) Else Call WriteINI("INFO", "Password", "", (ClientConfigurationFile))
    End If
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then MsgBox "Il n'y a pas de personnage à cette emplacement!": Exit Sub

    Value = MsgBox("Es-tu certains de vouloir éffacer ce personnage?", vbYesNo, Game_Name)
    
    If Value = vbYes Then Call MenuState(MENU_STATE_DELCHAR)
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then MsgBox "Il n'y a pas de personnage à cette emplacement!": Exit Sub
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Public Sub ChangeScreenSettings(lWidth As Integer, lHeight As Integer, lColors As Integer)
Dim tDevMode As DEVMODE, lTemp As Long, lIndex As Long

lIndex = 0

Do
    lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
    If lTemp = 0 Then Exit Do
    lIndex = lIndex + 1
    With tDevMode
        If .dmPelsWidth = lWidth And .dmPelsHeight = lHeight And .dmBitsPerPel = lColors Then lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY): Exit Do
    End With
Loop

End Sub

Private Sub splash_Timer()
frmsplash.Visible = False
splash.Enabled = False
End Sub

Private Sub tmr2_Timer()
If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 0 Then
    frmMirage.BorderStyle = 3
    frmMirage.WindowState = 0
    'frmMirage.StartUpPosition = 1
End If
If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 1 Then
    frmMirage.BorderStyle = 0
    frmMirage.WindowState = 2
    'frmMirage.StartUpPosition = 2
End If
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picConnect_Click
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picConnect_Click
End Sub
