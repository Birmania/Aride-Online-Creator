VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{984C2AF5-D8F3-11D7-8532-00E07DD46690}#1.0#0"; "HookMenuPlus.ocx"
Begin VB.Form frmMirage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   10260
   ClientLeft      =   2805
   ClientTop       =   -345
   ClientWidth     =   15270
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMirage.frx":17D2A
   ScaleHeight     =   684
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox Canal 
      Height          =   315
      ItemData        =   "frmMirage.frx":17E7C
      Left            =   3600
      List            =   "frmMirage.frx":17E8C
      TabIndex        =   140
      Top             =   9840
      Width           =   1215
   End
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   3600
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Description-"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   142
         ToolTipText     =   "Force/d�fense/vitesse requise pour �quipper l'objet"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Donne-"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   141
         ToolTipText     =   "Force/d�fense/vitesse requise pour �quipper l'objet"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Usure 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Usure : XXXX/XXXX"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   139
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Magi: XXXXX Speed: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   14
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Description de l'objet"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   12
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Requi�re-"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Force/d�fense/vitesse requise pour �quipper l'objet"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Nom de l'objet"
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picpics 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   900
      ScaleHeight     =   2625
      ScaleWidth      =   2505
      TabIndex        =   62
      Top             =   4620
      Visible         =   0   'False
      Width           =   2535
      Begin VB.PictureBox tmpsquete 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   720
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   63
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label minute 
            BackStyle       =   0  'Transparent
            Caption         =   "00:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   65
            ToolTipText     =   "Minutes restante avant la fin de la qu�te en cour"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label seconde 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   64
            ToolTipText     =   "Secondes restante avant la fin de la qu�te en cour"
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.PictureBox vieetc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   2385
         TabIndex        =   107
         Top             =   0
         Visible         =   0   'False
         Width           =   2385
         Begin VB.Label lblStats 
            Caption         =   "EXP :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   146
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblStats 
            Caption         =   "PM :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   145
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblStats 
            Caption         =   "PV :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   144
            Top             =   120
            Width           =   375
         End
         Begin VB.Label monnom 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "nom"
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
            Left            =   2040
            TabIndex        =   122
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label maclasse 
            BackStyle       =   0  'Transparent
            Caption         =   "classe"
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
            Left            =   2040
            TabIndex        =   121
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lblPoints 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "point"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   120
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "niv"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   119
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label lblSPEED 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "vitese"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   118
            Top             =   1800
            Width           =   465
         End
         Begin VB.Label lblMAGI 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "magi"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   117
            Top             =   1560
            Width           =   360
         End
         Begin VB.Label lblDEF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "def"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   116
            Top             =   2040
            Width           =   225
         End
         Begin VB.Label lblSTR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "force"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   115
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblEXP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   480
            TabIndex        =   114
            Top             =   600
            Width           =   1890
         End
         Begin VB.Shape Shape1 
            Height          =   180
            Left            =   480
            Top             =   600
            Width           =   1890
         End
         Begin VB.Label lblMP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CB884B&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   480
            TabIndex        =   113
            Top             =   360
            Width           =   1890
         End
         Begin VB.Label lblHP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   480
            TabIndex        =   112
            Top             =   120
            Width           =   1890
         End
         Begin VB.Label AddDef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2040
            TabIndex        =   111
            Top             =   2040
            Width           =   165
         End
         Begin VB.Label AddMagi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2040
            TabIndex        =   110
            Top             =   1560
            Width           =   165
         End
         Begin VB.Label AddSpeed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2040
            TabIndex        =   109
            Top             =   1800
            Width           =   165
         End
         Begin VB.Label AddStr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2040
            TabIndex        =   108
            Top             =   1320
            Width           =   165
         End
         Begin VB.Shape shpTNL 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   180
            Left            =   480
            Top             =   600
            Width           =   1905
         End
         Begin VB.Shape shpMP 
            BackColor       =   &H00CB884B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   480
            Top             =   360
            Width           =   1905
         End
         Begin VB.Shape shpHP 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   480
            Top             =   120
            Width           =   1905
         End
      End
      Begin VB.PictureBox picWhosOnline 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   2385
         TabIndex        =   123
         Top             =   0
         Visible         =   0   'False
         Width           =   2385
         Begin VB.ListBox lstOnline 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            ItemData        =   "frmMirage.frx":17EAC
            Left            =   0
            List            =   "frmMirage.frx":17EAE
            TabIndex        =   124
            Top             =   60
            Width           =   2350
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   2385
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   2385
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nom de la Guilde:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   240
            TabIndex        =   97
            Top             =   645
            Width           =   1050
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Votre access:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   480
            TabIndex        =   96
            Top             =   960
            Width           =   825
         End
         Begin VB.Label lblGuild 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1425
            TabIndex        =   95
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label lblRank 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rank"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1425
            TabIndex        =   94
            Top             =   975
            Width           =   1080
         End
         Begin VB.Label cmdLeave 
            BackStyle       =   0  'Transparent
            Caption         =   "Quitter la Guilde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   720
            TabIndex        =   93
            Top             =   2280
            Width           =   1110
         End
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   2385
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   2385
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1080
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   90
            Top             =   120
            Width           =   555
            Begin VB.PictureBox HelmetImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   91
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1680
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   88
            Top             =   720
            Width           =   555
            Begin VB.PictureBox ShieldImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   89
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1080
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   86
            Top             =   720
            Width           =   555
            Begin VB.PictureBox ArmorImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   87
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   480
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   84
            Top             =   720
            Width           =   555
            Begin VB.PictureBox WeaponImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   85
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1080
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   82
            Top             =   1320
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox LegsImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   83
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1680
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   80
            Top             =   1920
            Width           =   555
            Begin VB.PictureBox PetImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   81
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1680
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   78
            Top             =   1320
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox Ring2Image 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   79
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   480
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   76
            Top             =   1320
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox Ring1Image 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   77
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   480
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   74
            Top             =   1920
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox GlovesImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   75
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox AmuletImage2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1680
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   72
            Top             =   120
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox AmuletImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   73
               Top             =   0
               Width           =   495
            End
         End
      End
      Begin VB.PictureBox picPlayerSpells 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   167
         TabIndex        =   134
         Top             =   -60
         Visible         =   0   'False
         Width           =   2505
         Begin VB.ListBox lstSpells 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   2190
            ItemData        =   "frmMirage.frx":17EB0
            Left            =   60
            List            =   "frmMirage.frx":17EB2
            TabIndex        =   135
            Top             =   60
            Width           =   2325
         End
         Begin VB.Label lblCast 
            BackStyle       =   0  'Transparent
            Caption         =   "Lancer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   136
            Top             =   2280
            Width           =   375
         End
      End
      Begin VB.PictureBox picGuildAdmin 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   2385
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   2385
         Begin VB.CommandButton cmdAccess 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Changer l'Access"
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
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   1980
            Width           =   1815
         End
         Begin VB.CommandButton cmdDisown 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Faire quitter la Guilde"
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
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   1650
            Width           =   1815
         End
         Begin VB.CommandButton cmdMember 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Faire entrer dans la Guild"
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
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   1305
            Width           =   1815
         End
         Begin VB.CommandButton cmdTrainee 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Faire entrainer"
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
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   975
            Width           =   1815
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   720
            TabIndex        =   127
            Top             =   345
            Width           =   1575
         End
         Begin VB.TextBox txtAccess 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   720
            MaxLength       =   2
            TabIndex        =   126
            Top             =   585
            Width           =   1575
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nom:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   180
            TabIndex        =   133
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   150
            TabIndex        =   132
            Top             =   615
            Width           =   465
         End
      End
      Begin VB.PictureBox picquete 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   0
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox quetetxt 
            Appearance      =   0  'Flat
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Text            =   "frmMirage.frx":17EB4
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label artquete 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arreter la quete"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   720
            TabIndex        =   70
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label qt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   120
            TabIndex        =   69
            Top             =   2040
            Width           =   45
         End
         Begin VB.Label av 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   1200
            TabIndex        =   68
            Top             =   2040
            Width           =   45
         End
      End
      Begin VB.PictureBox picInv3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   0
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   163
         TabIndex        =   98
         Top             =   0
         Visible         =   0   'False
         Width           =   2440
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   0
            ScaleHeight     =   145
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   177
            TabIndex        =   102
            Top             =   0
            Width           =   2655
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   3735
               Left            =   0
               ScaleHeight     =   249
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   177
               TabIndex        =   103
               Top             =   0
               Width           =   2655
               Begin VB.PictureBox picInv 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   480
                  Index           =   0
                  Left            =   120
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   104
                  Top             =   120
                  Width           =   480
               End
               Begin VB.Shape EquipS 
                  BorderColor     =   &H0000FFFF&
                  BorderWidth     =   3
                  Height          =   540
                  Index           =   4
                  Left            =   0
                  Top             =   0
                  Width           =   540
               End
               Begin VB.Shape EquipS 
                  BorderColor     =   &H0000FFFF&
                  BorderWidth     =   3
                  Height          =   540
                  Index           =   0
                  Left            =   0
                  Top             =   0
                  Width           =   540
               End
               Begin VB.Shape EquipS 
                  BorderColor     =   &H0000FFFF&
                  BorderWidth     =   3
                  Height          =   540
                  Index           =   1
                  Left            =   -360
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Shape EquipS 
                  BorderColor     =   &H0000FFFF&
                  BorderWidth     =   3
                  Height          =   540
                  Index           =   2
                  Left            =   0
                  Top             =   120
                  Width           =   540
               End
               Begin VB.Shape EquipS 
                  BorderColor     =   &H0000FFFF&
                  BorderWidth     =   3
                  Height          =   540
                  Index           =   3
                  Left            =   0
                  Top             =   0
                  Width           =   540
               End
               Begin VB.Shape SelectedItem 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  Height          =   525
                  Left            =   105
                  Top             =   105
                  Width           =   525
               End
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   330
            Left            =   2640
            Max             =   100
            TabIndex        =   101
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox Down 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1365
            Picture         =   "frmMirage.frx":17EBA
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   100
            Top             =   2235
            Width           =   270
         End
         Begin VB.PictureBox Up 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   975
            Picture         =   "frmMirage.frx":18145
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   99
            Top             =   2235
            Width           =   270
         End
         Begin VB.Label lblUseItem 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Utiliser"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   15
            TabIndex        =   106
            Top             =   2265
            Width           =   690
         End
         Begin VB.Label lblDropItem 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Jeter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1830
            TabIndex        =   105
            Top             =   2265
            Width           =   795
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   160
            Y1              =   144
            Y2              =   144
         End
         Begin VB.Line Line2 
            X1              =   4
            X2              =   171
            Y1              =   144
            Y2              =   144
         End
      End
   End
   Begin VB.PictureBox Attributs 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6810
      Left            =   0
      ScaleHeight     =   454
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Maintenir le click gauche  pour d�placer,click droit pour position de default."
      Top             =   480
      Visible         =   0   'False
      Width           =   3540
      Begin VB.OptionButton optMapBorder 
         Caption         =   "Bord map"
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
         Left            =   1560
         TabIndex        =   147
         TabStop         =   0   'False
         ToolTipText     =   "Case o� le joueur doit marcher pour ouvrit la fen�tre de la banque"
         Top             =   3480
         Width           =   975
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Bloquer PNJ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Bloque seulement les PNJ(personnage non joueur)"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optBDir 
         Caption         =   "Bloque Direction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   138
         TabStop         =   0   'False
         ToolTipText     =   "Bloque seulement une ou plusieurs directions"
         Top             =   1800
         Width           =   1410
      End
      Begin VB.OptionButton optbtoit 
         Caption         =   "Bloque Toit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   137
         TabStop         =   0   'False
         ToolTipText     =   "Bloque le joueur mais garde les caract�ristique de l'attribut Toit"
         Top             =   3120
         Width           =   1170
      End
      Begin VB.OptionButton optBguilde 
         Caption         =   "Bloquer Guilde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Bloque seulement les joueurs qui ont une monture"
         Top             =   1560
         Width           =   1395
      End
      Begin VB.OptionButton opttoit 
         Caption         =   "Toit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Quand le jouer marche sur une case Toit toutes les couches frange 1,2 et 3 qui sont sur une case Toit autour de lui disparaisse"
         Top             =   2880
         Width           =   1170
      End
      Begin VB.OptionButton optBniv 
         Caption         =   "Bloquer Niv."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Bloque seulement a partir d'un certain Niveau"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.OptionButton optBmont 
         Caption         =   "Bloquer Monture"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Bloque seulement les joueurs qui ont une monture"
         Top             =   1320
         Width           =   1395
      End
      Begin VB.OptionButton optcoffre 
         Caption         =   "Coffre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Cr�e un coffre qui ne pourras �tre ouvert qu'avec un objet cl� s�lectionner ou un code"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.OptionButton optportecode 
         Caption         =   "Porte � code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Cr�e une porte qui ne pourras �tre ouvert qu'avec un cod�finit"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "T�l�portation"
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
         Left            =   1560
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "T�l�port le joueur au positions et a la carte choisie"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton OptBank 
         Caption         =   "Banque"
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
         Left            =   0
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Case o� le joueur doit marcher pour ouvrit la fen�tre de la banque"
         Top             =   4200
         Width           =   975
      End
      Begin VB.OptionButton optScripted 
         Caption         =   "Case Script�e"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Ex�cute le script de la case s�lectionner"
         Top             =   1680
         Width           =   1290
      End
      Begin VB.OptionButton optClassChange 
         Caption         =   "Chg de Classe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Change la classe du joueur"
         Top             =   2640
         Width           =   1200
      End
      Begin VB.OptionButton optNotice 
         Caption         =   "Avertissement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Avertissement sous forme de texte et/ou de bruit"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.OptionButton optDoor 
         Caption         =   "Porte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Cr�e une porte qui s'ouvrira � l'approche du joueur"
         Top             =   2640
         Width           =   960
      End
      Begin VB.OptionButton optSign 
         Caption         =   "Panneau"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Cr�e un panneau"
         Top             =   3480
         Width           =   1080
      End
      Begin VB.OptionButton optSprite 
         Caption         =   "Chg Sprite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Change l'apparence du joueur(Sprite = skin/habit)"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Jouer un son"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Joue un son quand le joueur passe sur la case"
         Top             =   600
         Width           =   1170
      End
      Begin VB.OptionButton optArena 
         Caption         =   "Ar�ne"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Case � mettre dans les ar�nes qui enl�vent les p�nalit�s PK (tuer des joueurs)"
         Top             =   840
         Width           =   1170
      End
      Begin VB.OptionButton optCBlock 
         Caption         =   "Bloquer Class"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Bloque seulement certaines classes"
         Top             =   840
         Width           =   1125
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Magasin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Case o� le joueur doit marcher pour ouvrit la fen�tre du magasin s�lectionner"
         Top             =   3960
         Width           =   810
      End
      Begin VB.OptionButton optKill 
         Caption         =   "Tuer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Tue un joueur"
         Top             =   1320
         Width           =   810
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Soins"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Soigne un joueur"
         Top             =   1080
         Width           =   1035
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Ouvrir une Porte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "En passant sur cette case le joueur ouvrira une porte s�lectionner par ses coordonner"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Bloquer"
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
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Bloque les joueur et le PNJ"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ffacer tout les attributs"
         Height          =   300
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   25
         ToolTipText     =   "Efface tout les attributs sur la carte"
         Top             =   4560
         Width           =   2175
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Objet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Pose l'objet s�lectionner au sol"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Porte � cl�"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Cr�e une porte qui ne pourras �tre ouvert qu'avec un objet cl� s�lectionner"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label atrib 
         Appearance      =   0  'Flat
         Caption         =   "Attributs :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox txtQ 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3540
      ScaleHeight     =   1545
      ScaleWidth      =   9465
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox TxtQ2 
         Height          =   1180
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "frmMirage.frx":183DD
         Top             =   185
         Width           =   8865
      End
      Begin VB.Label OK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9000
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11040
      Top             =   480
   End
   Begin VB.Timer timerbar 
      Interval        =   500
      Left            =   11520
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12000
      Top             =   480
   End
   Begin VB.Timer quetetimersec 
      Enabled         =   0   'False
      Left            =   12480
      Top             =   480
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12960
      Top             =   1440
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13440
      Top             =   1440
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   13920
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imagebouton 
      Left            =   14400
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":183E3
            Key             =   ""
            Object.Tag             =   "frange"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":18813
            Key             =   ""
            Object.Tag             =   "frange1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":18C5C
            Key             =   ""
            Object.Tag             =   "sol"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":18E04
            Key             =   ""
            Object.Tag             =   "mask1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":18FC0
            Key             =   ""
            Object.Tag             =   "mask"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1915C
            Key             =   ""
            Object.Tag             =   "enregistrer"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":19287
            Key             =   ""
            Object.Tag             =   "script"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":193B4
            Key             =   ""
            Object.Tag             =   "tester"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":19474
            Key             =   ""
            Object.Tag             =   "maskanim"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1960A
            Key             =   ""
            Object.Tag             =   "frange2anim"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":19A2A
            Key             =   ""
            Object.Tag             =   "frange3"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":19E51
            Key             =   ""
            Object.Tag             =   "frange3anim"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1A254
            Key             =   ""
            Object.Tag             =   "frange1anim"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1A661
            Key             =   ""
            Object.Tag             =   "mask2anim"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1A810
            Key             =   ""
            Object.Tag             =   "mask3"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1A9B4
            Key             =   ""
            Object.Tag             =   "mask3anim"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1AB4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1AC4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1ACF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1ADE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1B1A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1B573
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1B92D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1BB11
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1BF15
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1C0F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1C271
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMirage.frx":1C4DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   820
      ButtonWidth     =   847
      ButtonHeight    =   820
      Style           =   1
      ImageList       =   "imagebouton"
      HotImageList    =   "imagebouton"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   34
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tester la carte (F9)"
            Object.Tag             =   "1"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enregistrer la carte sur le serveur"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sol"
            ImageIndex      =   3
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Masque 1"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Masque 2"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Masque 3"
            ImageIndex      =   15
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Frange 1"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Frange 2"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Frange 3"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Animation Masque 1"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Animation Masque 2"
            ImageIndex      =   14
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Animation Masque 3"
            ImageIndex      =   16
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Animation Frange 1"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Animation Frange 2"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Animation Frange 3"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Annuler la derni�re action (CTRL+Z)"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "R�tablir la derni�re action (CTRL+U)"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Couches"
            ImageIndex      =   20
            Style           =   1
            Object.Width           =   1e-4
            Value           =   1
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Attributs"
            ImageIndex      =   21
            Style           =   1
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lumi�res"
            ImageIndex      =   22
            Style           =   1
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom 100%"
            ImageIndex      =   24
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom 50%"
            ImageIndex      =   25
            Style           =   1
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom 30%"
            ImageIndex      =   23
            Style           =   1
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remplir la carte par l'�l�ment graphique s�l�ctionn�"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pr�lever un �l�ment graphique de la carte (Maj+Supr)"
            ImageIndex      =   17
            Style           =   1
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Efface un �l�ment de la carte (Supr)"
            ImageIndex      =   26
            Style           =   1
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Supprimer la couche active sur toute la carte"
            ImageIndex      =   19
         EndProperty
      EndProperty
      MouseIcon       =   "frmMirage.frx":1C748
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   6480
      LargeChange     =   10
      Left            =   3300
      Max             =   512
      TabIndex        =   19
      Top             =   825
      Width           =   255
   End
   Begin HookMenuPlus.ctxHookMenu ctxHookMenu1 
      Left            =   2280
      Top             =   7680
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   10
      Bmp:1           =   "frmMirage.frx":1C8AA
      Mask:1          =   16777215
      Key:1           =   "#test"
      Bmp:2           =   "frmMirage.frx":1CACC
      Mask:2          =   16777215
      Key:2           =   "#envoicarte"
      Bmp:3           =   "frmMirage.frx":1D1DA
      Mask:3          =   16777215
      Key:3           =   "#enregcarte"
      Bmp:4           =   "frmMirage.frx":1D8E8
      Mask:4          =   16777215
      Key:4           =   "#quit"
      Bmp:5           =   "frmMirage.frx":1DFF6
      Mask:5          =   16777215
      Key:5           =   "#rempli"
      Bmp:6           =   "frmMirage.frx":1E218
      Mask:6          =   16777215
      Key:6           =   "#prelv"
      Bmp:7           =   "frmMirage.frx":1E926
      Mask:7          =   16777215
      Key:7           =   "#tp:1"
      Bmp:8           =   "frmMirage.frx":1F034
      Mask:8          =   16777215
      Key:8           =   "#tp:2"
      Bmp:9           =   "frmMirage.frx":1F742
      Mask:9          =   16777215
      Key:9           =   "#tp:3"
      Bmp:10          =   "frmMirage.frx":1FE50
      Mask:10         =   16777215
      Key:10          =   "#gom"
      UseSystemFont   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelMenuBorder   =   16744576
      SelMenuBack     =   16761024
      SelMenuFore     =   16777215
      CheckBack       =   16761024
      CheckFore       =   0
      SelCheckBack    =   16744576
      MenuBorder      =   16777215
      MenuBack        =   16761024
      MenuFore        =   0
      DisabledMenuBorder=   16744576
      DisabledMenuBack=   16761024
      DisabledMenuFore=   14737632
      MenuBarBack     =   14737632
      MenuPopupBack   =   16777215
      AutorSet        =   -1  'True
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10560
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.VScrollBar hautbas 
      Height          =   9330
      Left            =   15015
      Max             =   23
      Min             =   -1
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   480
      Value           =   -1
      Width           =   255
   End
   Begin VB.HScrollBar gauchedroite 
      Height          =   255
      Left            =   3540
      Max             =   24
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   9555
      Width           =   11475
   End
   Begin VB.TextBox txtMyTextBox 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   255
      TabIndex        =   2
      Top             =   9840
      Width           =   10335
   End
   Begin VB.ComboBox tilescmb 
      Height          =   315
      ItemData        =   "frmMirage.frx":2055E
      Left            =   0
      List            =   "frmMirage.frx":20574
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   495
      Width           =   3525
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   30
      Min             =   1
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7050
      Value           =   1
      Width           =   3300
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9075
      Left            =   3540
      ScaleHeight     =   605
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   765
      TabIndex        =   3
      Top             =   480
      Width           =   11475
      Begin VB.PictureBox mana 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9720
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lmana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CB884B&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape smana 
            BackColor       =   &H00CB884B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox vie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9720
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lvie 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape svie 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox xp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9720
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lexp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape sexp 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox ObjNm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   52
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Label OName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox picmapeditor 
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   3540
      ScaleHeight     =   8175
      ScaleWidth      =   7095
      TabIndex        =   18
      Top             =   360
      Width           =   7095
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   3270
      TabIndex        =   43
      Top             =   495
      Width           =   3270
   End
   Begin VB.ListBox lstIndex 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2370
      ItemData        =   "frmMirage.frx":205A8
      Left            =   60
      List            =   "frmMirage.frx":205AF
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   7440
      Width           =   3300
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   15
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   143
      Top             =   855
      Visible         =   0   'False
      Width           =   480
   End
   Begin WMPLibCtl.WindowsMediaPlayer mediaplayer 
      Height          =   240
      Left            =   12840
      TabIndex        =   49
      Top             =   2040
      Width           =   360
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   2
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   635
      _cy             =   423
   End
   Begin VB.Menu fichie 
      Caption         =   "Fichier"
      Begin VB.Menu test 
         Caption         =   "Tester"
         Shortcut        =   {F9}
      End
      Begin VB.Menu opti 
         Caption         =   "Options"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu envoicarte 
         Caption         =   "Envoyer la carte au serveur"
      End
      Begin VB.Menu enregcarte 
         Caption         =   "Enregistrer la carte dans le dossier"
      End
      Begin VB.Menu envserv 
         Caption         =   "Envoyez les �l�ments modifi�s hors ligne au serveur"
      End
      Begin VB.Menu bare 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "Quitter"
         Shortcut        =   +^{F12}
      End
   End
   Begin VB.Menu Editeurs 
      Caption         =   "Editeurs"
      Begin VB.Menu Editeurpatch 
         Caption         =   "Editeur de Patches"
      End
      Begin VB.Menu Editeurclas 
         Caption         =   "Editeur de Classes"
         Shortcut        =   ^R
      End
      Begin VB.Menu Editeursort 
         Caption         =   "Editeur de Sorts"
         Shortcut        =   ^S
      End
      Begin VB.Menu Editeurarea 
         Caption         =   "Editeur de Zones"
      End
      Begin VB.Menu Editeurobj 
         Caption         =   "Editeur d'Objets"
         Shortcut        =   ^O
      End
      Begin VB.Menu Editeurmags 
         Caption         =   "Editeur de Magasins"
         Shortcut        =   ^M
      End
      Begin VB.Menu Editeurpng 
         Caption         =   "Editeur de Pnj"
         Shortcut        =   ^P
      End
      Begin VB.Menu Editeurflech 
         Caption         =   "Editeur de Fl�ches"
         Shortcut        =   ^F
      End
      Begin VB.Menu Editeuremot 
         Caption         =   "Editeur d'Emoticons"
         Shortcut        =   ^E
      End
      Begin VB.Menu Editeurqut 
         Caption         =   "Editeur de Qu�tes"
         Shortcut        =   ^Q
      End
      Begin VB.Menu Editeurreve 
         Caption         =   "Editeur de Reves"
      End
      Begin VB.Menu Editeurcraft 
         Caption         =   "Editeur de Crafts"
         Shortcut        =   ^C
      End
      Begin VB.Menu Editcinem 
         Caption         =   "Editeur de Cinematique"
         Enabled         =   0   'False
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu editeurpet 
         Caption         =   "Editeurs de Familier"
      End
   End
   Begin VB.Menu comtest 
      Caption         =   "Commandes de Test"
      Enabled         =   0   'False
      Begin VB.Menu inv 
         Caption         =   "Inventaire"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sort 
         Caption         =   "Sorts"
         Shortcut        =   {F2}
      End
      Begin VB.Menu opt 
         Caption         =   "Options"
         Shortcut        =   {F3}
      End
      Begin VB.Menu guild 
         Caption         =   "Guilde"
         Shortcut        =   {F4}
      End
      Begin VB.Menu equip 
         Caption         =   "Equipement"
         Shortcut        =   {F5}
      End
      Begin VB.Menu creeguild 
         Caption         =   "Cr�er une guilde"
         Shortcut        =   {F6}
      End
      Begin VB.Menu adminiguild 
         Caption         =   "Administration de la guilde"
         Shortcut        =   {F7}
      End
      Begin VB.Menu vies 
         Caption         =   "Vie,Magie,..."
         Shortcut        =   {F11}
      End
      Begin VB.Menu nj 
         Caption         =   "Nuit/Jour"
         Shortcut        =   {F12}
      End
      Begin VB.Menu qutcour 
         Caption         =   "Qu�te en cours"
         Shortcut        =   ^{F6}
      End
   End
   Begin VB.Menu carte 
      Caption         =   "Carte"
      Begin VB.Menu propricarte 
         Caption         =   "Propri�t�s de la Carte"
      End
      Begin VB.Menu stopmusic 
         Caption         =   "Arr�ter la musique"
      End
      Begin VB.Menu scrshot 
         Caption         =   "Prendre une capture d'�cran de toute la carte"
      End
      Begin VB.Menu outi 
         Caption         =   "Outils"
         Begin VB.Menu rempli 
            Caption         =   "Remplir"
            Shortcut        =   %{BKSP}
         End
         Begin VB.Menu prelv 
            Caption         =   "Pr�l�vement d'une couche"
            Shortcut        =   +{DEL}
         End
         Begin VB.Menu gom 
            Caption         =   "Gommer"
            Shortcut        =   {DEL}
         End
      End
      Begin VB.Menu afich 
         Caption         =   "Affichage"
         Begin VB.Menu modscreen 
            Caption         =   "Mode screenshot"
         End
         Begin VB.Menu br 
            Caption         =   "-"
         End
         Begin VB.Menu zp 
            Caption         =   "Zoom +"
         End
         Begin VB.Menu zm 
            Caption         =   "Zoom -"
         End
         Begin VB.Menu bare2 
            Caption         =   "-"
         End
         Begin VB.Menu previsu 
            Caption         =   "Pr�visualisation"
            Checked         =   -1  'True
         End
         Begin VB.Menu grile 
            Caption         =   "Grille"
            Checked         =   -1  'True
         End
         Begin VB.Menu nuitjour 
            Caption         =   "Nuit"
         End
      End
      Begin VB.Menu tile 
         Caption         =   "Planches/Tiles"
         Begin VB.Menu Tiles 
            Caption         =   "Tiles0"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu types 
         Caption         =   "Types"
         Begin VB.Menu tp 
            Caption         =   "Couches"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu tp 
            Caption         =   "Attributs"
            Index           =   2
         End
         Begin VB.Menu tp 
            Caption         =   "Lumieres"
            Index           =   3
         End
      End
   End
   Begin VB.Menu modo 
      Caption         =   "Mod�ration"
      Begin VB.Menu quilgn 
         Caption         =   "Qui est en ligne?"
      End
      Begin VB.Menu pmodo 
         Caption         =   "Panneau de mod�ration"
      End
      Begin VB.Menu modoserv 
         Caption         =   "Mod�ration du serveur"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu admin 
      Caption         =   "Administration"
      Begin VB.Menu qui 
         Caption         =   "Qui est en ligne?"
      End
      Begin VB.Menu padmin 
         Caption         =   "Panneau d'administration"
      End
      Begin VB.Menu adminserv 
         Caption         =   "Administration du serveur"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu tuto 
      Caption         =   "Tutoriaux"
      Enabled         =   0   'False
      Begin VB.Menu intro 
         Caption         =   "Introduction � la cr�ation"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu tutbase 
         Caption         =   "Principes de base"
      End
      Begin VB.Menu tutcarte 
         Caption         =   "Comment faire une carte?"
      End
      Begin VB.Menu tutsort 
         Caption         =   "Comment cr�e un sort?"
      End
      Begin VB.Menu tutobj 
         Caption         =   "Comment cr�e un objet?"
      End
      Begin VB.Menu tutmaga 
         Caption         =   "Comment cr�e un magasin?"
      End
      Begin VB.Menu tutpng 
         Caption         =   "Comment cr�e un Pnj?"
      End
      Begin VB.Menu tutfleche 
         Caption         =   "Comment cr�e une fl�che?"
      End
      Begin VB.Menu tutemot 
         Caption         =   "Comment cr�e un �moticon?"
      End
      Begin VB.Menu tutquet 
         Caption         =   "Comment cr�e une qu�te?"
      End
   End
   Begin VB.Menu apop 
      Caption         =   "?"
      Begin VB.Menu credit 
         Caption         =   "Cr�dits"
      End
      Begin VB.Menu site 
         Caption         =   "Site officiel de Frog Creator"
      End
      Begin VB.Menu hscript 
         Caption         =   "Site d'aide pour les scripts"
      End
      Begin VB.Menu siteequp 
         Caption         =   "Site de l'�quipe de Frog"
      End
      Begin VB.Menu don 
         Caption         =   "Dons Gratuits!"
      End
   End
   Begin VB.Menu mclik 
      Caption         =   "Menuclik"
      Visible         =   0   'False
      Begin VB.Menu eff 
         Caption         =   "Effacer"
      End
      Begin VB.Menu copco 
         Caption         =   "Copier les coordon�es"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "edit"
      Visible         =   0   'False
      Begin VB.Menu coup 
         Caption         =   "Couper"
      End
      Begin VB.Menu copi 
         Caption         =   "Copier"
      End
      Begin VB.Menu colle 
         Caption         =   "Coller"
      End
   End
End
Attribute VB_Name = "frmMirage"
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
Dim SpellMemorized As Long
Dim KeyShift As Boolean
Dim nbcle As Long
Dim Quite As Boolean
Dim attender As String
Dim OldPCX As Long
Dim OldPCY As Long
Dim OldTiles As Long
Dim InPM As Boolean
Dim TCouche As Byte
Dim Couche As String
Dim CTRLDOWN As Boolean

Dim mclsToolTip As New clsToolTip

Private Sub AddDef_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMagi_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 2 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 3 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddStr_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Private Sub adminiguild_Click()
Dim V As Boolean
V = Not picGuildAdmin.Visible
    If Player(MyIndex).Guildaccess > 1 Then Call NetPic: frmMirage.picGuildAdmin.Visible = V
End Sub

Private Sub artquete_Click()
    Player(MyIndex).QueteEnCour = 0
    Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & SEP_CHAR & END_CHAR)
    picquete.Visible = False
End Sub

Private Sub Attributs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Attributs.Top = 33: Attributs.Left = 0: Exit Sub
    dr = True
    drx = x
    dry = y
End Sub

Private Sub Attributs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If dr Then Attributs.Refresh: DoEvents: If dr Then Call Attributs.Move(Attributs.Left + (x - drx), Attributs.Top + (y - dry))
If Attributs.Left > Me.Width Or Attributs.Top > Me.Height Then Attributs.Top = 33: Attributs.Left = 0: Exit Sub
End Sub

Private Sub Attributs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub cmdLeave_Click()
Dim Packet As String
    Packet = "GUILDLEAVE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdMember_Click()
Dim Packet As String
    If txtName.Text = vbNullString Then Exit Sub
    Packet = "GUILDMEMBER" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub colle_Click()
Dim f As Long
    If DonID <> lstIndex.ListIndex And lstIndex.ListIndex <> -1 Then
        If FileExiste("Maps\map" & DonID & ".aoc") Then Call FileCopy(App.Path & "\Maps\map" & DonID & ".aoc", App.Path & "\Maps\map" & lstIndex.ListIndex & ".aoc") Else Call SaveMap(DonID): Call FileCopy(App.Path & "\Maps\map" & DonID & ".aoc", App.Path & "\Maps\map" & lstIndex.ListIndex & ".aoc")
        'Call ViderTMap(lstIndex.ListIndex)
        Call ClearMap(lstIndex.ListIndex)
        If FileExiste("Maps\map" & lstIndex.ListIndex & ".aoc") Then
            FileName = App.Path & "\Maps\map" & lstIndex.ListIndex & ".aoc"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Map(lstIndex.ListIndex)
            Close #f
        End If
        Call SaveMap(lstIndex.ListIndex)
        Call SendTMap(lstIndex.ListIndex)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Map(lstIndex.ListIndex).name
        If DonTP = 1 Then
            If FileExiste("Maps\map" & DonID & ".aoc") Then Call Kill(App.Path & "\Maps\map" & DonID & ".aoc")
            'Call ViderTMap(DonID)
            Call ClearMap(DonID)
            Call SaveMap(DonID)
            Call SendTMap(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        End If
    End If
End Sub

Private Sub copco_Click()
CoordX = CurX
CoordY = CurY
CoordM = Player(MyIndex).Map
InPM = False
End Sub

Private Sub copi_Click()
    DonID = lstIndex.ListIndex
    DonTP = 2
End Sub

Private Sub coup_Click()
    DonID = lstIndex.ListIndex
    DonTP = 1
End Sub

Private Sub credit_Click()
frmpet.Show vbModeless, frmMirage
End Sub

Private Sub creeguild_Click()
frmGuild.Show vbModeless, frmMirage
End Sub

Private Sub don_Click()
ShellExecute Me.hWnd, "open", "http://www.aride-online.com", vbNullString, App.Path, 1
End Sub

Public Sub Editeurarea_Click()
Dim i As Integer
Call NetInEditor
If HORS_LIGNE = 1 Then
    InAreaEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    For i = 0 To MAX_AREAS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Areas(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
End If
End Sub

Private Sub Editeurclas_Click()
frmoptions.SSTab1.Tab = 3
frmoptions.Show
frmoptions.nbcls.Text = ReadINI("INFO", "MaxClasses", App.Path & "\Classes\info.ini")
Dim i As Long
Call frmoptions.clase.Clear
'clase.Text = "S�l�ctioner une classe"
For i = 0 To Val(frmoptions.nbcls.Text)
    Call frmoptions.clase.AddItem("Classe" & i, i)
Next i
If Val(ReadINI("INFO", "HPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.PV = ReadINI("INFO", "HPRegen", App.Path & "\config.ini")
If Val(ReadINI("INFO", "MPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.pm = ReadINI("INFO", "MPRegen", App.Path & "\config.ini")
If Val(ReadINI("INFO", "SPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.ps = ReadINI("INFO", "SPRegen", App.Path & "\config.ini")
If Val(ReadINI("CONFIG", "Scrolling", App.Path & "\config.ini")) > 0 Then frmoptions.defl = ReadINI("CONFIG", "Scrolling", App.Path & "\config.ini")
If Val(ReadINI("CONFIG", "Scripting", App.Path & "\config.ini")) > 0 Then frmoptions.script = ReadINI("CONFIG", "Scripting", App.Path & "\config.ini")
If Val(ReadINI("INFO", "GameName", App.Path & "\config.ini")) > 0 Then frmoptions.nom = ReadINI("INFO", "GameName", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")) > 0 Then frmoptions.mj = ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxitems", App.Path & "\config.ini")) > 0 Then frmoptions.mo = ReadINI("INFO", "Maxitems", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")) > 0 Then frmoptions.mpnj = ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxshops", App.Path & "\config.ini")) > 0 Then frmoptions.mm = ReadINI("INFO", "Maxshops", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxspells", App.Path & "\config.ini")) > 0 Then frmoptions.ms = ReadINI("INFO", "Maxspells", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")) > 0 Then frmoptions.mc = ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")) > 0 Then frmoptions.moc = ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxemots", App.Path & "\config.ini")) > 0 Then frmoptions.Me = ReadINI("INFO", "Maxemots", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")) > 0 Then frmoptions.mn = ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxquet", App.Path & "\config.ini")) > 0 Then frmoptions.mq = ReadINI("INFO", "Maxquet", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxguilds", App.Path & "\config.ini")) > 0 Then frmoptions.mg = ReadINI("INFO", "Maxguilds", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxjguild", App.Path & "\config.ini")) > 0 Then frmoptions.mjg = ReadINI("INFO", "Maxjguild", App.Path & "\config.ini")
If ReadINI("INFO", "motd", App.Path & "\config.ini") > vbNullString Then frmoptions.motd = ReadINI("INFO", "motd", App.Path & "\config.ini")
End Sub

Public Sub Editeurcraft_Click()
Dim i As Integer
Call NetInEditor
If HORS_LIGNE = 1 Then
    InCraftEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    For i = 0 To MAX_CRAFTS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Crafts(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditCraft Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeuremot_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InEmoticonEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    For i = 0 To MAX_EMOTICONS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Emoticons(i).Command)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditEmoticon Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurflech_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InArrowEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    For i = 1 To MAX_ARROWS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Arrows(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditArrow Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurmags_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InShopEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_SHOPS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Shop(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditShop Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurobj_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InItemsEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 0 To MAX_ITEMS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Item(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditItem Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurpatch_Click()
'frmPatch.Show
frmPatchEditor.Show
End Sub

Private Sub editeurpet_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InPetsEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_PETS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Pets(i).nom)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditPet Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurpng_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InNpcEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 0 To MAX_NPCS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Npc(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditNpc Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurqut_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InQuetesEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 1 To MAX_QUETES
        frmIndex.lstIndex.AddItem i & " : " & Trim$(quete(i).nom)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditQuetes Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeurreve_Click()
Call NetInEditor
If HORS_LIGNE = 1 Then
    InDreamEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    Call frmIndex.Load_Dreams
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditDream Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub Editeursort_Click()
Dim i As Long
Call NetInEditor
If HORS_LIGNE = 1 Then
    InSpellEditor = True
    frmIndex.Show vbModeless, frmMirage
    DonID = 0
    frmIndex.lstIndex.Clear
    ' Add the names
    For i = 0 To MAX_SPELLS
        frmIndex.lstIndex.AddItem i & " : " & Trim$(Spell(i).name)
    Next i
    frmIndex.lstIndex.ListIndex = 0
Else: If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then Call SendRequestEditSpell Else Call AddText("Vous n'�tes pas autoris� � faire cette action.", BrightRed)
End If
End Sub

Private Sub eff_Click()
Call EditorMouseDown(2, 1, (PotX + NewPlayerPicX), (PotY + NewPlayerPicY))
InPM = False
End Sub

Private Sub enregcarte_Click()
save = 0
Call SaveMap(Player(MyIndex).Map)
End Sub

Private Sub envoicarte_Click()
Dim x As Long
    x = MsgBox("�tes-vous s�r de vouloir enregistrer les changements de la carte ?", vbYesNo)
    If x = vbNo Then Exit Sub
    save = 0
    ScreenMode = False

    Call SaveLocalMap(Player(MyIndex).Map)
    Call EditorSend
End Sub

Private Sub envserv_Click()
Call EnvoieServeur
End Sub

Private Sub equip_Click()
Dim V As Boolean
    V = Not picEquip.Visible
    Call NetPic
    Call UpdateVisInv
    picEquip.Visible = V
    Picpics.Visible = V
End Sub

Private Sub Form_Activate()
gauchedroite.min = 0
gauchedroite.Max = Int(MaxMapX - (picScreen.Width / 32)) + 1
hautbas.min = -1
hautbas.Max = Int(MaxMapY - (picScreen.Height / 32))
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

If Val(ReadINI("CONFIG", "MapGrid", FileName)) = 0 Then frmMirage.grile.Checked = False Else frmMirage.grile.Checked = True
If Val(ReadINI("CONFIG", "PreVisu", FileName)) = 0 Then frmMirage.previsu.Checked = False Else frmMirage.previsu.Checked = True
Call InitMirageVars

Dim ctrl As Control
With mclsToolTip
    Call .Create(Me)
    
    .MaxTipWidth = 240
    .DelayTime(ttDelayShow) = 20000
    
    For Each ctrl In Controls
        Call .AddTool(ctrl)
    Next ctrl
    
    .ToolTipHeaderShow = True
    '.ToolText(frmMirage.picScreen) = "test"
End With

Quite = False
OldPCX = 0
OldPCY = 0
nbcle = 5
TempNum = 0
Dim Screenw As Long
If HORS_LIGNE = 1 Then Call InitMirage
'HScroll1.Max = (frmMirage.picBackSelect.Width / 32)
HScroll1.Width = (scrlPicture.Left)
'picBack.Width = (scrlPicture.Left + 17)
'gauchedroite.min = 0
'gauchedroite.Max = Int(MaxMapX - (picScreen.Width / 32)) + 1
'hautbas.min = -1
'hautbas.Max = Int(MaxMapY - (picScreen.Height / 32))
Call NetPic
Couche = "Sol"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Toolbar1.buttons(1).Image = 19 Then Exit Sub
    itmDesc.Visible = False
    InEditor = True
    'frmMirage.scrlPicture.Max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
'    frmMirage.picBack.Width = frmMirage.picBackSelect.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim attender As String

If GettingMap Then
    Cancel = True
Else
    If save = 1 Then
        attender = MsgBox("La carte n'a pas �t� enregistr�e ni sur le serveur ni sur le disque dur voulez vous vraiment quitter?", vbYesNo)
        If attender = vbNo Then Cancel = True: GoTo qui:
    End If
    If Val(ReadINI("modif", "carte" & Player(MyIndex).Map, App.Path & "\config.ini")) = 1 And HORS_LIGNE = 0 And Val(attender) <> vbYes Then
        attender = MsgBox("La carte n'a pas �t� enregistr�e sur le serveur voulez vous vraiment quitter?", vbYesNo)
        If attender = vbNo Then Cancel = True: GoTo qui:
    End If
    If HORS_LIGNE = 0 Then frmmsg.Show: frmMainMenu.Timer2.Enabled = True Else Call GameDestroy
End If
qui:
End Sub

Private Sub gauchedroite_Change()
If InEditor = False Then Exit Sub
Call SetPlayerX(MyIndex, ((picScreen.Width \ 32) \ 2) + gauchedroite.value)
End Sub

Private Sub gauchedroite_Scroll()
If InEditor = False Then Exit Sub
Call SetPlayerX(MyIndex, ((picScreen.Width \ 32) \ 2) + gauchedroite.value)
End Sub

Private Sub gom_Click()
    If frmMirage.MousePointer = 2 Then Call prelv_Click
    If frmMirage.MousePointer = 99 Then frmMirage.MousePointer = 0: frmMirage.Toolbar1.buttons(33).value = tbrUnpressed Else frmMirage.MousePointer = 99
End Sub

Private Sub grile_Click()
    If grile.Checked Then
        WriteINI "CONFIG", "MapGrid", 0, App.Path & "\Config\Account.ini"
        grile.Checked = False
        AccOpt.MapGrid = False
    Else
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\Config\Account.ini"
        grile.Checked = True
        AccOpt.MapGrid = True
    End If
End Sub

Private Sub guild_Click()
' Set Their Guild Name and Their Rank
Dim V As Boolean
V = Not Picture1.Visible
Call NetPic
lblGuild.Caption = GetPlayerGuild(MyIndex)
lblRank.Caption = GetPlayerGuildAccess(MyIndex)
Picture1.Visible = V
Picpics.Visible = V
End Sub

Private Sub hautbas_Change()
If Not InEditor Then Exit Sub
Call SetPlayerY(MyIndex, (Int(picScreen.Height / 32) / 2) + hautbas.value)
End Sub

Private Sub hautbas_GotFocus()
If Not InEditor Then Exit Sub
Call SetPlayerY(MyIndex, (Int(picScreen.Height / 32) / 2) + hautbas.value)
End Sub

Private Sub hautbas_Scroll()
If Not InEditor Then Exit Sub
Call SetPlayerY(MyIndex, (Int(picScreen.Height / 32) / 2) + hautbas.value)
End Sub

Private Sub hscript_Click()
ShellExecute Me.hWnd, "open", "http://www.aride-online.com", vbNullString, App.Path, 1
End Sub

Private Sub HScroll1_Change()
'HScroll1.Max = frmMirage.picBackSelect.Width / 32
'picBackSelect.Left = (frmMirage.HScroll1.value * PIC_Y) * -1 + 33
End Sub

Private Sub HScroll1_Scroll()
'HScroll1.Max = frmMirage.picBackSelect.Width / 32
'picBackSelect.Left = (frmMirage.HScroll1.value * PIC_Y) * -1 + 33
End Sub

Private Sub inv_Click()
    Dim V As Boolean
    V = Not picInv3.Visible
    Call NetPic
    Call UpdateVisInv
    picInv3.Visible = V
    Picpics.Visible = V
End Sub

Private Sub lblPoints_Change()
    If GetPlayerPOINTS(MyIndex) > 0 Then
        frmMirage.AddStr.Visible = True
        frmMirage.AddDef.Visible = True
        frmMirage.AddSpeed.Visible = True
        frmMirage.AddMagi.Visible = True
    Else
        frmMirage.AddStr.Visible = False
        frmMirage.AddDef.Visible = False
        frmMirage.AddSpeed.Visible = False
        frmMirage.AddMagi.Visible = False
    End If
End Sub

Private Sub lstIndex_DblClick()
'Dim sauve As String
'Dim attender As String
'If save = 1 Then
'    sauve = MsgBox("Vous �tes sur de vouloir changer de carte sans sauvegarder sur le disque dur?", vbYesNo)
'    If sauve = vbNo Then Exit Sub
'    GoTo oui:
'End If
'
'If Val(ReadINI("modif", "carte" & Player(MyIndex).Map, App.Path & "\config.ini")) = 1 And HORS_LIGNE = 0 Then
'    attender = MsgBox("La carte n'a pas �t� enregistr�e sur le serveur voulez vous vraiment changer de carte?", vbYesNo)
'    If attender = vbNo Then Exit Sub
'End If
'oui:
''If frmMapProperties.Visible Then Unload frmMapProperties
'frmmsg.Show
'GettingMap = True
'Call WriteINI("modif", "carte" & Player(MyIndex).Map, "0", App.Path & "\config.ini")
'save = 0
Dim sauve As String
Dim attender As String
If save = 1 Then
    sauve = MsgBox("Vous �tes sur de vouloir changer de carte sans sauvegarder sur le disque dur?", vbYesNo)
    If sauve = vbNo Then Exit Sub
    GoTo oui:
End If

If Val(ReadINI("modif", "carte" & Player(MyIndex).Map, App.Path & "\config.ini")) = 1 And HORS_LIGNE = 0 Then
    attender = MsgBox("La carte n'a pas �t� enregistr�e sur le serveur voulez vous vraiment changer de carte?", vbYesNo)
    If attender = vbNo Then Exit Sub
End If
oui:
Call GoToMap(lstIndex.ListIndex)
'Dim i As Long
'If HORS_LIGNE = 1 Then
'    If FileExiste("maps\map" & lstIndex.ListIndex & ".aoc") Then
'        Call SetPlayerMap(MyIndex, Val(lstIndex.ListIndex))
'        Call ChargerPnjs
'        Call ChargerObjets(MyIndex)
'        Call ChargerCarte(lstIndex.ListIndex)
'    Else
'        Call SaveMapVide(lstIndex.ListIndex)
'        Call SetPlayerMap(MyIndex, Val(lstIndex.ListIndex))
'        Call ChargerPnjs
'        Call ChargerObjets(MyIndex)
'        Call ChargerCarte(lstIndex.ListIndex)
'    End If
'    For i = 0 To 5
'        Call NetTempMap(i)
'    Next i
'    TempNum = 0
'    frmMirage.Toolbar1.buttons(20).Enabled = True
'    frmMirage.Toolbar1.buttons(21).Enabled = False
'    'Call SauvTemp
'Else
'    GettingMap = True
'    Call LoadMap(lstIndex.ListIndex)
'    Call SendData("WARPTO" & SEP_CHAR & lstIndex.ListIndex & SEP_CHAR & END_CHAR)
'    frmMirage.SetFocus
'    For i = 0 To 5
'        Call NetTempMap(i)
'    Next i
'    TempNum = 0
'    frmMirage.Toolbar1.buttons(20).Enabled = True
'    frmMirage.Toolbar1.buttons(21).Enabled = False
''    Call SauvTemp
'End If
End Sub

Private Sub GoToMap(ByVal mapNumber As Integer)
Dim i As Long


'If frmMapProperties.Visible Then Unload frmMapProperties
frmmsg.Show
GettingMap = True
Call WriteINI("modif", "carte" & Player(MyIndex).Map, "0", App.Path & "\config.ini")
save = 0

lstIndex.ListIndex = mapNumber

If HORS_LIGNE = 1 Then
    If FileExiste("maps\map" & mapNumber & ".aoc") Then
        Call SetPlayerMap(MyIndex, Val(mapNumber))
        Call ChargerPnjs
        Call ChargerObjets(MyIndex)
        Call ChargerCarte(mapNumber)
    Else
        Call SaveMapVide(mapNumber)
        Call SetPlayerMap(MyIndex, Val(mapNumber))
        Call ChargerPnjs
        Call ChargerObjets(MyIndex)
        Call ChargerCarte(mapNumber)
    End If
    Call ClearTempTile
    
    Call ReInitView
    
    For i = 0 To 5
        Call NetTempMap(i)
    Next i
    TempNum = 0
    frmMirage.Toolbar1.buttons(20).Enabled = True
    frmMirage.Toolbar1.buttons(21).Enabled = False
    'Call SauvTemp
Else
    GettingMap = True
    Call LoadMap(mapNumber)
    Call SendData("WARPTO" & SEP_CHAR & mapNumber & SEP_CHAR & END_CHAR)
    frmMirage.SetFocus
    For i = 0 To 5
        Call NetTempMap(i)
    Next i
    TempNum = 0
    frmMirage.Toolbar1.buttons(20).Enabled = True
    frmMirage.Toolbar1.buttons(21).Enabled = False
'    Call SauvTemp
End If
End Sub

Public Sub ReInitView()
    Call DestroyDirectX
    Call InitDirectX
    'Call InitSurfaces
    frmMirage.picScreen.Refresh
    
    Call InitPano(Player(MyIndex).Map)
    Call InitNightAndFog(Player(MyIndex).Map)
    Call StopMidi
End Sub

Private Sub lstIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If DonID > 0 Then colle.Enabled = True Else colle.Enabled = False
    Call PopupMenu(Edit)
End If
End Sub

Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim$(lstOnline.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstSpells_DblClick()
    If Player(MyIndex).Spell(lstSpells.ListIndex) > 0 Then
        SpellMemorized = lstSpells.ListIndex
        Call AddText("M�morisation de la magie r�ussit!", White)
    Else
        Call AddText("Aucune magie � m�moriser.", BrightRed)
    End If
End Sub

Private Sub modscreen_Click()
    If modscreen.Checked Then ScreenMode = False: modscreen.Checked = False Else ScreenMode = True: modscreen.Checked = True
End Sub

Private Sub nj_Click()
    If GameTime = TIME_DAY Then GameTime = TIME_NIGHT: Call InitNightAndFog(Player(MyIndex).Map) Else GameTime = TIME_DAY
End Sub

Private Sub nuitjour_Click()
    nuitjour.Checked = Not nuitjour.Checked
    Call InitNightAndFog(Player(MyIndex).Map)
End Sub

Private Sub ok_Click()
Dim i As Long
Dim msgb As String

If Player(MyIndex).QueteEnCour > 0 And Accepter = False Then
    msgb = MsgBox("Voulez-vous faire la qu�te propos�e?", vbYesNo, "qu�te")
    If msgb = vbYes Then
        Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & SEP_CHAR & END_CHAR)
        Accepter = True
    Else
        Player(MyIndex).QueteEnCour = 0
        Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & SEP_CHAR & END_CHAR)
        Accepter = False
    End If
End If
txtQ.Visible = False
End Sub

Private Sub opt_Click()
frmoptions.SSTab1.Tab = 0
If Val(ReadINI("INFO", "HPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.PV = ReadINI("INFO", "HPRegen", App.Path & "\config.ini")
If Val(ReadINI("INFO", "MPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.pm = ReadINI("INFO", "MPRegen", App.Path & "\config.ini")
If Val(ReadINI("INFO", "SPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.ps = ReadINI("INFO", "SPRegen", App.Path & "\config.ini")
If Val(ReadINI("CONFIG", "Scrolling", App.Path & "\config.ini")) > 0 Then frmoptions.defl = ReadINI("CONFIG", "Scrolling", App.Path & "\config.ini")
If Val(ReadINI("CONFIG", "Scripting", App.Path & "\config.ini")) > 0 Then frmoptions.script = ReadINI("CONFIG", "Scripting", App.Path & "\config.ini")
If Val(ReadINI("INFO", "GameName", App.Path & "\config.ini")) > 0 Then frmoptions.nom = ReadINI("INFO", "GameName", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")) > 0 Then frmoptions.mj = ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxitems", App.Path & "\config.ini")) > 0 Then frmoptions.mo = ReadINI("INFO", "Maxitems", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")) > 0 Then frmoptions.mpnj = ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxshops", App.Path & "\config.ini")) > 0 Then frmoptions.mm = ReadINI("INFO", "Maxshops", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxspells", App.Path & "\config.ini")) > 0 Then frmoptions.ms = ReadINI("INFO", "Maxspells", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")) > 0 Then frmoptions.mc = ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")) > 0 Then frmoptions.moc = ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxemots", App.Path & "\config.ini")) > 0 Then frmoptions.Me = ReadINI("INFO", "Maxemots", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")) > 0 Then frmoptions.mn = ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxquet", App.Path & "\config.ini")) > 0 Then frmoptions.mq = ReadINI("INFO", "Maxquet", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxguilds", App.Path & "\config.ini")) > 0 Then frmoptions.mg = ReadINI("INFO", "Maxguilds", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxjguild", App.Path & "\config.ini")) > 0 Then frmoptions.mjg = ReadINI("INFO", "Maxjguild", App.Path & "\config.ini")
If ReadINI("INFO", "motd", App.Path & "\config.ini") > vbNullString Then frmoptions.motd = ReadINI("INFO", "motd", App.Path & "\config.ini")
frmoptions.Show
End Sub

Private Sub optArena_Click()
    frmArena.scrlNum1.Max = MAX_MAPS
    frmArena.Show vbModeless
End Sub

Private Sub OptBank_Click()
Dim variable As String
variable = InputBox("Message d'accueil:", "Banque")
bankmsg = variable
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optBguilde_Click()
frmBguilde.Show vbModeless
End Sub

Private Sub optBlocked_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optBmont_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optBniv_Click()
frmBniv.Show vbModeless
End Sub

Private Sub optbtoit_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optCBlock_Click()
    frmBClass.scrlNum1.Max = Max_Classes
    frmBClass.scrlNum2.Max = Max_Classes
    frmBClass.scrlNum3.Max = Max_Classes
    frmBClass.Show vbModeless
End Sub

Private Sub optClassChange_Click()
    frmClassChange.scrlClass.Max = Max_Classes
    frmClassChange.scrlReqClass.Max = Max_Classes
    frmClassChange.Show vbModeless
End Sub

Private Sub optcoffre_Click()
frmcoffre.Show vbModeless
End Sub

Private Sub optBDir_Click()
frmBDir.Show
End Sub

Private Sub optDoor_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optHeal_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub opti_Click()
frmoptions.Show
frmoptions.nbcls.Text = ReadINI("INFO", "MaxClasses", App.Path & "\Classes\info.ini")
Dim i As Long
Call frmoptions.clase.Clear
For i = 0 To Val(frmoptions.nbcls.Text)
    Call frmoptions.clase.AddItem("Classe" & i, i)
Next i
If Val(ReadINI("INFO", "HPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.PV = ReadINI("INFO", "HPRegen", App.Path & "\config.ini")
If Val(ReadINI("INFO", "MPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.pm = ReadINI("INFO", "MPRegen", App.Path & "\config.ini")
If Val(ReadINI("INFO", "SPRegen", App.Path & "\config.ini")) > 0 Then frmoptions.ps = ReadINI("INFO", "SPRegen", App.Path & "\config.ini")
If Val(ReadINI("CONFIG", "Scrolling", App.Path & "\config.ini")) > 0 Then frmoptions.defl = ReadINI("CONFIG", "Scrolling", App.Path & "\config.ini")
If Val(ReadINI("CONFIG", "Scripting", App.Path & "\config.ini")) > 0 Then frmoptions.script = ReadINI("CONFIG", "Scripting", App.Path & "\config.ini")
If Val(ReadINI("INFO", "GameName", App.Path & "\config.ini")) > 0 Then frmoptions.nom = ReadINI("INFO", "GameName", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")) > 0 Then frmoptions.mj = ReadINI("INFO", "Maxplayers", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxitems", App.Path & "\config.ini")) > 0 Then frmoptions.mo = ReadINI("INFO", "Maxitems", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")) > 0 Then frmoptions.mpnj = ReadINI("INFO", "Maxnpcs", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxshops", App.Path & "\config.ini")) > 0 Then frmoptions.mm = ReadINI("INFO", "Maxshops", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxspells", App.Path & "\config.ini")) > 0 Then frmoptions.ms = ReadINI("INFO", "Maxspells", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")) > 0 Then frmoptions.mc = ReadINI("INFO", "Maxmaps", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")) > 0 Then frmoptions.moc = ReadINI("INFO", "Maxmapitems", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxemots", App.Path & "\config.ini")) > 0 Then frmoptions.Me = ReadINI("INFO", "Maxemots", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")) > 0 Then frmoptions.mn = ReadINI("INFO", "Maxlevel", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxquet", App.Path & "\config.ini")) > 0 Then frmoptions.mq = ReadINI("INFO", "Maxquet", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxguilds", App.Path & "\config.ini")) > 0 Then frmoptions.mg = ReadINI("INFO", "Maxguilds", App.Path & "\config.ini")
If Val(ReadINI("INFO", "Maxjguild", App.Path & "\config.ini")) > 0 Then frmoptions.mjg = ReadINI("INFO", "Maxjguild", App.Path & "\config.ini")
If ReadINI("INFO", "motd", App.Path & "\config.ini") > vbNullString Then frmoptions.motd = ReadINI("INFO", "motd", App.Path & "\config.ini")
End Sub

Private Sub optKill_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optNpcAvoid_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optportecode_Click()
    frmportecode.Show vbModeless
End Sub

Private Sub opttoit_Click()
On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModeless
End Sub

Private Sub optItem_Click()
    frmMapItem.scrlItem.value = 1
    frmMapItem.Show vbModeless
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModeless
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModeless
End Sub

Private Sub optNotice_Click()
    frmNotice.Show vbModeless
End Sub

Private Sub optScripted_Click()
    frmScript.Show vbModeless
End Sub

Private Sub optShop_Click()
    frmShop.scrlNum.Max = MAX_SHOPS
    frmShop.Show vbModeless
End Sub

Private Sub optSign_Click()
    frmSign.Show vbModeless
End Sub

Private Sub optSound_Click()
    frmSound.Show vbModeless
End Sub

Private Sub optSprite_Click()
    'frmSpriteChange.picSprite.Height = ((PIC_NPC1 * 32) * Screen.TwipsPerPixelY)
    frmSpriteChange.scrlItem.Max = MAX_ITEMS
    frmSpriteChange.Show vbModeless
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub

Private Sub padmin_Click()
frmadmin.Show
End Sub

Private Sub picInv_DblClick(Index As Integer)
'Dim d As Long
'
'If Player(MyIndex).inv(Inventory).num <= 0 Then Exit Sub
'
'Call SendUseItem(Inventory)
'
'For d = 1 To MAX_INV
'    If Player(MyIndex).inv(d).num > 0 Then If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then picInv(d - 1).Picture = LoadPicture()
'Next d
'Call UpdateVisInv
End Sub

Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Inventory = Index + 1
    frmMirage.SelectedItem.Top = frmMirage.picInv(Inventory - 1).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
        Call UpdateVisInv
    ElseIf Button = 2 Then
        Call DropItems
    End If
End Sub

Private Sub picInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim d As Long
'd = Index
'
'    If Player(MyIndex).inv(d + 1).num > 0 Then
'        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
'            If Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
'                itmDesc.Height = 17
'                itmDesc.Top = 224
'            Else
'                itmDesc.Height = 249
'                itmDesc.Top = 8
'            End If
'            descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
'        Else
'            If Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
'                itmDesc.Height = 161
'                itmDesc.Top = 96
'            Else
'                itmDesc.Height = 249
'                itmDesc.Top = 8
'            End If
'            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
'                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
'            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
'                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
'            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
'                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
'            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
'                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
'            ElseIf GetPlayerPetSlot(MyIndex) = d + 1 Then
'                descName.Caption = Trim$(Pets(GetPlayerInvItemNum(MyIndex, d + 1)).nom) & " (worn)"
'            Else
'                If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable <> 0 Then
'                    descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
'                Else
'                    descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name)
'                End If
'            End If
'        End If
'        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Force"
'        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " D�fense"
'        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DexReq & " Dext�rit�"
'        descHpMp.Caption = "PV: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " PM: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " End: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
'        descSD.Caption = "FOR: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
'        descMS.Caption = "Science: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSci & " Dexterit�: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDex
'        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1 <= 0 Then Usure.Caption = "Usure : Ind." Else Usure.Caption = "Usure : " & GetPlayerInvItemDur(MyIndex, d + 1) & "/" & Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1
'        desc.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
'        descName.ForeColor = Item(GetPlayerInvItemNum(MyIndex, d + 1)).NCoul
'
'        itmDesc.Visible = True
'    Else
'        itmDesc.Visible = False
'    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long
If ConOff = True And Not InEditor Then Exit Sub
    Call CheckInput(0, KeyCode, Shift)
        
    If KeyCode = vbKeyF2 Then frmPlayerHelp.Show: frmPlayerHelp.SetFocus
      
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData("cast" & SEP_CHAR & SpellMemorized & SEP_CHAR & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Vous ne pouvez lancer un sort en marchant!", BrightRed)
                End If
            End If
        Else
            Call AddText("Aucune magie m�moriser.", BrightRed)
        End If
    Else
        Call CheckInput(0, KeyCode, Shift)
    End If
    
    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)
        i = 0
        ii = 0
        Do
            If FileExiste("Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, 8, 6, 634, 479)
        i = 0
        ii = 0
        Do
            If FileExiste("Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyEnd Then
    d = GetPlayerDir(MyIndex)
        If Player(MyIndex).Moving = NO Then
            If Player(MyIndex).dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
                If d <> DIR_LEFT Then Call SendPlayerDir
            ElseIf Player(MyIndex).dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
                If d <> DIR_UP Then Call SendPlayerDir
            ElseIf Player(MyIndex).dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
                If d <> DIR_RIGHT Then Call SendPlayerDir
            ElseIf Player(MyIndex).dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
                If d <> DIR_DOWN Then Call SendPlayerDir
            End If
        End If
    End If
     KeyShift = False
    If KeyCode = vbKeyControl Then CTRLDOWN = False
End Sub

Private Sub picScreen_GotFocus()
On Error Resume Next
    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
    If (Button = 1) And InEditor And Not InPM And Me.MousePointer <> 99 Then Call SauvTemp: Call EditorMouseDown(Button, Shift, (x + NewPlayerPicX), (y + NewPlayerPicY))
    If Button = 1 And Not InEditor And Not InPM Then Call PlayerSearch(Button, Shift, (x + NewPlayerPicX), (y + NewPlayerPicY))
    If Button = 1 And InEditor And Me.MousePointer = 99 Then Call SauvTemp: Call EditorMouseDown(2, 1, (PotX + NewPlayerPicX), (PotY + NewPlayerPicY))
    If Button = 2 And InEditor Then InPM = True: Call PopupMenu(mclik)
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 1) And InEditor And Not InPM And Me.MousePointer <> 99 Then Call EditorMouseDown(Button, Shift, (x + NewPlayerPicX), (y + NewPlayerPicY))
    If Button = 1 And InEditor And Me.MousePointer = 99 Then Call EditorMouseDown(2, 1, (PotX + NewPlayerPicX), (PotY + NewPlayerPicY))
        CurX = ((x + NewPlayerPicX) \ 32 / VZoom * 3)
        CurY = ((y + NewPlayerPicY) \ 32 / VZoom * 3)
        PotX = x
        PotY = y
        frmMirage.Caption = "Editeur pour le jeu : " & Trim$(GAME_NAME) & " Mettez votre souris sur un �l�ment pour plus de d�tails."
        frmMirage.Caption = frmMirage.Caption & "X: " & Int(CurX) & " Y: " & Int(CurY)
        
        Dim borders As String
        Dim direction As String
        Dim i As Integer
        borders = ""
        For i = 0 To GetMapBordersCount(Player(MyIndex).Map) - 1
            With Map(Player(MyIndex).Map).borders(i)
                If .XSource = CurX And .YSource = CurY Then
                    If .DirectionSource = DIR_DOWN Then
                        direction = "Bas"
                    ElseIf .DirectionSource = DIR_UP Then
                        direction = "Haut"
                    ElseIf .DirectionSource = DIR_RIGHT Then
                        direction = "Droite"
                    ElseIf .DirectionSource = DIR_LEFT Then
                        direction = "Gauche"
                    End If
                    
                    borders = borders & vbCrLf & direction & " : " & .MapDestination & " : " & .XDestination & "/" & .YDestination
                End If
            End With
        Next i
        
        With mclsToolTip
            .ToolText(frmMirage.picScreen) = "X: " & Int(CurX) & " Y: " & Int(CurY) & " " & Couche & borders
        End With
'        With mclsToolTip
'            If .HasToolTip(frmMirage.picScreen) Then
'                .RemoveTool frmMirage.picScreen
'            Else
'                .ToolTipHeaderShow = False
'                .AddTool frmMirage.picScreen, "This is a long tooltip"
'                .ToolTipHeaderShow = True
'            End If
'        End With
        'frmMirage.picScreen.ToolTipText = "X: " & Int(CurX) & " Y: " & Int(CurY) & " " & Couche & borders
    If Button = 2 And InEditor Then InPM = True: Call PopupMenu(mclik)
    If CurX <> OldPCX Or CurY <> OldPCY Then Call CaseChange(CurX, CurY): OldPCX = CurX: OldPCY = CurY
End Sub

Private Sub picScreen_Click()
InPM = False
'If GetBorderSize(SourceBorder) > 0 Then
Debug.Print "debut test"
If SourceBorder.Count > 0 Then
    If SourceBorderMap = -1 Then
        Debug.Print "fin test"
        Dim value As String
        value = InputBox("Aller vers quelle map ?")
        If value <> "" And IsNumeric(value) Then
            If Int(value) >= 0 And Int(value) <= MAX_MAPS Then
                SourceBorderDirection = frmSelectDirection.ShowWithResult
            
                SourceBorderMap = Player(MyIndex).Map
                Dim firstPos As clsPosition
                Dim lastPos As clsPosition
                
                Set firstPos = SourceBorder(1)
                Set lastPos = SourceBorder(SourceBorder.Count)
                
                If firstPos.x > lastPos.x Or firstPos.y > lastPos.y Then
                    'Reorder SourceBorders
                    Dim i As Integer
                    
                    Set DestinationBorder = SourceBorder
                    Set SourceBorder = New Collection
                    
                    For i = 0 To DestinationBorder.Count - 1
                        SourceBorder.add DestinationBorder(DestinationBorder.Count - i)
                    Next i
                    
                    Set DestinationBorder = Nothing
                End If
                
                Call GoToMap(value)
                Exit Sub
            End If
        End If
        Set SourceBorder = Nothing
        Set SourceBorder = New Collection
        'End If
    '    Set SourceBorder = Nothing
    '    Set SourceBorder = New Collection
    Else ' Source already selectionned
        For i = 1 To SourceBorder.Count
            With Map(Player(MyIndex).Map)
                ReDim Preserve .borders(0 To GetMapBordersCount(Player(MyIndex).Map)) As BorderRec
                With .borders(GetMapBordersCount(Player(MyIndex).Map) - 1)
                    .XSource = DestinationBorder(i).x
                    .YSource = DestinationBorder(i).y
                    If SourceBorderDirection = DIR_DOWN Then
                        .DirectionSource = DIR_UP
                    ElseIf SourceBorderDirection = DIR_UP Then
                        .DirectionSource = DIR_DOWN
                    ElseIf SourceBorderDirection = DIR_LEFT Then
                        .DirectionSource = DIR_RIGHT
                    ElseIf SourceBorderDirection = DIR_RIGHT Then
                        .DirectionSource = DIR_LEFT
                    End If
                    
                    .MapDestination = SourceBorderMap
                    .XDestination = SourceBorder(i).x
                    .YDestination = SourceBorder(i).y
                End With
            End With
            
            With Map(SourceBorderMap)
                ReDim Preserve .borders(0 To GetMapBordersCount(SourceBorderMap)) As BorderRec
                With .borders(GetMapBordersCount(SourceBorderMap) - 1)
                    .XSource = SourceBorder(i).x
                    .YSource = SourceBorder(i).y
                    .DirectionSource = SourceBorderDirection
                    .MapDestination = Player(MyIndex).Map
                    .XDestination = DestinationBorder(i).x
                    .YDestination = DestinationBorder(i).y
                End With
            End With
        Next i

        Debug.Print "begin clear"
        Set SourceBorder = Nothing
        Set SourceBorder = New Collection
        Set DestinationBorder = Nothing
        Set DestinationBorder = New Collection
        SourceBorderMap = -1
        SourceBorderDirection = -1
        Debug.Print "end clear"
    End If
End If
End Sub

Private Sub picTiles_Click(Index As Integer)
    Dim i As Integer
    EditorSet = frmMirage.picTiles(Index).DataField

    For i = 0 To frmMirage.picTiles.Count - 1
        'frmMirage.picTiles(i).AutoRedraw = False
        frmMirage.picTiles(i).Refresh
        'frmMirage.picTiles(i).AutoRedraw = True
    Next i
    frmMirage.picTiles(Index).AutoRedraw = False
    frmMirage.picTiles(Index).Line (0, 0)-(31, 31), vbRed, B
    frmMirage.picTiles(Index).AutoRedraw = True
    'frmMirage.picTiles(Index).DrawMode = vbXorPen
    'frmMirage.picTiles(Index).DrawWidth = 1
    'frmMirage.picTiles(Index).Scale (15, 15), 15, vbRed
    'frmMirage.picTiles(Index).Circle (15, 15), 15, vbRed
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub pmodo_Click()
frmadmin.Show
End Sub

Private Sub prelv_Click()
    If frmMirage.MousePointer = 99 Then Call gom_Click
    If frmMirage.MousePointer = 2 Then frmMirage.MousePointer = 0: frmMirage.Toolbar1.buttons(32).value = tbrUnpressed Else frmMirage.MousePointer = 2
End Sub

Private Sub previsu_Click()
    If previsu.Checked Then
        WriteINI "CONFIG", "PreVisu", 0, App.Path & "\Config\Account.ini"
        previsu.Checked = False
        AccOpt.CPreVisu = False
    Else
        WriteINI "CONFIG", "PreVisu", 1, App.Path & "\Config\Account.ini"
        previsu.Checked = True
        AccOpt.CPreVisu = True
        Call PreVisua
    End If
End Sub

Private Sub propricarte_Click()
    frmMapProperties.Show vbModal, frmMirage
    'frmMapProperties.InitMPr
    InProprieter = True
    save = 1
    Call WriteINI("modif", "carte" & Player(MyIndex).Map, "1", App.Path & "\config.ini")
End Sub

Private Sub quetetimersec_Timer()
Dim Queten As Long

Queten = Val(Player(MyIndex).QueteEnCour)
If Queten <= 0 Then Exit Sub
If quete(Queten).Temps > 0 And Player(MyIndex).QueteEnCour > 0 Then

    Seco = Seco - 1
    If Seco <= 0 And Minu > 0 Then
        Seco = 59
        seconde.Caption = Seco
        Minu = Minu - 1
        If Len(CStr(Minu)) > 2 Then Minute.Caption = Minu & ":" Else Minute.Caption = "0" & Minu & ":"
    End If

    If Seco <= 0 And Minu <= 0 Then
        seconde.Caption = 0
        Call MsgBox("La qu�te : " & Trim$(quete(Queten).nom) & " est termin�e, le temps est �coul�")
        Player(MyIndex).QueteEnCour = 0
        quetetimersec.Enabled = False
        tmpsquete.Visible = False
    End If

    If Len(CStr(Seco)) > 2 Then seconde.Caption = Seco Else seconde.Caption = "0" & Seco
Else
    Player(MyIndex).QueteEnCour = 0
    tmpsquete.Visible = False
    quetetimersec.Enabled = False
End If

End Sub

Private Sub qui_Click()
If picWhosOnline.Visible Then
    Call NetPic
Else
    Call NetPic
    Call SendOnlineList
    picWhosOnline.Visible = True
    Picpics.Visible = True
End If
End Sub

Private Sub quilgn_Click()
If picWhosOnline.Visible Then
    Call NetPic
Else
    Call NetPic
    Call SendOnlineList
    picWhosOnline.Visible = True
    Picpics.Visible = True
End If
End Sub

Private Sub quit_Click()
    If save = 1 Then
        attender = MsgBox("La carte n'a pas �t� enregistr�e ni sur le serveur ni sur le disque dur voulez vous vraiment quitter?", vbYesNo)
        If attender = vbNo Then Exit Sub
        GoTo quit:
    End If
    If Val(ReadINI("modif", "carte" & Player(MyIndex).Map, App.Path & "\config.ini")) = 1 And HORS_LIGNE = 0 And Val(attender) <> vbYes Then
        attender = MsgBox("La carte n'a pas �t� enregistr�e sur le serveur voulez vous vraiment quitter?", vbYesNo)
        If attender = vbNo Then Exit Sub
    End If
quit:
Quite = True
Call GameDestroy
End Sub

Private Sub qutcour_Click()
Dim V As Boolean
V = Not picquete.Visible
If Player(MyIndex).QueteEnCour > 0 Then Call NetPic: frmMirage.picquete.Visible = V: Picpics.Visible = V
End Sub

Private Sub rempli_Click()
Dim y As Long
Dim x As Long

x = MsgBox("Es-tu sur de vouloir remplir la carte?", vbYesNo)
If x = vbNo Then Exit Sub

Call SauvTemp
If frmMirage.tp(1).Checked = True Then
    For y = 0 To MaxMapY
        For x = 0 To MaxMapX
            With Map(Player(MyIndex).Map).tile(x, y)
                If frmMirage.Toolbar1.buttons(5).value = tbrPressed Then
                    '.Ground = EditorTileY * TilesInSheets + EditorTileX
                    .Ground = EditorSet
                    '.GroundSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(6).value = tbrPressed Then
                    '.Mask = EditorTileY * TilesInSheets + EditorTileX
                    .Mask = EditorSet
                    '.MaskSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(13).value = tbrPressed Then
                    '.Anim = EditorTileY * TilesInSheets + EditorTileX
                    .Anim = EditorSet
                    '.AnimSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(7).value = tbrPressed Then
                    '.Mask2 = EditorTileY * TilesInSheets + EditorTileX
                    .Mask2 = EditorSet
                    '.Mask2Set = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(14).value = tbrPressed Then
                    '.M2Anim = EditorTileY * TilesInSheets + EditorTileX
                    .M2Anim = EditorSet
                    '.M2AnimSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(8).value = tbrPressed Then '<--
                    '.Mask3 = EditorTileY * TilesInSheets + EditorTileX
                    .Mask3 = EditorSet
                    '.Mask3Set = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(15).value = tbrPressed Then '<--
                    '.M3Anim = EditorTileY * TilesInSheets + EditorTileX
                    .M3Anim = EditorSet
                    '.M3AnimSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(9).value = tbrPressed Then
                    '.Fringe = EditorTileY * TilesInSheets + EditorTileX
                    .Fringe = EditorSet
                    '.FringeSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(16).value = tbrPressed Then
                    '.FAnim = EditorTileY * TilesInSheets + EditorTileX
                    .FAnim = EditorSet
                    '.FAnimSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(10).value = tbrPressed Then
                    '.Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                    .Fringe2 = EditorSet
                    '.Fringe2Set = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(17).value = tbrPressed Then
                    '.F2Anim = EditorTileY * TilesInSheets + EditorTileX
                    .F2Anim = EditorSet
                    '.F2AnimSet = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(11).value = tbrPressed Then '<--
                    '.Fringe3 = EditorTileY * TilesInSheets + EditorTileX
                    .Fringe3 = EditorSet
                    '.Fringe3Set = EditorSet
                ElseIf frmMirage.Toolbar1.buttons(18).value = tbrPressed Then '<--
                    '.F3Anim = EditorTileY * TilesInSheets + EditorTileX
                    .F3Anim = EditorSet
                    '.F3AnimSet = EditorSet
                End If
            End With
        Next x
    Next y
ElseIf frmMirage.tp(2).Checked = True Then
    For y = 0 To MaxMapY
        For x = 0 To MaxMapX
            'Call SetTile(Map(Player(MyIndex).Map).tile(x, y))
            Call SetTile(x, y)
        Next x
    Next y
ElseIf frmMirage.tp(3).Checked = True Then
    For y = 0 To MaxMapY
        For x = 0 To MaxMapX
            Map(Player(MyIndex).Map).tile(x, y).Light = EditorTileY * TilesInSheets + EditorTileX
        Next x
    Next y
End If
End Sub

Private Sub scrlPicture_Change()
Call EditorTileScroll
End Sub

Private Sub scrlPicture_Scroll()
Call EditorTileScroll
End Sub

Private Sub scrshot_Click()
OldVZoom = VZoom
VZoom = 9
ScreenDC = True
End Sub

Private Sub site_Click()
ShellExecute Me.hWnd, "open", "http://www.aride-online.com", vbNullString, App.Path, 1
End Sub

Private Sub siteequp_Click()
ShellExecute Me.hWnd, "open", "http://www.aride-online.com", vbNullString, App.Path, 1
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then Call IncomingData(bytesTotal)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ConOff = True And Not InEditor Then Exit Sub
    Call HandleKeypresses(KeyAscii)
    If (KeyAscii = vbKeyReturn) Then KeyAscii = 0
    'Raccourcie +/-
    If KeyAscii = 43 And InEditor Then
        Call zp_Click
    ElseIf KeyAscii = 45 And InEditor Then
        Call zm_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ConOff = True And Not InEditor Then Exit Sub
    Call CheckInput(1, KeyCode, Shift)
    If InEditor Then
        If KeyCode = vbKeyUp And hautbas.value > -1 Then hautbas.value = hautbas.value - 1
        If KeyCode = vbKeyDown And hautbas.value < hautbas.Max Then hautbas.value = hautbas.value + 1
        If KeyCode = vbKeyLeft And gauchedroite.value > 0 Then gauchedroite.value = gauchedroite.value - 1
        If KeyCode = vbKeyRight And gauchedroite.value < gauchedroite.Max Then gauchedroite.value = gauchedroite.value + 1
        'Raccourcie pgup et pgdown
        If KeyCode = vbKeyPageDown Then
            picScreen.SetFocus
            Call PageBas
        ElseIf KeyCode = vbKeyPageUp Then
            picScreen.SetFocus
            Call PageHaut
        End If
    End If
    If KeyCode = vbKeyShift Then KeyShift = True
    If KeyCode = vbKeyControl Then CTRLDOWN = True
    If KeyCode = vbKeyZ And CTRLDOWN Then Call Toolbar1_ButtonClick(Toolbar1.buttons(20))
    If KeyCode = vbKeyU And CTRLDOWN Then Call Toolbar1_ButtonClick(Toolbar1.buttons(21))
End Sub

Private Sub sort_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
End Sub

Private Sub stopmusic_Click()
Call StopMidi
End Sub

Private Sub test_Click()
Call Tester
End Sub

Private Sub Tiles_Click(Index As Integer)
    MsgBox "test"
    If Tiles(Index).Checked = False Then
        Tiles(Index).Checked = True
        'frmMirage.picBackSelect.Picture = LoadPNG(App.Path & "\GFX\Tiles" & index & ".png")
        EditorSet = Index
        Call AffTilesPic(EditorSet, frmMirage.scrlPicture.value * PIC_Y)
        'frmMirage.scrlPicture.Max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
'        HScroll1.Max = frmMirage.picBackSelect.Width / 32
'        frmMirage.picBack.Width = frmMirage.picBackSelect.Width
        frmMirage.tilescmb.ListIndex = Index
    End If
    
'    Dim i As Byte
'    For i = 0 To ExtraSheets
'        If i <> Index Then Tiles(i).Checked = False
'    Next i
End Sub

Private Sub tilescmb_Click()
    Dim i, x As Integer
    
    'frmMirage.picBackSelect.Picture = LoadPNG(App.Path & "\GFX\Tiles" & Val(tilescmb.ListIndex) & ".png")
    EditorSet = Val(tilescmb.ListIndex)
    Call AffTilesPic(EditorSet, frmMirage.scrlPicture.value * PIC_Y)
'    Tiles(tilescmb.ListIndex).Checked = True
'    HScroll1.Max = frmMirage.picBackSelect.Width / 32
    'frmMirage.scrlPicture.max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
    'frmMirage.picBack.Width = frmMirage.picBackSelect.Width
    
'    For i = 0 To ExtraSheets
'        If i <> tilescmb.ListIndex Then Tiles(i).Checked = False
'    Next i
    
    'Set DD_Temp = LoadImage(App.Path & "\GFX\tiles" & EditorSet & ".png", DD, DDSD_Temp)
    'SetMaskColorFromPixel DD_Temp, 0, 0

    'frmMirage.picTiles(0).Visible = False
    For i = 0 To frmMirage.picTiles.Count - 1
        'MsgBox i
        'Unload frmMirage.picTiles(i)
        frmMirage.picTiles(i).Visible = False
    Next i
    
    Dim TileFiles() As String
    Call ListFiles(App.Path & "\GFX\Tiles\" & tilescmb.Text, TileFiles)
    If GetArraySize(TileFiles) > 0 Then
        'ReDim DDSD_TilesTemp(0 To GetArraySize(TileFiles) - 1)
        For i = 0 To GetArraySize(TileFiles) - 1
            If i + 1 > frmMirage.picTiles.Count Then
                Load frmMirage.picTiles(i)
            End If
            x = Int(i / 7)
    '        frmMirage.picTiles(i).Top = 8 + 40 * x
    '        frmMirage.picTiles(i).Left = 8 + (i - x * 3) * 40
            frmMirage.picTiles(i).Top = frmMirage.picTiles(0).Top + 32 * x
            frmMirage.picTiles(i).Left = frmMirage.picTiles(0).Left + (i - x * 7) * 32
            
            DDSD_TilesTemp.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            DDSD_TilesTemp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_Temp = LoadImage(App.Path & "\GFX\Tiles\" & tilescmb.Text & "\" & TileFiles(i), DD, DDSD_TilesTemp)
            SetMaskColorFromPixel DD_Temp, 0, 0
            
            Call AffSurfPic(DD_Temp, frmMirage.picTiles(i), 0, 0)
            
            frmMirage.picTiles(i).DataField = Left(Right(TileFiles(i), Len(TileFiles(i)) - 5), Len(TileFiles(i)) - 9)
            frmMirage.picTiles(i).Visible = True
        Next i
    End If
    'Call PreVisua
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If mediaplayer.URL > vbNullString Then If mediaplayer.Controls.currentPosition = 0 And mediaplayer.currentMedia.name = Mid$(Map(Player(MyIndex).Map).Music, 1, Len(Map(Player(MyIndex).Map).Music) - 4) Then Call frmMirage.mediaplayer.Controls.Play
End Sub

Private Sub Timer2_Timer()
Call GameDestroy
End Sub

Private Sub timerbar_Timer()
If frmMirage.Visible = True Then
    
    If Player(MyIndex).y < ((picScreen.Height / 32) \ 2) Then
        hautbas.value = hautbas.min
    Else
        If ((picScreen.Height \ 32) \ 2) + (Player(MyIndex).y - (picScreen.Height \ 32)) < hautbas.Max Then hautbas.value = ((picScreen.Height \ 32) \ 2) + (Player(MyIndex).y - (picScreen.Height \ 32)) Else hautbas.value = hautbas.Max
    End If
    
    If Player(MyIndex).x < ((picScreen.Width \ 32) \ 2) Then
        gauchedroite.value = gauchedroite.min
    Else
        If ((picScreen.Width \ 32) \ 2) + 1 + (Player(MyIndex).x - (picScreen.Width \ 32)) < gauchedroite.Max Then gauchedroite.value = ((picScreen.Width \ 32) \ 2) + 1 + (Player(MyIndex).x - (picScreen.Width \ 32)) Else gauchedroite.value = gauchedroite.Max
    End If
End If
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then tmrRainDrop.Enabled = False: Exit Sub
    If BLT_RAIN_DROPS > 0 Then If DropRain(BLT_RAIN_DROPS).Randomized = False Then Call RNDRainDrop(BLT_RAIN_DROPS)
    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1
    If tmrRainDrop.Interval > 30 Then tmrRainDrop.Interval = tmrRainDrop.Interval - 10
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then tmrSnowDrop.Enabled = False: Exit Sub
    If BLT_SNOW_DROPS > 0 Then If DropSnow(BLT_SNOW_DROPS).Randomized = False Then Call RNDSnowDrop(BLT_SNOW_DROPS)
    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1
    If tmrSnowDrop.Interval > 30 Then tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Numbut As Long
Dim i As Long

Numbut = Button.Index

If Numbut < 0 Or Numbut > 34 Then Exit Sub

If Numbut = 1 Then Call Tester

If Numbut = 2 Then
    If GettingMap Then Exit Sub
    If HORS_LIGNE = 1 Then Call enregcarte_Click Else Call envoicarte_Click
    Call SendData("mapreport" & SEP_CHAR & END_CHAR)
    Exit Sub
End If

If Numbut = 3 Then
Dim PathServ As String
    If LCase$(dir$(Mid$(App.Path, 1, Len(App.Path) - 7) & "Serveur", vbDirectory)) <> "serveur" Then
        Call MsgBox("Le Dossier du serveur est introuvable!!", , "Erreur")
    Else
        PathServ = Mid$(App.Path, 1, Len(App.Path) - 7) & "Serveur"
        Call Shell(PathServ & "\Editeur de script.exe", vbNormalFocus)
    End If
Exit Sub
End If

If Numbut > 4 And Numbut < 19 Then
    Couche = Toolbar1.buttons(Numbut).ToolTipText
    For i = 5 To 18
        If i <> Numbut Then Toolbar1.buttons(i).value = tbrUnpressed
    Next i
    Numbut = 0
    For i = 5 To 18
        If Toolbar1.buttons(i).value = tbrPressed Then Numbut = i: TCouche = i
    Next i
    If Numbut <= 0 Then Toolbar1.buttons(5).value = tbrPressed
    Exit Sub
End If

If Numbut = 20 Then
'   TODO : Voir l'utilit�

'    If TempNum = 5 Then Toolbar1.buttons(20).Enabled = False: Exit Sub
'    If Val(TempMap(TempNum + 1).Revision) <> -1 Then
'        If TempNum = 0 Then TempMap(0) = Map(Player(MyIndex).Map)
'        TempNum = TempNum + 1
'        Map(Player(MyIndex).Map) = TempMap(TempNum)
'        If TempNum = 5 Then Toolbar1.buttons(20).Enabled = False Else If TempMap(TempNum + 1).Revision <> -1 Then Toolbar1.buttons(20).Enabled = True Else Toolbar1.buttons(20).Enabled = False
'        Toolbar1.buttons(21).Enabled = True
'    End If
End If

If Numbut = 21 Then
    If TempNum = 0 Then Toolbar1.buttons(21).Enabled = False: Exit Sub
    TempNum = TempNum - 1
    Map(Player(MyIndex).Map) = TempMap(TempNum)
    If TempNum = 0 Then Toolbar1.buttons(21).Enabled = False Else Toolbar1.buttons(21).Enabled = True
    If TempNum < 5 Then Toolbar1.buttons(20).Enabled = True
End If

If Numbut = 23 And tp(1).Checked = False Then
    Toolbar1.buttons(24).value = tbrUnpressed
    Toolbar1.buttons(25).value = tbrUnpressed
    Call tp_Click(1)
    tp(1).Checked = True
    Toolbar1.buttons(Numbut).value = tbrPressed
    Exit Sub
ElseIf Numbut = 23 Then Toolbar1.buttons(Numbut).value = tbrPressed
End If

If Numbut = 24 And tp(2).Checked = False Then
    Toolbar1.buttons(23).value = tbrUnpressed
    Toolbar1.buttons(25).value = tbrUnpressed
    Call tp_Click(2)
    tp(2).Checked = True
    Toolbar1.buttons(Numbut).value = tbrPressed
    Exit Sub
ElseIf Numbut = 24 Then Toolbar1.buttons(Numbut).value = tbrPressed
End If

If Numbut = 25 And tp(3).Checked = False Then
    Toolbar1.buttons(23).value = tbrUnpressed
    Toolbar1.buttons(24).value = tbrUnpressed
    Call tp_Click(3)
    Toolbar1.buttons(Numbut).value = tbrPressed
    Exit Sub
ElseIf Numbut = 25 Then Toolbar1.buttons(Numbut).value = tbrPressed
End If

If Numbut = 27 Then VZoom = 3: picScreen.Refresh: Toolbar1.buttons(28).value = tbrUnpressed: Toolbar1.buttons(29).value = tbrUnpressed

If Numbut = 28 Then VZoom = 6: picScreen.Refresh: Toolbar1.buttons(27).value = tbrUnpressed: Toolbar1.buttons(29).value = tbrUnpressed

If Numbut = 29 Then VZoom = 9: picScreen.Refresh: Toolbar1.buttons(28).value = tbrUnpressed: Toolbar1.buttons(27).value = tbrUnpressed

If Numbut = 31 Then Call rempli_Click

If Numbut = 32 Then Call prelv_Click

If Numbut = 33 Then Call gom_Click

If Numbut = 34 Then Call EditorClearLayer

End Sub

Private Sub tp_Click(Index As Integer)
Dim i As Byte

Attributs.Visible = False

If Index = 1 Then
ElseIf Index = 2 Then
    Attributs.Visible = True
ElseIf Index = 3 Then
End If

For i = 1 To 3
    If i <> Index Then tp(i).Checked = False
Next i

'Dim i As Byte
'
'    tp(Index).Checked = True
'    If Index = 1 Then
'        If tp(1).Checked = True Then
'            Attributs.Visible = False
''            For i = 1 To ExtraSheets
''                Tiles(i).Checked = False
''            Next i
'            Tiles(OldTiles).Checked = True
'            'frmMirage.picBackSelect.Picture = LoadPNG(App.Path & "\GFX\Tiles" & OldTiles & ".png")
'            EditorSet = OldTiles
'            Call AffTilesPic(EditorSet, frmMirage.scrlPicture.value * PIC_Y)
'            tilescmb.ListIndex = OldTiles
'            'frmMirage.scrlPicture.Max = ((DDSD_Tile(EditorSet).lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
'            frmMirage.picBack.Width = frmMirage.picBackSelect.Width
'            tile.Enabled = True
'            For i = 5 To 18
'                Toolbar1.buttons(i).Enabled = True
'            Next i
'            For i = 26 To 34
'                Toolbar1.buttons(i).Enabled = True
'            Next i
'        End If
'    ElseIf Index = 2 Then
'        If tp(2).Checked = True Then
'            Attributs.Visible = True
'            frmMirage.shpSelected.Width = 32
'            frmMirage.shpSelected.Height = 32
'            tile.Enabled = True
'            For i = 5 To 30
'                If i <> 20 And i <> 21 And i <> 23 And i <> 24 And i <> 25 And i <> 27 And i <> 28 And i <> 29 Then Toolbar1.buttons(i).Enabled = False
'            Next i
'            Toolbar1.buttons(34).Enabled = False
'        End If
'    Else
'        If tp(3).Checked = True Then
'            Attributs.Visible = False
''            For i = 0 To ExtraSheets - 1
''                If Tiles(i).Checked = True Then OldTiles = i
''                Tiles(i).Checked = False
''            Next i
''            Tiles(ExtraSheets).Checked = True
''            tilescmb.ListIndex = ExtraSheets
''            For i = 0 To ExtraSheets
''                If i <> ExtraSheets Then frmMirage.Tiles(i).Checked = False
''            Next i
'            'frmMirage.picBackSelect.Picture = LoadPNG(App.Path & "\GFX\Tiles" & 6 & ".png")
'            EditorSet = ExtraSheets
'            Call AffOutilPic(frmMirage.scrlPicture.value * PIC_Y)
'            frmMirage.scrlPicture.Max = ((DDSD_Outil.lHeight - frmMirage.picBackSelect.Height) \ PIC_Y)
'            frmMirage.picBack.Width = frmMirage.picBackSelect.Width
'            tile.Enabled = False
'            For i = 5 To 18
'                Toolbar1.buttons(i).Enabled = True
'            Next i
'            For i = 26 To 34
'                Toolbar1.buttons(i).Enabled = True
'            Next i
'        End If
'    End If
'
'    For i = 1 To 3
'        If i <> Index Then tp(i).Checked = False
'    Next i
'
End Sub

Private Sub lblUseItem_Click()
'Dim d As Long
'
'If Player(MyIndex).inv(Inventory).num <= 0 Then Exit Sub
'
'Call SendUseItem(Inventory)
'
'For d = 1 To MAX_INV
'    If Player(MyIndex).inv(d).num > 0 Then If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then picInv(d - 1).Picture = LoadPicture()
'Next d
'Call UpdateVisInv
End Sub

Private Sub lblDropItem_Click()
    Call DropItems
End Sub

Sub DropItems()
Dim InvNum As Long
Dim GoldAmount As String
On Error GoTo Done
If Inventory <= 0 Then Exit Sub

    InvNum = Inventory
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_MONTURE Then Call SendUseItem(Inventory)
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Empilable <> 0 Then
            If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                GoldAmount = InputBox("Combien de " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") voulez vous jeter?", "Jeter " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmMirage.Left, frmMirage.Top)
            Else
                GoldAmount = 1
            End If
            If IsNumeric(GoldAmount) Then Call SendDropItem(InvNum, GoldAmount)
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then MsgBox "Entrez un chiffre SVP!"
End Sub


Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & lstSpells.ListIndex & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Vous pouvez lancer un sort en marchant!", BrightRed)
            End If
        End If
    Else
        Call AddText("Aucuns sort ici.", BrightRed)
    End If
End Sub

Private Sub cmdAccess_Click()
Dim Packet As String
    If txtName.Text = vbNullString Or txtAccess.Text = vbNullString Or Not IsNumeric(txtAccess.Text) Then Exit Sub
    Packet = "GUILDCHANGEACCESS" & SEP_CHAR & txtName.Text & SEP_CHAR & txtAccess.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdDisown_Click()
Dim Packet As String
    If txtName.Text = vbNullString Then Exit Sub
    Packet = "GUILDDISOWN" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdTrainee_Click()
Dim Packet As String
    If txtName.Text = vbNullString Then Exit Sub
    Packet = "GUILDTRAINEE" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub txtMyTextBox_Change()
If txtMyTextBox <> vbNullString And InEditor Then
    If Mid$(txtMyTextBox.Text, Len(txtMyTextBox), 1) = "+" Or Mid$(txtMyTextBox.Text, Len(txtMyTextBox), 1) = "-" Then txtMyTextBox.Text = Mid$(txtMyTextBox.Text, 1, Len(txtMyTextBox) - 1)
    If Mid$(txtMyTextBox.Text, 1, 1) = "+" Or Mid$(txtMyTextBox.Text, 1, 1) = "-" Then txtMyTextBox.Text = Mid$(txtMyTextBox.Text, 2)
End If
End Sub

Private Sub txtQ_KeyPress(KeyAscii As Integer)
txtQ.Visible = False
 KeyAscii = 0
End Sub

Private Sub TxtQ2_KeyPress(KeyAscii As Integer)
txtQ.Visible = False
KeyAscii = 0
End Sub

Private Sub Up_Click()
If VScroll1.value = 0 Then Exit Sub
    VScroll1.value = VScroll1.value - 1
    Picture9.Top = VScroll1.value * -PIC_Y
End Sub

Private Sub Down_Click()
Dim x As Byte
x = Int(MAX_INV / 8)
x = x + 1
If (x * 8) < MAX_INV Then x = x + 1
If VScroll1.value = x Then Exit Sub
    VScroll1.value = VScroll1.value + 1
    Picture9.Top = VScroll1.value * -PIC_Y
End Sub

'Private Sub picBackSelect_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyShift Then KeyShift = True
'End Sub
'
'Private Sub picBackSelect_KeyUp(KeyCode As Integer, Shift As Integer)
'    KeyShift = False
'End Sub

'Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
'    If Button = 1 Then
'        If KeyShift = False Then
'            Call EditorChooseTile(Button, Shift, x, y)
'            shpSelected.Width = 32
'            shpSelected.Height = 32
'        Else
'            EditorTileX = (x \ PIC_X)
'            EditorTileY = (y \ PIC_Y)
'
'            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
'                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
'                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
'            Else
'                If shpSelected.Width > PIC_X Then
'                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
'                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
'                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
'                    End If
'                End If
'            End If
'
'            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
'                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
'                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
'            Else
'                If shpSelected.Height > PIC_Y Then
'                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
'                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
'                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
'                    End If
'                End If
'            End If
'            EditorTileX = (shpSelected.Left \ PIC_X)
'            EditorTileY = (shpSelected.Top \ PIC_Y) + frmMirage.scrlPicture.value
'        End If
'    End If
'
'    If frmMirage.tp(2).Checked = True Then shpSelected.Width = 32: shpSelected.Height = 32
'    If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call PreVisua
'    If Button = 2 And Not frmTile.Visible Then Call AffSurfPic(DD_TileSurf(EditorSet), frmTile.picTile, 0, 0): frmTile.Defile.Max = Int((DDSD_Tile(EditorSet).lHeight - frmTile.picTile.Height) \ PIC_Y): frmTile.Defile.value = scrlPicture.value: frmTile.shpSelected.Width = shpSelected.Width: frmTile.shpSelected.Height = shpSelected.Height: frmTile.Show vbModeless, frmMirage
'    'EditorTileX = ((shpSelected.Left + PIC_X) \ PIC_X)
'    'EditorTileY = ((shpSelected.Top + PIC_Y) \ PIC_Y)
'End Sub
'
'Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 1 Then
'       ' If KeyShift = False Then
'        '    Call EditorChooseTile(Button, Shift, x, y)
'         '   shpSelected.Width = 32
'          '  shpSelected.Height = 32
'           ' If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call PreVisua
'        'Else
'            EditorTileX = (x \ PIC_X)
'            EditorTileY = (y \ PIC_Y)
'
'            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
'                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
'                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
'            Else
'                If shpSelected.Width > PIC_X Then
'                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
'                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
'                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
'                    End If
'                End If
'            End If
'
'            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
'                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
'                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
'            Else
'                If shpSelected.Height > PIC_Y Then
'                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
'                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
'                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
'                    End If
'                End If
'            End If
'            EditorTileX = (shpSelected.Left \ PIC_X)
'            EditorTileY = (shpSelected.Top \ PIC_Y) + frmMirage.scrlPicture.value
'            If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call PreVisua
'        'End If
'    End If
'
'    If frmMirage.tp(2).Checked = True Then shpSelected.Width = 32: shpSelected.Height = 32
'End Sub

Private Sub vies_Click()
Dim V As Boolean
V = Not vieetc.Visible
Call NetPic
vieetc.Visible = V
Picpics.Visible = V
End Sub

Public Sub NetPic()
picquete.Visible = False
picGuildAdmin.Visible = False
picInv3.Visible = False
Picture1.Visible = False
picEquip.Visible = False
picPlayerSpells.Visible = False
picWhosOnline.Visible = False
vieetc.Visible = False
Picpics.Visible = False
End Sub

Private Sub zm_Click()
If VZoom < 9 Then
    Toolbar1.buttons(27).value = tbrUnpressed
    Toolbar1.buttons(28).value = tbrUnpressed
    Toolbar1.buttons(29).value = tbrUnpressed
    VZoom = VZoom + 3
    picScreen.Refresh
    Toolbar1.buttons(27 + ((VZoom \ 3) - 1)).value = tbrPressed
End If
End Sub

Private Sub zp_Click()
If VZoom > 3 Then
    Toolbar1.buttons(27).value = tbrUnpressed
    Toolbar1.buttons(28).value = tbrUnpressed
    Toolbar1.buttons(29).value = tbrUnpressed
    VZoom = VZoom - 3
    picScreen.Refresh
    Toolbar1.buttons(27 + ((VZoom \ 3) - 1)).value = tbrPressed
End If
End Sub

Private Sub PageBas()
Dim i As Long
    If TCouche <= 0 Then TCouche = 5: Exit Sub
    If TCouche = 5 Then Exit Sub
    If TCouche = 13 Then TCouche = TCouche - 1
    TCouche = TCouche - 1
    For i = 5 To 18
        Toolbar1.buttons(i).value = tbrUnpressed
    Next i
    Toolbar1.buttons(TCouche).value = tbrPressed
    Couche = Toolbar1.buttons(TCouche).ToolTipText
    frmMirage.picScreen.ToolTipText = "X: " & Int(CurX) & " Y: " & Int(CurY) & " " & Couche
End Sub

Private Sub PageHaut()
Dim i As Long
    If TCouche <= 0 Then TCouche = 5
    If TCouche = 18 Then Exit Sub
    If TCouche = 11 Then TCouche = TCouche + 1
    TCouche = TCouche + 1
    For i = 5 To 18
        Toolbar1.buttons(i).value = tbrUnpressed
    Next i
    Toolbar1.buttons(TCouche).value = tbrPressed
    Couche = Toolbar1.buttons(TCouche).ToolTipText
    frmMirage.picScreen.ToolTipText = "X: " & Int(CurX) & " Y: " & Int(CurY) & " " & Couche
End Sub

