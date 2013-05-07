VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   10020
   ClientLeft      =   4455
   ClientTop       =   750
   ClientWidth     =   12000
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
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   668
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Visible         =   0   'False
   Begin VB.PictureBox dragDropPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   840
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   266
      Top             =   120
      Width           =   480
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
      Left            =   9720
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   57
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Description de l'objet"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   69
         ToolTipText     =   "Nom de l'objet"
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Requière-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   68
         ToolTipText     =   "Force/défense/vitesse requise pour équipper l'objet"
         Top             =   240
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
         TabIndex        =   67
         Top             =   480
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
         TabIndex        =   66
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label descDex 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Dex"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   65
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Donne-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   64
         ToolTipText     =   "Se que vous apporte l'objet"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   63
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   62
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Description-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   60
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Science: XXXXX Dex: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   59
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Usure 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Usure : XXXX/XXXX"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   58
         Top             =   2160
         Width           =   2655
      End
   End
   Begin VB.Frame picInvitation 
      Height          =   1455
      Index           =   3
      Left            =   9840
      TabIndex        =   245
      Top             =   120
      Width           =   1935
      Begin VB.TextBox messageInvitation 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   248
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdOui 
         Caption         =   "Oui"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   247
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdNon 
         Caption         =   "Non"
         Height          =   315
         Index           =   3
         Left            =   1200
         TabIndex        =   246
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.PictureBox pictTouche 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   1920
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   116
      Top             =   -240
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton cmdOTA 
         Caption         =   "Annuler"
         Height          =   255
         Left            =   6840
         TabIndex        =   117
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton cmdRegisterCommands 
         Caption         =   "Ok"
         Height          =   255
         Left            =   6120
         TabIndex        =   118
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   13
         Left            =   4920
         TabIndex        =   293
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   4920
         TabIndex        =   292
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   4920
         TabIndex        =   291
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   4920
         TabIndex        =   290
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   4920
         TabIndex        =   289
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   4920
         TabIndex        =   288
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   4920
         TabIndex        =   287
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   4920
         TabIndex        =   286
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   4920
         TabIndex        =   285
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   4920
         TabIndex        =   284
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   4920
         TabIndex        =   283
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   4920
         TabIndex        =   282
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   4920
         TabIndex        =   281
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblCommandRac 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   4920
         TabIndex        =   280
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCommandAction 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   279
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblCommandRamasser 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   278
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblCommandCourir 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   277
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblCommandAttaque 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   276
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblCommandDroite 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   275
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblCommandGauche 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   274
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblCommandBas 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   273
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblCommandHaut 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   272
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 14 :"
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
         Index           =   20
         Left            =   3960
         TabIndex        =   237
         Top             =   3390
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 13 :"
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
         Index           =   19
         Left            =   3960
         TabIndex        =   236
         Top             =   3150
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 12 :"
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
         Index           =   18
         Left            =   3960
         TabIndex        =   235
         Top             =   2910
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 11 :"
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
         Index           =   17
         Left            =   3960
         TabIndex        =   234
         Top             =   2670
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 10 :"
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
         Index           =   16
         Left            =   3960
         TabIndex        =   233
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 8 :"
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
         Index           =   15
         Left            =   3960
         TabIndex        =   232
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 7 :"
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
         Index           =   14
         Left            =   3960
         TabIndex        =   231
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 6 :"
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
         Index           =   13
         Left            =   3960
         TabIndex        =   230
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 5 :"
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
         Index           =   12
         Left            =   3960
         TabIndex        =   229
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 4 :"
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
         Index           =   11
         Left            =   3960
         TabIndex        =   228
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 3 :"
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
         Index           =   10
         Left            =   3960
         TabIndex        =   227
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 2 :"
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
         Index           =   9
         Left            =   3960
         TabIndex        =   226
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 1 :"
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
         Index           =   8
         Left            =   3960
         TabIndex        =   225
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Action :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   224
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ramasser :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   223
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Courir :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   222
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Attaque :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   221
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Droite :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   220
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gauche :"
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
         Index           =   2
         Left            =   120
         TabIndex        =   219
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bas :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   218
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "- Touche de Jeu -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   240
         TabIndex        =   122
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "- Touche de Racourci -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3960
         TabIndex        =   121
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Haut :"
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
         Index           =   0
         Left            =   120
         TabIndex        =   120
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 9 :"
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
         Left            =   3960
         TabIndex        =   119
         Top             =   2190
         Width           =   855
      End
   End
   Begin VB.Frame picInvitation 
      Height          =   1455
      Index           =   2
      Left            =   9840
      TabIndex        =   206
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdNon 
         Caption         =   "Non"
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   209
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdOui 
         Caption         =   "Oui"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   208
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox messageInvitation 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   207
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame picInvitation 
      Height          =   1455
      Index           =   1
      Left            =   9840
      TabIndex        =   201
      Top             =   120
      Width           =   1935
      Begin VB.TextBox messageInvitation 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   204
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdOui 
         Caption         =   "Oui"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   203
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdNon 
         Caption         =   "Non"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   202
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   9120
      Picture         =   "frmMirage.frx":000C
      ScaleHeight     =   263.107
      ScaleMode       =   0  'User
      ScaleWidth      =   2565
      TabIndex        =   139
      Top             =   1800
      Visible         =   0   'False
      Width           =   2595
      Begin VB.PictureBox Picture16 
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
         Left            =   960
         Picture         =   "frmMirage.frx":21A3C
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   249
         Top             =   2520
         Width           =   270
      End
      Begin VB.PictureBox backPPMana 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   170
         Index           =   1
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   2175
         TabIndex        =   157
         Top             =   1560
         Width           =   2175
         Begin VB.Label lblPPMana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PM : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   165
            Index           =   1
            Left            =   0
            TabIndex        =   158
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape shpPPMana 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox backPPLife 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   170
         Index           =   1
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   2175
         TabIndex        =   155
         Top             =   1292
         Width           =   2175
         Begin VB.Label lblPPLife 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PV : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   1
            Left            =   0
            TabIndex        =   156
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape shpPPLife 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox backPPMana 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   170
         Index           =   0
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   2175
         TabIndex        =   153
         Top             =   840
         Width           =   2175
         Begin VB.Label lblPPMana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PM : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   165
            Index           =   0
            Left            =   0
            TabIndex        =   154
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape shpPPMana 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox backPPLife 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   170
         Index           =   0
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   2175
         TabIndex        =   151
         Top             =   600
         Width           =   2175
         Begin VB.Label lblPPLife 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PV : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   0
            Left            =   0
            TabIndex        =   152
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape shpPPLife 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox backPPLife 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   170
         Index           =   2
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   2175
         TabIndex        =   149
         Top             =   2040
         Width           =   2175
         Begin VB.Label lblPPLife 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PV : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   2
            Left            =   0
            TabIndex        =   150
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape shpPPLife 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture15 
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
         Left            =   1320
         Picture         =   "frmMirage.frx":21CC7
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   142
         Top             =   2520
         Width           =   270
      End
      Begin VB.PictureBox backPPMana 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   170
         Index           =   2
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   2175
         TabIndex        =   147
         Top             =   2280
         Width           =   2175
         Begin VB.Label lblPPMana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PM : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   165
            Index           =   2
            Left            =   0
            TabIndex        =   148
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape shpPPMana 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Label lbl_party_numPage 
         BackStyle       =   0  'Transparent
         Caption         =   "1/1"
         Height          =   255
         Left            =   720
         TabIndex        =   251
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " Page :                                  "
         Height          =   375
         Left            =   120
         MousePointer    =   5  'Size
         TabIndex        =   250
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rejoindre/Quitter le groupe"
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
         Left            =   0
         TabIndex        =   146
         Top             =   2760
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblPPName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   240
         TabIndex        =   145
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblPPName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   240
         TabIndex        =   144
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblPPName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   240
         TabIndex        =   143
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   1920
         TabIndex        =   141
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   2280
         TabIndex        =   140
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   7125
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   92
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   91
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   6045
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   90
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   5505
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   89
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   4965
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   88
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   4425
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   87
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   86
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   3345
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   85
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   2805
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   84
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2265
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   83
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   82
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   81
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   645
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   80
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   105
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   79
      Top             =   9315
      Width           =   480
   End
   Begin VB.ComboBox Canal 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmMirage.frx":21F5F
      Left            =   120
      List            =   "frmMirage.frx":21F6F
      TabIndex        =   78
      Text            =   "Carte"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1335
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   70
      Top             =   8760
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Timer quetetimersec 
      Enabled         =   0   'False
      Left            =   9240
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   8400
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.PictureBox Picturesprite 
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
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   9000
      ScaleHeight     =   359
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   94
      Top             =   1680
      Visible         =   0   'False
      Width           =   2625
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
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
         Left            =   120
         TabIndex        =   95
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton CmdoptTouche 
         Caption         =   "Configurer les touches"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CheckBox chknobj 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nom des objets aux sol (quand la souris le survole)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   110
         ToolTipText     =   "Petit barre afficher au dessu de vous"
         Top             =   960
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerbar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mini barre de vie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   720
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkplayername 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chknpcname 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Noms"
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
         Left            =   120
         TabIndex        =   107
         Top             =   1440
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkbubblebar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bulles de dialogue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chknpcbar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Affichés leur mini barre de vie"
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
         Left            =   120
         TabIndex        =   105
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerdamage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dégâts affichés au dessus de la tête"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   104
         Top             =   480
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.CheckBox chknpcdamage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dégâts affichés au dessus de la tête"
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
         Left            =   120
         TabIndex        =   103
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkmusic 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Musique"
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
         Left            =   120
         TabIndex        =   102
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chksound 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Effets sonores"
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
         Left            =   120
         TabIndex        =   101
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkAutoScroll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Défilement automatique"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   3960
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualiser le thème"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   4
         TabIndex        =   98
         Top             =   3600
         Value           =   6
         Width           =   2055
      End
      Begin VB.CheckBox chkLowEffect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Désactiver les effets avancé"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   4245
         Width           =   2325
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de ligne écrite sur l'écran: 6"
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
         Left            =   120
         TabIndex        =   115
         Top             =   3420
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage du Joueur-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   0
         TabIndex        =   114
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Musique/Sons-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   0
         TabIndex        =   113
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage du Chat-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   0
         TabIndex        =   112
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage des NPCs-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   0
         TabIndex        =   111
         Top             =   1275
         Width           =   2655
      End
   End
   Begin VB.PictureBox PicMenuQuitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   6720
      Picture         =   "frmMirage.frx":21F8F
      ScaleHeight     =   1965
      ScaleWidth      =   3600
      TabIndex        =   123
      Top             =   -840
      Visible         =   0   'False
      Width           =   3600
      Begin VB.Label lblCdP 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   127
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDeco 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   126
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblQuitter 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   125
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3360
         TabIndex        =   124
         Top             =   0
         Width           =   255
      End
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
      Height          =   9120
      Left            =   0
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.PictureBox picInv3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   8760
         Picture         =   "frmMirage.frx":24B41
         ScaleHeight     =   265.778
         ScaleMode       =   0  'User
         ScaleWidth      =   2595
         TabIndex        =   252
         Top             =   6240
         Visible         =   0   'False
         Width           =   2595
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
            Height          =   1935
            Left            =   240
            ScaleHeight     =   129
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   137
            TabIndex        =   254
            Top             =   480
            Width           =   2055
            Begin VB.PictureBox Picture9 
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
               Height          =   7935
               Left            =   0
               ScaleHeight     =   529
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   126
               TabIndex        =   255
               Top             =   0
               Width           =   1890
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
                  TabIndex        =   256
                  Top             =   120
                  Width           =   480
               End
               Begin VB.Shape IDAD 
                  BorderColor     =   &H00008000&
                  BorderWidth     =   3
                  Height          =   510
                  Left            =   0
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   510
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
            Height          =   1935
            LargeChange     =   50
            Left            =   2280
            Max             =   350
            SmallChange     =   30
            TabIndex        =   253
            Top             =   480
            Width           =   255
         End
         Begin VB.Label lbl_deplacer_inv 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   0
            TabIndex        =   261
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lbl_reduire_inv 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1920
            TabIndex        =   260
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbl_fermer_inv 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2280
            TabIndex        =   259
            Top             =   0
            Width           =   255
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
            Left            =   240
            TabIndex        =   258
            Top             =   2640
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
            Left            =   1680
            TabIndex        =   257
            Top             =   2640
            Width           =   555
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   2640
            Y1              =   224.375
            Y2              =   224.375
         End
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5250
         Left            =   3600
         Picture         =   "frmMirage.frx":46571
         ScaleHeight     =   5250
         ScaleWidth      =   4500
         TabIndex        =   159
         Top             =   600
         Width           =   4500
         Begin VB.PictureBox Picture1 
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
            Height          =   555
            Left            =   2520
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   170
            Top             =   2640
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
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   171
               Top             =   30
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture5 
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
            Height          =   555
            Left            =   1920
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   168
            Top             =   2640
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
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   169
               Top             =   30
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture4 
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
            Height          =   555
            Left            =   1920
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   166
            Top             =   960
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
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   167
               Top             =   30
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture6 
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
            Height          =   555
            Left            =   1320
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   164
            Top             =   1560
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
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   165
               Top             =   30
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture7 
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
            Height          =   555
            Left            =   2520
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   162
            Top             =   1560
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
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   163
               Top             =   30
               Width           =   495
            End
         End
         Begin VB.PictureBox Picsprt 
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
            Height          =   555
            Left            =   1920
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   160
            Top             =   1560
            Width           =   555
            Begin VB.PictureBox Picsprts 
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
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   161
               Top             =   30
               Width           =   495
            End
         End
         Begin VB.Label lblLANGBonus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   271
            ToolTipText     =   "Points permettant d'augmenter vos chances d'esquive"
            Top             =   4440
            Width           =   270
         End
         Begin VB.Label lblSCIBonus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   270
            ToolTipText     =   "Points permettant d'augmenter vos sorts disponibles "
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label lblDEXBonus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   269
            ToolTipText     =   "Points permettant d'augmenter vos sorts disponibles "
            Top             =   3960
            Width           =   270
         End
         Begin VB.Label lblDEFBonus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   2040
            TabIndex        =   268
            ToolTipText     =   "Points permettant d'augmenter votre résistance et vos chances de bloquer"
            Top             =   3720
            Width           =   270
         End
         Begin VB.Label lblSTRBonus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   267
            ToolTipText     =   "Points permettant d'augmenter vos dégâts et vos chances de coup critique"
            Top             =   3480
            Width           =   270
         End
         Begin VB.Label AddDex 
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
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   263
            Top             =   3960
            Width           =   165
         End
         Begin VB.Label lblDEX 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   262
            ToolTipText     =   "Points permettant d'augmenter vos sorts disponibles "
            Top             =   3960
            Width           =   270
         End
         Begin VB.Label fermer 
            BackStyle       =   0  'Transparent
            Height          =   135
            Left            =   3720
            TabIndex        =   184
            Top             =   5040
            Width           =   735
         End
         Begin VB.Label ifermer 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   255
            Left            =   4200
            TabIndex        =   182
            Top             =   0
            Width           =   255
         End
         Begin VB.Label ireduire 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   255
            Left            =   3960
            TabIndex        =   183
            Top             =   0
            Width           =   255
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            MousePointer    =   5  'Size
            Top             =   0
            Width           =   4545
         End
         Begin VB.Label lblPoints 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   300
            Left            =   1680
            TabIndex        =   181
            Top             =   4750
            Width           =   435
         End
         Begin VB.Label lblLANG 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   180
            ToolTipText     =   "Points permettant d'augmenter vos chances d'esquive"
            Top             =   4485
            Width           =   270
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
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   179
            Top             =   3480
            Width           =   165
         End
         Begin VB.Label AddLang 
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
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   178
            Top             =   4455
            Width           =   165
         End
         Begin VB.Label AddSci 
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
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   177
            Top             =   4215
            Width           =   165
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
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   2640
            TabIndex        =   176
            Top             =   3720
            Width           =   165
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   600
            TabIndex        =   175
            Top             =   360
            Width           =   240
         End
         Begin VB.Label lblSTR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   174
            ToolTipText     =   "Points permettant d'augmenter vos dégâts et vos chances de coup critique"
            Top             =   3480
            Width           =   270
         End
         Begin VB.Label lblDEF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   1560
            TabIndex        =   173
            ToolTipText     =   "Points permettant d'augmenter votre résistance et vos chances de bloquer"
            Top             =   3720
            Width           =   270
         End
         Begin VB.Label lblSCI 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   172
            ToolTipText     =   "Points permettant d'augmenter vos sorts disponibles "
            Top             =   4215
            Width           =   270
         End
      End
      Begin VB.PictureBox picPlayerSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   3240
         Picture         =   "frmMirage.frx":4B833
         ScaleHeight     =   199
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   4920
         Visible         =   0   'False
         Width           =   2595
         Begin VB.VScrollBar VScroll2 
            Height          =   1935
            LargeChange     =   50
            Left            =   2280
            Max             =   120
            SmallChange     =   30
            TabIndex        =   192
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   120
            ScaleHeight     =   1935
            ScaleWidth      =   2175
            TabIndex        =   189
            Top             =   480
            Width           =   2175
            Begin VB.PictureBox Picture11 
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
               Height          =   7935
               Left            =   120
               ScaleHeight     =   529
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   126
               TabIndex        =   190
               Top             =   0
               Width           =   1890
               Begin VB.PictureBox picspell 
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
                  TabIndex        =   191
                  Top             =   120
                  Width           =   480
               End
               Begin VB.Shape SDAD 
                  BorderColor     =   &H00008000&
                  BorderWidth     =   3
                  Height          =   510
                  Left            =   105
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   510
               End
            End
         End
         Begin VB.Label lbl_reduire_PicPlayerSpells 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   195
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbl_fermer_PicPlayerSpells 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   194
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lbl_deplacer_PicPlayerSpells 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   193
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox picCraft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   1080
         ScaleHeight     =   3585
         ScaleWidth      =   7065
         TabIndex        =   210
         Top             =   4200
         Visible         =   0   'False
         Width           =   7095
         Begin MSComctlLib.ProgressBar ProgressBarCraft 
            Height          =   255
            Left            =   600
            TabIndex        =   243
            Top             =   3000
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Timer TimerCraft 
            Left            =   120
            Top             =   3120
         End
         Begin VB.Frame Frame1 
            Caption         =   "Matériaux :"
            Height          =   1575
            Index           =   0
            Left            =   120
            TabIndex        =   212
            Top             =   1080
            Width           =   3135
            Begin VB.PictureBox picMaterial 
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
               TabIndex        =   214
               Top             =   240
               Width           =   480
            End
            Begin VB.Shape MaterialState 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   525
               Index           =   0
               Left            =   105
               Top             =   220
               Width           =   525
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Produits :"
            Height          =   1575
            Index           =   1
            Left            =   3840
            TabIndex        =   213
            Top             =   1080
            Width           =   3135
            Begin VB.PictureBox picProduct 
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
               TabIndex        =   215
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Label lblLaunchCraft 
            Caption         =   "lblLaunchCraft"
            Height          =   375
            Left            =   4680
            TabIndex        =   244
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lbl_fermer_craft 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Left            =   6840
            TabIndex        =   242
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblCraftName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "lblCraftName"
            Height          =   255
            Left            =   2040
            TabIndex        =   241
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.PictureBox picCrafts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   720
         ScaleHeight     =   3375
         ScaleWidth      =   4215
         TabIndex        =   216
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdSelectCraft 
            Caption         =   "Sélectionner"
            Height          =   375
            Left            =   1560
            TabIndex        =   238
            Top             =   2760
            Width           =   1095
         End
         Begin VB.ListBox lstCraft 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            Height          =   1785
            Left            =   240
            TabIndex        =   217
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label lbl_fermer_crafts 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   255
            Left            =   3960
            TabIndex        =   240
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Liste des patrons :"
            Height          =   255
            Left            =   240
            TabIndex        =   239
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame picInvitation 
         Height          =   1455
         Index           =   0
         Left            =   9840
         TabIndex        =   128
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton cmdNon 
            Caption         =   "Non"
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   131
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdOui 
            Caption         =   "Oui"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   130
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox messageInvitation 
            Appearance      =   0  'Flat
            Height          =   855
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   129
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.PictureBox picRightClick 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   240
         Picture         =   "frmMirage.frx":6D263
         ScaleHeight     =   263.107
         ScaleMode       =   0  'User
         ScaleWidth      =   2565
         TabIndex        =   132
         Top             =   5160
         Visible         =   0   'False
         Width           =   2595
         Begin VB.Label lbl_adopt_rightclick 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Adopter"
            Height          =   375
            Left            =   360
            TabIndex        =   264
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lbl_trade_rightclick 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Echanger"
            Height          =   375
            Left            =   360
            TabIndex        =   205
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label nom 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   360
            TabIndex        =   138
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lbl_groupe_rightclick 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inviter dans le groupe"
            Height          =   375
            Left            =   360
            TabIndex        =   137
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lbl_fermer_rightclick 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2280
            TabIndex        =   136
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbl_reduire_rightclick 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1920
            TabIndex        =   135
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbl_deplacer_rightclick 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   0
            TabIndex        =   134
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lbl_discussion_rightclick 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Demander une discussion privée"
            Height          =   495
            Left            =   360
            TabIndex        =   133
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label lbl_abandon_rightclick 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abandonner"
            Height          =   375
            Left            =   360
            TabIndex        =   265
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.PictureBox picWhosOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
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
         Height          =   2985
         Left            =   4440
         Picture         =   "frmMirage.frx":8EC93
         ScaleHeight     =   2985
         ScaleWidth      =   2595
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   5280
         Visible         =   0   'False
         Width           =   2595
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
            ItemData        =   "frmMirage.frx":B06C3
            Left            =   200
            List            =   "frmMirage.frx":B06C5
            TabIndex        =   197
            Top             =   480
            Width           =   2220
         End
         Begin VB.Label lbl_fermer_picWhosOnline 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   200
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbl_reduire_picWhosOnline 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   199
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbl_deplacer_picWhosOnline 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   198
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox tmpsquete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   185
         Top             =   5400
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
            TabIndex        =   187
            ToolTipText     =   "Minutes restante avant la fin de la quête en cour"
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
            TabIndex        =   186
            ToolTipText     =   "Secondes restante avant la fin de la quête en cour"
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.PictureBox xp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   45
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
            TabIndex        =   46
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
      Begin VB.PictureBox mana 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   43
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
            TabIndex        =   44
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
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   41
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
            TabIndex        =   42
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
      Begin VB.PictureBox ObjNm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   38
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Label OName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Timer Timer1 
         Left            =   7320
         Top             =   0
      End
      Begin VB.Timer tmrSnowDrop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8760
         Top             =   0
      End
      Begin VB.Timer tmrRainDrop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8280
         Top             =   0
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
         Left            =   9240
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame fra_fenetre 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   2985
         Left            =   8640
         TabIndex        =   1
         Top             =   5640
         Width           =   2595
         Begin VB.Frame fraCarte 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
            Begin VB.Image imgCarte 
               Height          =   2295
               Left            =   420
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.PictureBox picEquip 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
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
            Left            =   120
            ScaleHeight     =   167
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   159
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
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
               TabIndex        =   14
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
                  TabIndex        =   15
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
               TabIndex        =   12
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
                  TabIndex        =   13
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
               TabIndex        =   10
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
                  TabIndex        =   11
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
               TabIndex        =   8
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
                  TabIndex        =   9
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
               TabIndex        =   6
               Top             =   1920
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox BootsImage 
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
                  TabIndex        =   7
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
               TabIndex        =   4
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
                  TabIndex        =   5
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox picItems 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
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
               Height          =   2.25000e5
               Left            =   2400
               Picture         =   "frmMirage.frx":B06C7
               ScaleHeight     =   2.23636e5
               ScaleMode       =   0  'User
               ScaleWidth      =   477.091
               TabIndex        =   3
               Top             =   2760
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin VB.PictureBox picGuildAdmin 
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
            Height          =   2505
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   2385
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
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
               TabIndex        =   28
               Top             =   585
               Width           =   1575
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
               TabIndex        =   27
               Top             =   345
               Width           =   1575
            End
            Begin VB.CommandButton cmdTrainee 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Recruter"
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
               TabIndex        =   26
               Top             =   975
               Width           =   1815
            End
            Begin VB.CommandButton cmdMember 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Recruter (comme recruteur)"
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
               TabIndex        =   25
               Top             =   1305
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
               TabIndex        =   24
               Top             =   1650
               Width           =   1815
            End
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
               TabIndex        =   23
               Top             =   1980
               Width           =   1815
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
               TabIndex        =   30
               Top             =   615
               Width           =   465
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
               TabIndex        =   29
               Top             =   360
               Width           =   345
            End
         End
         Begin VB.PictureBox picGuild 
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
            Height          =   2505
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   2385
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
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
               TabIndex        =   20
               Top             =   975
               Width           =   1080
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
               TabIndex        =   19
               Top             =   660
               Width           =   1065
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
               TabIndex        =   18
               Top             =   960
               Width           =   825
            End
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
               TabIndex        =   17
               Top             =   645
               Width           =   1050
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
               TabIndex        =   21
               Top             =   2280
               Width           =   1110
            End
         End
         Begin VB.Label lblmaskinvferm 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2280
            TabIndex        =   50
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblmaskinvmin 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2040
            TabIndex        =   49
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblmaskinv 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   0
            MousePointer    =   5  'Size
            TabIndex        =   36
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label ffermer 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2280
            TabIndex        =   48
            Top             =   0
            Width           =   375
         End
         Begin VB.Label freduire 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   1920
            TabIndex        =   47
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image3 
            Height          =   2985
            Left            =   0
            Picture         =   "frmMirage.frx":210009
            Top             =   0
            Width           =   2595
         End
      End
      Begin VB.PictureBox txtQ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         Picture         =   "frmMirage.frx":229483
         ScaleHeight     =   1545
         ScaleWidth      =   9510
         TabIndex        =   33
         Top             =   7200
         Visible         =   0   'False
         Width           =   9540
         Begin VB.TextBox TxtQ2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1065
            Left            =   158
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Text            =   "frmMirage.frx":2592D5
            Top             =   180
            Width           =   9160
         End
         Begin VB.Label OK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9120
            TabIndex        =   35
            Top             =   1360
            Width           =   495
         End
      End
      Begin VB.PictureBox picquete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4290
         Left            =   480
         Picture         =   "frmMirage.frx":2592DD
         ScaleHeight     =   286
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   3345
         Begin VB.TextBox quetetxt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   3015
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label qf 
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   3045
            TabIndex        =   56
            Top             =   0
            Width           =   285
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
            Left            =   1440
            TabIndex        =   55
            Top             =   2040
            Width           =   45
         End
         Begin VB.Label qt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   54
            Top             =   3600
            Width           =   1020
         End
         Begin VB.Label artquete 
            BackStyle       =   0  'Transparent
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
            Height          =   405
            Left            =   1440
            TabIndex        =   53
            Top             =   3840
            Width           =   1845
         End
      End
   End
   Begin MSWinsockLib.Winsock SocketTCP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SocketUDPSend 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label menu_craft 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   10200
      TabIndex        =   211
      Top             =   9360
      Width           =   375
   End
   Begin VB.Label menu_quete 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9240
      TabIndex        =   93
      ToolTipText     =   "Quetes"
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label menu_quit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   77
      ToolTipText     =   "Quitter"
      Top             =   9360
      Width           =   345
   End
   Begin VB.Label menu_equ 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8760
      TabIndex        =   76
      ToolTipText     =   "Equipements"
      Top             =   9240
      Width           =   345
   End
   Begin VB.Label menu_guild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   75
      ToolTipText     =   "Guilde"
      Top             =   9360
      Width           =   405
   End
   Begin VB.Label menu_opt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11160
      TabIndex        =   74
      ToolTipText     =   "Options"
      Top             =   9360
      Width           =   420
   End
   Begin VB.Label menu_who 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10560
      TabIndex        =   73
      ToolTipText     =   "Qui est en ligne ?"
      Top             =   9360
      Width           =   540
   End
   Begin VB.Label menu_sort 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8280
      TabIndex        =   72
      ToolTipText     =   "Sorts"
      Top             =   9240
      Width           =   465
   End
   Begin VB.Label menu_inv 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7920
      TabIndex        =   71
      ToolTipText     =   "Inventaire"
      Top             =   9240
      Width           =   315
   End
   Begin VB.Image Interface 
      Height          =   900
      Left            =   0
      Picture         =   "frmMirage.frx":2881DF
      Top             =   9120
      Width           =   12000
   End
   Begin WMPLibCtl.WindowsMediaPlayer Mediaplayer 
      Height          =   720
      Left            =   12360
      TabIndex        =   37
      Top             =   4560
      Width           =   480
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
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
      _cx             =   847
      _cy             =   1270
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

'1024X768

'Mouse Wheel Begin
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const WM_DESTROY As Long = &H2

Private hWnd_APIwindow As Long

Public WithEvents cSubclasserHooker As clsMouseWheel
Attribute cSubclasserHooker.VB_VarHelpID = -1
'Mouse Wheel End

' Cursor
Public Ctrl2Handle As Collection
Public SysCursHandle As Collection
Public Curs2Handle As Long
' Cursor End



' Right click on listbox
Public m_sngLBRowHeight  As Single

Private SpellMemorized As Long
Public DragImg As Long
Public DragX As Long
Public DragY As Long
Private OldPCX As Long
Private OldPCY As Long
Private twippx As Long
Private twippy As Long
Private Lon As Long
Private Hau As Long
Private hauteurEcran As Long
Private largeurEcran As Long
Private valeur As Long
'Variable du menu clique droit
'Private playerIndex As Long



Private Sub AddDef_Click()
    If Player(MyIndex).FreePoints > 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CAddDef
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub AddDex_Click()
    If Player(MyIndex).FreePoints > 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CAddDex
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub AddLang_Click()
    If Player(MyIndex).FreePoints > 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CAddLang
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub AddSci_Click()
    If Player(MyIndex).FreePoints > 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CAddSci
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub AddStr_Click()
    If Player(MyIndex).FreePoints > 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CAddStr
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub ArmorImage_DblClick()
    Call TakeOutArmor
End Sub

Public Sub TakeOutArmor()
    If Player(MyIndex).ArmorSlot.num >= 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CTakeOutArmor
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub ArmorImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Player(MyIndex).ArmorSlot.num >= 0 Then
        Call DisplayDescription(Player(MyIndex).ArmorSlot)
    Else
        frmMirage.itmDesc.Visible = False
    End If
End Sub

Private Sub artquete_Click()
    Player(MyIndex).QueteEnCour = 0
'    Accepter = False
    Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & SEP_CHAR & END_CHAR)
    frmMirage.picquete.Visible = False
    If quetetimersec.Enabled Then
        quetetimersec.Enabled = False
        tmpsquete.Visible = False
    End If
End Sub

Private Sub cbtr1_Change()

End Sub

Private Sub chkLowEffect_Click()
    WriteINI "CONFIG", "LowEffect", chkLowEffect.Value, ClientConfigurationFile
End Sub

Private Sub chknobj_Click()
    WriteINI "CONFIG", "NomObjet", chknobj.Value, ClientConfigurationFile
End Sub

Private Sub chksound_Click()
    WriteINI "CONFIG", "Sound", chksound.Value, ClientConfigurationFile
End Sub

Private Sub chkbubblebar_Click()
    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.Value, ClientConfigurationFile
End Sub

Private Sub chknpcbar_Click()
    WriteINI "CONFIG", "NpcBar", chknpcbar.Value, ClientConfigurationFile
End Sub

Private Sub chknpcdamage_Click()
    WriteINI "CONFIG", "NPCDamage", chknpcdamage.Value, ClientConfigurationFile
End Sub

Private Sub chknpcname_Click()
    WriteINI "CONFIG", "NPCName", chknpcname.Value, ClientConfigurationFile
End Sub

Private Sub chkplayerbar_Click()
    WriteINI "CONFIG", "PlayerBar", chkplayerbar.Value, ClientConfigurationFile
End Sub

Private Sub chkplayerdamage_Click()
    WriteINI "CONFIG", "PlayerDamage", chkplayerdamage.Value, ClientConfigurationFile
End Sub

Private Sub chkAutoScroll_Click()
    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.Value, ClientConfigurationFile
End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.Value, ClientConfigurationFile
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.Value, ClientConfigurationFile
    If MyIndex <= 0 Then Exit Sub
    'Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
    Call PlayMidi(App.Path & "\" & Trim$(Map.Music))
End Sub

Private Sub cmdRegisterCommands_Click()
    Dim FileName As String
    Dim i As Integer

    UserCommand("haut") = CInt(lblCommandHaut.DataField)
    UserCommand("bas") = CInt(lblCommandBas.DataField)
    UserCommand("gauche") = CInt(lblCommandGauche.DataField)
    UserCommand("droite") = CInt(lblCommandDroite.DataField)
    UserCommand("attaque") = CInt(lblCommandAttaque.DataField)
    UserCommand("courir") = CInt(lblCommandCourir.DataField)
    UserCommand("ramasser") = CInt(lblCommandRamasser.DataField)
    UserCommand("action") = CInt(lblCommandAction.DataField)
    
    For i = 0 To 13
        UserCommand("rac" & (i + 1)) = CInt(lblCommandRac(i).DataField)
    Next i
    
    'FileName = App.Path & "\Config\Option.ini"

    Call WriteINI("COMMAND", "haut", UserCommand("haut"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "bas", UserCommand("bas"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "gauche", UserCommand("gauche"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "droite", UserCommand("droite"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "attaque", UserCommand("attaque"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "courir", UserCommand("courir"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "ramasser", UserCommand("ramasser"), OptionConfigurationFile)
    Call WriteINI("COMMAND", "action", UserCommand("action"), OptionConfigurationFile)
    For i = 0 To 13
        Call WriteINI("COMMAND", "rac" & (i + 1), UserCommand("rac" & (i + 1)), OptionConfigurationFile)
    Next i
    
    pictTouche.Visible = False
End Sub

Private Sub HelmetImage_DblClick()
    Call TakeOutHelmet
End Sub

Sub TakeOutHelmet()
    If Player(MyIndex).HelmetSlot.num >= 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CTakeOutHelmet
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub HelmetImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Player(MyIndex).HelmetSlot.num >= 0 Then
        Call DisplayDescription(Player(MyIndex).HelmetSlot)
    Else
        frmMirage.itmDesc.Visible = False
    End If
End Sub

Private Sub itmDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub lbl_abandon_rightclick_Click()
    Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CAbandon
    
    SendData Packet.ToArray()
    Set Packet = Nothing
    
    picRightClick.Visible = False
End Sub

Private Sub lbl_adopt_rightclick_Click()
    Dim Packet As clsBuffer

    Set Packet = New clsBuffer
    
    Packet.WriteLong CAdopt
    Packet.WriteInteger frmMirage.nom.DataField
    
    SendData Packet.ToArray()
    Set Packet = Nothing
    
    picRightClick.Visible = False
End Sub

Private Sub lblCommandAction_DblClick()
    Call prepareCommandChange(lblCommandAction)
End Sub

Private Sub lblCommandAttaque_DblClick()
    Call prepareCommandChange(lblCommandAttaque)
End Sub

Private Sub lblCommandBas_DblClick()
    Call prepareCommandChange(lblCommandBas)
End Sub

Private Sub lblCommandCourir_DblClick()
    Call prepareCommandChange(lblCommandCourir)
End Sub

Private Sub lblCommandDroite_DblClick()
    Call prepareCommandChange(lblCommandDroite)
End Sub

Private Sub lblCommandGauche_DblClick()
    Call prepareCommandChange(lblCommandGauche)
End Sub

Private Sub lblCommandHaut_DblClick()
    Call prepareCommandChange(lblCommandHaut)
End Sub

Private Function resetCommandChange(ByRef lbl As Variant) As Boolean
    resetCommandChange = False

    If lbl.ForeColor = vbRed Then
        ConOff = ConOff - 1
        lbl.ForeColor = vbBlack
        resetCommandChange = True
    End If
End Function

Private Sub prepareCommandChange(ByRef lbl As Label)
    Dim i As Byte
    Dim currentLabel As Variant

    For Each currentLabel In UserCommandLabel
        Call resetCommandChange(currentLabel)
    Next

    ConOff = ConOff + 1
    lbl.ForeColor = vbRed
End Sub

Private Sub setCommandChange(ByVal keyCode As Integer)
    Dim currentLabel As Label

    ' Test if key now already taken
    For Each currentLabel In UserCommandLabel
        If currentLabel.DataField = keyCode Then
            If resetCommandChange(currentLabel) Then
                currentLabel.Caption = optTouche(keyCode)
                currentLabel.DataField = keyCode
            Else
                MsgBox "Touche déjà utilisé. Veuillez en choisir une autre."
            End If
            Exit Sub
        End If
    Next

    For Each currentLabel In UserCommandLabel
        If resetCommandChange(currentLabel) Then
            currentLabel.Caption = optTouche(keyCode)
            currentLabel.DataField = keyCode
            Exit For
        End If
    Next
End Sub

Private Sub lblCommandRac_DblClick(Index As Integer)
    Call prepareCommandChange(lblCommandRac(Index))
End Sub

Private Sub lblCommandRamasser_DblClick()
    Call prepareCommandChange(lblCommandRamasser)
End Sub

Private Sub lblLaunchCraft_Click()
    If CanBuild Then
        ' Progress bar
        ProgressBarCraft.Min = 0
        ProgressBarCraft.Max = 100
        ProgressBarCraft.Value = 0
        TimerCraft.Interval = 20
        TimerCraft.Enabled = True
    End If
End Sub

Private Sub cmdLeave_Click()
Dim Packet As clsBuffer
'    Set Packet = New clsBuffer
'
'    Packet.WriteLong CLeaveGuild
'
'    SendData Packet.ToArray()
'    Set Packet = Nothing

    lblGuild.Caption = vbNullString
    lblRank.Caption = 0
End Sub

Private Sub cmdMember_Click()
    ' TODO
End Sub

Private Sub cmdNon_Click(Index As Integer)
    picInvitation(Index).Visible = False
End Sub

Private Sub CmdoptTouche_Click()
Dim i As Byte
    lblCommandHaut.Caption = optTouche(UserCommand("haut"))
    lblCommandHaut.DataField = UserCommand("haut")
    lblCommandBas.Caption = optTouche(UserCommand("bas"))
    lblCommandBas.DataField = UserCommand("bas")
    lblCommandGauche.Caption = optTouche(UserCommand("gauche"))
    lblCommandGauche.DataField = UserCommand("gauche")
    lblCommandDroite.Caption = optTouche(UserCommand("droite"))
    lblCommandDroite.DataField = UserCommand("droite")
    lblCommandAttaque.Caption = optTouche(UserCommand("attaque"))
    lblCommandAttaque.DataField = UserCommand("attaque")
    lblCommandCourir.Caption = optTouche(UserCommand("courir"))
    lblCommandCourir.DataField = UserCommand("courir")
    lblCommandRamasser.Caption = optTouche(UserCommand("ramasser"))
    lblCommandRamasser.DataField = UserCommand("ramasser")
    lblCommandAction.Caption = optTouche(UserCommand("action"))
    lblCommandAction.DataField = UserCommand("action")

    For i = 0 To 13
        lblCommandRac(i).Caption = optTouche(UserCommand("rac" & (i + 1)))
        lblCommandRac(i).DataField = UserCommand("rac" & (i + 1))
    Next i
    
pictTouche.Visible = True
pictTouche.SetFocus
End Sub

Private Sub cmdOTA_Click()
    Dim currentLabel As Variant
    For Each currentLabel In UserCommandLabel
        Call resetCommandChange(currentLabel)
    Next
    
    pictTouche.Visible = False
End Sub

Private Sub cmdOui_Click(Index As Integer)
    Dim Packet As clsBuffer
    
    Select Case Index
        Case PARTY_MESSAGE
            Call SendJoinParty
        Case CHAT_MESSAGE
'            Set Packet = New clsBuffer
'
'            Packet.WriteLong CPlayerChatAccept
'
'            SendData Packet.ToArray()
'            Set Packet = Nothing
        Case TRADE_MESSAGE
'            Set Packet = New clsBuffer
'
'            Packet.WriteLong CTradeAccept
'
'            SendData Packet.ToArray()
'            Set Packet = Nothing
        Case SLEEP_MESSAGE
'            Set Packet = New clsBuffer
'
'            Packet.WriteLong CSleepAccept
'
'            SendData Packet.ToArray()
'            Set Packet = Nothing
    End Select
    picInvitation(Index).Visible = False
End Sub

Private Sub cmdSelectCraft_Click()
    Dim i, Position, X As Integer
    Dim craftNum As Integer

    If lstCraft.ListIndex >= 0 Then
        CraftLoad lstCraft.ItemData(lstCraft.ListIndex)
    End If
End Sub

Private Sub Command1_Click()
picOptions.Visible = False
Call InitAccountOpt
End Sub

Private Sub Command2_Click()
Call Form_Load
End Sub

Private Sub fermer_Click()
    picInfo.Visible = False
End Sub

Private Sub ffermer_Click()
fra_fenetre.Visible = False
End Sub

Private Sub Form_GotFocus()
Picsprt.Height = 48
Picsprts.Height = 48
On Error Resume Next
txtMyTextBox.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Long, X As Integer
Dim Ending As String
Dim Qq As Long
Dim ctl As Control

    Call SetIcon(Me)

    RemoveBorder lstCraft.hwnd
    RemoveBorder lstOnline.hwnd

    Set Ctrl2Handle = New Collection
    Call Ctrl2Handle.Add(picScreen)
    Call Ctrl2Handle.Add(VScroll1)
    Call Ctrl2Handle.Add(lstCraft)
    Call Ctrl2Handle.Add(cmdSelectCraft)
    Call Ctrl2Handle.Add(chkbubblebar)
    Call Ctrl2Handle.Add(Me)
    Call Ctrl2Handle.Add(fra_fenetre)
    Call Ctrl2Handle.Add(scrlBltText)
    

    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
        If FileExist(App.Path & Rep_Theme & "\Jeu\Text" & Ending) Then txtQ.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\text" & Ending)
        If FileExist(App.Path & Rep_Theme & "\info" & Ending) Then frmMirage.Picture = LoadPNG(App.Path & Rep_Theme & "\info" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Jeu\Info" & Ending) Then picInfo.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Info" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Jeu\inventaire" & Ending) Then Image3.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\inventaire" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Jeu\Interface" & Ending) Then Interface.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Interface" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Jeu\Carte" & Ending) Then imgcarte.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Carte" & Ending)
        If FileExist(App.Path & Rep_Theme & "\Jeu\quitter" & Ending) Then PicMenuQuitter.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\quitter" & Ending)
        
    Next i
    
    twippy = Screen.TwipsPerPixelY
    twippx = Screen.TwipsPerPixelX

    lblName.Font = ReadINI("POLICE", "Police", (ColorConfigurationFile))
    lblName.FontSize = ReadINI("POLICE", "PoliceSize", (ColorConfigurationFile))
    txtMyTextBox.Font = ReadINI("POLICE", "PoliceChat", (ColorConfigurationFile))
    
    'Init as in previous modHandleData
    For i = 1 To MAX_INV
        'Debug.Print "inventory : " & I
        If Loading = False Then Load frmMirage.picInv(i)

        X = Int(i / 3)
        frmMirage.picInv(i).Top = 8 + 40 * X
        frmMirage.picInv(i).Left = 8 + (i - X * 3) * 40
        'Debug.Print "Indice : " & I & " Top : " & frmMirage.picInv(I).Top & " Left : " & frmMirage.picInv(I).Left
        frmMirage.picInv(i).Visible = True
    Next
    
    'Init as in previous modHandleData
    For i = 1 To MAX_MATERIALS
        ' Load materials
        If Loading = False Then Load frmMirage.picMaterial(i)

        X = Int(i / 5)
        frmMirage.picMaterial(i).Top = frmMirage.picMaterial(0).Top + 600 * X
        frmMirage.picMaterial(i).Left = frmMirage.picMaterial(0).Left + (i - X * 5) * 600
        frmMirage.picMaterial(i).Visible = True

        ' Load materials states
        If Loading = False Then Load frmMirage.MaterialState(i)

        X = Int(i / 5)
        frmMirage.MaterialState(i).Top = frmMirage.MaterialState(0).Top + 600 * X
        frmMirage.MaterialState(i).Left = frmMirage.MaterialState(0).Left + (i - X * 5) * 600
        frmMirage.MaterialState(i).Visible = True

        ' Load products
        If Loading = False Then Load frmMirage.picProduct(i)

        X = Int(i / 5)
        frmMirage.picProduct(i).Top = frmMirage.picProduct(0).Top + 600 * X
        frmMirage.picProduct(i).Left = frmMirage.picProduct(0).Left + (i - X * 5) * 600
        frmMirage.picProduct(i).Visible = True
    Next

    frmMirage.Picture9.Height = frmMirage.picInv(MAX_INV).Top + 40
    Debug.Print "height : " & frmMirage.Picture9.Height

    For i = 1 To MAX_PLAYER_SKILLS
        If Loading = False Then Load frmMirage.picspell(i)

        X = Int(i / 3)
        frmMirage.picspell(i).Top = 8 + 40 * X
        frmMirage.picspell(i).Left = 8 + (i - X * 3) * 40
        frmMirage.picspell(i).Visible = True
    Next
    
    frmMirage.dragDropPicture.Visible = False
    
    picInfo.Visible = False
    fra_fenetre.Visible = False
    For i = 0 To IMSG_COUNT - 1
        picInvitation(i).Visible = False
    Next
    
    ' Record mousewheel
    Set cSubclasserHooker = New clsMouseWheel
    cSubclasserHooker.Attach_Subclasser Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GettingMap Then Cancel = True Else Call GameDestroy
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub freduire_Click()
    If fra_fenetre.Height >= 2985 / 15 Then fra_fenetre.Height = 315 / 15 Else fra_fenetre.Height = 2985 / 15
End Sub

Private Sub ifermer_Click()
    picInfo.Visible = False
End Sub

Private Sub Image3_Click()
    fra_fenetre.Visible = False
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 3
DragX = X
DragY = Y
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picInfo, X, Y)
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub ireduire_Click()
    Call reduire(picInfo, 350)
End Sub

Private Sub Label19_Click()
    If picParty.Height >= 2985 / twippy Then picParty.Height = 315 / twippy Else picParty.Height = 2985 / twippy
End Sub

Private Sub Label27_Click()
    PicMenuQuitter.Visible = False
End Sub

Private Sub Label3_Click()
    If Player(MyIndex).partyIndex > -1 Then SendLeaveParty: picParty.Visible = False Else picParty.Visible = False
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 1
DragX = X
DragY = Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picParty, X, Y)
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub Label6_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label8_Click()
    Dim i As Integer
    picParty.Visible = False

    For i = 1 To MAX_PLAYERS
        Player(i).partyIndex = -1
    Next i

    Call SendLeaveParty
End Sub

Private Sub lbl_deplacer_inv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 1
    DragX = X
    DragY = Y
End Sub

Private Sub lbl_deplacer_inv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picInv3, X, Y)
End Sub

Private Sub lbl_deplacer_inv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub lbl_deplacer_PicPlayerSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 1
    DragX = X
    DragY = Y
End Sub

Private Sub lbl_deplacer_PicPlayerSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picPlayerSpells, X, Y)
End Sub

Private Sub lbl_deplacer_PicPlayerSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub lbl_deplacer_picWhosOnline_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 1
    DragX = X
    DragY = Y
End Sub

Private Sub lbl_deplacer_picWhosOnline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picWhosOnline, X, Y)
End Sub

Private Sub lbl_deplacer_picWhosOnline_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub lbl_deplacer_rightclick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 1
    DragX = X
    DragY = Y
End Sub

Private Sub lbl_deplacer_rightclick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picRightClick, X, Y)
End Sub

Private Sub lbl_deplacer_rightclick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub lbl_discussion_rightclick_Click()
'    Dim Packet As clsBuffer
'
'    Set Packet = New clsBuffer
'
'    Packet.WriteLong CPlayerChat
'    Packet.WriteInteger frmMirage.nom.DataField
'
'    SendData Packet.ToArray()
'    Set Packet = Nothing

    picRightClick.Visible = False
End Sub

Private Sub lbl_fermer_crafts_Click()
    picCrafts.Visible = False
End Sub

Private Sub lbl_fermer_inv_Click()
    picInv3.Visible = False
End Sub

Private Sub lbl_fermer_craft_Click()
    picCraft.Visible = False
End Sub

Private Sub lbl_fermer_PicPlayerSpells_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub lbl_fermer_picWhosOnline_Click()
    picWhosOnline.Visible = False
End Sub

Private Sub lbl_fermer_rightclick_Click()
    picRightClick.Visible = False
End Sub

Private Sub lbl_groupe_rightclick_Click()
    If Player(MyIndex).partyIndex = -1 Then
        Dim partyName As String
        partyName = InputBox("Nom du groupe :", "Nom du groupe")
        If partyName <> vbNullString Then
            Call SendRequestParty(frmMirage.nom.Caption, partyName)
        End If
    Else
        Call SendRequestParty(frmMirage.nom.Caption)
    End If
    picRightClick.Visible = False
End Sub

Private Sub lbl_reduire_inv_Click()
    Call reduire(picInv3, 199)
End Sub

Private Sub lbl_reduire_PicPlayerSpells_Click()
    Call reduire(picPlayerSpells, 199)
End Sub

Private Sub lbl_reduire_picWhosOnline_Click()
    Call reduire(picWhosOnline, 199)
End Sub

Private Sub lbl_reduire_rightclick_Click()
    Call reduire(picRightClick, 199)
End Sub

Private Sub lbl_trade_rightclick_Click()
    Call SendTradeRequest(frmMirage.nom.DataField)
    
    picRightClick.Visible = False
End Sub

Private Sub lblCdP_Click()
    Call SendData("CHANGECHAR" & SEP_CHAR & END_CHAR)
    frmMirage.Visible = False
    frmMainMenu.Visible = True
    frmMainMenu.fraPers.Visible = True
    frmsplash.Visible = False
    PicMenuQuitter.Visible = False
End Sub

Private Sub lblDeco_Click()
Dim i As Integer
    Call RestoreCursor
    InGame = False
    deco = True
    PicMenuQuitter.Visible = False
    frmMainMenu.Visible = True
    If frmMainMenu.Check1.Value = 1 Then If FileExist(App.Path & "\Music\mainmenu.mid") Then Call PlayMidi(App.Path & "\Music\mainmenu.mid") Else Call PlayMidi(App.Path & "\Music\mainmenu.mp3")
    frmMainMenu.fraLogin.Visible = True
    frmMainMenu.fraPers.Visible = False
    frmMirage.tmpsquete.Visible = False
    frmMirage.quetetimersec.Enabled = False
    frmMirage.Visible = False
    frmMirage.SocketTCP.Close
End Sub

Private Sub lblmaskinvferm_Click()
fra_fenetre.Visible = False
End Sub

Private Sub lblmaskinvmin_Click()
    If fra_fenetre.Height >= 2985 / twippy Then fra_fenetre.Height = 315 / twippy Else fra_fenetre.Height = 2985 / twippy
End Sub

Private Sub lblmaskinv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 2
DragX = X
DragY = Y
End Sub

Private Sub lblmaskinv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragImg = 2 Then fra_fenetre.Top = fra_fenetre.Top + ((Y / twippy) - (DragY / twippy)): fra_fenetre.Left = fra_fenetre.Left + ((X / twippx) - (DragX / twippx))
End Sub

Private Sub lblmaskinv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub



Private Sub lblmaskmenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 1
DragX = X
DragY = Y
End Sub



Private Sub lblmaskmenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub lblQuitter_Click()
    Call GameDestroy
End Sub

Private Sub lstOnline_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngListRow As Long
     
    If Button = 2 Then
        lngListRow = (Y / m_sngLBRowHeight) + frmMirage.lstOnline.TopIndex - 1
        If lngListRow > (frmMirage.lstOnline.ListCount - 1) Then lngListRow = frmMirage.lstOnline.ListCount - 1
        'If lngListRow < 0 Then lngListRow = 0
        If lngListRow >= 0 Then
            frmMirage.lstOnline.Selected(lngListRow) = True
            
            Dim pos As POINTAPI

            pos = GetMousePosition
            
            hauteurEcran = picScreen.Height - picWhosOnline.Height
            largeurEcran = picScreen.Width - picWhosOnline.Width
        
            frmMirage.picRightClick.Visible = False
        
            frmMirage.nom.Caption = Trim$(lstOnline.Text)
            frmMirage.nom.DataField = frmMirage.lstOnline.ItemData(frmMirage.lstOnline.ListIndex)
            frmMirage.picRightClick.Top = pos.Y
            If frmMirage.picRightClick.Top > hauteurEcran Then
                frmMirage.picRightClick.Top = hauteurEcran
            End If
            frmMirage.picRightClick.Left = pos.X
            If frmMirage.picRightClick.Left > largeurEcran Then
                frmMirage.picRightClick.Left = largeurEcran
            End If
            frmMirage.picRightClick.Visible = True
        End If
    End If
End Sub

Private Sub menu_craft_Click()
    If picCrafts.Visible Then
        picCrafts.Visible = False
    Else
        Call CraftsLoad
    
        picCrafts.Visible = True
    End If
End Sub

Private Sub menu_equ_Click()
    If picInfo.Visible Then
        picInfo.Visible = False
    Else
        Call UpdateVisInv
        picInfo.Visible = True
        Picsprt.Height = (48 + 4) * twippy
        Picsprts.Height = (48)
        If Picsprts.Height <= 32 Then Picture5.Top = 2160 Else Picture5.Top = 2640
        Call AffSurfPic(DD_SpriteSurf(Player(MyIndex).sprite), Picsprts, 0, 0)
        'Call BitBlt(Picsprts.hDC, 0, 0, PIC_X, PIC_Y * PIC_NPC1, Picturesprite.hDC, 3 * PIC_X, Val(Player(MyIndex).Sprite) * (PIC_Y * PIC_NPC1), SRCCOPY)
    End If
End Sub

Private Sub menu_guild_Click()
' Set Their Guild Name and Their Rank
If fra_fenetre.Visible = True And picGuild.Visible = True Then fra_fenetre.Visible = False Else fra_fenetre.Visible = True
picParty.Visible = (fra_fenetre.Visible And (Player(MyIndex).partyIndex > -1))
Label3.Visible = picParty.Visible
If picParty.Visible Then
    Dim i As Integer, C As Byte
    If lblPPName(0).Tag <= lblPPName(2).Tag Or lblPPName(2).Caption <> vbNullString Then
        For i = (Val(lblPPName(2).Tag) + 1) To MAX_PLAYERS
            If IsPlaying(i) And Player(i).partyIndex = Player(MyIndex).partyIndex And C < 3 And i <> MyIndex Then
                C = C + 1
                lblPPName(C - 1).Tag = i
            End If
        Next
        For i = 0 To 2
            lblPPName(i).Visible = (i < C)
            backPPLife(i).Visible = lblPPName(i).Visible
            backPPMana(i).Visible = lblPPName(i).Visible
            If lblPPName(i).Visible Then
                lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).STP / Player(Val(lblPPName(i).Tag)).MaxSTP * backPPMana(i).Width
            End If
        Next
    End If
End If
frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
picGuild.Visible = True
End Sub

Private Sub menu_inv_Click()
    If picInv3.Visible = True Then
        picInv3.Visible = False
    Else
        Call UpdateVisInv
        picInv3.Visible = True
    End If
End Sub

Private Sub menu_opt_Click()
    If picOptions.Visible Then
        picOptions.Visible = False
    Else
        picOptions.Visible = True
    End If
End Sub

Private Sub menu_quete_Click()
    If Player(MyIndex).QueteEnCour > 0 Then frmMirage.picquete.Visible = True: frmMirage.quetetxt.Text = quete(Player(MyIndex).QueteEnCour).description
End Sub

Private Sub menu_quit_Click()
If PicMenuQuitter.Visible Then PicMenuQuitter.Visible = False Else PicMenuQuitter.Visible = True
End Sub

Private Sub menu_sort_Click()
If picPlayerSpells.Visible = True Then
    picPlayerSpells.Visible = False
Else
    Call Affspell
    picPlayerSpells.Visible = True
End If
End Sub

Private Sub menu_who_Click()
    If picWhosOnline.Visible Then
        picWhosOnline.Visible = False
    Else
        Call SendOnlineList
        picWhosOnline.Visible = True
    End If
End Sub

Private Sub OK_Click()
Dim i As Long
Dim msgb As String
Dim Packet As clsBuffer


If Player(MyIndex).QueteEnCour > 0 And frmMirage.OK.Tag = MSG_TYPE_QUEST Then
    msgb = MsgBox("Voulez-vous faire la Quete proposer?", vbYesNo, "Quete")
        If msgb = vbYes Then

            If quete(Player(MyIndex).QueteEnCour).Type = QUETE_TYPE_APORT Then
                i = FindOpenInvSlot(quete(Player(MyIndex).QueteEnCour).Datas(0))
                If i = 0 Then
                    Call AddText("Ton inventaire est plein tu ne peut pas faire cette quête!", red)
                    Player(MyIndex).QueteEnCour = 0
                    Exit Sub
                End If
            End If
            
            If quete(Player(MyIndex).QueteEnCour).Temps > 0 Then Call SetQuestTime(quete(Player(MyIndex).QueteEnCour).Temps)
            
'            Set Packet = New clsBuffer
'
'            Packet.WriteLong CQuestLaunch
'            Packet.WriteInteger Player(MyIndex).QueteEnCour
'
'            SendData Packet.ToArray()
'
'            Set Packet = Nothing
            
        Else
            Player(MyIndex).QueteEnCour = 0

        End If
End If
txtQ.Visible = False
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub picInv_DblClick(Index As Integer)
Dim d As Long

If Player(MyIndex).Inv(Index).num < 0 Or Player(MyIndex).Inv(Index).num > MAX_ITEMS Then Exit Sub

Call SendUseItem(Index)
End Sub

Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMirage.SelectedItem.Top = frmMirage.picInv(Index).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Index).Left - 1
    
    If Button = 1 Then
        If Player(MyIndex).Inv(Index).num < 0 Or Player(MyIndex).Inv(Index).num > MAX_ITEMS Then
            Exit Sub
        End If
        If dragAndDrop.Index = -1 Then
            dragAndDrop.Type = 2
            dragAndDrop.Index = Index
        End If
        
    ElseIf Button = 2 Then

    ElseIf Button = 3 Then

    End If
End Sub

Private Sub ProcessDrag()
    Dim pos As POINTAPI

    pos = GetMousePosition

    frmMirage.dragDropPicture.Top = pos.Y + 15
    frmMirage.dragDropPicture.Left = pos.X + 15

    'Peut etre inutile
    'frmMirage.dragDropPicture.Refresh

    If frmMirage.dragDropPicture.Visible = False Then
        If dragAndDrop.Type = 1 Then
            With Player(MyIndex)
                Call AffSurfPic(DD_ItemSurf, frmMirage.dragDropPicture, (skill(.skill(dragAndDrop.Index)).SkillIco - (skill(.skill(dragAndDrop.Index)).SkillIco \ 6) * 6) * PIC_X, (skill(.skill(dragAndDrop.Index)).SkillIco \ 6) * PIC_Y)
            End With
        ElseIf dragAndDrop.Type = 2 Then
            With Player(MyIndex).Inv(dragAndDrop.Index)
                Call AffSurfPic(DD_ItemSurf, frmMirage.dragDropPicture, (item(.num).Pic - (item(.num).Pic \ 6) * 6) * PIC_X, (item(.num).Pic \ 6) * PIC_Y)
            End With
        End If
        frmMirage.itmDesc.Visible = False
        frmMirage.dragDropPicture.Visible = True
    End If
End Sub

Private Sub picInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragAndDrop.Index > -1 Then
        Call ProcessDrag
    Else
        Dim d As Long
        d = Index
        
        If Player(MyIndex).Inv(d).num >= 0 Then
            Call DisplayDescription(Player(MyIndex).Inv(d))
        Else
            frmMirage.itmDesc.Visible = False
        End If
    End If
End Sub

Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long
Dim PX As Long
Dim PY As Long
Dim Cod As String
Dim TP As Long
Dim ctl As Control
Dim Packet As clsBuffer

    If ConOff > 0 Then
        'Test command setting
        Call setCommandChange(keyCode)
        
        Exit Sub
    End If
    
    Call CheckInput(0, keyCode, Shift)
    
    If (keyCode = UserCommand.item("action")) Then
    
        On Error Resume Next
        PX = 0
        PY = 0
        If Player(MyIndex).Y - 1 > -1 And PX = 0 And PY = 0 Then
            TP = Map.tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type
            If TP = TILE_TYPE_COFFRE Or TP = TILE_TYPE_PORTE_CODE And Player(MyIndex).dir = DIR_UP Then PX = 0: PY = -1
        End If
                
        If Player(MyIndex).Y + 1 < MaxMapY + 1 And PX = 0 And PY = 0 Then
            TP = Map.tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type
            If TP = TILE_TYPE_COFFRE Or TP = TILE_TYPE_PORTE_CODE And Player(MyIndex).dir = DIR_DOWN Then PX = 0: PY = 1
        End If
                
        If Player(MyIndex).X - 1 > -1 And PX = 0 And PY = 0 Then
            TP = Map.tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type
            If TP = TILE_TYPE_COFFRE Or TP = TILE_TYPE_PORTE_CODE And Player(MyIndex).dir = DIR_LEFT Then PX = -1: PY = 0
        End If
        
        If Player(MyIndex).X + 1 < MaxMapX + 1 And PX = 0 And PY = 0 Then
            TP = Map.tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type
            If TP = TILE_TYPE_COFFRE Or TP = TILE_TYPE_PORTE_CODE And Player(MyIndex).dir = DIR_RIGHT Then PX = 1: PY = 0
        End If
        
        If PX <> 0 Or PY <> 0 Then
        With Map.tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY)
            If .Strings(0) > vbNullString And TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                Cod = InputBox("Veuillez entre le mot de passe :", "Code")
                If Cod = .Strings(0) Then
                    TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = YES
                    
'                    Set Packet = New clsBuffer
'                    Packet.WriteLong COuvrir
'                    Packet.WriteInteger (GetPlayerX(MyIndex) + PX)
'                    Packet.WriteInteger (GetPlayerY(MyIndex) + PY)
'                    SendData Packet.ToArray()
'                    Set Packet = Nothing


                    If .Type = TILE_TYPE_COFFRE Then
                        i = FindOpenInvSlot(Val(.Datas(2)))
                        If i > 0 Then
                            Call SetPlayerInvItemNum(MyIndex, i, Val(.Datas(2)))
                            Call SetPlayerInvItemValue(MyIndex, i, GetPlayerInvItemValue(MyIndex, i) + 1)
                            Call SetPlayerInvItemDur(MyIndex, i, item(Val(.Datas(2))).Datas(0))
                            Call UpdateVisInv
'                            Set Packet = New clsBuffer
'                            Packet.WriteLong CCoffre
'                            Packet.WriteInteger i
'                            Packet.WriteInteger Val(.Datas(2))
'                            Packet.WriteInteger 1
'                            Packet.WriteLong item(Val(.Datas(2))).Datas(0)
'                            SendData Packet.ToArray()
'                            Set Packet = Nothing
                        End If
                    End If
                Else
                    Call MsgBox("Mauvais code!", vbCritical)
                End If
            End If
        End With
        End If
        
        If GetPlayerY(MyIndex) - 1 > 0 And GetPlayerY(MyIndex) - 1 < MaxMapY Then
            With Map.tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1)
            If .Type = TILE_TYPE_SIGN And Player(MyIndex).dir = DIR_UP Then
                If Trim$(.Strings(0)) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.Strings(0)))
                If Trim$(.Strings(1)) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.Strings(1)))
                If Trim$(.Strings(2)) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.Strings(2)))
                Exit Sub
            End If
            End With
        End If
    End If
    
    If Not frmMirage.txtMyTextBox.Visible Then
        If keyCode = UserCommand.item("rac1") Then Call useRac(0)
        If keyCode = UserCommand.item("rac2") Then Call useRac(1)
        If keyCode = UserCommand.item("rac3") Then Call useRac(2)
        If keyCode = UserCommand.item("rac4") Then Call useRac(3)
        If keyCode = UserCommand.item("rac5") Then Call useRac(4)
        If keyCode = UserCommand.item("rac6") Then Call useRac(5)
        If keyCode = UserCommand.item("rac7") Then Call useRac(6)
        If keyCode = UserCommand.item("rac8") Then Call useRac(7)
        If keyCode = UserCommand.item("rac9") Then Call useRac(8)
        If keyCode = UserCommand.item("rac10") Then Call useRac(9)
        If keyCode = UserCommand.item("rac11") Then Call useRac(10)
        If keyCode = UserCommand.item("rac12") Then Call useRac(11)
        If keyCode = UserCommand.item("rac13") Then Call useRac(12)
        If keyCode = UserCommand.item("rac14") Then Call useRac(13)
    End If
    
    If keyCode = vbKeyEscape Then
        If PicMenuQuitter.Visible Then PicMenuQuitter.Visible = False Else PicMenuQuitter.Visible = True
    End If

    ' The Guild Maker
    If keyCode = vbKeyF5 Then
        If Player(MyIndex).Guildaccess > 1 Then fra_fenetre.Visible = True: frmMirage.picGuildAdmin.Visible = True
    End If
    
    ' The Guild Creator
    If keyCode = vbKeyF6 Then frmGuild.txtName = GetPlayerName(MyIndex): frmGuild.Show vbModeless, frmMirage
    
    'quete desc
    If keyCode = vbKeyF7 Then
        If Player(MyIndex).QueteEnCour > 0 Then fra_fenetre.Visible = False: frmMirage.picquete.Visible = True: frmMirage.quetetxt.Text = quete(Player(MyIndex).QueteEnCour).description
    End If
    
    If keyCode = vbKeyF8 Then frmPlayerHelp.Show
    
    If keyCode = vbKeyF9 Then If Player(MyIndex).Access > 0 Then frmadmin.Show
    
    If keyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
        Else
            Call AddText("Aucune magie mémoriser.", BrightRed)
        End If
    End If
    
    If keyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)
        i = 0
        ii = 0
        Do
            If FileExist(App.Path & "\Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    ElseIf keyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, 8, 6, 634, 479)
        i = 0
        ii = 0
        Do
            If FileExist(App.Path & "\Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    End If
    
    If keyCode = vbKeyEnd Then
    d = GetPlayerDir(MyIndex)
    End If
End Sub

Private Sub picInv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragAndDrop.Index >= 0 Then
        Dim i As Integer
        Dim ctl As Control
        Dim bHandled As Boolean
        Dim bOver As Boolean

        Dim pos As POINTAPI

        GetCursorPos pos

        If (frmMirage.picScreen.Visible And cSubclasserHooker.IsOver(frmMirage.picScreen.hwnd, pos.X, pos.Y)) Then
            'Peut etre un drop d'objet si la souris n'est sur aucun autre control
            For Each ctl In Me.Controls
                On Error Resume Next
                bOver = (ctl.Visible And cSubclasserHooker.IsOver(ctl.hwnd, pos.X, pos.Y))
                On Error GoTo 0
                
                If bOver Then
                    ' Test if we are out of every control
                    If ctl <> frmMirage.picScreen Then
                        Debug.Print "ctl : " & ctl
                        bHandled = True
                    End If
                    
        
                End If
                If bHandled Then Exit For
            Next ctl
            
            If Not bHandled Then 'Out of every control (excep picScreen)
                Call DropItem(dragAndDrop.Index)
                'dragAndDrop = -1
                Call ResetDragAndDrop
                'dragDropPicture.Visible = False
                Exit Sub
            End If
        End If
    
        For i = 0 To picRac.UBound
            If cSubclasserHooker.IsOver(picRac(i).hwnd, pos.X, pos.Y) Then
                'Debug.Print "ok rac"
                rac(i, 0) = dragAndDrop.Index
                rac(i, 1) = 2
                
                Call AffRac
                
                Call ResetDragAndDrop
                Call SaveRac
                'dragAndDrop = -1
                'dragDropPicture.Visible = False
                Exit Sub
            End If
        Next i
    
        If cSubclasserHooker.IsOver(Picture8.hwnd, pos.X, pos.Y) Then
            For i = 0 To picInv.UBound
                If (picInv(i).Visible And cSubclasserHooker.IsOver(picInv(i).hwnd, pos.X, pos.Y)) Then
                    Dim Packet As clsBuffer

                    Set Packet = New clsBuffer

                    Packet.WriteLong CMoveInventoryItem
                    Packet.WriteByte dragAndDrop.Index
                    Packet.WriteByte i
                    Packet.WriteInteger Player(MyIndex).Inv(dragAndDrop.Index).num
                    Packet.WriteInteger Player(MyIndex).Inv(dragAndDrop.Index).Value

                    SendData Packet.ToArray()

                    Set Packet = Nothing

                    Call ResetDragAndDrop
                    'dragAndDrop = -1
                    'dragDropPicture.Visible = False
                    Exit Sub
                End If
            Next i
        End If
        
        Call ResetDragAndDrop
        'dragAndDrop = -1
        'dragDropPicture.Visible = False
    End If
End Sub

Private Sub picMaterial_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Long
Dim itemNum As Integer
d = Index
    itemNum = picMaterial(d).DataField
    
    If itemNum > 0 Then
        
        If item(itemNum).Type = ITEM_TYPE_CURRENCY And Trim$(item(itemNum).desc) = vbNullString Then
            itmDesc.Height = 17
            itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            itmDesc.Left = fra_fenetre.Left
        ElseIf Trim$(item(itemNum).desc) = vbNullString Then
            itmDesc.Height = 161
            itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            itmDesc.Left = fra_fenetre.Left
        ElseIf Trim$(item(itemNum).desc) > vbNullString Then
            itmDesc.Height = 249
            itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            itmDesc.Left = fra_fenetre.Left
        End If

        descName.Caption = Trim$(item(itemNum).name) & "(" & GetPlayerInvItemTotalValue(MyIndex, Crafts(lblCraftName.DataField).Materials(d).itemNum) & "/" & Crafts(lblCraftName.DataField).Materials(d).Count & ")"

        If item(itemNum).Type = ITEM_TYPE_PET Then
            descStr.Caption = Npc(Pets(item(itemNum).Datas(0)).num).Str & " Force"
            descDef.Caption = Npc(Pets(item(itemNum).Datas(0)).num).Def & " Défense"
        Else
            descStr.Caption = item(itemNum).StrReq & " Force"
            descDef.Caption = item(itemNum).DefReq & " Défense"
        End If
        descDex.Caption = item(itemNum).DexReq & " Dexterité"
        descHpMp.Caption = "PV: " & item(itemNum).AddHP & " PM: " & item(itemNum).AddSLP & " End: " & item(itemNum).AddSTP
        descSD.Caption = "FOR: " & item(itemNum).AddStr & " Def: " & item(itemNum).AddDef
        descMS.Caption = "Science: " & item(itemNum).AddSci & " Dexterité: " & item(itemNum).AddDex
        If (item(itemNum).Type >= ITEM_TYPE_WEAPON) And (item(itemNum).Type <= ITEM_TYPE_SHIELD) Then
            If item(itemNum).Datas(0) <= 0 Then Usure.Caption = "Usure : Ind." Else Usure.Caption = "Usure max : " & item(itemNum).Datas(0)
        End If
        desc.Caption = Trim$(item(itemNum).desc)
        descName.ForeColor = item(itemNum).NCoul
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub



Private Sub picProduct_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Long
Dim itemNum As Integer
d = Index
    itemNum = picProduct(d).DataField
    
    If itemNum > 0 Then
        If item(itemNum).Type = ITEM_TYPE_CURRENCY And Trim$(item(itemNum).desc) = vbNullString Then
            itmDesc.Height = 17
            itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            itmDesc.Left = fra_fenetre.Left
        ElseIf Trim$(item(itemNum).desc) = vbNullString Then
            itmDesc.Height = 161
            itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            itmDesc.Left = fra_fenetre.Left
        ElseIf Trim$(item(itemNum).desc) > vbNullString Then
            itmDesc.Height = 249
            itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            itmDesc.Left = fra_fenetre.Left
        End If

        descName.Caption = Trim$(item(itemNum).name) & "(" & Crafts(lblCraftName.DataField).Products(d).Count & ")"

        If item(itemNum).Type = ITEM_TYPE_PET Then
            descStr.Caption = Npc(Pets(item(itemNum).Datas(0)).num).Str & " Force"
            descDef.Caption = Npc(Pets(item(itemNum).Datas(0)).num).Def & " Défense"
        Else
            descStr.Caption = item(itemNum).StrReq & " Force"
            descDef.Caption = item(itemNum).DefReq & " Défense"
        End If
        descDex.Caption = item(itemNum).DexReq & " Dexterité"
        descHpMp.Caption = "PV: " & item(itemNum).AddHP & " PM: " & item(itemNum).AddSLP & " End: " & item(itemNum).AddSTP
        descSD.Caption = "FOR: " & item(itemNum).AddStr & " Def: " & item(itemNum).AddDef
        descMS.Caption = "Science: " & item(itemNum).AddSci & " Dexterité: " & item(itemNum).AddDex
        If (item(itemNum).Type >= ITEM_TYPE_WEAPON) And (item(itemNum).Type <= ITEM_TYPE_SHIELD) Then
            If item(itemNum).Datas(0) <= 0 Then Usure.Caption = "Usure : Ind." Else Usure.Caption = "Usure max : " & item(itemNum).Datas(0)
        End If
        desc.Caption = Trim$(item(itemNum).desc)
        descName.ForeColor = item(itemNum).NCoul
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub PicOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 1
    DragX = X
    DragY = Y
End Sub

Private Sub PicOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picOptions, X, Y)
End Sub

Private Sub PicOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub picquete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 5
DragX = X
DragY = Y
End Sub

Private Sub picquete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(picquete, X, Y)
End Sub

Private Sub picquete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub picRac_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Q As Long
Dim Qq As Long
Dim d As Byte
    If Button = 1 Then
        Call useRac(Index)
    End If
    If Button = 2 Then

    End If

    SDAD.Visible = False
    IDAD.Visible = False
End Sub

Private Sub picScreen_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    hauteurEcran = picScreen.Height - picRightClick.Height
    largeurEcran = picScreen.Width - picRightClick.Width

    If Button = 1 Then
        If Player(MyIndex).castedSpell > -1 Then
            Call SendUseSkill(Player(MyIndex).castedSpell, CurX, CurY)
            Exit Sub
        End If
    ElseIf Button = 2 Then
        frmMirage.picRightClick.Visible = False
        frmMirage.lbl_groupe_rightclick.Visible = False
        frmMirage.lbl_discussion_rightclick.Visible = False
        frmMirage.lbl_trade_rightclick.Visible = False
        frmMirage.lbl_adopt_rightclick.Visible = False
        frmMirage.lbl_abandon_rightclick.Visible = False
        
        Dim Target() As Integer
        Target = FindIndexAtPos(GetPlayerMap(MyIndex), CurX, CurY)
        
        If Target(1) >= 0 Then
            frmMirage.nom.DataField = Target(0)
            'playerIndex = I
            frmMirage.picRightClick.Top = Int(Y / PIC_Y) * PIC_Y + 50
            If frmMirage.picRightClick.Top > hauteurEcran Then
                frmMirage.picRightClick.Top = hauteurEcran
            End If
            frmMirage.picRightClick.Left = Int(X / PIC_X) * PIC_X
            If frmMirage.picRightClick.Left > largeurEcran Then
                frmMirage.picRightClick.Left = largeurEcran
            End If
            'Packet.WriteByte Target(1) 'Target Type
            'Packet.WriteInteger Target(0) 'Target index
            If Target(1) = PLAYER_TYPE Then
                frmMirage.picRightClick.DataField = PLAYER_TYPE
                frmMirage.nom.Caption = GetPlayerName(Target(0))

                If Target(0) <> MyIndex Then
                    frmMirage.lbl_groupe_rightclick.Visible = True
                    frmMirage.lbl_discussion_rightclick.Visible = True
                    frmMirage.lbl_trade_rightclick.Visible = True
                    frmMirage.picRightClick.Visible = True
                End If
            ElseIf Target(1) = PET_TYPE Then
                frmMirage.picRightClick.DataField = PET_TYPE
                frmMirage.nom.Caption = "Compagnon de : " & Player(frmMirage.nom.DataField).name
            
                If MyIndex = frmMirage.nom.DataField Then
                    frmMirage.lbl_abandon_rightclick.Visible = True
                End If
                frmMirage.picRightClick.Visible = True
            ElseIf Target(1) = NPC_TYPE Then
                frmMirage.picRightClick.DataField = NPC_TYPE
                frmMirage.nom.Caption = Npc(MapNpc(Target(0)).num).name

                If Pets(MyIndex).num = -1 Then
                    frmMirage.lbl_adopt_rightclick.Visible = True
                End If
                frmMirage.picRightClick.Visible = True
            End If
        End If
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not GettingMap And Not IsDead Then
    CurX = ((X + NewPlayerPicX - XMapPadding) \ 32)
    CurY = ((Y + NewPlayerPicY - YMapPadding) \ 32)
    
    If CurX < MinMapX Then
        CurX = MinMapX
    ElseIf CurX > MaxMapX Then
        CurX = MaxMapX
    End If
    
    If CurY < MinMapY Then
        CurY = MinMapY
    ElseIf CurY > MaxMapY Then
        CurY = MaxMapY
    End If

    PotX = X
    PotY = Y
    
    If CurX <> OldPCX Or CurY <> OldPCY Then Call CaseChange(CurX, CurY): OldPCX = CurX: OldPCY = CurY
    'End If
    itmDesc.Visible = False
End If
End Sub

Private Sub picspell_Click(Index As Integer)
    Call UseSkill(Index)
End Sub

Private Sub picspell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If dragAndDrop.Index = -1 Then
            dragAndDrop.Type = 1
            dragAndDrop.Index = Index
        End If
    End If
    If Button = 2 Then

    End If
End Sub

Private Sub picspell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dragAndDrop.Index > -1 Then
        Call ProcessDrag
    End If
End Sub

Private Sub picspell_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim pos As POINTAPI
    GetCursorPos pos
    
    For i = 0 To picRac.UBound
        If cSubclasserHooker.IsOver(picRac(i).hwnd, pos.X, pos.Y) Then
            'Debug.Print "ok rac"
            rac(i, 0) = dragAndDrop.Index
            rac(i, 1) = 1
            
            Call AffRac
            
            Call ResetDragAndDrop

            Exit Sub
        End If
    Next i
    
    Call ResetDragAndDrop
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub qf_Click()
picquete.Visible = False
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
    If Len(Str$(Minu)) > 2 Then Minute.Caption = Minu & ":" Else Minute.Caption = "0" & Minu & ":"
End If
If Seco <= 0 And Minu <= 0 Then
    seconde.Caption = 0
    Call MsgBox("La quête : " & Trim$(quete(Queten).nom) & " est terminer, le temps est écouler")
    Player(MyIndex).QueteEnCour = 0
    quetetimersec.Enabled = False
    tmpsquete.Visible = False
End If

If Len(Str$(Seco)) > 2 Then seconde.Caption = Seco Else seconde.Caption = "0" & Seco
Else
Player(MyIndex).QueteEnCour = 0
tmpsquete.Visible = False
quetetimersec.Enabled = False
End If

End Sub

Private Sub scrlBltText_Change()
Dim i As Long
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).time = i
    Next i
    
    MAX_BLT_LINE = scrlBltText.Value
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    lblLines.Caption = "Nbr de ligne écrite sur l'écran: " & scrlBltText.Value
End Sub

Private Sub ShieldImage_DblClick()
    Call TakeOutShield
End Sub

Public Sub TakeOutShield()
    If Player(MyIndex).ShieldSlot.num >= 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CTakeOutShield
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub ShieldImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Player(MyIndex).ShieldSlot.num >= 0 Then
        Call DisplayDescription(Player(MyIndex).ShieldSlot)
    Else
        frmMirage.itmDesc.Visible = False
    End If
End Sub

Private Sub SocketTCP_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then Call IncomingTCPData(bytesTotal)
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then Call IncomingData(bytesTotal)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ConOff > 0 Then Exit Sub

    Call HandleKeypresses(KeyAscii)
    If (KeyAscii = vbKeyReturn) Then KeyAscii = 0
    If (KeyAscii = UserCommand.item("action")) Then KeyAscii = 0
    If KeyAscii = vbKeyEscape Then
        If fra_fenetre.Visible = True Then fra_fenetre.Visible = False
        If picInfo.Visible = True Then picInfo.Visible = False
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
    If ConOff > 0 Then Exit Sub
    If txtMyTextBox.Visible = False Then
        Call CheckInput(1, keyCode, Shift)
    End If
    On Error Resume Next
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Mediaplayer.URL > vbNullString Then
    If Mediaplayer.Controls.currentPosition = 0 And Mediaplayer.currentMedia.name = Mid$(Map.Music, 1, Len(Map.Music) - 4) Then Call frmMirage.Mediaplayer.Controls.play
End If
End Sub

Private Sub Timer2_Timer()
    Call AffRac
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

Sub DropItem(ByVal InvNum As Long)
Dim GoldAmount As String
On Error GoTo Done

If InvNum < 0 Then Exit Sub
   
    If GetPlayerInvItemNum(MyIndex, InvNum) >= 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, InvNum)).Empilable <> 0 Then
            GoldAmount = InputBox("Combien de " & Trim$(item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") voulez vous jeter?", "Jeter " & Trim$(item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmMirage.Left, frmMirage.Top)
            If IsNumeric(GoldAmount) Then Call SendDropItem(InvNum, GoldAmount)
        Else
            Call SendDropItem(InvNum, 1)
        End If
    End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then MsgBox "Trop grande quantiter(erreur du logiciel)"
End Sub

Private Sub cmdAccess_Click()
Dim Packet As clsBuffer
    If txtName.Text = vbNullString Or txtAccess.Text = vbNullString Or Not IsNumeric(txtAccess.Text) Then Exit Sub
'    Set Packet = New clsBuffer
'
'    Packet.WriteLong CSetMemberAccess
'    Packet.WriteInteger FindPlayer(txtName.Text)
'    Packet.WriteInteger txtAccess.Text
'
'    SendData Packet.ToArray()
'    Set Packet = Nothing
End Sub

Private Sub cmdDisown_Click()
'Dim Packet As clsBuffer
'    If txtName.Text = vbNullString Then Exit Sub
'    Set Packet = New clsBuffer
'
'    Packet.WriteLong CFireMember
'    Packet.WriteInteger FindPlayer(txtName.Text)
'
'    SendData Packet.ToArray
'    Set Packet = Nothing
End Sub

Private Sub txtQ_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: txtQ.Visible = False
End Sub

Private Sub txtQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 4
DragX = X
DragY = Y
End Sub

Private Sub txtQ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call deplacer(txtQ, X, Y)
End Sub

Private Sub txtQ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub TxtQ2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: txtQ.Visible = False
End Sub

Private Sub deplacer(P_Picture As PictureBox, X As Single, Y As Single)

If DragImg <> 0 Then
    hauteurEcran = picScreen.Height - P_Picture.Height
    largeurEcran = picScreen.Width - P_Picture.Width
    
    valeur = P_Picture.Top + ((Y / twippy) - (DragY / twippy))
    If valeur >= 0 And valeur <= hauteurEcran Then
        P_Picture.Top = valeur
    ElseIf valeur < 0 Then
        P_Picture.Top = 0
    ElseIf valeur > hauteurEcran Then
        P_Picture.Top = hauteurEcran
    End If
    valeur = P_Picture.Left + ((X / twippx) - (DragX / twippx))
    If valeur >= 0 And valeur <= largeurEcran Then
        P_Picture.Left = valeur
    ElseIf valeur < 0 Then
        P_Picture.Left = 0
    ElseIf valeur > largeurEcran Then
        P_Picture.Left = largeurEcran
    End If
End If
End Sub

Sub reduire(P_Picture As PictureBox, P_Size As Integer)
    If P_Picture.Height > 315 / twippy Then P_Picture.Height = 315 / twippy Else P_Picture.Height = P_Size
End Sub

Sub CraftLoad(craftNum)
    Dim i As Integer

    lblCraftName.Caption = Trim$(Crafts(craftNum).name)
    lblCraftName.DataField = craftNum
    For i = 0 To MAX_MATERIALS
        ' Clear the objet display
        frmMirage.picMaterial(i).Picture = LoadPicture()
        frmMirage.picProduct(i).Picture = LoadPicture()
        
        If i < GetNbMaterials(craftNum) Then
            ' Setting the objet index in the data field
            frmMirage.picMaterial(i).DataField = Crafts(craftNum).Materials(i).itemNum
            ' Display the object and the state
            If GetPlayerInvItemTotalValue(MyIndex, Crafts(craftNum).Materials(i).itemNum) >= Crafts(craftNum).Materials(i).Count Then
                MaterialState(i).BorderColor = &HFF00&
            Else
                MaterialState(i).BorderColor = &HFF&
            End If
            MaterialState(i).Visible = True
            
            Call AffSurfPic(DD_ItemSurf, frmMirage.picMaterial(i), (item(Crafts(craftNum).Materials(i).itemNum).Pic - (item(Crafts(craftNum).Materials(i).itemNum).Pic \ 6) * 6) * PIC_X, (item(Crafts(craftNum).Materials(i).itemNum).Pic \ 6) * PIC_Y)
        Else
            frmMirage.picMaterial(i).DataField = -1
            ' Don't show the state
            MaterialState(i).Visible = False
        End If
        
        ' Setting the objet index in the data field
        If i < GetNbProducts(craftNum) Then
            ' Setting the objet index in the data field
            frmMirage.picProduct(i).DataField = Crafts(craftNum).Products(i).itemNum
            ' Display the object
            Call AffSurfPic(DD_ItemSurf, frmMirage.picProduct(i), (item(Crafts(craftNum).Products(i).itemNum).Pic - (item(Crafts(craftNum).Products(i).itemNum).Pic \ 6) * 6) * PIC_X, (item(Crafts(craftNum).Products(i).itemNum).Pic \ 6) * PIC_Y)
        Else
            frmMirage.picProduct(i).DataField = -1
        End If
    Next i
    
    picCraft.Visible = True
End Sub

Sub CraftsLoad()
    Dim i As Integer
    Dim oldPosition As Integer
    
    If picCraft.Visible Then
        oldPosition = lstCraft.ListIndex
    End If
    
    lstCraft.Clear
    
    If Player(MyIndex).Crafts.Count > 0 Then
        For i = 1 To Player(MyIndex).Crafts.Count
            lstCraft.AddItem (Crafts(Player(MyIndex).Crafts(i)).name)
            lstCraft.ItemData(lstCraft.ListCount - 1) = Player(MyIndex).Crafts(i)
        Next i
    End If
    
    If lstCraft.ListCount > oldPosition Then
        lstCraft.ListIndex = oldPosition
    Else
        If lstCraft.ListIndex > 0 Then
            lstCraft.ListIndex = 0
        End If
    End If
End Sub

Private Sub TimerCraft_Timer()
    ProgressBarCraft.Value = ProgressBarCraft.Value + 1
    If ProgressBarCraft.Value >= ProgressBarCraft.Max Then
        Dim Packet As clsBuffer
        
        'Disable timer
        ProgressBarCraft.Value = 0
        TimerCraft.Enabled = False
        
        If CanBuild Then
            Set Packet = New clsBuffer
            
            Packet.WriteLong CExecuteCraft
            Packet.WriteInteger lblCraftName.DataField
            
            SendData Packet.ToArray()
            Set Packet = Nothing
        End If
    End If
End Sub

Function CanBuild()
    Dim i, J, K As Integer
    Dim itemNum As Integer
    Dim TakenSlot(0 To MAX_INV) As Integer
    
    CanBuild = True
    For i = 0 To GetNbMaterials(lblCraftName.DataField) - 1
        If Crafts(lblCraftName.DataField).Materials(i).itemNum > 0 Then
            If GetPlayerInvItemTotalValue(MyIndex, Crafts(lblCraftName.DataField).Materials(i).itemNum) < Crafts(lblCraftName.DataField).Materials(i).Count Then
                CanBuild = False
            End If
        End If
    Next i
    
    If Not CanBuild Then
        MsgBox "Il vous manque des matériaux !"
        Exit Function
    End If
    
    ' Init all taken slot to false
    For i = 0 To MAX_INV
        TakenSlot(i) = False
    Next i
    
    For i = 0 To GetNbProducts(lblCraftName.DataField) - 1
        itemNum = Crafts(lblCraftName.DataField).Products(i).itemNum
        If itemNum > 0 Then
            CanBuild = False
            
            If item(itemNum).Empilable Then
                ' If currency then check to see if they already have an instance of the item and add it to that
                For J = 1 To MAX_INV
                    If GetPlayerInvItemNum(MyIndex, J) = itemNum Then
                        CanBuild = True
                        Exit For
                    End If
                Next J
            End If
            
            If Not CanBuild Then ' Not a currency object or not already in inventory
                For K = 1 To Crafts(lblCraftName.DataField).Products(i).Count
                    CanBuild = False
                    For J = 0 To MAX_INV
                        ' Try to find an open free slot
                        If GetPlayerInvItemNum(MyIndex, J) <= 0 And Not TakenSlot(J) Then
                            CanBuild = True
                            TakenSlot(J) = True
                            Exit For
                        End If
                    Next J
                Next K
            End If
            If Not CanBuild Then
                Exit For
            End If
        End If
    Next i
    
    If Not CanBuild Then
        MsgBox "Vous n'avez pas assez de place dans l'inventaire pour recevoir tous les produits."
        Exit Function
    End If
End Function

Private Sub VScroll1_Change()
    Picture9.Top = VScroll1.Value * -1 * 0.7
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub

Private Sub VScroll2_Change()
    Picture11.Top = VScroll2.Value * -1 * 20
End Sub

Private Sub VScroll2_Scroll()
    VScroll2_Change
End Sub

Private Sub WeaponImage_DblClick()
    Call TakeOutWeapon
End Sub

Public Sub TakeOutWeapon()
    If Player(MyIndex).WeaponSlot.num >= 0 Then
        Dim Packet As clsBuffer
    
        Set Packet = New clsBuffer
        
        Packet.WriteLong CTakeOutWeapon
        
        SendData Packet.ToArray()
        
        Set Packet = Nothing
    End If
End Sub

Private Sub WeaponImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Player(MyIndex).WeaponSlot.num >= 0 Then
        Call DisplayDescription(Player(MyIndex).WeaponSlot)
    Else
        frmMirage.itmDesc.Visible = False
    End If
End Sub

Private Sub cSubclasserHooker_SubclassMessage(ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    Dim sMessage As String
    Dim ctl As Control
    Dim bHandled As Boolean
    Dim bOver As Boolean
    Dim Xpos As Long
    Dim Ypos As Long
    
    Xpos = lParam And 65535
    Ypos = lParam / 65536
    
    If Message = &H7B Then ' 123
        If Player(MyIndex).castedSpell > -1 Then
            Player(MyIndex).castedSpell = -1
            
            Call RestoreCursor
        End If
    ElseIf Message = &H20A Then

      For Each ctl In Me.Controls
        ' Is the mouse over the control
        On Error Resume Next
        bOver = (ctl.Visible And cSubclasserHooker.IsOver(ctl.hwnd, Xpos, Ypos))
        On Error GoTo 0
        
        If bOver Then
          ' If so, respond accordingly
          bHandled = True
          Select Case True
          
            Case ctl = frmMirage.picInv3
                Dim newValue As Integer
                If wParam = &H780000 Then
                    newValue = frmMirage.VScroll1.Value - 50
                ElseIf wParam = &HFF880000 Then
                    newValue = frmMirage.VScroll1.Value + 50
                End If
                
                If newValue < frmMirage.VScroll1.Min Then
                    newValue = frmMirage.VScroll1.Min
                ElseIf newValue > frmMirage.VScroll1.Max Then
                    newValue = frmMirage.VScroll1.Max
                End If
                
                frmMirage.VScroll1.Value = newValue
    
            Case Else
              bHandled = False
    
          End Select
          If bHandled Then Exit Sub
        End If
        bOver = False
      Next ctl

    End If
    
End Sub

Public Sub RefreshStatsBonus()
    If Player(MyIndex).StrBonus > 0 Then
        frmMirage.lblSTRBonus.Caption = "+" & Player(MyIndex).StrBonus
        frmMirage.lblSTRBonus.ForeColor = vbBlue
    ElseIf Player(MyIndex).StrBonus < 0 Then
        frmMirage.lblSTRBonus.Caption = Player(MyIndex).StrBonus
        frmMirage.lblSTRBonus.ForeColor = vbRed
    Else
        frmMirage.lblSTRBonus.Caption = ""
    End If
    
    If Player(MyIndex).DefBonus > 0 Then
        frmMirage.lblDEFBonus.Caption = "+" & Player(MyIndex).DefBonus
        frmMirage.lblDEFBonus.ForeColor = vbBlue
    ElseIf Player(MyIndex).DefBonus < 0 Then
        frmMirage.lblDEFBonus.Caption = Player(MyIndex).DefBonus
        frmMirage.lblDEFBonus.ForeColor = vbRed
    Else
        frmMirage.lblDEFBonus.Caption = ""
    End If
    
    If Player(MyIndex).DexBonus > 0 Then
        frmMirage.lblDEXBonus.Caption = "+" & Player(MyIndex).DexBonus
        frmMirage.lblDEXBonus.ForeColor = vbBlue
    ElseIf Player(MyIndex).DexBonus < 0 Then
        frmMirage.lblDEXBonus.Caption = Player(MyIndex).DexBonus
        frmMirage.lblDEXBonus.ForeColor = vbRed
    Else
        frmMirage.lblDEXBonus.Caption = ""
    End If
    
    If Player(MyIndex).SciBonus > 0 Then
        frmMirage.lblSCIBonus.Caption = "+" & Player(MyIndex).SciBonus
        frmMirage.lblSCIBonus.ForeColor = vbBlue
    ElseIf Player(MyIndex).SciBonus < 0 Then
        frmMirage.lblSCIBonus.Caption = Player(MyIndex).SciBonus
        frmMirage.lblSCIBonus.ForeColor = vbRed
    Else
        frmMirage.lblSCIBonus.Caption = ""
    End If
    
    If Player(MyIndex).LangBonus > 0 Then
        frmMirage.lblLANGBonus.Caption = "+" & Player(MyIndex).LangBonus
        frmMirage.lblLANGBonus.ForeColor = vbBlue
    ElseIf Player(MyIndex).LangBonus < 0 Then
        frmMirage.lblLANGBonus.Caption = Player(MyIndex).LangBonus
        frmMirage.lblLANGBonus.ForeColor = vbRed
    Else
        frmMirage.lblLANGBonus.Caption = ""
    End If
End Sub
