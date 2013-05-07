VERSION 5.00
Begin VB.Form frmbank 
   BorderStyle     =   0  'None
   Caption         =   "Banque"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmbank.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   672
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picture3 
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
      Height          =   735
      Left            =   6720
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   59
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label dur 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Durabilité"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   915
         TabIndex        =   62
         Top             =   480
         Width           =   705
      End
      Begin VB.Label nb 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   975
         TabIndex        =   61
         Top             =   240
         Width           =   585
      End
      Begin VB.Label nom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1095
         TabIndex        =   60
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   22
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   24
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   12
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   21
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   20
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   8
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   23
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   16
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   4
      Left            =   3120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   19
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   18
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   17
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   15
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   14
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   13
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   11
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   10
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   9
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   7
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   6
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   5
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   3
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   2
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   20
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   58
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   10
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   57
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   30
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   56
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   25
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   55
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   15
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   54
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   5
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   53
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   29
      Left            =   8160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   52
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   28
      Left            =   7440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   51
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   27
      Left            =   6720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   50
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   26
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   49
      Top             =   4200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   24
      Left            =   8160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   48
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   23
      Left            =   7440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   47
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   22
      Left            =   6720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   46
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   21
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   45
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   19
      Left            =   8160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   44
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   18
      Left            =   7440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   43
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   17
      Left            =   6720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   42
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   16
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   41
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   14
      Left            =   8160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   40
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   13
      Left            =   7440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   39
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   12
      Left            =   6720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   11
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   2400
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   9
      Left            =   8160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   8
      Left            =   7440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   35
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   7
      Left            =   6720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   6
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   4
      Left            =   8160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   3
      Left            =   7440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   31
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   2
      Left            =   6720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   1
      Left            =   6000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   1
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   5985
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   945
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label inve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventaire :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   795
   End
   Begin VB.Label coffre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coffre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   510
   End
   Begin VB.Label jeter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Label OK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label jinv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4980
      Width           =   1575
   End
End
Attribute VB_Name = "frmbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SInv As Long
Dim SCof As Long
Dim DInv As Boolean
Dim DCof As Boolean

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dr = True
drx = X
dry = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
dr = False
drx = 0
dry = 0
End Sub

Public Sub jeter_Click()
Dim Packet As clsBuffer
Dim ival As Long

If SCof = 0 Then Call MsgBox("Veuillez séléctionner un slot dans le coffre!!", vbCritical, "Erreur"): Exit Sub

If CoffreTmp(SCof).Numeros <= 0 Then Call MsgBox("Aucun objet dans le slot" & SCof & " du coffre!!!", vbCritical, "Erreur"): Exit Sub

ival = CoffreTmp(SCof).valeur

If ival <= 0 Then ival = 1

cont = MsgBox("Voulez vous vraiment jeter " & ival & Trim$(item(CoffreTmp(SCof).Numeros).name) & " du coffre?? il sera supprimé définitivement!!", vbYesNo, "Demande")

Set Packet = New clsBuffer

If cont = vbYes Then
    Set Packet = New clsBuffer
    
    Packet.WriteLong CChangeSafeItem
    Packet.WriteInteger SCof
    Packet.WriteInteger 0
    Packet.WriteInteger 0
    Packet.WriteInteger 0
    
    SendData Packet.ToArray()
    Set Packet = Nothing
    'Packet = "MODIFCOFFRE" & SEP_CHAR & SCof & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR: Call SendData(Packet)
End If

CoffreTmp(SCof).Numeros = 0
CoffreTmp(SCof).valeur = 0
CoffreTmp(SCof).Durabiliter = 0
frmbank.ActPic
'Packet = "COFFREITEM" & SEP_CHAR & END_CHAR

'Call SendData(Packet)
End Sub

'Public Sub jeter_Click()
'Dim Packet As String
'Dim ival As Long
'
'If SCof = 0 Then Call MsgBox("Veuillez séléctionner un slot dans le coffre!!", vbCritical, "Erreur"): Exit Sub
'
'If CoffreTmp(SCof).Numeros <= 0 Then Call MsgBox("Aucun objet dans le slot" & SCof & " du coffre!!!", vbCritical, "Erreur"): Exit Sub
'
'ival = CoffreTmp(SCof).valeur
'
'If ival <= 0 Then ival = 1
'
'cont = MsgBox("Voulez vous vraiment jeter " & ival & Trim$(Item(CoffreTmp(SCof).Numeros).name) & " du coffre?? il sera supprimé définitivement!!", vbYesNo, "Demande")
'
'If cont = vbYes Then Packet = "MODIFCOFFRE" & SEP_CHAR & SCof & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR: Call SendData(Packet)
'
'Packet = "COFFREITEM" & SEP_CHAR & END_CHAR
'
'Call SendData(Packet)
'End Sub

Private Sub jinv_Click()
Dim Packet As String
Dim ival As Long

If SInv = 0 Then Call MsgBox("Veuillez séléctionner un slot dans l'inventaire!!", vbCritical, "Erreur"): Exit Sub

If Val(GetPlayerInvItemNum(MyIndex, SInv)) <= 0 Then Call MsgBox("Aucun objet dans le slot" & SInv & " de l'inventaire!!!", vbCritical, "Erreur"): Exit Sub

If GetPlayerInvItemNum(MyIndex, SInv) > 0 And GetPlayerInvItemNum(MyIndex, SInv) <= MAX_ITEMS Then
    If item(GetPlayerInvItemNum(MyIndex, SInv)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, SInv)).Empilable <> 0 Then
        ival = InputBox("Combien de " & Trim$(item(GetPlayerInvItemNum(MyIndex, SInv)).name) & "(" & GetPlayerInvItemValue(MyIndex, SInv) & ") voulez vous jeter?", "Jeter " & Trim$(item(GetPlayerInvItemNum(MyIndex, SInv)).name), 0, frmMirage.Left, frmMirage.Top)
    Else
        ival = 1
    End If
End If
'ival = GetPlayerInvItemValue(MyIndex, SInv)
If IsNumeric(ival) Then
    'If ival <= 0 Then ival = 1
    
    cont = MsgBox("Voulez vous vraiment jeter " & ival & Trim$(item(GetPlayerInvItemNum(MyIndex, SInv)).name) & " de l'inventaire?? il sera supprimer définitivement!!", vbYesNo, "Demande")
    
    If cont = vbYes Then
        Call SendDestroyItem(SInv, ival)
        'Packet = "MODIFINV" & SEP_CHAR & SInv & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & SCof & SEP_CHAR & END_CHAR: Call SendData(Packet)
    End If
    'Packet = "COFFREITEM" & SEP_CHAR & END_CHAR
    '
    'Call SendData(Packet)
    
    frmbank.ActPic
End If
End Sub

Public Sub Form_Load()
Dim Packet As clsBuffer
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist(App.Path & Rep_Theme & "\Jeu\Bank" & Ending) Then Me.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Bank" & Ending)
    Next i
'Picturesprite.Picture = LoadPNG(App.Path & "\GFX\items.png", True)

inve.Caption = "Inventaire de " & GetPlayerName(MyIndex) & " :"
coffre.Caption = "Coffre de " & GetPlayerName(MyIndex) & " :"

Set Packet = New clsBuffer

Packet.WriteLong CGetSafeItem

SendData Packet.ToArray()
Set Packet = Nothing

'Packet = "COFFREITEM" & SEP_CHAR & END_CHAR

'Call SendData(Packet)
DInv = False
DCof = False
Call ActPic
End Sub

Public Sub MoveItem(ByVal TargetSlot As Long, ByVal SourceSlot As Long, ByVal MoveType As Long)
    Dim TargetNum, SourceNum As Integer
    Dim TargetVal, SourceVal As Long
    Dim TargetDur, SourceDur As Long
    Dim Packet As clsBuffer

    ' Control the slot number
    If MoveType = MOVE_TO_INV Then
        If TargetSlot > 24 Or TargetSlot < 1 Then Exit Sub
        If SourceSlot > 30 Or SourceSlot < 1 Then Exit Sub
    ElseIf (MoveType = MOVE_TO_SAFE) Then
        If SourceSlot > 24 Or SourceSlot < 1 Then Exit Sub
        If TargetSlot > 30 Or TargetSlot < 1 Then Exit Sub
    End If
    
    TargetNum = GetPlayerInvItemNum(MyIndex, TargetSlot)
    TargetVal = GetPlayerInvItemValue(MyIndex, TargetSlot)
    TargetDur = GetPlayerInvItemDur(MyIndex, TargetSlot)
    
    SourceNum = CoffreTmp(SourceSlot).Numeros
    SourceVal = CoffreTmp(SourceSlot).valeur
    SourceDur = CoffreTmp(SourceSlot).Durabiliter
    
    If TargetVal < 0 Then TargetVal = 0
    If SourceVal < 0 Then SourceVal = 0
    
    If SourceNum <= 0 Then Call MsgBox("Aucun objet dans le slot " & SourceSlot & " source!!!", vbCritical, "Erreur"): Exit Sub
    
    If TargetNum <> 0 Then
        If item(TargetNum).Type <> ITEM_TYPE_CURRENCY Or item(TargetNum).Empilable = 0 Then
            Call MsgBox("Il y a déja un objet dans le slot " & TargetSlot & " de destination!!", vbCritical, "Erreur")
            Exit Sub
        Else
            If TargetNum <> SourceNum Then
                Call MsgBox("Il y a déja un objet dans le slot " & TargetSlot & " de destination!!", vbCritical, "Erreur")
                Exit Sub
            End If
        End If
    End If
    
    If item(SourceNum).Type = ITEM_TYPE_CURRENCY Or item(SourceNum).Empilable <> 0 Then
    
        Nbi = InputBox("Combiens d'objet(s) voulez-vous mettre dans le coffre?", "Demande")
    
        If Not IsNumeric(Nbi) Then Call MsgBox("Entrez un nombre SVP!!", vbCritical, "Erreur"): Exit Sub
        
        If Val(Nbi) > SourceVal Then Call MsgBox("Valeur supérieur au nombre d'objet!!", vbCritical, "Erreur"): Exit Sub
        
        If Val(Nbi) <= 0 Then Exit Sub
    
        If TargetNum = SourceNum Or TargetNum = 0 Then
        
            Set Packet = New clsBuffer
        
            If MoveType = MOVE_TO_INV Then
                Packet.WriteLong CChangeInv
            ElseIf MoveType = MOVE_TO_SAFE Then
                CoffreTmp(TargetSlot).Numeros = SourceNum
                CoffreTmp(TargetSlot).valeur = TargetVal + Val(Nbi)
                CoffreTmp(TargetSlot).Durabiliter = SourceDur
            
                Packet.WriteLong CChangeSafeItem
            End If
    
            Packet.WriteInteger TargetSlot
            Packet.WriteLong SourceNum
            Packet.WriteLong TargetVal + Val(Nbi)
            Packet.WriteLong SourceDur
    
            'Packet = "DANSINV" & SEP_CHAR & SlotI & SEP_CHAR & Cnum & SEP_CHAR & ival + Val(Nbi) & SEP_CHAR & Cdur & SEP_CHAR & SlotC & SEP_CHAR & END_CHAR
    
            SendData Packet.ToArray()
            Set Packet = Nothing
            'Call SendData(Packet)
            
            If MoveType = MOVE_TO_INV Then
                If SourceVal - Val(Nbi) > 0 Then
                    CoffreTmp(SourceSlot).valeur = SourceVal - Val(Nbi)
                Else ' Sourceval - Val(Nbi) = 0
                    CoffreTmp(SourceSlot).Numeros = 0
                    CoffreTmp(SourceSlot).valeur = 0
                    CoffreTmp(SourceSlot).Durabiliter = 0
                End If
                Set Packet = New clsBuffer
                Packet.WriteLong CChangeSafeItem
                
                Packet.WriteInteger SourceSlot
                Packet.WriteLong CoffreTmp(SourceSlot).Numeros
                Packet.WriteLong CoffreTmp(SourceSlot).valeur
                Packet.WriteLong CoffreTmp(SourceSlot).Durabiliter
                    
                SendData Packet.ToArray()
                Set Packet = Nothing
            ElseIf MoveType = MOVE_TO_SAFE Then
                If SourceVal - Val(Nbi) > 0 Then
                    Set Packet = New clsBuffer
                    
                    Packet.WriteLong CChangeInv
                    
                    Packet.WriteInteger SourceSlot
                    Packet.WriteInteger SourceNum
                    Packet.WriteLong SourceVal - Val(Nbi)
                    Packet.WriteLong SourceDur
                    
                    SendData Packet.ToArray()
                    Set Packet = Nothing
                Else ' Sourceval - Val(Nbi) = 0
                    Set Packet = New clsBuffer
                    
                    Packet.WriteLong CChangeInv
                    
                    Packet.WriteInteger SourceSlot
                    Packet.WriteInteger 0
                    Packet.WriteLong 0
                    Packet.WriteLong 0
                    
                    SendData Packet.ToArray()
                    Set Packet = Nothing
                End If
            End If
        End If
    Else
        Set Packet = New clsBuffer
    
        If MoveType = MOVE_TO_INV Then
            Packet.WriteLong CChangeInv
        ElseIf MoveType = MOVE_TO_SAFE Then
            Packet.WriteLong CChangeSafeItem
        End If

        Packet.WriteInteger TargetSlot
        Packet.WriteLong SourceNum
        Packet.WriteLong SourceVal
        Packet.WriteLong SourceDur
        
        SendData Packet.ToArray()
        Set Packet = Nothing
        
        If MoveType = MOVE_TO_INV Then
            CoffreTmp(SourceSlot).Numeros = 0
            CoffreTmp(SourceSlot).valeur = 0
            CoffreTmp(SourceSlot).Durabiliter = 0
            
            Set Packet = New clsBuffer
            Packet.WriteLong CChangeSafeItem
            
            Packet.WriteInteger SourceSlot
            Packet.WriteLong CoffreTmp(SourceSlot).Numeros
            Packet.WriteLong CoffreTmp(SourceSlot).valeur
            Packet.WriteLong CoffreTmp(SourceSlot).Durabiliter
                
            SendData Packet.ToArray()
            Set Packet = Nothing
        ElseIf MoveType = MOVE_TO_SAFE Then
            Set Packet = New clsBuffer
            
            Packet.WriteLong CChangeInv
            
            Packet.WriteInteger SourceSlot
            Packet.WriteInteger 0
            Packet.WriteLong 0
            Packet.WriteLong 0
            
            SendData Packet.ToArray()
            Set Packet = Nothing
        End If
    End If
    
    Call ActPic
End Sub

'Private Sub DansInv(ByVal SlotI As Long, ByVal SlotC As Long)
'Dim Packet As String
'Dim Inum As Long
'Dim ival As Long
'Dim Idur As Long
'Dim Cnum As Long
'Dim Cval As Long
'Dim Cdur As Long
'Dim Nbi As String
'
'If SlotI > 24 Or SlotI < 1 Then Exit Sub
'If SlotC > 30 Or SlotC < 1 Then Exit Sub
'
'Inum = GetPlayerInvItemNum(MyIndex, SlotI)
'ival = GetPlayerInvItemValue(MyIndex, SlotI)
'Idur = GetPlayerInvItemDur(MyIndex, SlotI)
'
'Cnum = CoffreTmp(SlotC).Numeros
'Cval = CoffreTmp(SlotC).valeur
'Cdur = CoffreTmp(SlotC).Durabiliter
'
'If Cval < 0 Then Cval = 0
'If ival < 0 Then ival = 0
'
'If Cnum <= 0 Then Call MsgBox("Aucun objet dans le slot " & SlotC & " du coffre!!!", vbCritical, "Erreur"): Exit Sub
'
'If Inum <> 0 Then
'    If Item(Inum).Type <> ITEM_TYPE_CURRENCY Or Item(Inum).Empilable = 0 Then
'        Call MsgBox("Il y a déja un objet dans le slot " & SlotI & " de l'inventaire!!", vbCritical, "Erreur")
'        Exit Sub
'    Else
'        If Cnum <> Inum Then
'            Call MsgBox("Il y a déja un objet dans le slot " & SlotI & " de l'inventaire!!", vbCritical, "Erreur")
'            Exit Sub
'        End If
'    End If
'End If
'
'If Item(Cnum).Type = ITEM_TYPE_CURRENCY Or Item(Cnum).Empilable <> 0 Then
'
'    Nbi = InputBox("Combiens d'objet(s) voulez-vous métre dans le coffre?", "Demande")
'
'    If IsNumeric(Nbi) = False Then Call MsgBox("Entrez un nombre SVP!!", vbCritical, "Erreur"): Exit Sub
'
'    If Val(Nbi) > Cval Then Call MsgBox("Valeur supérieur au nombre d'objet!!", vbCritical, "Erreur"): Exit Sub
'
'    If Val(Nbi) <= 0 Then Exit Sub
'
'End If
'
'If Item(Cnum).Type = ITEM_TYPE_CURRENCY Or Item(Cnum).Empilable <> 0 Then
'
'    If Inum = Cnum Or Inum = 0 Then
'
'        Packet = "DANSINV" & SEP_CHAR & SlotI & SEP_CHAR & Cnum & SEP_CHAR & ival + Val(Nbi) & SEP_CHAR & Cdur & SEP_CHAR & SlotC & SEP_CHAR & END_CHAR
'
'        Call SendData(Packet)
'
'        Call SetPlayerInvItemValue(MyIndex, SlotI, ival + Val(Nbi))
'        Call SetPlayerInvItemNum(MyIndex, SlotI, Cnum)
'        Call SetPlayerInvItemDur(MyIndex, SlotI, Cdur)
'
'        If Cval - Val(Nbi) > 0 Then
'            CoffreTmp(SlotC).Numeros = Cnum
'            CoffreTmp(SlotC).valeur = Cval - Val(Nbi)
'            CoffreTmp(SlotC).Durabiliter = Cdur
'
'            Packet = "MODIFCOFFRE" & SEP_CHAR & SlotC & SEP_CHAR & Cnum & SEP_CHAR & Cval - Val(Nbi) & SEP_CHAR & Cdur & SEP_CHAR & SlotI & SEP_CHAR & END_CHAR
'
'            Call SendData(Packet)
'        Else
'            CoffreTmp(SlotC).Numeros = 0
'            CoffreTmp(SlotC).valeur = 0
'            CoffreTmp(SlotC).Durabiliter = 0
'
'            Packet = "MODIFCOFFRE" & SEP_CHAR & SlotC & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & SlotI & SEP_CHAR & END_CHAR
'
'            Call SendData(Packet)
'        End If
'    End If
'
'Else
'
'    Packet = "DANSINV" & SEP_CHAR & SlotI & SEP_CHAR & Cnum & SEP_CHAR & Cval & SEP_CHAR & Cdur & SEP_CHAR & SlotC & SEP_CHAR & END_CHAR
'
'    Call SendData(Packet)
'
'    Call SetPlayerInvItemValue(MyIndex, SlotI, Cval)
'    Call SetPlayerInvItemNum(MyIndex, SlotI, Cnum)
'    Call SetPlayerInvItemDur(MyIndex, SlotI, Cdur)
'
'    CoffreTmp(SlotC).Numeros = 0
'    CoffreTmp(SlotC).valeur = 0
'    CoffreTmp(SlotC).Durabiliter = 0
'
'    Packet = "MODIFCOFFRE" & SEP_CHAR & SlotC & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & SlotI & SEP_CHAR & END_CHAR
'
'    Call SendData(Packet)
'End If
'
'Packet = "COFFREITEM" & SEP_CHAR & END_CHAR
'
'Call SendData(Packet)
'
'Call ActPic
'End Sub
'
'Public Sub DansCoffre(ByVal SlotC As Long, ByVal SlotI As Long)
'Dim Packet As String
'Dim Inum As Long
'Dim ival As Long
'Dim Idur As Long
'Dim Cnum As Long
'Dim Cval As Long
'Dim Cdur As Long
'Dim Nbi As String
'
'If SlotI > 24 Or SlotI < 1 Then Exit Sub
'If SlotC > 30 Or SlotC < 1 Then Exit Sub
'
'Inum = GetPlayerInvItemNum(MyIndex, SlotI)
'ival = GetPlayerInvItemValue(MyIndex, SlotI)
'Idur = GetPlayerInvItemDur(MyIndex, SlotI)
'
'Cnum = CoffreTmp(SlotC).Numeros
'Cval = CoffreTmp(SlotC).valeur
'Cdur = CoffreTmp(SlotC).Durabiliter
'
'If SlotI = GetPlayerHelmetSlot(MyIndex) Then Call SetPlayerHelmetSlot(MyIndex, 0)
'If SlotI = GetPlayerArmorSlot(MyIndex) Then Call SetPlayerArmorSlot(MyIndex, 0)
'If SlotI = GetPlayerShieldSlot(MyIndex) Then Call SetPlayerShieldSlot(MyIndex, 0)
'If SlotI = GetPlayerWeaponSlot(MyIndex) Then Call SetPlayerWeaponSlot(MyIndex, 0)
'
'If Cval < 0 Then Cval = 0
'If ival < 0 Then ival = 0
'
'If Inum <= 0 Then Call MsgBox("Aucun objet dans le slot " & SlotI & " de l'inventaire!!!", vbCritical, "Erreur"): Exit Sub
'
'If Cnum <> 0 Then
'    If Item(Cnum).Type <> ITEM_TYPE_CURRENCY Or Item(Cnum).Empilable = 0 Then
'        Call MsgBox("Il y a déja un objet dans le slot " & SlotC & " du coffre!!", vbCritical, "Erreur")
'        Exit Sub
'    ElseIf Cnum <> Inum Then
'        Call MsgBox("Il y a déja un objet dans le slot " & SlotC & " du coffre!!", vbCritical, "Erreur")
'        Exit Sub
'    End If
'End If
'
'If Item(Inum).Type = ITEM_TYPE_CURRENCY Or Item(Inum).Empilable <> 0 Then
'
'    Nbi = InputBox("Combiens d'objet(s) voulez-vous métre dans le coffre?", "Demande")
'
'    If IsNumeric(Nbi) = False Then Call MsgBox("Entrez un nombre SVP!!", vbCritical, "Erreur"): Exit Sub
'
'    If Val(Nbi) > ival Then Call MsgBox("Valeur supérieur au nombre d'objet!!", vbCritical, "Erreur"): Exit Sub
'
'    If Val(Nbi) <= 0 Then Exit Sub
'
'End If
'
'If Item(Inum).Type = ITEM_TYPE_CURRENCY Or Item(Inum).Empilable <> 0 Then
'    If Inum = Cnum Or Cnum = 0 Then
'
'        Packet = "DANSCOFFRE" & SEP_CHAR & SlotC & SEP_CHAR & GetPlayerInvItemNum(MyIndex, SlotI) & SEP_CHAR & Cval + Val(Nbi) & SEP_CHAR & GetPlayerInvItemDur(MyIndex, SlotI) & SEP_CHAR & SlotI & SEP_CHAR & END_CHAR
'
'        Call SendData(Packet)
'
'        CoffreTmp(SlotC).Numeros = GetPlayerInvItemNum(MyIndex, SlotI)
'        CoffreTmp(SlotC).valeur = Cval + Val(Nbi)
'        CoffreTmp(SlotC).Durabiliter = GetPlayerInvItemDur(MyIndex, SlotI)
'
'        If Val(GetPlayerInvItemValue(MyIndex, SlotI)) - Val(Nbi) > 0 Then
'
'            Packet = "MODIFINV" & SEP_CHAR & SlotI & SEP_CHAR & Inum & SEP_CHAR & (Val(GetPlayerInvItemValue(MyIndex, SlotI)) - Val(Nbi)) & SEP_CHAR & Idur & SEP_CHAR & SlotC & SEP_CHAR & END_CHAR
'
'            Call SendData(Packet)
'
'            Call SetPlayerInvItemValue(MyIndex, SlotI, Val(GetPlayerInvItemValue(MyIndex, SlotI)) - Val(Nbi))
'        Else
'
'            Call SetPlayerInvItemValue(MyIndex, SlotI, 0)
'            Call SetPlayerInvItemNum(MyIndex, SlotI, 0)
'            Call SetPlayerInvItemDur(MyIndex, SlotI, 0)
'
'            Packet = "MODIFINV" & SEP_CHAR & SlotI & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & SlotC & SEP_CHAR & END_CHAR
'
'            Call SendData(Packet)
'        End If
'    End If
'
'Else
'
'    Packet = "DANSCOFFRE" & SEP_CHAR & SlotC & SEP_CHAR & GetPlayerInvItemNum(MyIndex, SlotI) & SEP_CHAR & GetPlayerInvItemValue(MyIndex, SlotI) & SEP_CHAR & GetPlayerInvItemDur(MyIndex, SlotI) & SEP_CHAR & SlotI & SEP_CHAR & END_CHAR
'
'    Call SendData(Packet)
'
'    CoffreTmp(SlotC).Numeros = GetPlayerInvItemNum(MyIndex, SlotI)
'    CoffreTmp(SlotC).valeur = GetPlayerInvItemValue(MyIndex, SlotI)
'    CoffreTmp(SlotC).Durabiliter = GetPlayerInvItemDur(MyIndex, SlotI)
'
'    Call SetPlayerInvItemValue(MyIndex, SlotI, 0)
'    Call SetPlayerInvItemNum(MyIndex, SlotI, 0)
'    Call SetPlayerInvItemDur(MyIndex, SlotI, 0)
'
'    Packet = "MODIFINV" & SEP_CHAR & SlotI & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & SlotC & SEP_CHAR & END_CHAR
'
'    Call SendData(Packet)
'End If
'
'Packet = "COFFREITEM" & SEP_CHAR & END_CHAR
'
'Call SendData(Packet)
'
'Call ActPic
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.Left + (X - drx), Me.Top + (Y - dry))
If Me.Left > Screen.Width Or Me.Top > Screen.Height Then Me.Top = Screen.Height \ 2: Me.Left = Screen.Width \ 2
Picture3.Visible = False
End Sub

Private Sub OK_Click()
SInv = 0
SCof = 0
frmMirage.txtQ.Visible = False
Unload Me
End Sub

Private Sub Picture1_Click(Index As Integer)
SInv = Index
Shape3.Visible = True
Shape3.Left = Picture1(SInv).Left - 1
Shape3.Top = Picture1(SInv).Top - 1
DCof = False
DInv = False
End Sub

Private Sub Picture1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If DInv = True Then DInv = False: Exit Sub
DCof = False
SInv = Index
Call MoveItem(SInv, SCof, MOVE_TO_INV)
End Sub

Private Sub Picture1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If DInv = False And DCof = False Then
    SInv = Index
    SCof = 0
    DInv = True
    DCof = False
End If
Shape3.Visible = True
Shape3.Left = Picture1(Index).Left - 1
Shape3.Top = Picture1(Index).Top - 1
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Inum As Long
Dim ival As Long
Dim Idur As String

Inum = GetPlayerInvItemNum(MyIndex, Index)
ival = GetPlayerInvItemValue(MyIndex, Index)
Idur = Str$(GetPlayerInvItemDur(MyIndex, Index))

If Val(Idur) <= 0 Then Idur = "Ind."

If ival = 0 Then ival = 1

If Inum > 0 Then
    Picture3.Top = Picture1(Index).Top + 32
    Picture3.Left = Picture1(Index).Left - 40
    nom.Caption = "Nom : " & Trim$(item(Inum).name)
    nb.Caption = "  Nombre : " & ival
    dur.Caption = "  Durabilité : " & Idur
    Picture3.Visible = True
End If

End Sub

Private Sub Picture2_Click(Index As Integer)
SCof = Index
Shape1.Visible = True
Shape1.Left = Picture2(SCof).Left - 1
Shape1.Top = Picture2(SCof).Top - 1
DCof = False
DInv = False
End Sub

Private Sub Picture2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If DCof = True Then DCof = False: Exit Sub
DInv = False
SCof = Index
Call MoveItem(SCof, SInv, MOVE_TO_SAFE)
SInv = 0
SCof = 0
End Sub

Private Sub Picture2_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If DCof = False And DInv = False Then
    SCof = Index
    SInv = 0
    DCof = True
    DInv = False
End If
Shape1.Visible = True
Shape1.Left = Picture2(Index).Left - 1
Shape1.Top = Picture2(Index).Top - 1
End Sub

Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cnum As Long
Dim Cval As Long
Dim Cdur As String

Cnum = CoffreTmp(Index).Numeros
Cval = CoffreTmp(Index).valeur
Cdur = Str$(CoffreTmp(Index).Durabiliter)

If Val(Cdur) <= 0 Then Cdur = "Ind."

If Cval = 0 Then Cval = 1

If Cnum > 0 Then
    Picture3.Top = Picture2(Index).Top + 32
    Picture3.Left = Picture2(Index).Left - 89
    nom.Caption = "Nom : " & Trim$(item(Cnum).name)
    nb.Caption = "  Nombre : " & Cval
    dur.Caption = "  Durabilité : " & Cdur
    Picture3.Visible = True
End If
End Sub

Sub ActPic()
For i = 1 To 24
    Inums = GetPlayerInvItemNum(MyIndex, i)
    If Val(Inums) > 0 Then Call AffSurfPic(DD_ItemSurf, Picture1(i), (item(Inums).Pic - (item(Inums).Pic \ 6) * 6) * PIC_X, (item(Inums).Pic \ 6) * PIC_Y) Else Call Picture1(i).Cls
Next i

For i = 1 To 30
    Inums = CoffreTmp(i).Numeros
    If Val(Inums) > 0 Then Call AffSurfPic(DD_ItemSurf, Picture2(i), (item(Inums).Pic - (item(Inums).Pic \ 6) * 6) * PIC_X, (item(Inums).Pic \ 6) * PIC_Y) Else Call Picture2(i).Cls
Next i

End Sub

