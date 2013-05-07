VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Principal"
   ClientHeight    =   5100
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   5085
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":000C
   ScaleHeight     =   5100
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   2160
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
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
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton picConnect 
      Caption         =   "Connexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
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
      Height          =   225
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   2355
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2355
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   195
   End
   Begin VB.CommandButton picCredits 
      Caption         =   "Cr�dits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton picQuit 
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4440
      Top             =   3480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editer Hors Ligne"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   12
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de compte:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enregistrer Infos"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblss 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�tat du serveur :"
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
      Left            =   720
      TabIndex        =   9
      Top             =   4350
      Width           =   1170
   End
   Begin VB.Label version 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version de l'�diteur : 0.6"
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
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   1875
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "hgjhgj"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1920
      TabIndex        =   0
      Top             =   4320
      Width           =   2565
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

txtName.Text = Trim$(ReadINI("INFO", "Account", App.Path & "\Config\Account.ini"))
txtPassword.Text = Trim$(ReadINI("INFO", "Password", App.Path & "\Config\Account.ini"))
If Trim$(txtPassword.Text) <> vbNullString Then Check1.value = Checked Else Check1.value = Unchecked
txtName.SelStart = Len(txtName.Text)
status.ForeColor = vbRed
status.Caption = "Recherche en cours..."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
GettingMap = False
frmmsg.Show
frmMirage.Timer2.Enabled = True
End Sub

Private Sub picConnect_Click()
If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
    If Len(Trim$(txtName.Text)) < 3 Or Len(Trim$(txtPassword.Text)) < 3 Then MsgBox "Votre nom et votre mot de pass doit faire au minimum 3 caract�re de long": Exit Sub
    Call MenuState(MENU_STATE_LOGIN)
    Call WriteINI("INFO", "Account", txtName.Text, (App.Path & "\Config\Account.ini"))
    AccOpt.InfName = txtName.Text
    If Check1.value = Checked Then Call WriteINI("INFO", "Password", txtPassword.Text, (App.Path & "\Config\Account.ini")): AccOpt.InfPass = txtPassword.Text Else Call WriteINI("INFO", "Password", vbNullString, (App.Path & "\Config\Account.ini")): AccOpt.InfPass = vbNullString
End If
End Sub

Private Sub picCredits_Click()
    frmpet.Visible = True
    Me.Visible = False
End Sub

Private Sub picQuit_Click()
    frmmsg.Show
    Call GameDestroy
End Sub

Private Sub Timer1_Timer()
    If ConnectToServer Then status.ForeColor = vbGreen: status.Caption = "Connect�" Else status.ForeColor = vbRed: status.Caption = "D�connect�"
End Sub

Private Sub Timer2_Timer()
    Call GameDestroy
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picConnect_Click
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picConnect_Click
End Sub
