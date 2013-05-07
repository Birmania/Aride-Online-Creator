VERSION 5.00
Begin VB.Form frmFlash 
   BorderStyle     =   0  'None
   Caption         =   "Évènement de Flash"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Check 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      TabIndex        =   0
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check_Timer()
'    If Flash.CurrentFrame > 0 Then
'        If Flash.CurrentFrame >= Flash.TotalFrames - 1 Then
'            Flash.FrameNum = 0
'            Flash.Stop
'            Check.Enabled = False
'            WriteINI "CONFIG", "Music", frmMirage.chkmusic.Value, ClientConfigurationFile
'            'Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
'            Call PlayMidi(App.Path & "\" & Trim$(Map.Music))
'            WriteINI "CONFIG", "Sound", frmMirage.chksound.Value, ClientConfigurationFile
'            Unload Me
'        End If
'    End If
End Sub

Private Sub Form_Load()

Dim i As Long
Dim Ending As String



For i = 1 To 3
    If i = 1 Then Ending = ".gif"
    If i = 2 Then Ending = ".jpg"
    If i = 3 Then Ending = ".png"
 
    If FileExist(App.Path & Rep_Theme & "\Jeu\flash" & Ending) Then Me.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\flash" & Ending)
Next i

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dr = True
drx = X
dry = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.Left + (X - drx), Me.Top + (Y - dry))
If Me.Left > Screen.Width Or Me.Top > Screen.Height Then Me.Top = Screen.Height \ 2: Me.Left = Screen.Width \ 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub Label1_Click()
'    Flash.FrameNum = 0
'    Flash.Stop
'    Check.Enabled = False
'    WriteINI "CONFIG", "Music", frmMirage.chkmusic.Value, ClientConfigurationFile
'    'Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
'    Call PlayMidi(App.Path & "\" & Trim$(Map.Music))
'    WriteINI "CONFIG", "Sound", frmMirage.chksound.Value, ClientConfigurationFile
'    Unload Me
End Sub

