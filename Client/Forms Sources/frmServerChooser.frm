VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmServerChooser 
   BorderStyle     =   0  'None
   Caption         =   "Sélection du Serveur"
   ClientHeight    =   2610
   ClientLeft      =   2805
   ClientTop       =   1365
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmServerChooser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServerChooser.frx":000C
   ScaleHeight     =   2610
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRafraichir 
      Caption         =   "Rafraichir"
      Height          =   255
      Left            =   1860
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ListBox lstServers 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Quitter"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Connecter"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   60
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmServerChooser"
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

Public Path As String
Public Extension As String
Private Sub cmdCancel_Click()
    Call GameDestroy
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If lstServers.ListCount <= 0 Then Exit Sub

    If lstServers.ListIndex = -1 Then
        lstServers.ListIndex = 0
    End If

    GAME_IP = ReadINI("SERVER" & lstServers.ListIndex, "IP", App.Path & "\Config\Serveur.ini")
    GAME_PORT = Val(ReadINI("SERVER" & lstServers.ListIndex, "PORT", App.Path & "\Config\Serveur.ini"))
    frmMainMenu.Show
    Call frmMainMenu.txtName.SetFocus
    Unload Me
End Sub

Private Sub CmdRafraichir_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
Dim FileName As String
Dim i As Long
Dim Ending As String
Dim Buffer As clsBuffer
Dim Servers() As String
Dim Server As Variant

Call SetIcon(Me)

For i = 1 To 4
    If i = 1 Then Ending = ".gif"
    If i = 2 Then Ending = ".jpg"
    If i = 3 Then Ending = ".png"
    If i = 4 Then Ending = ".bmp"

    If FileExist(App.Path & Rep_Theme & "\Login\choix_serveur" & Ending) Then frmServerChooser.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\choix_serveur" & Ending)
Next i
    frmServerChooser.Visible = True

    FileName = App.Path & "\Config\Serveur.ini"
    i = 0
    'C = 0
    CHECK_WAIT = False
    lstServers.Clear
   
    cmdOk.Enabled = False
    CmdRafraichir.Enabled = False
    
    Me.MousePointer = 13

    CLIENT_PORT = Val(ReadINI("CONFIG", "PORT", ClientConfigurationFile))
    frmMirage.SocketTCP.LocalPort = CLIENT_PORT
    
    Servers = ReadINISections(FileName)
 
    For Each Server In Servers


        If Not CHECK_WAIT Then
            If ReadINI(CStr(Server), "IP", FileName) <> vbNullString And ReadINI(CStr(Server), "PORT", FileName) <> vbNullString Then
                GAME_IP = ReadINI(CStr(Server), "IP", FileName)
                GAME_PORT = Val(ReadINI(CStr(Server), "PORT", FileName))
                Set Buffer = New clsBuffer
                Buffer.WriteLong CFindServer

                lstServers.AddItem ReadINI(CStr(Server), "Name", FileName)
                If CheckServerStatus Then CHECK_WAIT = True: SendDataTCP Buffer.ToArray()
                Set Buffer = Nothing
                i = i + 1
            End If
        End If

        Do While CHECK_WAIT And IsConnected
            Sleep 10
            DoEvents
        Loop
        If Not IsConnected Then
            lstServers.List(lstServers.ListCount - 1) = lstServers.List(lstServers.ListCount - 1) & " - Fermé"
        End If
    Next
    
    frmMirage.SocketTCP.Close
    cmdOk.Enabled = True
    CmdRafraichir.Enabled = True
    Me.MousePointer = 0
    
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
    Call cmdCancel_Click
End Sub

Private Sub lstServers_DblClick()
    If cmdOk.Enabled = True Then Call cmdOk_Click
End Sub
