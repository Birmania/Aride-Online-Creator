VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmReport 
   Caption         =   "Oups..."
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDontSendReport 
      Caption         =   "Ne pas envoyer le rapport"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdSendReport 
      Caption         =   "Envoyer le rapport"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oups... Une erreur est survenue !"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wasShown As Boolean

Private Sub cmdDontSendReport_Click()
    Unload Me
End Sub

Private Sub cmdSendReport_Click()
    On Error GoTo HandleError
    Dim webReporter As String
    Dim updaterConfigurationFile As String
    
    updaterConfigurationFile = App.Path & "\Config\Updater.ini"
    
    If FileExist(ErrorLogFile) Then
        If Not FileExist(updaterConfigurationFile) Then
            webReporter = "ftp://aride-online.com"
        Else
            webReporter = ReadINI("UPDATER", "WebReporter", updaterConfigurationFile)
        End If
        'host = "ftp://" & Right(webReporter, Len(webReporter) - 7)
        
        'Inet.AccessType = icDirect
        Inet.Protocol = icFTP
        Inet.URL = webReporter
        Inet.Username = "Players"
        Inet.Password = ""
        'Inet.RemoteHost = host
        'Inet.RemotePort = 21
        'MsgBox "PUT """ & ErrorLogFile & """ """ & GetTickCount & ".txt"""
        Call Inet.Execute(, "PUT """ & ErrorLogFile & """ """ & GetTickCount & ".txt""")
        Do While Inet.StillExecuting
            DoEvents
            'Call Sleep(1000)
        Loop
        
        Call Inet.Execute(, "CLOSE")
    End If

    Unload Me

'Error handler
    Exit Sub
HandleError:
    MsgBox "Impossible d'envoyer le rapport..."

    Unload Me
End Sub

Private Sub Form_Load()
    wasShown = True
End Sub
