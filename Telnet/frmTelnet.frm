VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTelnet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet 1.0 - Not connected"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmTelnet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   3240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
      Begin VB.Menu mnuRemoteSystem 
         Caption         =   "&Remote System ..."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConnectSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit          Alt+F4"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy &All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'global  prevText As String


Private Sub Form_Load()
    Initialize
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Socket.Close
    End
End Sub

Private Sub mnuAbout_Click()
    Me.Enabled = False
    frmAbout.Show
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText (txtContent.SelText)
End Sub

Private Sub mnuCopyAll_Click()
    Clipboard.Clear
    Clipboard.SetText txtContent.Text
End Sub

Private Sub mnuDisconnect_Click()
    Initialize
End Sub

Private Sub mnuExit_Click()
    Socket.Close
    End
End Sub



Private Sub mnuRemoteSystem_Click()
    Me.Enabled = False
    frmConnect.Show
End Sub


Private Sub mnuSelectAll_Click()
    txtContent.SelStart = 0
    txtContent.SelLength = Len(txtContent.Text)
End Sub

Private Sub Socket_Close()
    MsgBox "Connection to host lost"
    Socket.Close
    Socket.RemoteHost = 0
    Socket.RemotePort = 0
    Initialize
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim newStr As String
    Socket.GetData newStr, vbString
    'MsgBox Len(newStr)
    txtContent.Text = txtContent.Text + newStr
    prevText = txtContent.Text
    txtContent.SelLength = Len(txtContent.Text)
    txtContent.SelText = txtContent.Text
    
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Connection failed.", vbOKOnly + vbCritical, "Telnet"
    frmTelnet.Caption = "Telnet 1.0 - Not connected"
    ChangeMenu (True)
End Sub



Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyTab
            KeyCode = 0
        Case vbKeyLeft
            KeyCode = 0
        Case vbKeyUp
            KeyCode = 0
        Case vbKeyPageUp
            KeyCode = 0
        Case vbKeyDelete
            KeyCode = 0
    End Select
    
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    Dim newStr As String
    
    If mnuRemoteSystem.Enabled = True Then
        Exit Sub
    End If
    If Len(txtContent.SelText) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        newStr = Right(txtContent.Text, Len(txtContent.Text) - Len(prevText))
        prevText = txtContent.Text
        SendMessage (newStr)
    End If
    
    
End Sub

Private Sub txtContent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        MsgBox "Right click disabled"
    End If
End Sub

Public Sub SendMessage(str As String)
    If Socket.State = sckConnected Then
        Socket.SendData str & vbCrLf
    Else
        MsgBox "dis"
    End If
    
End Sub
