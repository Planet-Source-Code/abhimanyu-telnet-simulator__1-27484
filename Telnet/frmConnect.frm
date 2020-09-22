VERSION 5.00
Begin VB.Form frmConnect 
   Caption         =   "Connect"
   ClientHeight    =   2190
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   3480
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPort 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cboHostName 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmConnect.frx":0CCA
      Left            =   1560
      List            =   "frmConnect.frx":0CCC
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   600
   End
   Begin VB.Label lblHostName 
      AutoSize        =   -1  'True
      Caption         =   "Host Name:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cboHostName_GotFocus()
    Dim i As Integer
    i = cboHostName.ListCount
    If i > 0 Then
        i = Len(cboHostName.List(i - 1))
        cboHostName.SelStart = 0
        cboHostName.SelLength = i
    End If
End Sub

Private Sub cboHostName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cboPort.SetFocus
    End If
End Sub

Private Sub cboPort_GotFocus()
    Dim i As Integer
    i = cboPort.ListCount
    If i > 0 Then
        i = Len(cboPort.List(i - 1))
        cboPort.SelStart = 0
        cboPort.SelLength = i
    End If
End Sub

Private Sub cboPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdConnect_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Visible = False
    frmTelnet.Enabled = True
    frmTelnet.SetFocus
End Sub

Private Sub cmdConnect_Click()
    Dim host As String
    Dim port As Integer
    Dim i As Integer
On Error GoTo err
    
    frmTelnet.Socket.Close
    If Len(cboHostName.Text) < 1 Then
        MsgBox "No host name specified.", , "Connect"
        cboHostName.SetFocus
        Exit Sub
    End If
    
    host = Trim(cboHostName.Text)
    If Len(cboPort.Text) < 1 Then
        port = 80
    Else
        port = Val(cboPort.Text)
    End If
    frmTelnet.Socket.connect host, port
    frmTelnet.Caption = "Telnet 1.0 - " & host
    ChangeMenu (False)
    Me.Visible = False
    frmTelnet.MousePointer = 11
    Do While frmTelnet.Socket.State = sckConnecting
        DoEvents
    Loop
    frmTelnet.MousePointer = 0
    frmTelnet.Enabled = True
    frmTelnet.txtContent.Locked = False
    Call addItem(CStr(host), CStr(port))
    frmTelnet.SetFocus
    Exit Sub
err:
    MsgBox err.Description, vbOKOnly + vbCritical, "Connection failed"
    ChangeMenu (True)
    Initialize
    Exit Sub
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    cboHostName.SetFocus
    i = cboHostName.ListCount
    If (i > 0) Then
        cboHostName.Text = cboHostName.List(i - 1)
    End If
    i = cboPort.ListCount
    If (i > 0) Then
        cboPort.Text = cboPort.List(i - 1)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTelnet.Enabled = True
End Sub




Public Sub addItem(item1 As String, item2 As String)
    Dim i As Integer
    Dim j As Integer
    Dim found As Integer
    found = 0
    i = cboHostName.ListCount
    For j = 0 To i
        If (cboHostName.List(j) = item1) Then
            found = 1
        End If
    Next
    If (found = 0) Then
        cboHostName.addItem (item1)
    End If
    
    found = 0
    i = cboPort.ListCount
    For j = 0 To i
        If (cboPort.List(j) = item2) Then
            found = 1
        End If
    Next
    If (found = 0) Then
        cboPort.addItem (item2)
    End If
End Sub
