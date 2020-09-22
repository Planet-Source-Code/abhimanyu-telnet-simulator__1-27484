Attribute VB_Name = "Module1"
Option Explicit
Public prevText As String



Public Sub ChangeMenu(connect As Boolean)
    frmTelnet.mnuRemoteSystem.Enabled = connect
    frmTelnet.mnuDisconnect.Enabled = Not connect
End Sub

Public Sub Initialize()
    frmTelnet.txtContent.Locked = True
    frmTelnet.txtContent.Text = ""
    prevText = ""
    frmTelnet.Caption = "Telnet 1.0 - Not conneted"
    frmTelnet.Socket.Close
    ChangeMenu (True)
End Sub
