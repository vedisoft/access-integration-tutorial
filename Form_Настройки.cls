VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Настройки"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub connectButton_Click()
    Dim server As String
    Dim password As String
    Me.server.SetFocus
    server = Me.server.Text
    Me.password.SetFocus
    password = Me.password.Text
    
    Call prostie_zvonki_wrapper.SetConnectionSettings(server, password)
    
    Me.connectButton.SetFocus
    prostie_zvonki_wrapper.Connect
End Sub

Private Sub disconnectButton_Click()
    prostie_zvonki_wrapper.Disconnect
End Sub

Private Sub Form_Load()
    If (prostie_zvonki_wrapper Is Nothing) Then
        CreateWrapper
    End If
    State = prostie_zvonki_wrapper.GetConnectionState
    UpdateState (State)
End Sub

Sub UpdateState(State As Boolean)
    Me.General_Page.SetFocus
    Me.connectButton.Enabled = Not State
    Me.disconnectButton.Enabled = State
    If (State) Then
        Me.statusLabel.Caption = "подключен"
    Else
        Me.statusLabel.Caption = "отключен"
    End If
    Me.server.SetFocus
    Me.server.Text = prostie_zvonki_wrapper.GetServer
    Me.password.SetFocus
    Me.password.Text = prostie_zvonki_wrapper.GetPassword
    Me.General_Page.SetFocus
End Sub
