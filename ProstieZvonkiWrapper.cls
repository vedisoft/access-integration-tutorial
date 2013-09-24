VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProstieZvonkiWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents prostie_zvonki_lib As CTIControlX
Attribute prostie_zvonki_lib.VB_VarHelpID = -1

Public Function Initialize(Phone As String)
    Set prostie_zvonki_lib = New CTIControlX
    On Error GoTo Errhandler
    Call prostie_zvonki_lib.Connect("127.0.0.1:10150", Phone, "Excel", "a9548dc6-4e09-4faa-8cfa-8b5fbbb03087", Environ("LocalAppData") & "\Ведисофт\Excel\ProtocolLib_log.log", 2, 5000)
    Exit Function
Errhandler:
    MsgBox ("Can't connect to server")
End Function

Public Sub Class_Terminate()
    If Not IsNull(prostie_zvonki_lib) Then
        Call prostie_zvonki_lib.Disconnect
    End If
End Sub

Public Sub MakeCall(Phone As String)
    Call prostie_zvonki_lib.Call("123", Phone)
End Sub

Private Sub prostie_zvonki_lib_OnTransferredCall(ByVal CallID As Long, ByVal src As String, ByVal dst As String)
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT [Имя], [Фамилия] FROM Контакты WHERE Контакты.[Рабочий телефон] = '" & src & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        MsgBox ("Incoming call from " & rs.Fields(0).Value & " " & rs.Fields(1).Value) 'show client name
    Else
        MsgBox ("Incoming call from " & src) 'can't find client, show number only
    End If
    Call rs.Close
End Sub


Private Sub prostie_zvonki_lib_OnTransferRequest(ByVal CallID As Long, ByVal from As String)
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT Менеджеры.[Телефон] FROM Менеджеры INNER JOIN Контакты ON Менеджеры.[Код] = Контакты.[Менеджер] WHERE Контакты.[Рабочий телефон] = '" & from & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        Call prostie_zvonki_lib.Transfer(CallID, rs.Fields(0).Value) 'call with number, if manager can handle call
    Else
        Call prostie_zvonki_lib.Transfer(CallID, "") 'call with empty string, if manager can't handle call
    End If
    Call rs.Close
End Sub
