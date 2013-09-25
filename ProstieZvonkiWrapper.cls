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

Const guidKey = "HKEY_CURRENT_USER\Software\Vedisoft\Access\GUID"

Dim clientID As String

Public Function Initialize(client_id As String)
    Dim guid As String
    Dim logPath As String
    
    clientID = client_id
    
    'add unique client guid to registry
    If (RegKeyExists(guidKey)) Then
        guid = RegKeyRead(guidKey)
    Else
        guid = GenGuid()
        Call RegKeySave(guidKey, guid)
    End If
    
    'create log directory
    logPath = Environ("LocalAppData") & "\Ведисофт\Access\"
    If Dir(logPath) = "" Then
        MkDir logPath
    End If
    Set prostie_zvonki_lib = New CTIControlX
    
    On Error GoTo Errhandler
    Call prostie_zvonki_lib.Connect("127.0.0.1:10150", clientID, "Access", guid, _
                                            logPath & "ProtocolLib_log.log", 2, 5000)
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
    Call prostie_zvonki_lib.Call(clientID, Phone)
End Sub

Function GenGuid() As String
    Dim TypeLib As Object
    Dim guid As String
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    guid = TypeLib.guid
    ' format is {24DD18D4-C902-497F-A64B-28B2FA741661}
    guid = Replace(guid, "{", "")
    guid = Replace(guid, "}", "")
    GenGuid = guid
End Function


Function RegKeyRead(i_RegKey As String) As String
    Dim myWS As Object
    On Error Resume Next
    Set myWS = CreateObject("WScript.Shell")
    RegKeyRead = myWS.RegRead(i_RegKey)
End Function

Function RegKeyExists(i_RegKey As String) As Boolean
    Dim myWS As Object
    On Error GoTo ErrorHandler
    Set myWS = CreateObject("WScript.Shell")
    myWS.RegRead i_RegKey
    RegKeyExists = True
    Exit Function
ErrorHandler:
    RegKeyExists = False
End Function

Sub RegKeySave(i_RegKey As String, _
               i_Value As String, _
      Optional i_Type As String = "REG_SZ")
    Dim myWS As Object
    Set myWS = CreateObject("WScript.Shell")
    myWS.RegWrite i_RegKey, i_Value, i_Type
End Sub

Private Sub prostie_zvonki_lib_OnTransferredCall(ByVal CallID As Long, _
                                                ByVal src As String, ByVal dst As String)
    If (dst <> clientID) Then
        End Sub
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT [Имя], [Фамилия] FROM Контакты" & _
            "WHERE Контакты.[Рабочий телефон] = '" & src & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        'show client name
        MsgBox ("Incoming call from " & rs.Fields(0).Value & " " & rs.Fields(1).Value)
    Else
        'can't find client, show number only
        MsgBox ("Incoming call from " & src)
    End If
    Call rs.Close
End Sub

Private Sub prostie_zvonki_lib_OnTransferRequest(ByVal CallID As Long, ByVal from As String)
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT Менеджеры.[Телефон] FROM Менеджеры" & _
            " INNER JOIN Контакты ON Менеджеры.[Код] = Контакты.[Менеджер]" & _
            " WHERE Контакты.[Рабочий телефон] = '" & from & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        'call with number, if manager can handle call
        Call prostie_zvonki_lib.Transfer(CallID, rs.Fields(0).Value)
    Else
        'call with empty string, if manager can't handle call
        Call prostie_zvonki_lib.Transfer(CallID, "")
    End If
    Call rs.Close
End Sub

