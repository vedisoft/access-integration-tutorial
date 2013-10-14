VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Выбор текущего Менеджера"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    If (prostie_zvonki_wrapper Is Nothing) Then
        Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
    End If
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT [Телефон] FROM Менеджеры" & _
            " WHERE [Телефон] = '" & prostie_zvonki_wrapper.GetManagerPhone & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        Me.Manager.Value = rs.Fields(0).Value
    End If
    Call rs.Close
End Sub
