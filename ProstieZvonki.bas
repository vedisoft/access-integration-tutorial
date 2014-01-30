Attribute VB_Name = "ProstieZvonki"
Option Compare Database
Option Explicit

Public prostie_zvonki_wrapper As ProstieZvonkiWrapper

Public Sub CreateWrapper()
    Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
End Sub

Public Function Init_Prostie_Zvonki(ManagerPhone As String)
    If (prostie_zvonki_wrapper Is Nothing) Then
        CreateWrapper
    End If
    prostie_zvonki_wrapper.SetManagerPhone (ManagerPhone)
End Function

Public Function MakeCall(phone As String)
    Call prostie_zvonki_wrapper.MakeCall(phone)
End Function


