Attribute VB_Name = "ProstieZvonki"
Option Compare Database
Option Explicit

Public prostie_zvonki_wrapper As ProstieZvonkiWrapper

Public Function Init_Prostie_Zvonki(ManagerPhone As String)
    Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
    Call prostie_zvonki_wrapper.Initialize(ManagerPhone)
End Function

Public Function MakeCall(Phone As String)
    Call prostie_zvonki_wrapper.MakeCall(Phone)
End Function

