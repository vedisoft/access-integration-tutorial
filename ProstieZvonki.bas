Attribute VB_Name = "ProstieZvonki"
Option Compare Database
Option Explicit

Public prostie_zvonki_wrapper As ProstieZvonkiWrapper

Public Function Init_Prostie_Zvonki(managerPhone As String)
    Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
    Call prostie_zvonki_wrapper.Initialize(managerPhone)
End Function

Public Function MakeCall(Phone As String)
    Call prostie_zvonki_wrapper.MakeCall(Phone)
End Function

