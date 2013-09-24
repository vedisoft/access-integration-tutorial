Attribute VB_Name = "ProstieZvonki"
Option Compare Database
Option Explicit

Public prostie_zvonki_wrapper As ProstieZvonkiWrapper

Public Const Phone = "123" 'current user phone number

Public Function Init_Prostie_Zvonki()
    Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
    Call prostie_zvonki_wrapper.Initialize(Phone)
End Function

Public Function MakeCall(Phone As String)
    Call prostie_zvonki_wrapper.MakeCall(Phone)
End Function

