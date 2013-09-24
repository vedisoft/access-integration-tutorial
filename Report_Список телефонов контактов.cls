VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Список телефонов контактов"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Business_Phone_Click()
    If Not Me.Business_Phone.Value = "" Then
        Call ProstieZvonki.MakeCall(Me.Business_Phone.Value)
    End If
End Sub
