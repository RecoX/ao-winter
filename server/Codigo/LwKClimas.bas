Attribute VB_Name = "LwKClimas"
'********************************Modulo Climas*********************************
'Author: Manuel (Lorwik)
'Last Modification: 17/11/2011
'Controla el clima y lo envia al cliente.
'******************************************************************************

Option Explicit

'******************************************************************************
'Sorteo Del Clima
'******************************************************************************
Public Function SortearClima()
    If Hour(Now) >= 6 And Hour(Now) < 12 Then
        Call Ma�ana
        frmMain.Clima.Caption = "Clima: Ma�ana"
    ElseIf Hour(Now) >= 12 And Hour(Now) < 18 Then
        Call Dia
        frmMain.Clima.Caption = "Clima: MedioDia"
    ElseIf Hour(Now) >= 18 And Hour(Now) < 20 Then
        Call Tarde
        frmMain.Clima.Caption = "Clima: Tarde"
    ElseIf Hour(Now) >= 20 And Hour(Now) < 6 Then
        Call Noche
        frmMain.Clima.Caption = "Clima: Noche"
    End If
End Function

'******************************************************************************
'Enviamos la Ma�ana
'******************************************************************************
Public Function Ma�ana()
Dim UserIndex As Integer
Dim i As Long
Anocheceria = 0
For i = 1 To LastUser
    Call writeNoche(i, 0)
Next i
End Function

'******************************************************************************
'Enviamos el Dia
'******************************************************************************
Public Function Dia()
Dim UserIndex As Integer
Dim i As Long
Anocheceria = 1
For i = 1 To LastUser
    Call writeNoche(i, 1)
Next i
End Function

'******************************************************************************
'Enviamos la Tarde
'******************************************************************************
Public Function Tarde()
Dim UserIndex As Integer
Dim i As Long
Anocheceria = 2
For i = 1 To LastUser
    Call writeNoche(i, 2)
Next i
End Function

'******************************************************************************
'Enviamos la Noche
'******************************************************************************
Public Function Noche()
Dim UserIndex As Integer
Dim i As Long
Anocheceria = 3
For i = 1 To LastUser
    Call writeNoche(i, 3)
Next i
End Function
