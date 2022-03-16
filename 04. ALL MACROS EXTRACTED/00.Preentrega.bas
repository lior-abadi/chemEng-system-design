"Public Sub Dim_De_Var()

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = Abs(ActiveSheet.Range(""R15""))


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Xoutlet tal que tres celdas luego de FM(i+3)-FM<Err, cuiadado con el DeltaW
For i = 28 To Lrow
j = i + 5
    RowM1 = ""I"" & i
    RowM2 = ""I"" & j
    error = Abs(ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value)
    If error < allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
Else
Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

ActiveSheet.Range(""O16"").Select

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()
Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty


End Sub

Private Sub CommandButton3_Click()
Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub

Private Sub CommandButton4_Click()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowObj As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
For i = 28 To Lrow
    RowM1 = ""I"" & i
    RowObj = ActiveSheet.Range(""S8"")
    error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
Else
Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

ActiveSheet.Range(""O16"").Select


With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton5_Click()
Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton6_Click()
Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String
Dim T0, T1, StepT, iT, contador As Double

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1
SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If


'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowObj As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
For i = 28 To Lrow
    RowM1 = ""I"" & i
    RowObj = ActiveSheet.Range(""S8"")
    error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
Else
Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = VolR


contador = contador + 1
Next



ActiveSheet.Range(""O16"").Select


With Application
.Calculation = xlManual
End With

End Sub" "Sub Macro2()
'
' Macro2 Macro
'

'
    Range(""V19"").Select
    Selection.Copy
    Range(""U13"").Select
    ActiveSheet.Paste
End Sub" "Sub Macro1()
'
' Macro1 Macro
'

'
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range(""S10"").GoalSeek Goal:=0.1, ChangingCell:=Range(""H2"")
    Range(""S10"").Select
End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As String

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""I"" & j
    RowM1 = ""I"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value

     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Exportación de Xeq a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout

contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = VolR
ActiveSheet.Range(CeldaResTf).Offset(0, 1) = ActiveSheet.Range(""AF23"").Value


contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Sout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As Double

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""K"" & j
    RowM1 = ""K"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout


contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O21"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el FFormaldehydo deseado
   RowObj = ActiveSheet.Range(""S8"")
     
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
     MsgBox error
    Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 3 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub




Private Sub CommandButton5_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
    RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
        Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim Selecc As Double
Dim CeldaResTf As String
Dim Numceltf As Double

Selecc = ActiveSheet.Range(""U24"").Value
Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf

Select Case Selecc
Case Is = 1
ActiveSheet.Range(CeldaResTf) = Wreq
Case Is = 2
ActiveSheet.Range(CeldaResTf) = Xout
Case Is = 3
ActiveSheet.Range(CeldaResTf) = Sout
Case Is = 4
ActiveSheet.Range(CeldaResTf) = VolR
End Select



contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With
End Sub

Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

If i < 30 Then
i = i + 58
Else
End If

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents

'Exportación de Wreq a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Wreq
ActiveSheet.Range(CeldaResTf).Offset(0, 1) = ActiveSheet.Range(""AF23"").Value


contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton8_Click()
With Application
.Calculation = xlAutomatic
 End With
 
With Application
.Calculation = xlManual
 End With
 
End Sub" "Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
 End With
 
With Application
.Calculation = xlManual
 End With
 
End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As String

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""I"" & j
    RowM1 = ""I"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value

     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout

contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = VolR

contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Sout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As Double

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""K"" & j
    RowM1 = ""K"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout


contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O21"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el FFormaldehydo deseado
   RowObj = ActiveSheet.Range(""S8"")
     
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
     MsgBox error
    Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 3 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub




Private Sub CommandButton5_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
    RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
        Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim Selecc As Double
Dim CeldaResTf As String
Dim Numceltf As Double

Selecc = ActiveSheet.Range(""U24"").Value
Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf

Select Case Selecc
Case Is = 1
ActiveSheet.Range(CeldaResTf) = Wreq
Case Is = 2
ActiveSheet.Range(CeldaResTf) = Xout
Case Is = 3
ActiveSheet.Range(CeldaResTf) = Sout
Case Is = 4
ActiveSheet.Range(CeldaResTf) = VolR
End Select



contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With
End Sub

Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

If i < 30 Then
i = i + 58
Else
End If

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents

'Exportación de Wreq a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Wreq
ActiveSheet.Range(CeldaResTf).Offset(0, 1) = ActiveSheet.Range(""AM23"").Value


contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton8_Click()
With Application
.Calculation = xlAutomatic
 End With
 
With Application
.Calculation = xlManual
 End With
 
End Sub"
"Public Sub Dim_De_Var()

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = Abs(ActiveSheet.Range(""R15""))


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Xoutlet tal que tres celdas luego de FM(i+3)-FM<Err, cuiadado con el DeltaW
For i = 28 To Lrow
j = i + 5
    RowM1 = ""I"" & i
    RowM2 = ""I"" & j
    error = Abs(ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value)
    If error < allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
Else
Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

ActiveSheet.Range(""O16"").Select

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()
Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty


End Sub

Private Sub CommandButton3_Click()
Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub

Private Sub CommandButton4_Click()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowObj As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
For i = 28 To Lrow
    RowM1 = ""I"" & i
    RowObj = ActiveSheet.Range(""S8"")
    error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
Else
Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

ActiveSheet.Range(""O16"").Select


With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton5_Click()
Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton6_Click()
Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String
Dim T0, T1, StepT, iT, contador As Double

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1
SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If


'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowObj As String
Dim allowE, step1 As Double

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
For i = 28 To Lrow
    RowM1 = ""I"" & i
    RowObj = ActiveSheet.Range(""S8"")
    error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
Else
Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = VolR


contador = contador + 1
Next



ActiveSheet.Range(""O16"").Select


With Application
.Calculation = xlManual
End With

End Sub" "Sub Macro2()
'
' Macro2 Macro
'

'
    Range(""V19"").Select
    Selection.Copy
    Range(""U13"").Select
    ActiveSheet.Paste
End Sub" "Sub Macro1()
'
' Macro1 Macro
'

'
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range(""S10"").GoalSeek Goal:=0.1, ChangingCell:=Range(""H2"")
    Range(""S10"").Select
End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As String

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""I"" & j
    RowM1 = ""I"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value

     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Exportación de Xeq a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout

contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = VolR
ActiveSheet.Range(CeldaResTf).Offset(0, 1) = ActiveSheet.Range(""AF23"").Value


contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Sout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As Double

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""K"" & j
    RowM1 = ""K"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout


contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O21"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el FFormaldehydo deseado
   RowObj = ActiveSheet.Range(""S8"")
     
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
     MsgBox error
    Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 3 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub




Private Sub CommandButton5_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
    RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
        Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim Selecc As Double
Dim CeldaResTf As String
Dim Numceltf As Double

Selecc = ActiveSheet.Range(""U24"").Value
Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf

Select Case Selecc
Case Is = 1
ActiveSheet.Range(CeldaResTf) = Wreq
Case Is = 2
ActiveSheet.Range(CeldaResTf) = Xout
Case Is = 3
ActiveSheet.Range(CeldaResTf) = Sout
Case Is = 4
ActiveSheet.Range(CeldaResTf) = VolR
End Select



contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With
End Sub

Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

If i < 30 Then
i = i + 58
Else
End If

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents

'Exportación de Wreq a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Wreq
ActiveSheet.Range(CeldaResTf).Offset(0, 1) = ActiveSheet.Range(""AF23"").Value


contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton8_Click()
With Application
.Calculation = xlAutomatic
 End With
 
With Application
.Calculation = xlManual
 End With
 
End Sub" "Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
 End With
 
With Application
.Calculation = xlManual
 End With
 
End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As String

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""I"" & j
    RowM1 = ""I"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value

     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout

contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""I"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -6).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O20"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O19"") = VolR

'Eliminar datos extra para calcular ok IDQ

'Exportación de VolR a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = VolR

contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O20"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub" "Public Sub Dim_De_Var()

Dim Lrow, aux1 As Double
Dim Cell0, CellN, Lcell, CellErase0 As String
Dim SR1, SRAfill, SRErase As String

Dim Xout, Sout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Dim RowM1, RowM2 As String
Dim allowE, step1 As Double

Dim RowObj As Double

Dim T0, T1, StepT, iT, contador As Double

Dim Nfinal As Double
Dim CellF0, CellF1, RangoFinal As String

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   
For i = 28 To Lrow Step 1
j = 1 + i
    RowM2 = ""K"" & j
    RowM1 = ""K"" & i
     error = ActiveSheet.Range(RowM2).Value - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Metano a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Xout


contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton2_Click()


'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & 1048576
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell

ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

Lrow = Empty
aux1 = Empty
Cell0 = Empty
CellN = Empty
Lcell = Empty
CellErase0 = Empty
SR1 = Empty
SRAfill = Empty
SRErase = Empty
Xout = Empty
Fmout = Empty
Wreq = Empty
Wreqaux = Empty
VolR = Empty
i = Empty
n = Empty
j = Empty
i1 = Empty
n2 = Empty
RowM1 = Empty
RowM2 = Empty
allowE = Empty
step1 = Empty
RowObj = Empty
T0 = Empty
T1 = Empty
StepT = Empty
iT = Empty
contador = Empty
Nfinal = Empty
CellF0 = Empty
CellF1 = Empty
RangoFinal = Empty




End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O21"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el FFormaldehydo deseado
   RowObj = ActiveSheet.Range(""S8"")
     
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
     MsgBox error
    Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 3 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents


'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub




Private Sub CommandButton5_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
    RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
      If error <= allowE Then
        Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

'Exportación de datos a tabla de resultados
Dim Selecc As Double
Dim CeldaResTf As String
Dim Numceltf As Double

Selecc = ActiveSheet.Range(""U24"").Value
Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf

Select Case Selecc
Case Is = 1
ActiveSheet.Range(CeldaResTf) = Wreq
Case Is = 2
ActiveSheet.Range(CeldaResTf) = Xout
Case Is = 3
ActiveSheet.Range(CeldaResTf) = Sout
Case Is = 4
ActiveSheet.Range(CeldaResTf) = VolR
End Select



contador = contador + 1
Next

'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With
End Sub

Private Sub CommandButton6_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents
End If

'Asignación de Rango de T
T0 = ActiveSheet.Range(""T20"").Value
T1 = ActiveSheet.Range(""T21"").Value
StepT = ActiveSheet.Range(""T22"").Value
contador = 0

'Comienzo de simulación en barrido

For iT = T0 To T1 Step StepT

ActiveSheet.Range(""O3"").Value = iT

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell

allowE = ActiveSheet.Range(""R15"")

Dim error As Double

'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

'Fila tal que se encuentre el Fmetanol deseado
   RowObj = ActiveSheet.Range(""S8"")
For i = 28 To Lrow
    RowM1 = ""K"" & i
     error = RowObj - ActiveSheet.Range(RowM1).Value
     If error <= allowE Then
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  Exit For
  Xout = ActiveSheet.Range(RowM1).Offset(0, -8).Value
  ActiveSheet.Range(""O16"") = Xout
  ActiveSheet.Range(""O21"") = i
  
Else
End If
Next i

'F Formaldehido a la salida
Fmout = ActiveSheet.Range(RowM1)
ActiveSheet.Range(""O17"") = Fmout

'W Cat
Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
ActiveSheet.Range(""O19"") = VolR

'S A la salida
Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
ActiveSheet.Range(""O20"") = Sout

If i < 30 Then
i = i + 58
Else
End If

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents

'Exportación de Wreq a tabla de resultados
Dim CeldaResTf As String
Dim Numceltf As Double

Numceltf = contador + ActiveSheet.Range(""U23"").Value
CeldaResTf = ActiveSheet.Range(""T23"") & Numceltf
ActiveSheet.Range(CeldaResTf) = Wreq
ActiveSheet.Range(CeldaResTf).Offset(0, 1) = ActiveSheet.Range(""AM23"").Value


contador = contador + 1
Next



'Finalización de sim
ActiveSheet.Range(""O16"").Select




With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton7_Click()

With Application
.Calculation = xlAutomatic
End With

If Not IsEmpty(ActiveSheet.Range(""A31"")) Then
'Limpieza de datos en caso de no darle clear antes
'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SRErase = CellErase0 & "":"" & Lcell
ActiveSheet.Range(SRErase).ClearContents
ActiveSheet.Range(""O16:O21"").ClearContents

End If

'Simulación del reactor. Comienzo.

'Asignacion de celdas de rango
aux1 = ActiveSheet.Range(""P7"").Value + 1
Cell0 = ActiveSheet.Range(""O7"") & ActiveSheet.Range(""P7"")
CellN = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P8"")
Lcell = ActiveSheet.Range(""O8"") & ActiveSheet.Range(""P9"")
Lrow = ActiveSheet.Range(""P9"")
CellErase0 = ActiveSheet.Range(""O7"") & aux1

SR1 = Cell0 & "":"" & CellN
SRAfill = Cell0 & "":"" & Lcell


'Simulación de reactor
ActiveSheet.Range(SR1).Select
Selection.AutoFill Destination:=Range(SRAfill), Type:=xlFillDefault

With Application
.Calculation = xlManual
End With

End Sub

Private Sub CommandButton8_Click()
With Application
.Calculation = xlAutomatic
 End With
 
With Application
.Calculation = xlManual
 End With
 
End Sub"
