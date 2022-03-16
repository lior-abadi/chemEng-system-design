"Public Sub ConfigDoc()
With Application
.Calculation = xlManual
End Sub
" "'Dims de Balance de Masa 1 reactor
Public AlimFresca(1 To 8) As Double
Public EntradaReact(1 To 8) As Double
Public SalidaReact(1 To 8) As Double
Public Producto(1 To 8) As Double
Public Purga(1 To 8) As Double
Public Reciclo(1 To 8) As Double
Public FracPurga As Double
Public iA, jA As Integer
Public BalanceGlobal, ErrorBalance As Double
Public CeldaFresca, CeldaER, CeldaSalR, CeldaProd, CeldaPurg, CeldaRec As String

'Dims de Reactor 1
Public Lrow, aux1 As Double
Public Cell0, CellN, Lcell, CellErase0 As String
Public SR1, SRAfill, SRErase As String

Public Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Public RowM1, RowM2 As String
Public allowE, step1 As Double

Public RowObj As String

Public T0, T1, StepT, iT, contador As Double

Public Nfinal As Double
Public CellF0, CellF1, RangoFinal As String

'Dims de Reactor 2
Public Sout As Double
Public AlimFresca2(1 To 12) As Double
Public EntradaReact2(1 To 12) As Double
Public Reciclo2(1 To 12) As Double

" "Private Sub CommandButton1_Click()

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

Call ClearAll1

End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O21"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()
Call Reactor2
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
 
End Sub" "Public Sub HojaConfig()

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()
End Sub
Private Sub CommandButton2_Click()

With Application
.Calculation = xlAutomatic
 End With

'Alim Fresca asignación valores
For iA = 1 To 8
jA = iA + 30
CeldaFresca = ""D"" & jA
AlimFresca(iA) = ActiveSheet.Range(CeldaFresca).Value
Next iA

'Reciclo asignación valores
For iA = 1 To 8
jA = iA + 12
CeldaRec = ""R"" & jA
Reciclo(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA


'Primer paso, suponer entrada reactor = entrada planta + reciclo anterior de seed
For iA = 1 To 8
jA = iA + 30
CeldaER = ""I"" & jA
EntradaReact(iA) = AlimFresca(iA) + Reciclo(iA)
ActiveSheet.Range(CeldaER) = EntradaReact(iA)
Next iA


'Ir a hoja de Reactor 1
Worksheets(""1R"").Activate

'Simular Reactor
Call ReactorAd1BM

'Volver a hoja BM
Worksheets(""CN + BM R1"").Activate

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N40"").Value
ErrorBalance = ActiveSheet.Range(""N41"").Value

'Loop para que haga el balance hasta que BAL de Masa Global
'sea menor al error admitido (Bal debe ser 0)

Do While BalanceGlobal > ErrorBalance

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N40"").Value

'Continuar con el balance de masa
'Reciclo asignación valores
For iA = 1 To 8
jA = iA + 12
CeldaRec = ""R"" & jA
Reciclo(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA

For iA = 1 To 8
EntradaReact(iA) = Reciclo(iA) + AlimFresca(iA)
Next iA

For iA = 1 To 8
jA = iA + 30
CeldaER = ""I"" & jA
ActiveSheet.Range(CeldaER) = EntradaReact(iA)
Next iA

'Ir a hoja de Reactor 1
Worksheets(""1R"").Activate

'Simular Reactor
Call ReactorAd1BM

'Volver a hoja BM
Worksheets(""CN + BM R1"").Activate
Debug.Print BalanceGlobal
Loop

With Application
.Calculation = xlManual
 End With
 
End Sub

" "Public Sub SetManual()

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

Call ClearAll1
End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()
Call ReactorAd1
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

End Sub" "Public Sub ReactorAd1()

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

'Parámetros
Fmout = ActiveSheet.Range(RowM1)
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value

'Celdas
ActiveSheet.Range(""O17"") = Fmout
ActiveSheet.Range(""O18"").Value = Wreq
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

End Sub" "Public Sub ClearAll1()

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




End Sub" "Public Sub ReactorAd1BM()

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

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

If i > 30 Then
Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents
End If

With Application
.Calculation = xlManual
End With

End Sub
" "Public Sub Reactor2()

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
'Fmout = ActiveSheet.Range(RowM1)
'ActiveSheet.Range(""O17"") = Fmout

'W Cat
'Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
'ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
'VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
'ActiveSheet.Range(""O19"") = VolR

'S A la salida
'Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
'ActiveSheet.Range(""O20"") = Sout

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
" "Public Sub HojaConfig()

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()
With Application
.Calculation = xlAutomatic
End With

Dim RestaObj, CeldaNBP As String

RestaObj = ActiveSheet.Range(""K3"")
CeldaNBP = ActiveSheet.Range(""K4"")

Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range(RestaObj).GoalSeek Goal:=0, ChangingCell:=Range(CeldaNBP)
    
With Application
.Calculation = xlManual
End With
    
End Sub
Private Sub CommandButton2_Click()

With Application
.Calculation = xlAutomatic
 End With

'Alim Fresca asignación valores
For iA = 1 To 12
jA = iA + 30
CeldaFresca = ""D"" & jA
AlimFresca2(iA) = ActiveSheet.Range(CeldaFresca).Value
Next iA

'Reciclo asignación valores
For iA = 1 To 12
jA = iA + 6
CeldaRec = ""R"" & jA
Reciclo2(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA

'Primer paso, suponer entrada reactor = entrada planta + reciclo anterior de seed
For iA = 1 To 12
jA = iA + 30
CeldaER = ""I"" & jA
EntradaReact2(iA) = AlimFresca2(iA) + Reciclo2(iA)
ActiveSheet.Range(CeldaER) = EntradaReact2(iA)
Next iA

'Ir a hoja de Reactor 2
Worksheets(""2R"").Activate

'Simular Reactor
Call Reactor2

'Volver a hoja BM
Worksheets(""CN + BM R2"").Activate

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N46"").Value
ErrorBalance = ActiveSheet.Range(""N47"").Value

'Loop para que haga el balance hasta que BAL de Masa Global
'sea menor al error admitido (Bal debe ser 0)

Do While BalanceGlobal > ErrorBalance

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N46"").Value

'Continuar con el balance de masa
'Reciclo asignación valores
For iA = 1 To 12
jA = iA + 6
CeldaRec = ""R"" & jA
Reciclo2(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA

For iA = 1 To 12
EntradaReact2(iA) = Reciclo2(iA) + AlimFresca2(iA)
Next iA

For iA = 1 To 12
jA = iA + 30
CeldaER = ""I"" & jA
ActiveSheet.Range(CeldaER) = EntradaReact2(iA)
Next iA

'Ir a hoja de Reactor 2
Worksheets(""2R"").Activate

'Simular Reactor
Call Reactor2

'Volver a hoja BM
Worksheets(""CN + BM R2"").Activate
Loop

With Application
.Calculation = xlManual
 End With
 
End Sub

" "Private Sub CommandButton1_Click()
With Application
.Calculation = xlAutomatic
End With

Dim RestaObj, CeldaNBP As String

RestaObj = ActiveSheet.Range(""K3"")
CeldaNBP = ActiveSheet.Range(""K4"")

Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range(RestaObj).GoalSeek Goal:=0, ChangingCell:=Range(CeldaNBP)
    
With Application
.Calculation = xlManual
End With
    
End Sub" "Private Sub CommandButton1_Click()
'Calcular la cant de agua fresca tal que
' se cumpla la solubilidad deseada.

 SolverOk SetCell:=""$S$39"", MaxMinVal:=2, ValueOf:=0, ByChange:=""$P$39:$P$40"", _
        Engine:=1, EngineDesc:=""GRG Nonlinear""
    SolverOk SetCell:=""$S$39"", MaxMinVal:=2, ValueOf:=0, ByChange:=""$P$39:$P$40"", _
        Engine:=1, EngineDesc:=""GRG Nonlinear""
    SolverSolve

End Sub

Private Sub CommandButton2_Click()

'Agua
Range(""S50"").GoalSeek Goal:=0, ChangingCell:=Range(""P40"")
'Metanol
Range(""S52"").GoalSeek Goal:=0, ChangingCell:=Range(""AD14"")
'Formaldehído
Range(""S54"").GoalSeek Goal:=0, ChangingCell:=Range(""AD16"")
'DME
Range(""S55"").GoalSeek Goal:=0, ChangingCell:=Range(""AD17"")

End Sub

Private Sub CommandButton3_Click()
'Bce de Energía en tope de torre
Range(""O5"").GoalSeek Goal:=0, ChangingCell:=Range(""M24"")
End Sub

Private Sub CommandButton4_Click()
'Caudal de agua de dilución
Range(""AX45"").GoalSeek Goal:=0, ChangingCell:=Range(""AM36"")
End Sub

Private Sub CommandButton5_Click()
'Temperatura final
Range(""AZ44"").GoalSeek Goal:=0, ChangingCell:=Range(""AU43"")
End Sub" "Sub Macro4()
'
' Macro4 Macro
'

'
   
End Sub"
"Public Sub ConfigDoc()
With Application
.Calculation = xlManual
End Sub
" "'Dims de Balance de Masa 1 reactor
Public AlimFresca(1 To 8) As Double
Public EntradaReact(1 To 8) As Double
Public SalidaReact(1 To 8) As Double
Public Producto(1 To 8) As Double
Public Purga(1 To 8) As Double
Public Reciclo(1 To 8) As Double
Public FracPurga As Double
Public iA, jA As Integer
Public BalanceGlobal, ErrorBalance As Double
Public CeldaFresca, CeldaER, CeldaSalR, CeldaProd, CeldaPurg, CeldaRec As String

'Dims de Reactor 1
Public Lrow, aux1 As Double
Public Cell0, CellN, Lcell, CellErase0 As String
Public SR1, SRAfill, SRErase As String

Public Xout, Fmout, Wreq, Wreqaux, VolR, i, n, j, i1, n2, error As Double
Public RowM1, RowM2 As String
Public allowE, step1 As Double

Public RowObj As String

Public T0, T1, StepT, iT, contador As Double

Public Nfinal As Double
Public CellF0, CellF1, RangoFinal As String

'Dims de Reactor 2
Public Sout As Double
Public AlimFresca2(1 To 12) As Double
Public EntradaReact2(1 To 12) As Double
Public Reciclo2(1 To 12) As Double

" "Private Sub CommandButton1_Click()

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

Call ClearAll1

End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O21"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()
Call Reactor2
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
 
End Sub" "Public Sub HojaConfig()

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()
End Sub
Private Sub CommandButton2_Click()

With Application
.Calculation = xlAutomatic
 End With

'Alim Fresca asignación valores
For iA = 1 To 8
jA = iA + 30
CeldaFresca = ""D"" & jA
AlimFresca(iA) = ActiveSheet.Range(CeldaFresca).Value
Next iA

'Reciclo asignación valores
For iA = 1 To 8
jA = iA + 12
CeldaRec = ""R"" & jA
Reciclo(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA


'Primer paso, suponer entrada reactor = entrada planta + reciclo anterior de seed
For iA = 1 To 8
jA = iA + 30
CeldaER = ""I"" & jA
EntradaReact(iA) = AlimFresca(iA) + Reciclo(iA)
ActiveSheet.Range(CeldaER) = EntradaReact(iA)
Next iA


'Ir a hoja de Reactor 1
Worksheets(""1R"").Activate

'Simular Reactor
Call ReactorAd1BM

'Volver a hoja BM
Worksheets(""CN + BM R1"").Activate

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N40"").Value
ErrorBalance = ActiveSheet.Range(""N41"").Value

'Loop para que haga el balance hasta que BAL de Masa Global
'sea menor al error admitido (Bal debe ser 0)

Do While BalanceGlobal > ErrorBalance

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N40"").Value

'Continuar con el balance de masa
'Reciclo asignación valores
For iA = 1 To 8
jA = iA + 12
CeldaRec = ""R"" & jA
Reciclo(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA

For iA = 1 To 8
EntradaReact(iA) = Reciclo(iA) + AlimFresca(iA)
Next iA

For iA = 1 To 8
jA = iA + 30
CeldaER = ""I"" & jA
ActiveSheet.Range(CeldaER) = EntradaReact(iA)
Next iA

'Ir a hoja de Reactor 1
Worksheets(""1R"").Activate

'Simular Reactor
Call ReactorAd1BM

'Volver a hoja BM
Worksheets(""CN + BM R1"").Activate
Debug.Print BalanceGlobal
Loop

With Application
.Calculation = xlManual
 End With
 
End Sub

" "Public Sub SetManual()

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

Call ClearAll1
End Sub

Private Sub CommandButton3_Click()

Nfinal = ActiveSheet.Range(""O20"")
CellF0 = ActiveSheet.Range(""O7"") & Nfinal
CellF1 = ActiveSheet.Range(""O8"") & Nfinal
RangoFinal = CellF0 & "":"" & CellF1

ActiveSheet.Range(RangoFinal).Select

End Sub
Private Sub CommandButton4_Click()
Call ReactorAd1
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

End Sub" "Public Sub ReactorAd1()

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

'Parámetros
Fmout = ActiveSheet.Range(RowM1)
Wreq = ActiveSheet.Range(RowM1).Offset(0, -8).Value
VolR = ActiveSheet.Range(RowM1).Offset(0, -7).Value

'Celdas
ActiveSheet.Range(""O17"") = Fmout
ActiveSheet.Range(""O18"").Value = Wreq
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

End Sub" "Public Sub ClearAll1()

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




End Sub" "Public Sub ReactorAd1BM()

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

'Eliminar datos extra para calcular ok IDQ
Dim Eridqfila As Double
Dim ErIDQ1, ErTot As String

If i > 30 Then
Eridqfila = 1 + i
ErIDQ1 = ""A"" & Eridqfila
ErTot = ErIDQ1 & "":"" & Lcell
ActiveSheet.Range(ErTot).ClearContents
End If

With Application
.Calculation = xlManual
End With

End Sub
" "Public Sub Reactor2()

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
'Fmout = ActiveSheet.Range(RowM1)
'ActiveSheet.Range(""O17"") = Fmout

'W Cat
'Wreq = ActiveSheet.Range(RowM1).Offset(0, -10).Value
'ActiveSheet.Range(""O18"").Value = Wreq

'V de Reactor
'VolR = ActiveSheet.Range(RowM1).Offset(0, -9).Value
'ActiveSheet.Range(""O19"") = VolR

'S A la salida
'Sout = ActiveSheet.Range(RowM1).Offset(0, -7).Value
'ActiveSheet.Range(""O20"") = Sout

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
" "Public Sub HojaConfig()

With Application
.Calculation = xlManual
End With

End Sub
Private Sub CommandButton1_Click()
With Application
.Calculation = xlAutomatic
End With

Dim RestaObj, CeldaNBP As String

RestaObj = ActiveSheet.Range(""K3"")
CeldaNBP = ActiveSheet.Range(""K4"")

Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range(RestaObj).GoalSeek Goal:=0, ChangingCell:=Range(CeldaNBP)
    
With Application
.Calculation = xlManual
End With
    
End Sub
Private Sub CommandButton2_Click()

With Application
.Calculation = xlAutomatic
 End With

'Alim Fresca asignación valores
For iA = 1 To 12
jA = iA + 30
CeldaFresca = ""D"" & jA
AlimFresca2(iA) = ActiveSheet.Range(CeldaFresca).Value
Next iA

'Reciclo asignación valores
For iA = 1 To 12
jA = iA + 6
CeldaRec = ""R"" & jA
Reciclo2(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA

'Primer paso, suponer entrada reactor = entrada planta + reciclo anterior de seed
For iA = 1 To 12
jA = iA + 30
CeldaER = ""I"" & jA
EntradaReact2(iA) = AlimFresca2(iA) + Reciclo2(iA)
ActiveSheet.Range(CeldaER) = EntradaReact2(iA)
Next iA

'Ir a hoja de Reactor 2
Worksheets(""2R"").Activate

'Simular Reactor
Call Reactor2

'Volver a hoja BM
Worksheets(""CN + BM R2"").Activate

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N46"").Value
ErrorBalance = ActiveSheet.Range(""N47"").Value

'Loop para que haga el balance hasta que BAL de Masa Global
'sea menor al error admitido (Bal debe ser 0)

Do While BalanceGlobal > ErrorBalance

'Definir Balance Global y Error el balance
BalanceGlobal = ActiveSheet.Range(""N46"").Value

'Continuar con el balance de masa
'Reciclo asignación valores
For iA = 1 To 12
jA = iA + 6
CeldaRec = ""R"" & jA
Reciclo2(iA) = ActiveSheet.Range(CeldaRec).Value
Next iA

For iA = 1 To 12
EntradaReact2(iA) = Reciclo2(iA) + AlimFresca2(iA)
Next iA

For iA = 1 To 12
jA = iA + 30
CeldaER = ""I"" & jA
ActiveSheet.Range(CeldaER) = EntradaReact2(iA)
Next iA

'Ir a hoja de Reactor 2
Worksheets(""2R"").Activate

'Simular Reactor
Call Reactor2

'Volver a hoja BM
Worksheets(""CN + BM R2"").Activate
Loop

With Application
.Calculation = xlManual
 End With
 
End Sub

" "Private Sub CommandButton1_Click()
With Application
.Calculation = xlAutomatic
End With

Dim RestaObj, CeldaNBP As String

RestaObj = ActiveSheet.Range(""K3"")
CeldaNBP = ActiveSheet.Range(""K4"")

Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range(RestaObj).GoalSeek Goal:=0, ChangingCell:=Range(CeldaNBP)
    
With Application
.Calculation = xlManual
End With
    
End Sub" "Private Sub CommandButton1_Click()
'Calcular la cant de agua fresca tal que
' se cumpla la solubilidad deseada.

 SolverOk SetCell:=""$S$39"", MaxMinVal:=2, ValueOf:=0, ByChange:=""$P$39:$P$40"", _
        Engine:=1, EngineDesc:=""GRG Nonlinear""
    SolverOk SetCell:=""$S$39"", MaxMinVal:=2, ValueOf:=0, ByChange:=""$P$39:$P$40"", _
        Engine:=1, EngineDesc:=""GRG Nonlinear""
    SolverSolve

End Sub

Private Sub CommandButton2_Click()

'Agua
Range(""S50"").GoalSeek Goal:=0, ChangingCell:=Range(""P40"")
'Metanol
Range(""S52"").GoalSeek Goal:=0, ChangingCell:=Range(""AD14"")
'Formaldehído
Range(""S54"").GoalSeek Goal:=0, ChangingCell:=Range(""AD16"")
'DME
Range(""S55"").GoalSeek Goal:=0, ChangingCell:=Range(""AD17"")

End Sub

Private Sub CommandButton3_Click()
'Bce de Energía en tope de torre
Range(""O5"").GoalSeek Goal:=0, ChangingCell:=Range(""M24"")
End Sub

Private Sub CommandButton4_Click()
'Caudal de agua de dilución
Range(""AX45"").GoalSeek Goal:=0, ChangingCell:=Range(""AM36"")
End Sub

Private Sub CommandButton5_Click()
'Temperatura final
Range(""AZ44"").GoalSeek Goal:=0, ChangingCell:=Range(""AU43"")
End Sub" "Sub Macro4()
'
' Macro4 Macro
'

'
   
End Sub"
