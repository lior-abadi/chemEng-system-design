"Private Sub CommandButton1_Click()
Call Compresion

End Sub

Private Sub CommandButton2_Click()
Call ClearAll
End Sub" "Option Explicit
Public Sub Compresion()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 'Dim de objetos de corriente
 Dim str1 As UniSimDesign.ProcessStream
 Dim str2 As UniSimDesign.ProcessStream
 Dim str3 As UniSimDesign.ProcessStream
 Dim str5 As UniSimDesign.ProcessStream
 Dim str6 As UniSimDesign.ProcessStream
 
 'Dim de objetos de operacion
 Dim K1 As UniSimDesign.CompressOp
 Dim K2 As UniSimDesign.CompressOp
 Dim Cooler As UniSimDesign.CoolerOp
  
 
 Dim CeldaP1, CeldaP2, CeldaT2, CeldaPK1, CeldaT3 As String
 Dim CeldaP5, CeldaP6, CeldaT6, CeldaPK2 As String
 
 Dim UnitSet(2) As String
 
 Dim i, inicio, fin As Integer
 
 'Link a UniSimDesign
   'Link a UniSim
    Set UniApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set SimCase = UniApp.ActiveDocument
    If SimCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
  'Set de Unidades:
    UnitSet(0) = ActiveSheet.Range(""A25"") 'Presion
    UnitSet(1) = ActiveSheet.Range(""D25"") 'Temperatura
    UnitSet(2) = ActiveSheet.Range(""E25"") 'Potencia
      
   'Link a corrientes y equipos
  Set str1 = SimCase.Flowsheet.MaterialStreams.Item(""1"")
  Set str2 = SimCase.Flowsheet.MaterialStreams.Item(""2"")
  Set str3 = SimCase.Flowsheet.MaterialStreams.Item(""3"")
  Set str5 = SimCase.Flowsheet.MaterialStreams.Item(""5"")
  Set str6 = SimCase.Flowsheet.MaterialStreams.Item(""6"")
  Set K1 = SimCase.Flowsheet.Operations.Item(""K1"")
  Set K2 = SimCase.Flowsheet.Operations.Item(""K2"")
  Set Cooler = SimCase.Flowsheet.Operations.Item(""Cooler"")
  
  'Delta P en Cooler
  Cooler.PressureDrop.SetValue ActiveSheet.Range(""B22""), UnitSet(0)
    
  inicio = ActiveSheet.Range(""B7"")
  fin = ActiveSheet.Range(""B8"")
    
        
  For i = inicio To fin
  CeldaP1 = ""A"" & i
  CeldaP2 = ""B"" & i
  CeldaT2 = ""D"" & i
  CeldaPK1 = ""E"" & i
  CeldaT3 = ""G"" & i
  CeldaP5 = ""H"" & i
  CeldaP6 = ""I"" & i
  CeldaT6 = ""K"" & i
  CeldaPK2 = ""L"" & i
    
        
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
        
   'Llevar datos a Unisim
  str1.Pressure.SetValue ActiveSheet.Range(CeldaP1), UnitSet(0)
  str2.Pressure.SetValue ActiveSheet.Range(CeldaP2), UnitSet(0)
  str3.Temperature.SetValue ActiveSheet.Range(CeldaT3), UnitSet(1)
  str6.Pressure.SetValue ActiveSheet.Range(CeldaP6), UnitSet(0)
       
    
  'Encender Solver
  SimCase.Solver.CanSolve = True
  
  'Traer datos de Simulador
  ActiveSheet.Range(CeldaT2) = str2.Temperature.GetValue(UnitSet(1))
  ActiveSheet.Range(CeldaPK1) = K1.Energy.GetValue(UnitSet(2))
  ActiveSheet.Range(CeldaP5) = str5.Pressure.GetValue(UnitSet(0))
  ActiveSheet.Range(CeldaT6) = str6.Temperature.GetValue(UnitSet(1))
  ActiveSheet.Range(CeldaPK2) = K2.Energy.GetValue(UnitSet(2))
  
  
  Next
  
   
End Sub" "Option Explicit
Public Sub ClearAll()

Dim desde, hasta, rangoErase As String

desde = ActiveSheet.Range(""B7"")
  hasta = ActiveSheet.Range(""B8"")
  
  rangoErase = ""B"" & desde & "":"" & ""C"" & hasta

Range(rangoErase).Select
    Selection.ClearContents

End Sub" "Private Sub CommandButton1_Click()
Call Compresion3

End Sub

Private Sub CommandButton2_Click()
Call ClearAll
End Sub" "Option Explicit
Public Sub Compresion3()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 'Dim de objetos de corriente
 Dim str1 As UniSimDesign.ProcessStream
 Dim str2 As UniSimDesign.ProcessStream
 Dim str3 As UniSimDesign.ProcessStream
 Dim str5 As UniSimDesign.ProcessStream
 Dim str6 As UniSimDesign.ProcessStream
 Dim str7 As UniSimDesign.ProcessStream
 Dim str9 As UniSimDesign.ProcessStream
 Dim str10 As UniSimDesign.ProcessStream
 
 'Dim de objetos de operacion
 Dim K1 As UniSimDesign.CompressOp
 Dim K2 As UniSimDesign.CompressOp
 Dim K3 As UniSimDesign.CompressOp
 Dim Cooler As UniSimDesign.CoolerOp
 Dim Cooler2 As UniSimDesign.CoolerOp
 
 Dim CeldaP1, CeldaP2, CeldaT2, CeldaPK1, CeldaT3 As String
 Dim CeldaP5, CeldaP6, CeldaT6, CeldaPK2 As String
 Dim CeldaP9, CeldaP10, CeldaT10, CeldaT9, CeldaPK3, CeldaT7 As String
 
 Dim UnitSet(2) As String
 
 Dim i, inicio, fin As Integer
 
 'Link a UniSimDesign
   'Link a UniSim
    Set UniApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set SimCase = UniApp.ActiveDocument
    If SimCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
  'Set de Unidades:
    UnitSet(0) = ActiveSheet.Range(""A25"") 'Presion
    UnitSet(1) = ActiveSheet.Range(""D25"") 'Temperatura
    UnitSet(2) = ActiveSheet.Range(""E25"") 'Potencia
      
   'Link a corrientes y equipos
  Set str1 = SimCase.Flowsheet.MaterialStreams.Item(""1"")
  Set str2 = SimCase.Flowsheet.MaterialStreams.Item(""2"")
  Set str3 = SimCase.Flowsheet.MaterialStreams.Item(""3"")
  Set str5 = SimCase.Flowsheet.MaterialStreams.Item(""5"")
  Set str6 = SimCase.Flowsheet.MaterialStreams.Item(""6"")
  Set str7 = SimCase.Flowsheet.MaterialStreams.Item(""7"")
  Set str9 = SimCase.Flowsheet.MaterialStreams.Item(""9"")
  Set str10 = SimCase.Flowsheet.MaterialStreams.Item(""10"")
  Set K1 = SimCase.Flowsheet.Operations.Item(""K1"")
  Set K2 = SimCase.Flowsheet.Operations.Item(""K2"")
  Set K3 = SimCase.Flowsheet.Operations.Item(""K3"")
  Set Cooler = SimCase.Flowsheet.Operations.Item(""Cooler"")
  Set Cooler2 = SimCase.Flowsheet.Operations.Item(""Cooler2"")
  
  'Delta P en Cooler
  Cooler.PressureDrop.SetValue ActiveSheet.Range(""B22""), UnitSet(0)
  Cooler2.PressureDrop.SetValue ActiveSheet.Range(""B22""), UnitSet(0)
    
  inicio = ActiveSheet.Range(""B7"")
  fin = ActiveSheet.Range(""B8"")
    
        
  For i = inicio To fin
  CeldaP1 = ""A"" & i
  CeldaP2 = ""B"" & i
  CeldaT2 = ""D"" & i
  CeldaPK1 = ""E"" & i
  CeldaT3 = ""G"" & i
  CeldaP5 = ""H"" & i
  CeldaP6 = ""I"" & i
  CeldaT6 = ""K"" & i
  CeldaPK2 = ""L"" & i
  CeldaT7 = ""N"" & i
  CeldaP9 = ""O"" & i
  CeldaP10 = ""P"" & i
  CeldaT10 = ""R"" & i
  CeldaPK3 = ""S"" & i
    
        
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
        
   'Llevar datos a Unisim
  str1.Pressure.SetValue ActiveSheet.Range(CeldaP1), UnitSet(0)
  str2.Pressure.SetValue ActiveSheet.Range(CeldaP2), UnitSet(0)
  str3.Temperature.SetValue ActiveSheet.Range(CeldaT3), UnitSet(1)
  str6.Pressure.SetValue ActiveSheet.Range(CeldaP6), UnitSet(0)
  str7.Temperature.SetValue ActiveSheet.Range(CeldaT7), UnitSet(1)
  str10.Pressure.SetValue ActiveSheet.Range(CeldaP10), UnitSet(0)
    
  'Encender Solver
  SimCase.Solver.CanSolve = True
  
  'Traer datos de Simulador
  ActiveSheet.Range(CeldaT2) = str2.Temperature.GetValue(UnitSet(1))
  ActiveSheet.Range(CeldaPK1) = K1.Energy.GetValue(UnitSet(2))
  ActiveSheet.Range(CeldaP5) = str5.Pressure.GetValue(UnitSet(0))
  ActiveSheet.Range(CeldaT6) = str6.Temperature.GetValue(UnitSet(1))
  ActiveSheet.Range(CeldaPK2) = K2.Energy.GetValue(UnitSet(2))
  ActiveSheet.Range(CeldaP9) = str9.Pressure.GetValue(UnitSet(0))
  ActiveSheet.Range(CeldaT10) = str10.Temperature.GetValue(UnitSet(1))
  ActiveSheet.Range(CeldaPK3) = K3.Energy.GetValue(UnitSet(2))
  
  Next
  
   
End Sub
"
