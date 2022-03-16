"Private Sub CommandButton1_Click()
Call Tflash

End Sub

Private Sub CommandButton2_Click()
Call ClearAll
End Sub" "Option Explicit
Public Sub Tflash()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 'Dim de objetos de corriente
 Dim str1 As UniSimDesign.ProcessStream
 Dim str2 As UniSimDesign.ProcessStream
 Dim str3 As UniSimDesign.ProcessStream
 Dim str4 As UniSimDesign.ProcessStream
 Dim CompMolarFr4, CompMolarFr3  As Variant
 Dim MeOHIndex, CO2index As Integer
 Dim Componentes As Components
 
 'Dim de objetos de operacion
  Dim Cooler As UniSimDesign.CoolerOp

 
 Dim CeldaT2, CeldaCO24, CeldaCO23 As String
 Dim CeldaM4, CeldaM3, CeldaX3, CeldaX4 As String
 
 Dim UnitSet(2) As String
 
 Dim i, j, inicio, fin As Integer
 
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
    UnitSet(0) = ActiveSheet.Range(""C22"") 'Presion
    UnitSet(1) = ActiveSheet.Range(""A26"") 'Temperatura
    UnitSet(2) = ActiveSheet.Range(""C26"") 'Caudal Molar
      
   'Link a corrientes y equipos
  Set str1 = SimCase.Flowsheet.MaterialStreams.Item(""1"")
  Set str2 = SimCase.Flowsheet.MaterialStreams.Item(""2"")
  Set str3 = SimCase.Flowsheet.MaterialStreams.Item(""3"")
  Set str4 = SimCase.Flowsheet.MaterialStreams.Item(""4"")
  Set Cooler = SimCase.Flowsheet.Operations.Item(""Cooler"")
  
  'Delta P en Cooler
  Cooler.PressureDrop.SetValue ActiveSheet.Range(""B22""), UnitSet(0)
    
  inicio = ActiveSheet.Range(""B7"")
  fin = ActiveSheet.Range(""B8"")
  
  'Referenciado a propiedades molares de corriente
  Set Componentes = SimCase.BasisManager.FluidPackages.Item(0).Components
     
          
  For i = inicio To fin
  CeldaT2 = ""A"" & i
  CeldaX4 = ""B"" & i
  CeldaM4 = ""C"" & i
  CeldaX3 = ""E"" & i
  CeldaM3 = ""F"" & i
  CeldaCO24 = ""H"" & i
  CeldaCO23 = ""I"" & i
    
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
        
   'Llevar datos a Unisim
  str2.Temperature.SetValue ActiveSheet.Range(CeldaT2), UnitSet(1)
          
    'Encender Solver
  SimCase.Solver.CanSolve = True
  
  CompMolarFr3 = str3.ComponentMolarFraction
  CompMolarFr4 = str4.ComponentMolarFraction
  For j = 0 To UBound(CompMolarFr4)
    Select Case Componentes.Item(j)
    Case Is = ""Methanol""
        MeOHIndex = j
    Case Is = ""CO2""
        CO2index = j
    End Select
    Next
  
  'Traer datos de Simulador
  ActiveSheet.Range(CeldaX3) = CompMolarFr3(MeOHIndex)
  ActiveSheet.Range(CeldaM3) = str3.MolarFlow.GetValue(UnitSet(2))
  
  ActiveSheet.Range(CeldaX4) = CompMolarFr4(MeOHIndex)
  ActiveSheet.Range(CeldaM4) = str4.MolarFlow.GetValue(UnitSet(2))
  
  ActiveSheet.Range(CeldaCO24) = CompMolarFr4(CO2index)
  ActiveSheet.Range(CeldaCO23) = CompMolarFr3(CO2index)
  
  
  
  Next
  
   
End Sub" "Option Explicit
Public Sub ClearAll()

Dim desde, hasta, rangoErase As String

desde = ActiveSheet.Range(""B7"")
  hasta = ActiveSheet.Range(""B8"")
  
  rangoErase = ""B"" & desde & "":"" & ""C"" & hasta

Range(rangoErase).Select
    Selection.ClearContents

End Sub"
