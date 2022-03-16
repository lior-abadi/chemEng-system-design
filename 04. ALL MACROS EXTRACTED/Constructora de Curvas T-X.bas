"Private Sub CommandButton1_Click()
Call CurvasTX1

End Sub" "Option Explicit
Public Sub CurvasTX1()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 'Dim de objetos de corriente
 Dim Comp1 As UniSimDesign.ProcessStream
 Dim Comp2 As UniSimDesign.ProcessStream
 Dim Tburb As UniSimDesign.ProcessStream
 Dim Troc As UniSimDesign.ProcessStream
 Dim Mezcla As UniSimDesign.ProcessStream
 
 Dim CeldaFC1, CeldaFC2 As String
 Dim CeldaTb, CeldaTr As String
 
 Dim UnitSet(2) As String
 
 Dim i As Integer
 
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
    UnitSet(0) = ActiveSheet.Range(""C3"") 'Presion
    UnitSet(1) = ActiveSheet.Range(""C4"") 'Caudal
    UnitSet(2) = ActiveSheet.Range(""B5"") 'Temperatura
    
   'Link a corrientes
  Set Comp1 = SimCase.Flowsheet.MaterialStreams.Item(""Comp1"")
  Set Comp2 = SimCase.Flowsheet.MaterialStreams.Item(""Comp2"")
  Set Tburb = SimCase.Flowsheet.MaterialStreams.Item(""Tburbuja"")
  Set Troc = SimCase.Flowsheet.MaterialStreams.Item(""Trocio"")
  Set Mezcla = SimCase.Flowsheet.MaterialStreams.Item(""Mezcla"")
  
  
  'Setear presión de evaluacion
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
  Mezcla.Pressure.SetValue ActiveSheet.Range(""B3""), UnitSet(0)
  'Encender solver de Unisim
        SimCase.Solver.CanSolve = True
        
  For i = 9 To 109
  CeldaFC1 = ""D"" & i
  CeldaFC2 = ""E"" & i
  CeldaTb = ""F"" & i
  CeldaTr = ""G"" & i
  
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
        
  'Llevar datos a Unisim
  Comp1.MolarFlow.SetValue ActiveSheet.Range(CeldaFC1), UnitSet(1)
  Comp2.MolarFlow.SetValue ActiveSheet.Range(CeldaFC2), UnitSet(1)
  
  'Encender Solver
  SimCase.Solver.CanSolve = True
  
  'Traer datos de T
  ActiveSheet.Range(CeldaTb) = Tburb.Temperature.GetValue(UnitSet(2))
  ActiveSheet.Range(CeldaTr) = Troc.Temperature.GetValue(UnitSet(2))
  
  Next
  
   
End Sub"
