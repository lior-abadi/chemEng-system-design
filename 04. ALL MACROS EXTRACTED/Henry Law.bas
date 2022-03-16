"Private Sub CommandButton1_Click()
Call CurvasSol1

End Sub" "Option Explicit
Public Sub CurvasSol1()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 'Dim de objetos de corriente
 Dim Gas As UniSimDesign.ProcessStream
 Dim Liq As UniSimDesign.ProcessStream
 Dim Saturada As UniSimDesign.ProcessStream
 
 
 Dim CeldaP, CeldaGMF As String
 
 Dim UnitSet(4) As String
 
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
    UnitSet(0) = ActiveSheet.Range(""B5"") 'Presion
    UnitSet(1) = ActiveSheet.Range(""C8"") 'Caudal molar
    UnitSet(2) = ActiveSheet.Range(""C4"") 'Caudal volumetrico
    UnitSet(3) = ActiveSheet.Range(""C3"") 'Temperatura
    UnitSet(4) = ActiveSheet.Range(""G1"") 'Masico
    
   'Link a corrientes
  Set Gas = SimCase.Flowsheet.MaterialStreams.Item(""Gas"")
  Set Liq = SimCase.Flowsheet.MaterialStreams.Item(""Liq"")
  Set Saturada = SimCase.Flowsheet.MaterialStreams.Item(""Saturada"")
   
  'Setear Temp y Caudal de Solvente de evaluacion
  
  'Apagar solver de Unisim
  SimCase.Solver.CanSolve = False
  Gas.Temperature.SetValue ActiveSheet.Range(""B3""), UnitSet(3)
  Liq.MolarFlow.SetValue ActiveSheet.Range(""B4""), UnitSet(1)
  
  'Encender solver de Unisim
  SimCase.Solver.CanSolve = True
        
  For i = 9 To 24
  CeldaP = ""B"" & i
  CeldaGMF = ""C"" & i
    
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
        
  Gas.MassFlow.SetValue 1000, UnitSet(4)
        
  'Llevar datos a Unisim
  Gas.Pressure.SetValue ActiveSheet.Range(CeldaP), UnitSet(0)
  
  
  'Encender Solver
  SimCase.Solver.CanSolve = True
  
  'Traer datos de Gas Molar Flow
  ActiveSheet.Range(CeldaGMF) = Gas.MolarFlow.GetValue(UnitSet(1))
  Next
  
   
End Sub"
