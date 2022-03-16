"Private Sub CommandButton1_Click()
Call DensidadCN

End Sub

Private Sub CommandButton2_Click()
Call ClearAll
End Sub" "Option Explicit
Public Sub DensidadCN()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 'Dim de objetos de corriente
 Dim Gas As UniSimDesign.ProcessStream
 
 Dim CeldaT, CeldaDens, CeldaZ As String

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
    UnitSet(0) = ActiveSheet.Range(""C5"") 'Presion
    UnitSet(1) = ActiveSheet.Range(""A13"") 'Temperatura
    UnitSet(2) = ActiveSheet.Range(""B13"") 'Densidad masica
      
   'Link a corrientes
  Set Gas = SimCase.Flowsheet.MaterialStreams.Item(""Gas"")
     
  'Setear Presion de evaluacion
    'Apagar solver de Unisim
  SimCase.Solver.CanSolve = False
  Gas.Pressure.SetValue ActiveSheet.Range(""B5""), UnitSet(0)
  
  'Encender solver de Unisim
  SimCase.Solver.CanSolve = True
  
  inicio = ActiveSheet.Range(""B7"")
  fin = ActiveSheet.Range(""B8"")
  
  
  
        
  For i = inicio To fin
  CeldaT = ""A"" & i
  CeldaDens = ""B"" & i
  CeldaZ = ""C"" & i
      
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
   'Llevar datos a Unisim
  Gas.Temperature.SetValue ActiveSheet.Range(CeldaT), UnitSet(1)
    
  'Encender Solver
  SimCase.Solver.CanSolve = True
  
  'Traer datos de Densidad y Z
  ActiveSheet.Range(CeldaDens) = Gas.MassDensity.GetValue(UnitSet(2))
  ActiveSheet.Range(CeldaZ) = Gas.ZFactor.GetValue
  
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
