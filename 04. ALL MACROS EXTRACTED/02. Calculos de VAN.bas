"Private Sub CommandButton1_Click()
Call OptimizCompr1
End Sub" "Option Explicit

Public Sub OptimizCompr1()



'Declare Variables------------------------------------------------------------------------------
    
    'UniSim Object variables
    Dim hyApp As UniSimDesign.Application
    Dim hyCase As UniSimDesign.SimulationCase
    
    'Corrientes
    Dim Exp2 As UniSimDesign.ProcessStream
    Dim Exp3 As UniSimDesign.ProcessStream
    Dim ExpQ1 As UniSimDesign.CoolerOp
    
    'Compresores
    Dim K100 As UniSimDesign.CompressOp
    Dim K101 As UniSimDesign.CompressOp
    
    'Auxiliares
    Dim i, j, FinalRow As Integer
    Dim ColumnaP, ColK100, ColQ1, ColPexp3, ColK101 As String
    Dim CeldaP, CeldaK100, CeldaQ1, CeldaPexp3, CeldaK101 As String

            
   'Set de Unidades
    Dim unitSet(1) As String
   
   'Procedimiento ----
   
   'Link a UniSimDesign
   'Link a UniSim
    Set hyApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set hyCase = hyApp.ActiveDocument
    If hyCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
    
    'Link a Corrientes:
    Set Exp2 = hyCase.Flowsheet.MaterialStreams.Item(""Exp2"")
    Set Exp3 = hyCase.Flowsheet.MaterialStreams.Item(""Exp3"")
    
    'Link a Equipos
    Set K100 = hyCase.Flowsheet.Operations.Item(""Kexp1"")
    Set K101 = hyCase.Flowsheet.Operations.Item(""Kexp2"")
    Set ExpQ1 = hyCase.Flowsheet.Operations.Item(""CoolExp1"")
    
    
    'Set de Unidades:
    unitSet(0) = ActiveSheet.Range(""A83"") 'Presión bar
    unitSet(1) = ActiveSheet.Range(""B83"") 'Energía kW
    
    'Seteo de posiciones
    FinalRow = ActiveSheet.Range(""I79"").Value
    ColumnaP = ActiveSheet.Range(""H80"")
    ColK100 = ""B""
    ColQ1 = ""F""
    ColPexp3 = ""G""
    ColK101 = ""H""
    
    'Barrido de presiones
    For i = 1 To FinalRow
    j = i + 83
    CeldaP = ColumnaP & j
    CeldaQ1 = ColQ1 & j
    CeldaK100 = ColK100 & j
    CeldaPexp3 = ColPexp3 & j
    CeldaK101 = ColK101 & j
     
    'Apagar solver de Unisim
        hyCase.Solver.CanSolve = False
        
    'Llevar info a UniSim
    Exp2.Pressure.SetValue ActiveSheet.Range(CeldaP).Value, unitSet(0)
          
    'Encender solver de Unisim
        hyCase.Solver.CanSolve = True
        
    'Leer Resultados de Unisim y llevarlos a Excel
    ActiveSheet.Range(CeldaK100).Value = K100.Energy.GetValue(unitSet(1))
    ActiveSheet.Range(CeldaQ1).Value = ExpQ1.Duty.GetValue(unitSet(1))
    ActiveSheet.Range(CeldaPexp3).Value = Exp3.Pressure.GetValue(unitSet(0))
    ActiveSheet.Range(CeldaK101).Value = K101.Energy.GetValue(unitSet(1))
    
    Next
    
End Sub
   "
"Private Sub CommandButton1_Click()
Call OptimizCompr1
End Sub" "Option Explicit

Public Sub OptimizCompr1()



'Declare Variables------------------------------------------------------------------------------
    
    'UniSim Object variables
    Dim hyApp As UniSimDesign.Application
    Dim hyCase As UniSimDesign.SimulationCase
    
    'Corrientes
    Dim Exp2 As UniSimDesign.ProcessStream
    Dim Exp3 As UniSimDesign.ProcessStream
    Dim ExpQ1 As UniSimDesign.CoolerOp
    
    'Compresores
    Dim K100 As UniSimDesign.CompressOp
    Dim K101 As UniSimDesign.CompressOp
    
    'Auxiliares
    Dim i, j, FinalRow As Integer
    Dim ColumnaP, ColK100, ColQ1, ColPexp3, ColK101 As String
    Dim CeldaP, CeldaK100, CeldaQ1, CeldaPexp3, CeldaK101 As String

            
   'Set de Unidades
    Dim unitSet(1) As String
   
   'Procedimiento ----
   
   'Link a UniSimDesign
   'Link a UniSim
    Set hyApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set hyCase = hyApp.ActiveDocument
    If hyCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
    
    'Link a Corrientes:
    Set Exp2 = hyCase.Flowsheet.MaterialStreams.Item(""Exp2"")
    Set Exp3 = hyCase.Flowsheet.MaterialStreams.Item(""Exp3"")
    
    'Link a Equipos
    Set K100 = hyCase.Flowsheet.Operations.Item(""Kexp1"")
    Set K101 = hyCase.Flowsheet.Operations.Item(""Kexp2"")
    Set ExpQ1 = hyCase.Flowsheet.Operations.Item(""CoolExp1"")
    
    
    'Set de Unidades:
    unitSet(0) = ActiveSheet.Range(""A83"") 'Presión bar
    unitSet(1) = ActiveSheet.Range(""B83"") 'Energía kW
    
    'Seteo de posiciones
    FinalRow = ActiveSheet.Range(""I79"").Value
    ColumnaP = ActiveSheet.Range(""H80"")
    ColK100 = ""B""
    ColQ1 = ""F""
    ColPexp3 = ""G""
    ColK101 = ""H""
    
    'Barrido de presiones
    For i = 1 To FinalRow
    j = i + 83
    CeldaP = ColumnaP & j
    CeldaQ1 = ColQ1 & j
    CeldaK100 = ColK100 & j
    CeldaPexp3 = ColPexp3 & j
    CeldaK101 = ColK101 & j
     
    'Apagar solver de Unisim
        hyCase.Solver.CanSolve = False
        
    'Llevar info a UniSim
    Exp2.Pressure.SetValue ActiveSheet.Range(CeldaP).Value, unitSet(0)
          
    'Encender solver de Unisim
        hyCase.Solver.CanSolve = True
        
    'Leer Resultados de Unisim y llevarlos a Excel
    ActiveSheet.Range(CeldaK100).Value = K100.Energy.GetValue(unitSet(1))
    ActiveSheet.Range(CeldaQ1).Value = ExpQ1.Duty.GetValue(unitSet(1))
    ActiveSheet.Range(CeldaPexp3).Value = Exp3.Pressure.GetValue(unitSet(0))
    ActiveSheet.Range(CeldaK101).Value = K101.Energy.GetValue(unitSet(1))
    
    Next
    
End Sub
   "
