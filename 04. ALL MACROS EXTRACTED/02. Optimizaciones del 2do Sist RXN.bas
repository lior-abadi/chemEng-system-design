"Private Sub CommandButton1_Click()
Call DataExtraer
End Sub" "Option Explicit
Public Sub DataExtraer()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase

'Link a UniSimDesign
   'Link a UniSim
    Set UniApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set SimCase = UniApp.ActiveDocument
    If SimCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
  'Dim de objetos de operacion
 Dim K21 As UniSimDesign.CompressOp
 Dim K22 As UniSimDesign.CompressOp
 Dim K23 As UniSimDesign.CompressOp
 
 Dim R21 As UniSimDesign.PFReactor
 Dim R22 As UniSimDesign.PFReactor
 Dim R23 As UniSimDesign.PFReactor
 
 'Link a Equipos
 Set K21 = SimCase.Flowsheet.Operations.Item(""K21"")
 Set K22 = SimCase.Flowsheet.Operations.Item(""K22"")
 Set K23 = SimCase.Flowsheet.Operations.Item(""K23"")
 
 Set R21 = SimCase.Flowsheet.Operations.Item(""R21"")
 Set R22 = SimCase.Flowsheet.Operations.Item(""R22"")
 Set R23 = SimCase.Flowsheet.Operations.Item(""R23"")
 
 'Dim de Corrientes
 Dim Producto As UniSimDesign.ProcessStream
 Dim Aire As UniSimDesign.ProcessStream
 Dim TopeFlash As UniSimDesign.ProcessStream
 Dim FondoFlash As UniSimDesign.ProcessStream
  
 'Link a Corrientes
 Set Producto = SimCase.Flowsheet.MaterialStreams.Item(""49"")
 Set TopeFlash = SimCase.Flowsheet.MaterialStreams.Item(""39"")
 Set FondoFlash = SimCase.Flowsheet.MaterialStreams.Item(""40"")
 Set Aire = SimCase.Flowsheet.MaterialStreams.Item(""23"")
 
 Dim UnitSet(4) As String
 
 'Dimensionado de Unidades
 UnitSet(0) = ActiveSheet.Range(""C13"") 'Caudal másico
 UnitSet(1) = ActiveSheet.Range(""C19"") 'Caudal Molar
 UnitSet(2) = ActiveSheet.Range(""G13"") 'Volumen
 UnitSet(3) = ActiveSheet.Range(""K17"") 'Potencia
 UnitSet(4) = ActiveSheet.Range(""S9"") 'Densidad
  
 'Caudales Másicos
  ActiveSheet.Range(""B13"") = Producto.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""B19"") = Aire.MassFlow.GetValue(UnitSet(0))
  
  ActiveSheet.Range(""R8"") = TopeFlash.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""R12"") = FondoFlash.MassFlow.GetValue(UnitSet(0))
  
  'Densidades
  ActiveSheet.Range(""R9"") = TopeFlash.MassDensity.GetValue(UnitSet(4))
  ActiveSheet.Range(""R13"") = FondoFlash.MassDensity.GetValue(UnitSet(4))
 
 'Volumenes
 ActiveSheet.Range(""F13"") = R21.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F14"") = R22.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F15"") = R23.TotalVolume.GetValue(UnitSet(2))
 
 
 'Potencias
 ActiveSheet.Range(""J17"") = K21.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J18"") = K22.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J19"") = K23.Energy.GetValue(UnitSet(3))
 
 
End Sub"
"Private Sub CommandButton1_Click()
Call DataExtraer
End Sub" "Option Explicit
Public Sub DataExtraer()

'Dim de objetos UniSim
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase

'Link a UniSimDesign
   'Link a UniSim
    Set UniApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set SimCase = UniApp.ActiveDocument
    If SimCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
  'Dim de objetos de operacion
 Dim K21 As UniSimDesign.CompressOp
 Dim K22 As UniSimDesign.CompressOp
 Dim K23 As UniSimDesign.CompressOp
 
 Dim R21 As UniSimDesign.PFReactor
 Dim R22 As UniSimDesign.PFReactor
 Dim R23 As UniSimDesign.PFReactor
 
 'Link a Equipos
 Set K21 = SimCase.Flowsheet.Operations.Item(""K21"")
 Set K22 = SimCase.Flowsheet.Operations.Item(""K22"")
 Set K23 = SimCase.Flowsheet.Operations.Item(""K23"")
 
 Set R21 = SimCase.Flowsheet.Operations.Item(""R21"")
 Set R22 = SimCase.Flowsheet.Operations.Item(""R22"")
 Set R23 = SimCase.Flowsheet.Operations.Item(""R23"")
 
 'Dim de Corrientes
 Dim Producto As UniSimDesign.ProcessStream
 Dim Aire As UniSimDesign.ProcessStream
 Dim TopeFlash As UniSimDesign.ProcessStream
 Dim FondoFlash As UniSimDesign.ProcessStream
  
 'Link a Corrientes
 Set Producto = SimCase.Flowsheet.MaterialStreams.Item(""49"")
 Set TopeFlash = SimCase.Flowsheet.MaterialStreams.Item(""39"")
 Set FondoFlash = SimCase.Flowsheet.MaterialStreams.Item(""40"")
 Set Aire = SimCase.Flowsheet.MaterialStreams.Item(""23"")
 
 Dim UnitSet(4) As String
 
 'Dimensionado de Unidades
 UnitSet(0) = ActiveSheet.Range(""C13"") 'Caudal másico
 UnitSet(1) = ActiveSheet.Range(""C19"") 'Caudal Molar
 UnitSet(2) = ActiveSheet.Range(""G13"") 'Volumen
 UnitSet(3) = ActiveSheet.Range(""K17"") 'Potencia
 UnitSet(4) = ActiveSheet.Range(""S9"") 'Densidad
  
 'Caudales Másicos
  ActiveSheet.Range(""B13"") = Producto.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""B19"") = Aire.MassFlow.GetValue(UnitSet(0))
  
  ActiveSheet.Range(""R8"") = TopeFlash.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""R12"") = FondoFlash.MassFlow.GetValue(UnitSet(0))
  
  'Densidades
  ActiveSheet.Range(""R9"") = TopeFlash.MassDensity.GetValue(UnitSet(4))
  ActiveSheet.Range(""R13"") = FondoFlash.MassDensity.GetValue(UnitSet(4))
 
 'Volumenes
 ActiveSheet.Range(""F13"") = R21.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F14"") = R22.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F15"") = R23.TotalVolume.GetValue(UnitSet(2))
 
 
 'Potencias
 ActiveSheet.Range(""J17"") = K21.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J18"") = K22.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J19"") = K23.Energy.GetValue(UnitSet(3))
 
 
End Sub"
