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
 Dim K11 As UniSimDesign.CompressOp
 Dim K12 As UniSimDesign.CompressOp
 Dim K13 As UniSimDesign.CompressOp
 
 Dim R11 As UniSimDesign.PFReactor
 Dim R12 As UniSimDesign.PFReactor
 Dim R13 As UniSimDesign.PFReactor
 
 'Link a Equipos
 Set K11 = SimCase.Flowsheet.Operations.Item(""K11"")
 Set K12 = SimCase.Flowsheet.Operations.Item(""K12"")
 Set K13 = SimCase.Flowsheet.Operations.Item(""K13"")
 
 Set R11 = SimCase.Flowsheet.Operations.Item(""R11"")
 Set R12 = SimCase.Flowsheet.Operations.Item(""R12"")
 Set R13 = SimCase.Flowsheet.Operations.Item(""R13"")
 
 'Dim de Corrientes
 Dim AlimFr As UniSimDesign.ProcessStream
 Dim Purga As UniSimDesign.ProcessStream
 Dim Producto As UniSimDesign.ProcessStream
 Dim Tope As UniSimDesign.ProcessStream
 
  
 'Link a Corrientes
 Set AlimFr = SimCase.Flowsheet.MaterialStreams.Item(""1"")
 Set Purga = SimCase.Flowsheet.MaterialStreams.Item(""15"")
 Set Producto = SimCase.Flowsheet.MaterialStreams.Item(""13"")
 Set Tope = SimCase.Flowsheet.MaterialStreams.Item(""14"")
 
 Dim UnitSet(4) As String
 
 'Dimensionado de Unidades
 UnitSet(0) = ActiveSheet.Range(""C13"") 'Caudal másico
 UnitSet(1) = ActiveSheet.Range(""C19"") 'Caudal Molar
 UnitSet(2) = ActiveSheet.Range(""G13"") 'Volumen
 UnitSet(3) = ActiveSheet.Range(""K17"") 'Potencia
 UnitSet(4) = ActiveSheet.Range(""S9"") 'Densidad masica
 
 'Dimensionado de indices auxiliares
 Dim XMeOHPurga, XMeOHProd, CaudalPu, CaudalPr As Double
 Dim MeOHIndex As Integer
 Dim CompMassFracPu, CompMassFracPr As Variant
 Dim Componentes As Components
 Dim j As Integer
 
 'Referenciado a propiedades molares de corriente
  Set Componentes = SimCase.BasisManager.FluidPackages.Item(0).Components
  CompMassFracPu = Purga.ComponentMassFraction
  CompMassFracPr = Producto.ComponentMassFraction
  
  For j = 0 To UBound(CompMassFracPu)
    Select Case Componentes.Item(j)
    Case Is = ""Methanol""
        MeOHIndex = j
    End Select
  Next
    
  
 'Extraer datos de UniSim
 'Fracciones másicas de Metanol
 XMeOHPurga = CompMassFracPu(MeOHIndex)
 XMeOHProd = CompMassFracPr(MeOHIndex)
  
 'Caudales
  CaudalPu = Purga.MassFlow.GetValue(UnitSet(0))
  CaudalPr = Producto.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""B13"") = CaudalPr * XMeOHProd
  ActiveSheet.Range(""B16"") = CaudalPu * XMeOHPurga
  
  ActiveSheet.Range(""B19"") = AlimFr.MolarFlow.GetValue(UnitSet(1))
  
  ActiveSheet.Range(""R8"") = Tope.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""R12"") = CaudalPr
 
 'Volumenes
 ActiveSheet.Range(""F13"") = R11.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F14"") = R12.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F15"") = R13.TotalVolume.GetValue(UnitSet(2))
 
 
 'Potencias
 ActiveSheet.Range(""J17"") = K11.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J18"") = K12.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J19"") = K13.Energy.GetValue(UnitSet(3))
 
 'Densidades
 ActiveSheet.Range(""R9"") = Tope.MassDensity.GetValue(UnitSet(4))
 ActiveSheet.Range(""R13"") = Producto.MassDensity.GetValue(UnitSet(4))
 
 
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
 Dim K11 As UniSimDesign.CompressOp
 Dim K12 As UniSimDesign.CompressOp
 Dim K13 As UniSimDesign.CompressOp
 
 Dim R11 As UniSimDesign.PFReactor
 Dim R12 As UniSimDesign.PFReactor
 Dim R13 As UniSimDesign.PFReactor
 
 'Link a Equipos
 Set K11 = SimCase.Flowsheet.Operations.Item(""K11"")
 Set K12 = SimCase.Flowsheet.Operations.Item(""K12"")
 Set K13 = SimCase.Flowsheet.Operations.Item(""K13"")
 
 Set R11 = SimCase.Flowsheet.Operations.Item(""R11"")
 Set R12 = SimCase.Flowsheet.Operations.Item(""R12"")
 Set R13 = SimCase.Flowsheet.Operations.Item(""R13"")
 
 'Dim de Corrientes
 Dim AlimFr As UniSimDesign.ProcessStream
 Dim Purga As UniSimDesign.ProcessStream
 Dim Producto As UniSimDesign.ProcessStream
 Dim Tope As UniSimDesign.ProcessStream
 
  
 'Link a Corrientes
 Set AlimFr = SimCase.Flowsheet.MaterialStreams.Item(""1"")
 Set Purga = SimCase.Flowsheet.MaterialStreams.Item(""15"")
 Set Producto = SimCase.Flowsheet.MaterialStreams.Item(""13"")
 Set Tope = SimCase.Flowsheet.MaterialStreams.Item(""14"")
 
 Dim UnitSet(4) As String
 
 'Dimensionado de Unidades
 UnitSet(0) = ActiveSheet.Range(""C13"") 'Caudal másico
 UnitSet(1) = ActiveSheet.Range(""C19"") 'Caudal Molar
 UnitSet(2) = ActiveSheet.Range(""G13"") 'Volumen
 UnitSet(3) = ActiveSheet.Range(""K17"") 'Potencia
 UnitSet(4) = ActiveSheet.Range(""S9"") 'Densidad masica
 
 'Dimensionado de indices auxiliares
 Dim XMeOHPurga, XMeOHProd, CaudalPu, CaudalPr As Double
 Dim MeOHIndex As Integer
 Dim CompMassFracPu, CompMassFracPr As Variant
 Dim Componentes As Components
 Dim j As Integer
 
 'Referenciado a propiedades molares de corriente
  Set Componentes = SimCase.BasisManager.FluidPackages.Item(0).Components
  CompMassFracPu = Purga.ComponentMassFraction
  CompMassFracPr = Producto.ComponentMassFraction
  
  For j = 0 To UBound(CompMassFracPu)
    Select Case Componentes.Item(j)
    Case Is = ""Methanol""
        MeOHIndex = j
    End Select
  Next
    
  
 'Extraer datos de UniSim
 'Fracciones másicas de Metanol
 XMeOHPurga = CompMassFracPu(MeOHIndex)
 XMeOHProd = CompMassFracPr(MeOHIndex)
  
 'Caudales
  CaudalPu = Purga.MassFlow.GetValue(UnitSet(0))
  CaudalPr = Producto.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""B13"") = CaudalPr * XMeOHProd
  ActiveSheet.Range(""B16"") = CaudalPu * XMeOHPurga
  
  ActiveSheet.Range(""B19"") = AlimFr.MolarFlow.GetValue(UnitSet(1))
  
  ActiveSheet.Range(""R8"") = Tope.MassFlow.GetValue(UnitSet(0))
  ActiveSheet.Range(""R12"") = CaudalPr
 
 'Volumenes
 ActiveSheet.Range(""F13"") = R11.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F14"") = R12.TotalVolume.GetValue(UnitSet(2))
 ActiveSheet.Range(""F15"") = R13.TotalVolume.GetValue(UnitSet(2))
 
 
 'Potencias
 ActiveSheet.Range(""J17"") = K11.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J18"") = K12.Energy.GetValue(UnitSet(3))
 ActiveSheet.Range(""J19"") = K13.Energy.GetValue(UnitSet(3))
 
 'Densidades
 ActiveSheet.Range(""R9"") = Tope.MassDensity.GetValue(UnitSet(4))
 ActiveSheet.Range(""R13"") = Producto.MassDensity.GetValue(UnitSet(4))
 
 
End Sub"
