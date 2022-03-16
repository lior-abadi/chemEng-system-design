"Private Sub CommandButton2_Click()
Call CoefDeActividad

End Sub" "Option Explicit

Public Sub CoefDeAct()

'Variables
    Dim UniApp As UniSimDesign.Application
    Dim SimCase As UniSimDesign.SimulationCase
    
    
' Paquete de propiedades
    Dim BasisManager As UniSimDesign.BasisManager
    Dim FluidPackage As UniSimDesign.FluidPackage
    Dim FluidPackages As UniSimDesign.FluidPackages
    Dim PropFluidPack As UniSimDesign.PropPackage
    Dim Uniquac As UniSimDesign.UNIQUACPropPkg
    Dim ActParam(3, 3) As UniSimDesign.ObjectVariable
    
    
    
    
'Variables auxiliares
    Dim j As Integer
    Dim MethForm As Double
    Dim Test As Variant

     
'Procedimiento ----
   
   'Link a UniSimDesign
   'Link a UniSim
    Set UniApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set SimCase = UniApp.ActiveDocument
    If SimCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
    'Link a FP
    Set BasisManager = SimCase.BasisManager
    Set FluidPackages = BasisManager.FluidPackages
     
    Set FluidPackage = FluidPackages.Item(0)
    
    'MsgBox FluidPackage.PropertyPackageName
    
    'Apagar solver de Unisim
   ' SimCase.Solver.CanSolve = False
    
    'Cambios de basis
    'SimCase.BasisManager.StartBasisChange
              
    Set Uniquac = FluidPackage.PropertyPackage
    Set ActParam(2, 0) = Uniquac.Aij(2, 0)
    
    MsgBox ActParam(2, 0)
    Uniquac.Aij(2, 0) = 1500
    
    'SimCase.BasisManager.EndBasisChange
  
   
   'Apagar solver de Unisim
   '     SimCase.Solver.CanSolve = True
      
     
     
    
    
        
'Leer valores

End Sub
" "Option Explicit

Public Sub CoefDeActividad()



'Variables
    Dim UniApp As UniSimDesign.Application
    Dim SimCase As UniSimDesign.SimulationCase
    Dim Fondo As UniSimDesign.ProcessStream
    Dim Formaldehido As UniSimDesign.ProcessStream
    Dim Agua As UniSimDesign.ProcessStream
    
    
    
    Dim CompMassFlow As Variant
    Dim j, i, k As Integer
    Dim UnitSet(1) As String
    Dim CaudalComp(2) As Double
    
    Dim ColT, ColF, ColW As String
    Dim CeldaT, CeldaF, CeldaW As String
    Dim FinalRow As Integer
       
    
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
    UnitSet(0) = ActiveSheet.Range(""C43"") 'Caudal masico kg/h
    UnitSet(1) = ActiveSheet.Range(""A59"") 'Temp C
    
   'Link a corrientes
  Set Fondo = SimCase.Flowsheet.MaterialStreams.Item(""Fondo"")
  Set Formaldehido = SimCase.Flowsheet.MaterialStreams.Item(""Formaldehido"")
  Set Agua = SimCase.Flowsheet.MaterialStreams.Item(""Agua"")
  
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
    
    FinalRow = ActiveSheet.Range(""D55"").Value
    ColT = ""A""
    ColF = ""C""
    ColW = ""D""
    
    
For i = 1 To FinalRow
    k = i + 59
    CeldaT = ColT & k
    CeldaF = ColF & k
    CeldaW = ColW & k
    
    'Llevar data a UniSim
   Formaldehido.Temperature.SetValue ActiveSheet.Range(CeldaT).Value, UnitSet(1)
   Agua.MassFlow.SetValue 150, UnitSet(0)
   
    
    
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = True
    
  'Extraer datos
  CompMassFlow = Fondo.ComponentMassFlow
  For j = 0 To UBound(CompMassFlow)
  CaudalComp(j) = CompMassFlow(j) * 3599.956034
  Next
  
  'Llevar a export sheet
  ActiveSheet.Range(""C40"").Value = CaudalComp(0)
    ActiveSheet.Range(CeldaF).Value = CaudalComp(0)
  
  ActiveSheet.Range(""D40"").Value = CaudalComp(1)
 ActiveSheet.Range(CeldaW).Value = CaudalComp(1)
  
Next
    
End Sub" "Option Explicit
Public Sub AntoineLit()


Dim CeldaChange, CeldaGoal As String
Dim k As Integer

For k = 8 To 106
CeldaChange = ActiveSheet.Range(""Y2"") & k
CeldaGoal = ActiveSheet.Range(""Y1"") & k

Range(CeldaGoal).GoalSeek Goal:=0, ChangingCell:=Range(CeldaChange)
Next k

End Sub" "Private Sub CommandButton1_Click()
Call TXCuaternaria1

End Sub" "Private Sub CommandButton1_Click()
Call AntoineLit

End Sub

Private Sub CommandButton2_Click()
Call TXUnisim1

End Sub" "Option Explicit
Public Sub TXUnisim1()

'Variables
    Dim UniApp As UniSimDesign.Application
    Dim SimCase As UniSimDesign.SimulationCase
    Dim Mezcla As UniSimDesign.ProcessStream
    Dim Formaldehido As UniSimDesign.ProcessStream
    Dim Agua As UniSimDesign.ProcessStream
    
    Dim CeldaAgua, CeldaT As String
    
    Dim i, j As Integer
    Dim UnitSet(1) As String
    
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
    UnitSet(0) = ActiveSheet.Range(""H7"") 'Caudal molar kgmole/h
    UnitSet(1) = ActiveSheet.Range(""I7"") 'Temp C
    
     'Link a corrientes
  Set Mezcla = SimCase.Flowsheet.MaterialStreams.Item(""Mezcla1"")
  Set Formaldehido = SimCase.Flowsheet.MaterialStreams.Item(""Formaldehido"")
  Set Agua = SimCase.Flowsheet.MaterialStreams.Item(""Agua"")
    
    
        
   For i = 8 To 106
   CeldaAgua = ""H"" & i
   CeldaT = ""I"" & i
   
   'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
   
   'Llevar data a UniSim
   Agua.MolarFlow.SetValue ActiveSheet.Range(CeldaAgua), UnitSet(0)
   
   'Apagar solver de Unisim
        SimCase.Solver.CanSolve = True
   
   
   'Extraer datos
    ActiveSheet.Range(CeldaT) = Mezcla.Temperature.GetValue(UnitSet(1))
    Next i
    

End Sub" "Option Explicit
Public Sub TXCuaternaria1()

Dim i As Integer
 
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 Dim Form2 As UniSimDesign.ProcessStream
 Dim MeOH2 As UniSimDesign.ProcessStream
 Dim Methy2 As UniSimDesign.ProcessStream
 Dim Agua2 As UniSimDesign.ProcessStream
 Dim Mezcla2 As UniSimDesign.ProcessStream
 
 
 Dim CeldaForm, CeldaMeOH, CeldaMethy, CeldaAgua, CeldaT As String
 Dim UnitSet(1) As String
 
 
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
    UnitSet(0) = ActiveSheet.Range(""L1"") 'Caudal molar kgmole/h
    UnitSet(1) = ActiveSheet.Range(""N4"") 'Temp K
    
    'Link a corrientes
  Set Form2 = SimCase.Flowsheet.MaterialStreams.Item(""Form2"")
  Set MeOH2 = SimCase.Flowsheet.MaterialStreams.Item(""MeOH2"")
  Set Methy2 = SimCase.Flowsheet.MaterialStreams.Item(""Methy2"")
  Set Agua2 = SimCase.Flowsheet.MaterialStreams.Item(""Agua2"")
  Set Mezcla2 = SimCase.Flowsheet.MaterialStreams.Item(""2"")
  
  For i = 5 To 44
  CeldaForm = ""J"" & i
  CeldaMeOH = ""K"" & i
  CeldaMethy = ""L"" & i
  CeldaAgua = ""M"" & i
  CeldaT = ""N"" & i
  
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
 
 'Llevar data a UniSim
   Form2.MolarFlow.SetValue ActiveSheet.Range(CeldaForm), UnitSet(0)
   MeOH2.MolarFlow.SetValue ActiveSheet.Range(CeldaMeOH), UnitSet(0)
   Methy2.MolarFlow.SetValue ActiveSheet.Range(CeldaMethy), UnitSet(0)
   Agua2.MolarFlow.SetValue ActiveSheet.Range(CeldaAgua), UnitSet(0)
 
 'Apagar solver de Unisim
        SimCase.Solver.CanSolve = True
        
       
 'Extraer datos
    ActiveSheet.Range(CeldaT) = Mezcla2.Temperature.GetValue(UnitSet(1))
    
    Next i
 
 
End Sub" "Private Sub CommandButton1_Click()
Call Antoine1erR

End Sub" "Option Explicit
Public Sub Antoine1erR()


Dim CeldaChange, CeldaGoal As String
Dim k As Integer

For k = 10 To 110
CeldaChange = ActiveSheet.Range(""H2"") & k
CeldaGoal = ActiveSheet.Range(""H3"") & k

Range(CeldaGoal).GoalSeek Goal:=0, ChangingCell:=Range(CeldaChange)
Next k

End Sub"
"Private Sub CommandButton2_Click()
Call CoefDeActividad

End Sub" "Option Explicit

Public Sub CoefDeAct()

'Variables
    Dim UniApp As UniSimDesign.Application
    Dim SimCase As UniSimDesign.SimulationCase
    
    
' Paquete de propiedades
    Dim BasisManager As UniSimDesign.BasisManager
    Dim FluidPackage As UniSimDesign.FluidPackage
    Dim FluidPackages As UniSimDesign.FluidPackages
    Dim PropFluidPack As UniSimDesign.PropPackage
    Dim Uniquac As UniSimDesign.UNIQUACPropPkg
    Dim ActParam(3, 3) As UniSimDesign.ObjectVariable
    
    
    
    
'Variables auxiliares
    Dim j As Integer
    Dim MethForm As Double
    Dim Test As Variant

     
'Procedimiento ----
   
   'Link a UniSimDesign
   'Link a UniSim
    Set UniApp = GetObject(, ""UniSimDesign.Application"")    'Solo Funciona si UniSim está abierto
    'Tomar datos de caso abierto actual
    Set SimCase = UniApp.ActiveDocument
    If SimCase Is Nothing Then
        MsgBox ""Una Simulación de UniSim debe estar abierta.""
        Exit Sub
    End If
    
    'Link a FP
    Set BasisManager = SimCase.BasisManager
    Set FluidPackages = BasisManager.FluidPackages
     
    Set FluidPackage = FluidPackages.Item(0)
    
    'MsgBox FluidPackage.PropertyPackageName
    
    'Apagar solver de Unisim
   ' SimCase.Solver.CanSolve = False
    
    'Cambios de basis
    'SimCase.BasisManager.StartBasisChange
              
    Set Uniquac = FluidPackage.PropertyPackage
    Set ActParam(2, 0) = Uniquac.Aij(2, 0)
    
    MsgBox ActParam(2, 0)
    Uniquac.Aij(2, 0) = 1500
    
    'SimCase.BasisManager.EndBasisChange
  
   
   'Apagar solver de Unisim
   '     SimCase.Solver.CanSolve = True
      
     
     
    
    
        
'Leer valores

End Sub
" "Option Explicit

Public Sub CoefDeActividad()



'Variables
    Dim UniApp As UniSimDesign.Application
    Dim SimCase As UniSimDesign.SimulationCase
    Dim Fondo As UniSimDesign.ProcessStream
    Dim Formaldehido As UniSimDesign.ProcessStream
    Dim Agua As UniSimDesign.ProcessStream
    
    
    
    Dim CompMassFlow As Variant
    Dim j, i, k As Integer
    Dim UnitSet(1) As String
    Dim CaudalComp(2) As Double
    
    Dim ColT, ColF, ColW As String
    Dim CeldaT, CeldaF, CeldaW As String
    Dim FinalRow As Integer
       
    
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
    UnitSet(0) = ActiveSheet.Range(""C43"") 'Caudal masico kg/h
    UnitSet(1) = ActiveSheet.Range(""A59"") 'Temp C
    
   'Link a corrientes
  Set Fondo = SimCase.Flowsheet.MaterialStreams.Item(""Fondo"")
  Set Formaldehido = SimCase.Flowsheet.MaterialStreams.Item(""Formaldehido"")
  Set Agua = SimCase.Flowsheet.MaterialStreams.Item(""Agua"")
  
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
    
    FinalRow = ActiveSheet.Range(""D55"").Value
    ColT = ""A""
    ColF = ""C""
    ColW = ""D""
    
    
For i = 1 To FinalRow
    k = i + 59
    CeldaT = ColT & k
    CeldaF = ColF & k
    CeldaW = ColW & k
    
    'Llevar data a UniSim
   Formaldehido.Temperature.SetValue ActiveSheet.Range(CeldaT).Value, UnitSet(1)
   Agua.MassFlow.SetValue 150, UnitSet(0)
   
    
    
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = True
    
  'Extraer datos
  CompMassFlow = Fondo.ComponentMassFlow
  For j = 0 To UBound(CompMassFlow)
  CaudalComp(j) = CompMassFlow(j) * 3599.956034
  Next
  
  'Llevar a export sheet
  ActiveSheet.Range(""C40"").Value = CaudalComp(0)
    ActiveSheet.Range(CeldaF).Value = CaudalComp(0)
  
  ActiveSheet.Range(""D40"").Value = CaudalComp(1)
 ActiveSheet.Range(CeldaW).Value = CaudalComp(1)
  
Next
    
End Sub" "Option Explicit
Public Sub AntoineLit()


Dim CeldaChange, CeldaGoal As String
Dim k As Integer

For k = 8 To 106
CeldaChange = ActiveSheet.Range(""Y2"") & k
CeldaGoal = ActiveSheet.Range(""Y1"") & k

Range(CeldaGoal).GoalSeek Goal:=0, ChangingCell:=Range(CeldaChange)
Next k

End Sub" "Private Sub CommandButton1_Click()
Call TXCuaternaria1

End Sub" "Private Sub CommandButton1_Click()
Call AntoineLit

End Sub

Private Sub CommandButton2_Click()
Call TXUnisim1

End Sub" "Option Explicit
Public Sub TXUnisim1()

'Variables
    Dim UniApp As UniSimDesign.Application
    Dim SimCase As UniSimDesign.SimulationCase
    Dim Mezcla As UniSimDesign.ProcessStream
    Dim Formaldehido As UniSimDesign.ProcessStream
    Dim Agua As UniSimDesign.ProcessStream
    
    Dim CeldaAgua, CeldaT As String
    
    Dim i, j As Integer
    Dim UnitSet(1) As String
    
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
    UnitSet(0) = ActiveSheet.Range(""H7"") 'Caudal molar kgmole/h
    UnitSet(1) = ActiveSheet.Range(""I7"") 'Temp C
    
     'Link a corrientes
  Set Mezcla = SimCase.Flowsheet.MaterialStreams.Item(""Mezcla1"")
  Set Formaldehido = SimCase.Flowsheet.MaterialStreams.Item(""Formaldehido"")
  Set Agua = SimCase.Flowsheet.MaterialStreams.Item(""Agua"")
    
    
        
   For i = 8 To 106
   CeldaAgua = ""H"" & i
   CeldaT = ""I"" & i
   
   'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
   
   'Llevar data a UniSim
   Agua.MolarFlow.SetValue ActiveSheet.Range(CeldaAgua), UnitSet(0)
   
   'Apagar solver de Unisim
        SimCase.Solver.CanSolve = True
   
   
   'Extraer datos
    ActiveSheet.Range(CeldaT) = Mezcla.Temperature.GetValue(UnitSet(1))
    Next i
    

End Sub" "Option Explicit
Public Sub TXCuaternaria1()

Dim i As Integer
 
 Dim UniApp As UniSimDesign.Application
 Dim SimCase As UniSimDesign.SimulationCase
 
 Dim Form2 As UniSimDesign.ProcessStream
 Dim MeOH2 As UniSimDesign.ProcessStream
 Dim Methy2 As UniSimDesign.ProcessStream
 Dim Agua2 As UniSimDesign.ProcessStream
 Dim Mezcla2 As UniSimDesign.ProcessStream
 
 
 Dim CeldaForm, CeldaMeOH, CeldaMethy, CeldaAgua, CeldaT As String
 Dim UnitSet(1) As String
 
 
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
    UnitSet(0) = ActiveSheet.Range(""L1"") 'Caudal molar kgmole/h
    UnitSet(1) = ActiveSheet.Range(""N4"") 'Temp K
    
    'Link a corrientes
  Set Form2 = SimCase.Flowsheet.MaterialStreams.Item(""Form2"")
  Set MeOH2 = SimCase.Flowsheet.MaterialStreams.Item(""MeOH2"")
  Set Methy2 = SimCase.Flowsheet.MaterialStreams.Item(""Methy2"")
  Set Agua2 = SimCase.Flowsheet.MaterialStreams.Item(""Agua2"")
  Set Mezcla2 = SimCase.Flowsheet.MaterialStreams.Item(""2"")
  
  For i = 5 To 44
  CeldaForm = ""J"" & i
  CeldaMeOH = ""K"" & i
  CeldaMethy = ""L"" & i
  CeldaAgua = ""M"" & i
  CeldaT = ""N"" & i
  
  'Apagar solver de Unisim
        SimCase.Solver.CanSolve = False
 
 'Llevar data a UniSim
   Form2.MolarFlow.SetValue ActiveSheet.Range(CeldaForm), UnitSet(0)
   MeOH2.MolarFlow.SetValue ActiveSheet.Range(CeldaMeOH), UnitSet(0)
   Methy2.MolarFlow.SetValue ActiveSheet.Range(CeldaMethy), UnitSet(0)
   Agua2.MolarFlow.SetValue ActiveSheet.Range(CeldaAgua), UnitSet(0)
 
 'Apagar solver de Unisim
        SimCase.Solver.CanSolve = True
        
       
 'Extraer datos
    ActiveSheet.Range(CeldaT) = Mezcla2.Temperature.GetValue(UnitSet(1))
    
    Next i
 
 
End Sub" "Private Sub CommandButton1_Click()
Call Antoine1erR

End Sub" "Option Explicit
Public Sub Antoine1erR()


Dim CeldaChange, CeldaGoal As String
Dim k As Integer

For k = 10 To 110
CeldaChange = ActiveSheet.Range(""H2"") & k
CeldaGoal = ActiveSheet.Range(""H3"") & k

Range(CeldaGoal).GoalSeek Goal:=0, ChangingCell:=Range(CeldaChange)
Next k

End Sub"
