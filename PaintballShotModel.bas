Attribute VB_Name = "Module2"
'Must download and add reference to the "CoolProp" library
'https://coolprop.org/coolprop/wrappers/Excel/index.html

Sub RunModel()
    Dim ReturnVal() As Double
    Dim TankEquilibriumTemp As Double
    Dim VCTemp As Double
    Dim TargetVCPressure As Double
    Dim AtmoPressure As Double
    Dim TankMass As Double
    Dim TankCP As Double
    Dim VCVol As Double
    Dim TankVol As Double
    Dim GasTemp As Double
    Dim TankTemp As Double
    Dim GasCv As Double
    Dim Fluidname As String
    Dim MolarU1 As Double
    Dim MolarU2 As Double
    Dim n1 As Double
    Dim n2 As Double
    Dim S2 As Double 'volume chamber entropy
    Dim i As Long
    Dim TankP As Double
    Dim VCTempMax As Double
    
    Dim filePath As String
    filePath = "C:\Users\Compooter\Documents\ModelOutput.csv"
    Open filePath For Output As #1
    
    Fluidname = "HEOS::Nitrogen"
    
    TankTemp = 300                       'K
    TankP = 4500 * 101325 / 14.696       'Pa
    VCTemp = 300                         'K
    TargetVCPressure = 115 * 101325 / 14.696    'Pa
    AtmoPressure = 101325                       'Pa
    TankMass = 1                                'kg
    TankCP = 900                                'J/kg K
    VCVol = 2.01628984537637E-05                'm^3
    TankVol = 0.001261804                       'm^3
    
    MolarU1 = CoolProp.PropsSI("UMOLAR", "T", TankTemp, "P", TankP, Fluidname)
    MolarU2 = CoolProp.PropsSI("UMOLAR", "T", VCTemp, "P", AtmoPressure, Fluidname)
    n1 = TankVol * CoolProp.PropsSI("DMOLAR", "T", TankTemp, "P", TankP, Fluidname)
    n2 = VCVol * CoolProp.PropsSI("DMOLAR", "T", VCTemp, "P", AtmoPressure, Fluidname)
    
    For i = 1 To 4000
        
        ReturnVal() = ExpansionCalc(TargetVCPressure, MolarU1, n1, TankVol, MolarU2, n2, VCVol, Fluidname)
        'Debug.Print ReturnVal(1, 1) & ";" & ReturnVal(1, 2) & ";" & ReturnVal(2, 1) & ";" & ReturnVal(2, 2)
        
        n1 = ReturnVal(1, 1)
        n2 = ReturnVal(1, 2)
        MolarU1 = ReturnVal(2, 1)
        MolarU2 = ReturnVal(2, 2)
        
        GasTemp = CoolProp.PropsSI("T", "UMOLAR", MolarU1, "DMOLAR", n1 / TankVol, Fluidname)
        GasCv = CoolProp.PropsSI("CVMOLAR", "UMOLAR", MolarU1, "DMOLAR", n1 / TankVol, Fluidname)
        
        TankEquilibriumTemp = (n1 * GasTemp * GasCv + TankTemp * TankMass * TankCP) / (n1 * GasCv + TankMass * TankCP)
        'Debug.Print TankEquilibriumTemp
        
        'Reinitialize and repeat
        TankTemp = TankEquilibriumTemp
        MolarU1 = CoolProp.PropsSI("UMOLAR", "T", TankTemp, "Dmolar", n1 / TankVol, Fluidname)
        'Assume isentropic expansion of volume chamber after each shot
        VCTempMax = CoolProp.PropsSI("T", "UMOLAR", MolarU2, "DMOLAR", n2 / VCVol, Fluidname)
        S2 = CoolProp.PropsSI("SMOLAR", "UMOLAR", MolarU2, "DMOLAR", n2 / VCVol, Fluidname)
        MolarU2 = CoolProp.PropsSI("UMOLAR", "SMOLAR", S2, "P", AtmoPressure, Fluidname)
        n2 = CoolProp.PropsSI("DMOLAR", "SMOLAR", S2, "P", AtmoPressure, Fluidname) * VCVol
        
        'Calculate tank P, if less than volume chamber final pressure exit iterations
        TankP = CoolProp.PropsSI("P", "T", TankTemp, "DMOLAR", n1 / TankVol, Fluidname)
        VCTemp = CoolProp.PropsSI("T", "SMOLAR", S2, "P", AtmoPressure, Fluidname)
        
        Print #1, i & ", " & n1 & "," & n2 & "," & TankTemp & "," & TankP & "," & VCTempMax & ", " & VCTemp
        
        If TankP < TargetVCPressure Then Exit For
    
    Next i
    
    Debug.Print i & " shots"
    Close #1
End Sub


Function ExpansionCalc(TargetPressure As Double, U01 As Double, n01 As Double, V1 As Double, U02 As Double, n02 As Double, V2 As Double, Fluidname As String) As Double()
    'Pressure in Pa, molar step size in moles
    Dim ReturnVal(1 To 2, 1 To 2) As Double
    Dim Stepsize As Double
    Dim MolesTransferred As Double
    Dim MaxSteps As Long
    Dim MolarU1 As Double
    Dim U1i As Double
    Dim MolarH1 As Double
    Dim n1 As Double
    
    Dim MolarU2 As Double
    Dim U2i As Double
    Dim n2 As Double
    
    Dim PreviousStepPressure As Double
    Dim StepPressure As Double
    Dim Overshoot As Double
    
    Dim i As Long

    Stepsize = -0.0001 'moles transferred from System 1 to System 2 per step
    MaxSteps = 1000
    
    'Calculate initial U, H, and density for System 1 and System 2
    
    MolarU1 = U01
    n1 = n01
    
    MolarU2 = U02
    n2 = n02
        
    'Iterate and see if pressure exceeds target pressure
    For i = 1 To MaxSteps
        U1i = MolarU1 * n1
        U2i = MolarU2 * n2
 
        MolarH1 = CoolProp.PropsSI("HMOLAR", "UMOLAR", MolarU1, "DMOLAR", n1 / V1, Fluidname)
        n1 = n1 + Stepsize
        n2 = n2 - Stepsize
                     
        MolarU1 = (U1i + MolarH1 * Stepsize) / n1
        MolarU2 = (U2i - MolarH1 * Stepsize) / n2
        MolesTransferred = MolesTransferred + Stepsize
        PreviousStepPressure = StepPressure
        StepPressure = CoolProp.PropsSI("P", "UMOLAR", MolarU2, "DMOLAR", n2 / V2, Fluidname)
        
        
        If StepPressure > TargetPressure Then
            'Calculate overshoot, then go back proportional to the overshoot and recalculate final U1 and U2
            Overshoot = -Stepsize * (StepPressure - TargetPressure) / (StepPressure - PreviousStepPressure)

             n1 = n1 + Overshoot
             n2 = n2 - Overshoot
             MolarU1 = (U1i + MolarH1 * (Stepsize + Overshoot)) / n1
             MolarU2 = (U2i - MolarH1 * (Stepsize + Overshoot)) / n2
             MolesTransferred = MolesTransferred + Overshoot

            Exit For
        End If
    Next i
    
    If i < MaxSteps Then
        ReturnVal(1, 1) = n1
        ReturnVal(1, 2) = n2
        ReturnVal(2, 1) = MolarU1
        ReturnVal(2, 2) = MolarU2
        ExpansionCalc = ReturnVal
    Else
        Debug.Print "Max steps exceeded"
    End If

End Function
