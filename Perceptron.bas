Attribute VB_Name = "Perceptron"
Sub Estimate_CP_Min_Sum_Error()
Attribute Estimate_CP_Min_Sum_Error.VB_ProcData.VB_Invoke_Func = " \n14"

    'ResetSolver
    SolverReset

    'Switch to Neural Network sheet
    Sheets("ANN").Select

    'Run Solver to Minimize Sum of Absolute Error
    SolverOptions MaxTime:=0, Iterations:=0, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=True, AssumeNonNeg:=False, Derivatives:=1
    SolverOptions PopulationSize:=750, RandomSeed:=0, MutationRate:=0.075, MultiStart _
        :=True, RequireBounds:=False, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=1, SolveWithout:=False, MaxTimeNoImp:=30
    SolverOk SetCell:="$AX$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$G$2:$AJ$2,$AP$2:$AT$2", Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve UserFinish:=True

    'Return to Prediction Sheet
    Sheets("Next-Day Prediction").Select

End Sub

Sub Estimate_CP_Min_Max_Error()

    'ResetSolver
    SolverReset

    'Switch to Neural Network sheet
    Sheets("ANN").Select

    'Run Solver to Minimize Max Absolute Error
    SolverOptions MaxTime:=0, Iterations:=0, Precision:=0.000001, Convergence:= _
        0.0001, StepThru:=False, Scaling:=True, AssumeNonNeg:=False, Derivatives:=1
    SolverOptions PopulationSize:=750, RandomSeed:=0, MutationRate:=0.075, MultiStart _
        :=True, RequireBounds:=False, MaxSubproblems:=0, MaxIntegerSols:=0, _
        IntTolerance:=1, SolveWithout:=False, MaxTimeNoImp:=30
    SolverOk SetCell:="$AY$2", MaxMinVal:=2, ValueOf:=0, ByChange:= _
        "$G$2:$AJ$2,$AP$2:$AT$2", Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve UserFinish:=True

    'Return to Prediction Sheet
    Sheets("Next-Day Prediction").Select

End Sub

