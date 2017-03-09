Attribute VB_Name = "Logistic"
' Logistic Regression
'
'
'
Option Explicit

Function Main()
    Dim OriginalWs As Worksheet, Ws As Worksheet
    Dim Predictor As Range, Objective As Range
    Dim N As Integer, P As Integer
    Dim rc As Integer, i As Integer

    
    MsgBox "Logistic!"
    
    Set OriginalWs = ActiveSheet
    Set Ws = Util.AddSummarySheet(After:=OriginalWs)
    
    With Ws
        .Rows(1).Insert
        .Cells(1, 1).Value = "Coefficient"
        .Range(.Cells(1, 2), .Cells(1, .UsedRange.Columns.Count)).Value = 0
        Set Predictor = .Range(.Cells(3, 2), .Cells(.UsedRange.Rows.Count, .UsedRange.Columns.Count))
        Set Objective = .Range(.Cells(3, 1), .Cells(.UsedRange.Rows.Count, 1))
        N = Predictor.Rows.Count
        P = Predictor.Columns.Count - 1
        
        .Cells(2, .UsedRange.Columns.Count + 1).Value = "yhat"
        With .Cells(3, .UsedRange.Columns.Count)
            .Formula = "=1/(1+EXP(-SUMPRODUCT(" & Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, P + 2)).Address _
                                                                             & "," _
                                                                             & Ws.Range(Ws.Cells(3, 2), Ws.Cells(3, P + 2)).Address(RowAbsolute:=False, ColumnAbsolute:=False) _
                                                                             & ")))"
            .AutoFill Destination:=Ws.Range(Ws.Cells(3, P + 3), Ws.Cells(N + 2, P + 3))
        End With
        
        .Cells(2, .UsedRange.Columns.Count + 1).Value = "loglikelihood"
        With .Cells(3, .UsedRange.Columns.Count)
            .Formula = "=A3*LN(" & .Offset(0, -1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")+(1-A3)*LN(1-" & .Offset(0, -1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            .AutoFill Destination:=Ws.Range(Ws.Cells(3, P + 4), Ws.Cells(N + 2, P + 4))
        End With
                          
       .Cells(1, P + 3).Value = "Sum of loglikelihood"
       .Cells(1, P + 4).Formula = "=-2*SUM(" & .Range(.Cells(3, P + 4), .Cells(N + 2, P + 4)).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
    End With
    
    rc = MsgBox("変数の標準化を行いますか？", vbYesNo + vbQuestion, "確認")
    If rc = vbYes Then
        For i = 3 To P + 2
            Set Predictor.Columns(i) = Util.Normalize(Predictor.Columns(i))
        Next i
    End If
End Function

Function CostFunction(y As Long, x() As Long)

End Function
