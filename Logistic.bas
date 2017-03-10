Attribute VB_Name = "Logistic"
' Logistic Regression
'
'
'
Option Explicit

Function Main()
    Dim OriginalWs As Worksheet, Ws As Worksheet
    Dim Predictor As Range, Objective As Range
    Dim Parameter As Range, Cost As Range
    Dim N As Integer, P As Integer
    Dim rc As Integer, i As Integer

    rc = MsgBox("ロジスティック回帰を行います。", vbYesNo + vbQuestion, "確認")
    If rc = vbNo Then
        MsgBox "処理がキャンセルされました。"
        End
    End If
    
    Set OriginalWs = ActiveSheet
    Set Ws = Util.AddSummarySheet(After:=OriginalWs)
    
    With Ws
        .Rows(1).Insert
        .Cells(1, 1).Value = "回帰係数："
        Set Parameter = .Range(.Cells(1, 2), .Cells(1, .UsedRange.Columns.Count))
        Parameter.Value = 0
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
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0.5"
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ColorIndex = 3
            End With
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0.5"
            With .FormatConditions(2).Interior
                .PatternColorIndex = xlAutomatic
                .ColorIndex = 5
            End With
            .AutoFill Destination:=Ws.Range(Ws.Cells(3, P + 3), Ws.Cells(N + 2, P + 3))
        End With
        
        .Cells(2, .UsedRange.Columns.Count + 1).Value = "対数尤度"
        'Set Cost = .Cells(3, .UsedRange.Columns.Count)
        With .Cells(3, .UsedRange.Columns.Count)
            .Formula = "=A3*LN(" & .Offset(0, -1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")+(1-A3)*LN(1-" & .Offset(0, -1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            .AutoFill Destination:=Ws.Range(Ws.Cells(3, P + 4), Ws.Cells(N + 2, P + 4))
        End With
                          
        .Cells(1, P + 3).Value = "対数尤度合計："
        Set Cost = .Cells(1, P + 4)
        Cost.Formula = "=-2*SUM(" & .Range(.Cells(3, P + 4), .Cells(N + 2, P + 4)).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
    End With
    
    Ws.Activate
    rc = MsgBox("独立変数の正規化を行いますか？", vbYesNo + vbQuestion, "確認")
    If rc = vbYes Then
        For i = 2 To P + 1
            'MsgBox Predictor.Columns(i).Address
            Util.Normalize Predictor.Columns(i)
        Next i
    End If
    
    MsgBox "最適化要件：対数尤度合計値を最小にする回帰係数を選択する。" & vbLf & _
                  "パラメータ：回帰係数（" & Parameter.Address & "）" & vbLf & _
                  "コスト関数：対数尤度合計値（" & Cost.Address & "）" _
                  , _
                  , "ソルバーを用いて最適化を行ってください。"
End Function

Function CostFunction(y As Long, x() As Long)

End Function
