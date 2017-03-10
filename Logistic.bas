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

    rc = MsgBox("���W�X�e�B�b�N��A���s���܂��B", vbYesNo + vbQuestion, "�m�F")
    If rc = vbNo Then
        MsgBox "�������L�����Z������܂����B"
        End
    End If
    
    Set OriginalWs = ActiveSheet
    Set Ws = Util.AddSummarySheet(After:=OriginalWs)
    
    With Ws
        .Rows(1).Insert
        .Cells(1, 1).Value = "��A�W���F"
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
        
        .Cells(2, .UsedRange.Columns.Count + 1).Value = "�ΐ��ޓx"
        'Set Cost = .Cells(3, .UsedRange.Columns.Count)
        With .Cells(3, .UsedRange.Columns.Count)
            .Formula = "=A3*LN(" & .Offset(0, -1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")+(1-A3)*LN(1-" & .Offset(0, -1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            .AutoFill Destination:=Ws.Range(Ws.Cells(3, P + 4), Ws.Cells(N + 2, P + 4))
        End With
                          
        .Cells(1, P + 3).Value = "�ΐ��ޓx���v�F"
        Set Cost = .Cells(1, P + 4)
        Cost.Formula = "=-2*SUM(" & .Range(.Cells(3, P + 4), .Cells(N + 2, P + 4)).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
    End With
    
    Ws.Activate
    rc = MsgBox("�Ɨ��ϐ��̐��K�����s���܂����H", vbYesNo + vbQuestion, "�m�F")
    If rc = vbYes Then
        For i = 2 To P + 1
            'MsgBox Predictor.Columns(i).Address
            Util.Normalize Predictor.Columns(i)
        Next i
    End If
    
    MsgBox "�œK���v���F�ΐ��ޓx���v�l���ŏ��ɂ����A�W����I������B" & vbLf & _
                  "�p�����[�^�F��A�W���i" & Parameter.Address & "�j" & vbLf & _
                  "�R�X�g�֐��F�ΐ��ޓx���v�l�i" & Cost.Address & "�j" _
                  , _
                  , "�\���o�[��p���čœK�����s���Ă��������B"
End Function

Function CostFunction(y As Long, x() As Long)

End Function
