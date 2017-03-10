Attribute VB_Name = "Util"
Function AddSummarySheet(After As Worksheet, Optional Name As String = "", Optional Residual As Boolean = True) As Worksheet
    On Error Resume Next
    Dim Ws As Worksheet, Target As Range
    Dim i As Integer, N As Integer, flag As Boolean
    Set Ws = Worksheets.Add(After:=After)
    If Not Name = "" Then
        Ws.Name = Name
    End If
    
    After.Activate
    Set Target = Application.InputBox("従属変数の範囲を選択してください", _
                                                          Title:="Logistic", _
                                                          Type:=8)
    If Err.Number > 0 Then
        MsgBox "処理がキャンセルされました。", vbExclamation
        End
    End If
    N = Target.Rows.Count
    Target.Copy Ws.Range("A2")
    Ws.Range("A1").Value = "y"
    
    flag = True
    Do While flag
        Set Target = Application.InputBox("独立変数の範囲を選択してください", _
                                                              Title:="Logistic", _
                                                              Type:=8)
        If Err.Number > 0 Then
            MsgBox "処理がキャンセルされました。", vbExclamation
            End
        ElseIf Target.Rows.Count <> N Then
            MsgBox "サンプル数が一致しません"
        Else
            flag = False
        End If
    Loop
    Target.Copy Ws.Range("B2")
    For i = 1 To Target.Columns.Count
      Ws.Cells(1, i + 1).Value = "x" & i
    Next
    
    If Residual Then
        Ws.Columns(2).Insert
        Ws.Cells(1, 2) = "x0"
        Ws.Range(Ws.Cells(2, 2), Ws.Cells(Target.Rows.Count + 1, 2)) = 1
    End If
    
    Set AddSummarySheet = Ws
End Function

Sub Normalize(ByVal Data As Range)
    Dim mu As Double, sigma As Double, i As Integer, j As Integer, c As Variant
    'tmp = Data
    mu = CDbl(WorksheetFunction.Average(Data))
    'sigma = Sqr(WorksheetFunction.var(tmp))
    sigma = CDbl(WorksheetFunction.StDev(Data))
    'For Each c In Data
    For i = 1 To Data.Rows.Count
        For j = 1 To Data.Columns.Count
            'MsgBox Data.Cells(i, j).Address & Data.Cells(i, j).Value
            Data.Cells(i, j).Value = WorksheetFunction.Standardize(Data.Cells(i, j).Value, mu, sigma)
           ' c.Cells() = WorksheetFunction.Standardize(CDbl(c.Value), mu, sigma)
        Next j
    Next i
End Sub

Sub StrToDbl()
    Dim Rng As Range, Denominator As Double, Numerator As Double, tmp As Double
    Dim i As Integer, j As Integer, x As Variant, flag As Boolean
    If Not TypeName(Selection) = "Range" Then
        MsgBox "セルが選択されていません"
        Exit Sub
    End If
    For i = 1 To Selection.Rows.Count
        flag = True
        For j = 1 To Selection.Columns.Count
            x = Selection.Cells(i, j).Value
            If x = "" Then
                flag = False
            ElseIf IsNumeric(x) = True Then
                tmp = CDbl(x)
            Else
                x = CStr(x)
                If InStr(x, "/") <> 0 Then
                    If IsNumeric(Split(x, "/")(0)) And IsNumeric(Split(x, "/")(1)) Then
                        tmp = CDbl(Split(x, "/")(0)) / CDbl(Split(x, "/")(1))
                    Else
                        flag = False
                    End If
                Else
                    flag = False
                End If
            End If
            'MsgBox tmp
            If flag Then
                Selection.Cells(i, j).Value = tmp
            End If
        Next j
    Next i
End Sub
