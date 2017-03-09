Attribute VB_Name = "Util"
Function AddSummarySheet(After As Worksheet, Optional Name As String = "", Optional Residual As Boolean = True) As Worksheet
    Dim Ws As Worksheet, Target As Range
    Dim i As Integer, N As Integer, flag As Boolean
    Set Ws = Worksheets.Add(After:=After)
    If Not Name = "" Then
        Ws.Name = Name
    End If
    
    After.Activate
    Set Target = Application.InputBox("目的変数の範囲を選択してください", _
                                                          Title:="Logistic", _
                                                          Type:=8)
    N = Target.Rows.Count
    Target.Copy Ws.Range("A2")
    Ws.Range("A1").Value = "y"
    
    flag = True
    Do While flag
        Set Target = Application.InputBox("従属変数の範囲を選択してください", _
                                                              Title:="Logistic", _
                                                              Type:=8)
        If Target.Rows.Count <> N Then
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

Function Normalize(Data As Range) As Range
    Dim mu As Double, sigma As Double, i As Integer, c As Variant
    'tmp = Data
    mu = CDbl(WorksheetFunction.Average(Data))
    'sigma = Sqr(WorksheetFunction.var(tmp))
    sigma = CDbl(WorksheetFunction.StDev(Data))
    'For i = 1 To Data.Rows.Count
    For Each c In Data
        'Data.Item(i).Value = WorksheetFunction.Standardize(Data.Item(i).Value, mu, sigma)
        c.Value = WorksheetFunction.Standardize(CDbl(c.Value), mu, sigma)
    Next
    Normalize = Data
End Function
