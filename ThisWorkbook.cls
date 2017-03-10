VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AddinInstall()
    Dim cbrCmd As CommandBar
    Dim cbcMenu As CommandBarControl
    Dim cbcRegression As CommandBarControl, cbcUtil As CommandBarControl
    
    Set cbrCmd = Application.CommandBars("Worksheet Menu Bar")
    Set cbcMenu = cbrCmd.Controls.Add(Type:=msoControlPopup)
    cbcMenu.Caption = "統計"
    
    Set cbcRegression = cbcMenu.Controls.Add(Type:=msoControlPopup)
    cbcRegression.Caption = "回帰"
    ' Logistic Regression
    ' Module: Logistic
    With cbcRegression.Controls.Add(Type:=msoControlButton)
        .Caption = "２値ロジスティック回帰"
        .OnAction = "Logistic.Main"
    End With
    
    Set cbcUtil = cbcMenu.Controls.Add(Type:=msoControlPopup)
    cbcUtil.Caption = "ユーティリティ"
    With cbcUtil.Controls.Add(Type:=msoControlButton)
        .Caption = "セル数値化"
        .OnAction = "Util.StrToDbl"
    End With
    
    ' Alert Initialize Stat Addin
    ReadMe
End Sub

Private Sub Workbook_AddinUninstall()
    On Error GoTo ErrHand
    With Application.CommandBars("Worksheet Menu Bar")
        .Controls("統計").Delete
    End With
    Exit Sub

ErrHand:
    If Err.Number = 5 Then
       Resume Next
    Else
       MsgBox Err.Description
    End If
End Sub

Sub ReadMe()
    MsgBox "統計用アドインを追加しました。"
End Sub
