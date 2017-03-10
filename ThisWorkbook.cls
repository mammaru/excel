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
    cbcMenu.Caption = "���v"
    
    Set cbcRegression = cbcMenu.Controls.Add(Type:=msoControlPopup)
    cbcRegression.Caption = "��A"
    ' Logistic Regression
    ' Module: Logistic
    With cbcRegression.Controls.Add(Type:=msoControlButton)
        .Caption = "�Q�l���W�X�e�B�b�N��A"
        .OnAction = "Logistic.Main"
    End With
    
    Set cbcUtil = cbcMenu.Controls.Add(Type:=msoControlPopup)
    cbcUtil.Caption = "���[�e�B���e�B"
    With cbcUtil.Controls.Add(Type:=msoControlButton)
        .Caption = "�Z�����l��"
        .OnAction = "Util.StrToDbl"
    End With
    
    ' Alert Initialize Stat Addin
    ReadMe
End Sub

Private Sub Workbook_AddinUninstall()
    On Error GoTo ErrHand
    With Application.CommandBars("Worksheet Menu Bar")
        .Controls("���v").Delete
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
    MsgBox "���v�p�A�h�C����ǉ����܂����B"
End Sub