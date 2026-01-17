Option Explicit

Public Sub OnCleanData(Control As IRibbonControl)
    Call MainModule.CleanData(Control)
End Sub

Public Sub OnSummaryStats(Control As IRibbonControl)
    Call MainModule.SummaryStats(Control)
End Sub

Public Sub OnAdvancedSummary(Control As IRibbonControl)
    Call MainModule.AdvancedSummary(Control)
End Sub

Public Sub OnAdvancedDuplicates(Control As IRibbonControl)
    Call MainModule.AdvancedDuplicates(Control)
End Sub

Public Sub OnJoinAppendData(Control As IRibbonControl)
    Call MainModule.JoinAppendData(Control)
End Sub

Public Sub OnGenerateReport(Control As IRibbonControl)
    Call MainModule.GenerateReport(Control)
End Sub

Public Sub OnFraudDetection(Control As IRibbonControl)
    Call MainModule.FraudDetection(Control)
End Sub
