Option Explicit

Public Function GetCustomUI(ByVal RibbonID As String) As String
    GetCustomUI = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" & _
                  "<ribbon><tabs><tab id='AuditXcelTab' label='AuditXcel AI'>" & _
                  "<group id='DataGroup' label='Data Tools'>" & _
                  "<button id='CleanButton' label='Clean Data' size='large' onAction='RibbonCallbacks.OnCleanData' />" & _
                  "<button id='AdvancedDupButton' label='Advanced Duplicates' size='large' onAction='RibbonCallbacks.OnAdvancedDuplicates' />" & _
                  "<button id='JoinAppendButton' label='Join/Append Data' size='large' onAction='RibbonCallbacks.OnJoinAppendData' />" & _
                  "</group>" & _
                  "<group id='AnalysisGroup' label='Analysis Tools'>" & _
                  "<button id='SummaryButton' label='Basic Summary' size='large' onAction='RibbonCallbacks.OnSummaryStats' />" & _
                  "<button id='AdvSummaryButton' label='Advanced Summary' size='large' onAction='RibbonCallbacks.OnAdvancedSummary' />" & _
                  "<button id='FraudButton' label='Fraud Detection' size='large' onAction='RibbonCallbacks.OnFraudDetection' />" & _
                  "</group>" & _
                  "<group id='ReportGroup' label='Reporting'>" & _
                  "<button id='ReportButton' label='Generate Report' size='large' onAction='RibbonCallbacks.OnGenerateReport' />" & _
                  "</group>" & _
                  "</tab></tabs></ribbon></customUI>"
End Function
