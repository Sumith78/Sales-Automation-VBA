Attribute VB_Name = "Module4"
Sub Run_Full_Automation()

    On Error GoTo HandleError

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Call Validate_And_Clean_Transactions
    Call Calculate_Trade_Metrics
    Call Generate_Exposure_Report
    Call Refresh_Dashboard

    Call Write_Audit_Log("SUCCESS")

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

HandleError:
    Call Write_Audit_Log("FAILED")
    MsgBox "Automation failed. Check Audit_Log.", vbCritical
    Resume Cleanup

End Sub

