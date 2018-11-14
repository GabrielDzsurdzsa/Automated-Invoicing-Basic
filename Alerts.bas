Attribute VB_Name = "Alerts"
'---------------------------------DECLARE ALERTS MODULE------------------------------------

'USED TO ENABLE / DISABLE EXCEL NOTIFICATIONS
'Pass test variable as true or false.
Sub Display_Notifications(TEST)
    Application.DisplayAlerts = TEST
    Application.ScreenUpdating = TEST
    'Maximize or minimize the excel window
    'Application.WindowState = xlMinimized ' Un-comment if you want to control minimized / maximized state of the Excel window.
End Sub

'USED TO DISPLAY ANY ALERT
'Pass your alert message, prompt_type (e.g. vbOK) and title.
Sub Alert(MESSAGE, PROMPT_TYPE, TITLE)
    MsgBox MESSAGE, PROMPT_TYPE, TITLE
End Sub

'USED TO SHOW PROGRESS
'pass your percentage completion as a integer or decimal.
Sub Update_Progress_Bar(PCTDONE)
    'Always should be UserForm1 (Unless edits are made to ProgressBar UserForm)
    With UserForm1
        .FrameProgress.Caption = PCTDONE 'Format(PCTDONE / 10, "0%")
        .LabelProgress.Width = PCTDONE * 2 '(.FrameProgress.Width)
    End With
    DoEvents
End Sub

'ADD ANY ADDITIONAL ALERT SUBROUTINES NEEDED HERE...
