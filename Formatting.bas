Attribute VB_Name = "Formatting"
'---------------------------------DECLARE FORMAT MODULE------------------------------------

'USED TO FORMAT WORK DATE
Function Format_Date(WORK_DATE, DATE_FORMAT)
    'Set workdate today and call format string yyy-MM-dd
    Format_Date = Format(WORK_DATE, DATE_FORMAT)
End Function

'ADD ANY ADDITIONAL FORMATTING FUNCTIONS NEEDED HERE...
