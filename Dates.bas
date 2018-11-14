Attribute VB_Name = "Dates"
'---------------------------------DECLARE DATES MODULE------------------------------------

'USED TO OUTPUT WORKING DATE / TODAY
Function Create_Date()
    'Set workdate by passing today's date and the format as string
    Create_Date = Format_Date(Date, "yyyy-MM-dd")
End Function

'ADD ANY ADDITIONAL DATE FUNCTIONS NEEDED HERE...
