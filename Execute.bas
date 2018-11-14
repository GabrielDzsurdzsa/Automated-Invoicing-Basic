Attribute VB_Name = "Execute"
'---------------------------------DECLARE EXECUTE MODULE------------------------------------

'PRIMARY EXECUTE SUBROUTINE
'CHAIN ALL OTHER SUBROUTINES AND FUNCTIONS IN SEQUENCE, BASED ON ORDER OF OPERATIONS
Sub Execute()
    Declare_Workbooks
    Declare_Work_Arrays
    Declare_Table_Objects
    Declare_Progress
    Declare_Alerts
    Declare_Cleanup
    Declare_Dates
    Declare_Folders
    Declare_Formatting
    Declare_Logging
    Declare_Mail
    Declare_Mail_Loop
    Declare_Work
On Error GoTo ErrorHandler
    Display_Notifications False
    WORK_DATE = Create_Date
    PATH = Check_Create_Output_Folder
    Set WB = ThisWorkbook
    LOG_FILE = Create_Log(PATH, "Automated_Invoicing_Basic")
    Open_Log LOG_FILE
    Write_To_Log Now & " - " & "Automated Invoicing Process Started!"
    Write_To_Log Now & " - " & "Invoice Data Mapping Started!"
    Set WS = WB.Sheets("Invoice Data")
    WS.Activate
    RETURN_ARRAY = Map_Invoice_Data(WS, "Table1")
    Write_To_Log Now & " - " & "Invoice Data Mapping Completed!"
    Write_To_Log Now & " - " & "Customer Invoice Building Started!"
    Set WS = WB.Sheets("Invoice Template")
    WS.Activate
    Build_Invoice PATH, WS, WB.Sheets("Invoice Data"), RETURN_ARRAY, "Table1"
    Write_To_Log Now & " - " & "Customer Invoice Building Completed!"
    Write_To_Log Now & " - " & "Invoice E-mail Delivery Started!"
    BUSINESS_EMAIL = WS.Range("A8").VALUE
    Loop_Through_Output BUSINESS_EMAIL, PATH, "xlsx", ActiveWorkbook.Sheets("Execute & Send").Range("J9"), ActiveWorkbook.Sheets("Execute & Send").Range("J11"), ActiveWorkbook.Sheets("Start").Range("BB156"), ActiveWorkbook.Sheets("Start").Range("BB157").VALUE
    Write_To_Log Now & " - " & "Invoice E-mail Delivery Completed!"
    Write_To_Log Now & " - " & "Process cleanup started!"
    CLEAN_ARRAY = Array("E5", "E6", "E7", "B10", "B11", "B12", "B13", "B14", "A17", "C17", "D17", "F17", "A20:E39")
    Clean WS, CLEAN_ARRAY
    Write_To_Log Now & " - " & "Invoice Data Summary Save Started!"
    Set WS = WB.Sheets("Invoice Data")
    WS.Activate
    Save_Summary WS, PATH, WORK_DATE
    Write_To_Log Now & " - " & "Invoice Data Summary Save Completed!"
    Delete_Table_Data "Invoice Data", "Table1"
    WB.Save
    Write_To_Log Now & " - " & "Process cleanup completed!"
    Write_To_Log Now & " - " & "Invoicing complete!"
    Display_Notifications True
    Alert "Process Complete. Review output and archive your files.", vbOK, "Gabrielcoder.ca - Automated Invoicing Basic"
    Unload UserForm1
ProcessExit:
    Unload UserForm1
    Exit Sub
ErrorHandler:
    Alert "The following error occured: " & Err.DESCRIPTION & " Exiting Subroutine!", vbOK, "Gabrielcoder.ca - Automated Invoicing Basic"
    Write_To_Log "The following error occured: " & Err.DESCRIPTION & " Exiting Subroutine!"
    Close_Log
    Resume ProcessExit
End Sub
