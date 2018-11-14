Attribute VB_Name = "Logging"
'---------------------------------DECLARE LOG MODULE------------------------------------
'This module contains logging subroutines

'Create log file if it does not exist
Function Create_Log(PATH, FILE_NAME)
    LOG_FILE = PATH & "\" & FILE_NAME & ".log"
    Create_Log = LOG_FILE
End Function

'Open log file
Sub Open_Log(LOG_FILE)
    Open LOG_FILE For Output As #1
End Sub

'Write to log
Sub Write_To_Log(MESSAGE)
    Write #1, MESSAGE
End Sub

'Close the logging file
Sub Close_Log()
    Close #1
End Sub

'ADD ANY ADDITIONAL LOGGING FUNCTIONS OR SUBROUTINES NEEDED HERE...
