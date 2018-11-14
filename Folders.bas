Attribute VB_Name = "Folders"
'---------------------------------DECLARE FOLDER WORK MODULE------------------------------------

'CHECK IF OUTPUT FOLDER EXISTS IN ROOT
'Create output folder if it does not exist at the root of this document
Function Check_Create_Output_Folder()
    'Set working path (uses application root, so if you put this document on your desktop, it will create a folder called Output)
    PATH = Application.ActiveWorkbook.PATH & "\Output" ' We'll name our output folder as output
    If Dir(PATH, vbDirectory) = "" Then
        Shell ("cmd /c mkdir """ & PATH & """")
    End If
    Check_Create_Output_Folder = PATH
End Function

'ADD ANY ADDITIONAL FOLDER FUNCTIONS NEEDED HERE...


