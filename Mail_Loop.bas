Attribute VB_Name = "Mail_Loop"
'---------------------------------DECLARE MAIL LOOP MODULE------------------------------------

'Loops through files in Output to e-mail out to customers
'Uses Update_Progress_Bar to show live loop progress
Sub Loop_Through_Output(BUSINESS_EMAIL, LOCATION, FILETYPE, USER, PASSWORD, DEFAULT_USER, DEFAULT_PASSWORD)
    PATH = Dir(LOCATION & "\*" & FILETYPE)
    Do While Len(PATH) > 0
        PATH_ARRAY = Split(PATH, "_")
        If (PATH <> "" And InStr(PATH, "@") > 0) Then
            COMPANY_NAME = PATH_ARRAY(0)
            CUSTOMER_NAME = PATH_ARRAY(3)
            INVOICE_NO = PATH_ARRAY(5)
            CUSTOMER_EMAIL = PATH_ARRAY(7)
            DUE_DATE = Replace(PATH_ARRAY(9), ".xlsx", "")
        End If
        If (COMPANY_NAME <> "" And CUSTOMER_NAME <> "" And INVOICE_NO <> "" And CUSTOMER_EMAIL <> "" And DUE_DATE <> "" And LOCATION <> "" And PATH <> "" And InStr(PATH, "@") > 0) Then
            Send USER, PASSWORD, DEFAULT_USER, DEFAULT_PASSWORD, "Start", 2, "smtp.gmail.com", 1, 465, True, BUSINESS_EMAIL, CUSTOMER_EMAIL, COMPANY_NAME & " Invoice - " & INVOICE_NO, CUSTOMER_NAME & ". Find your invoice attached to this e-mail. This invoice is due on " & DUE_DATE & " .", LOCATION & "\" & PATH
        End If
        'Update progress bar
        Update_Progress_Bar Replace(UserForm1.FrameProgress.Caption, "%", "")
        PATH = Dir
    Loop
End Sub

'ADD ANY ADDITIONAL MAIL_LOOP FUNCTIONS OR SUBROUTINES NEEDED HERE...
