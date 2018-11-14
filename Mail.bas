Attribute VB_Name = "Mail"
'---------------------------------DECLARE MAIL MODULE------------------------------------

'SEND INVOICE EMAIL
'Needs username and password range reference for customer manual input of gmail account, built in provided gmail account username and password
'As well as the sheet name, send_using, smtp_server, auth level, smtp port, ssl security, your business email, customer e-mail, subject, body and path
'Typically used in loops
Sub Send(USER, PASSWORD, DEFAULT_USER, DEFAULT_PASSWORD, SHEET_NAME, SEND_USING, SMTP_SERVER, AUTH, SMTP_PORT, USE_SSL, BUSINESS_EMAIL, CUSTOMER_EMAIL, SUBJECT, BODY, PATH)
    If USER = "" And PASSWORD = "" Then
        USER = DEFAULT_USER
        PASSWORD = DEFAULT_PASSWORD
    End If
    Set CDOCONFIG = CreateObject("CDO.Configuration")
    With CDOCONFIG.Fields
        .ITEM("http://schemas.microsoft.com/cdo/configuration/sendusing") = SEND_USING 'default 2
        .ITEM("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER  'default "smtp.gmail.com"
        .ITEM("http://schemas.microsoft.com/cdo/configuration/sendusername") = USER ' your SMPT relay username
        .ITEM("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PASSWORD ' your SMPT relay password
        .ITEM("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = AUTH ' default 1
        .ITEM("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT '465
        .ITEM("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = USE_SSL 'True
        .Update
    End With
    Set MSGONE = CreateObject("CDO.Message")
    Set MSGONE.Configuration = CDOCONFIG
    MSGONE.To = CUSTOMER_EMAIL
    MSGONE.from = BUSINESS_EMAIL
    MSGONE.Cc = BUSINESS_EMAIL
    MSGONE.SUBJECT = SUBJECT
    MSGONE.TextBody = BODY
    MSGONE.AddAttachment PATH
    MSGONE.Send
    Set MSGONE = Nothing
End Sub

'ADD ANY ADDITIONAL MAIL FUNCTIONS OR SUBROUTINES NEEDED HERE...
