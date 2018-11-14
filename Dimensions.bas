Attribute VB_Name = "Dimensions"
'---------------------------------DECLARE DIMENSIONS------------------------------------
'We want explicit dimension declaration, e.g. As Int, As String, As Object etc.
Option Explicit

'DECLARE WORKBOOK, ARRAYS, TABLE OBJECTS, PROGRESS & MODULE DIMENSIONS
'Workbooks
'Just your workbok and worksheet variables
Sub Declare_Workbooks()
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim SHEET_NAME As String
End Sub

'Work Arrays
'Typically Template and Data
Sub Declare_Work_Arrays()
    Dim TEMPLATE_ARRAY()
    Dim DATA_ARRAY()
End Sub

'Table Objects
'Any table objects or arrays that store multiple table objects
Sub Declare_Table_Objects()
    Dim TABLE_NAME As String
End Sub

'Progress
'Used in visual updating
Sub Declare_Progress()
    'Used in progress bar updating
    Dim COUNTER As Integer
    Dim ROWMAX As Integer
    Dim COLMAX As Integer
    Dim ROW As Integer
    Dim COLUMN As Integer
    Dim PCTDONE As Single
End Sub

'Alerts
'Used for all user prompting (except progress)
Sub Declare_Alerts()
    Dim TEST As Boolean
    Dim MESSAGE As String 'Also used in logging, mail & mail_loop modules
    Dim PROMPT_TYPE
    Dim TITLE As String
    Dim PCTDONE As Single
End Sub

'Cleanup
'Cleanup arrays and iteration objects
Sub Declare_Cleanup()
    Dim CLEAN_ARRAY()
    Dim CLEAN_ITEM As String
    Dim CLEAN_RANGE As Range
End Sub

'Dates
'Typically today's date
Sub Declare_Dates()
    Dim WORK_DATE As String
End Sub

'Folders
'Any folder work
Sub Declare_Folders()
    Dim PATH As String
    Dim FILE_NAME As String 'Also used in logging module
    Dim PATH_ARRAY As ListObject
End Sub

'Formatting
'Used in formatting functions
Sub Declare_Formatting()
    Dim DATE_FORMAT As String
End Sub

'Logging
'Used by all logging functions
Sub Declare_Logging()
    Dim LOG_FILE
End Sub

'Mail
'All objects used in mail subroutine
Sub Declare_Mail()
    'Needs Google mail account with ALLOW LESS SECURE APPS feature TURNED ON under your Google account's Sign In & Security
    Dim SEND_USING As Integer
    Dim SMTP_SERVER As String
    Dim USER As Range
    Dim PASSWORD As Range
    Dim DEFAULT_USER As String 'Range alphanumeric
    Dim DEFAULT_PASSWORD As String 'Range alphanumeric
    Dim AUTH As Integer
    Dim SMTP_PORT As Integer
    Dim USE_SSL As Boolean
    Dim CDOCONFIG As Mailer
    Dim MSGONE As Mailer
    Dim BUSINESS_EMAIL As String
    Dim CUSTOMER_EMAIL As String
    Dim SUBJECT As String
    Dim BODY As String
End Sub

'Mail loop
'We're looping through output folder to send e-mails, instead of calling Mail.Send for every outputed file on creation
'This ensures that in case of send failure, the outputs are at least created
Sub Declare_Mail_Loop()
    Dim LOCATION As String
    Dim FILETYPE As String
End Sub

'Work
Sub Declare_Work()
    'Used in Map_Invoice_Data to map invoice data
    Dim MAP_ARRAY() As Variant
    Dim TEMP_ARRAY As Variant
    Dim RETURN_ARRAY() As Variant
    'TEMPLATE FIELD RANGES
    'The template field names should be self-explanatory
    Dim INVOICE_NO As Range
    Dim INVOICE_DATE As Range
    Dim COMPANY_NAME As Range
    Dim CUSTOMER_ID As Range
    Dim CUSTOMER_NAME As Range
    Dim CUSTOMER_COMPANY_NAME As Range
    Dim CUSTOMER_STREET_ADDRESS As Range
    Dim CUSTOMER_CITY_ZIP_CODE As Range
    Dim CUSTOMER_PHONE As Range
    Dim SALESPERSON As Range
    Dim JOB As Range
    Dim PAYMENT_TERMS As Range
    Dim DUE_DATE As Range
    Dim QUANTITY As Range
    Dim DESCRIPTION As Range
    Dim UNIT_PRICE As Range
    Dim CUSTOMER_EMAIL
    Dim BUSINESS_EMAIL
    'INVOICE DATA TABLE OBJECTS
    Dim INVOICE_DATA As ListObject
    Dim INVOICE_DATA_ARRAY As Variant
    Dim X As Integer
    Dim INC As Integer
    Dim SINGLE_ITEM_RANGE As Range
End Sub

'DECLARE IMPORT OBJECTS
Sub Declare_Import()
    'Used in importing e-receipt data
    Dim THIS_ACTIVE As Workbook
    Dim THIS_ACTIVE_SHEET As Worksheet
    Dim CHOICE As Integer
    Dim PATHSELECTED As String
    Dim LIST_OBJ As Object
End Sub
