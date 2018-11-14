Attribute VB_Name = "Work"
'---------------------------------DECLARE WORK MODULE------------------------------------

'USED TO EXTRACT INVOICE DATA
'Returns table object or collection object as array
Function Map_Invoice_Data(WS, TABLE_NAME)
    Set INVOICE_DATA = WS.ListObjects(TABLE_NAME)
    TEMP_ARRAY = INVOICE_DATA.DataBodyRange
    INVOICE_DATA_ARRAY = Application.Transpose(TEMP_ARRAY)
    MAP_ARRAY = INVOICE_DATA_ARRAY
    Map_Invoice_Data = MAP_ARRAY
End Function

'USED TO LOOP THROUGH MAPPED INVOICE DATA ARRAY
'Uses  Update_Progress_Bar to update user realtime in the loop
Sub Build_Invoice(PATH, WS, WS_INVOICE, INVOICE_DATA_ARRAY, TABLE_NAME)
    INC = 20
    Set INVOICE_NO = WS.Range("E5")
    Set INVOICE_DATE = WS.Range("E6")
    Set CUSTOMER_ID = WS.Range("E7")
    Set CUSTOMER_NAME = WS.Range("B10")
    Set CUSTOMER_COMPANY_NAME = WS.Range("B11")
    Set CUSTOMER_STREET_ADDRESS = WS.Range("B12")
    Set CUSTOMER_CITY_ZIP_CODE = WS.Range("B13")
    Set CUSTOMER_PHONE = WS.Range("B14")
    Set SALESPERSON = WS.Range("A17")
    Set JOB = WS.Range("C17")
    Set PAYMENT_TERMS = WS.Range("D17")
    Set DUE_DATE = WS.Range("F17")
    Set BUSINESS_EMAIL = WS.Range("A8")
    Set COMPANY_NAME = WS.Range("A3")
    Set INVOICE_DATA = WS_INVOICE.ListObjects(TABLE_NAME)
    For X = 1 To INVOICE_DATA.DataBodyRange.Rows.Count
        INVOICE_NO.VALUE = INVOICE_DATA_ARRAY(1, X)
        INVOICE_DATE.VALUE = INVOICE_DATA_ARRAY(2, X)
        CUSTOMER_ID.VALUE = INVOICE_DATA_ARRAY(3, X)
        CUSTOMER_NAME.VALUE = INVOICE_DATA_ARRAY(4, X)
        CUSTOMER_COMPANY_NAME.VALUE = INVOICE_DATA_ARRAY(5, X)
        CUSTOMER_STREET_ADDRESS.VALUE = INVOICE_DATA_ARRAY(6, X)
        CUSTOMER_CITY_ZIP_CODE.VALUE = INVOICE_DATA_ARRAY(7, X) & "-" & INVOICE_DATA_ARRAY(8, X) & "-" & INVOICE_DATA_ARRAY(9, X)
        CUSTOMER_PHONE.VALUE = INVOICE_DATA_ARRAY(10, X)
        SALESPERSON.VALUE = INVOICE_DATA_ARRAY(11, X)
        JOB.VALUE = INVOICE_DATA_ARRAY(12, X)
        PAYMENT_TERMS.VALUE = INVOICE_DATA_ARRAY(13, X)
        DUE_DATE.VALUE = INVOICE_DATA_ARRAY(14, X)
        If (INVOICE_NO.VALUE <> "") Then
            If (X > 1) Then
                If (INVOICE_NO = INVOICE_DATA_ARRAY(1, X - 1) And JOB = INVOICE_DATA_ARRAY(12, X - 1)) Then
                    INC = INC + 1
                    Set QUANTITY = WS.Range("A" & INC)
                    QUANTITY.VALUE = INVOICE_DATA_ARRAY(15, X)
                    Set DESCRIPTION = WS.Range("B" & INC)
                    DESCRIPTION.VALUE = INVOICE_DATA_ARRAY(16, X)
                    Set UNIT_PRICE = WS.Range("E" & INC)
                    UNIT_PRICE.VALUE = INVOICE_DATA_ARRAY(17, X)
                Else
                    Set SINGLE_ITEM_RANGE = WS.Range("A20:E39")
                    SINGLE_ITEM_RANGE.ClearContents
                    Set QUANTITY = WS.Range("A20")
                    Set DESCRIPTION = WS.Range("B20")
                    Set UNIT_PRICE = WS.Range("E20")
                    QUANTITY.VALUE = INVOICE_DATA_ARRAY(15, X)
                    DESCRIPTION.VALUE = INVOICE_DATA_ARRAY(16, X)
                    UNIT_PRICE.VALUE = INVOICE_DATA_ARRAY(17, X)
                End If
                CUSTOMER_EMAIL = INVOICE_DATA_ARRAY(18, X)
            Else
                Set SINGLE_ITEM_RANGE = WS.Range("A20:E39")
                SINGLE_ITEM_RANGE.ClearContents
                Set QUANTITY = WS.Range("A20")
                Set DESCRIPTION = WS.Range("B20")
                Set UNIT_PRICE = WS.Range("E20")
                QUANTITY.VALUE = INVOICE_DATA_ARRAY(15, X)
                DESCRIPTION.VALUE = INVOICE_DATA_ARRAY(16, X)
                UNIT_PRICE.VALUE = INVOICE_DATA_ARRAY(17, X)
                CUSTOMER_EMAIL = INVOICE_DATA_ARRAY(18, X)
            End If
            WS.Copy
            Columns(7).EntireColumn.Hidden = True
            Columns(8).EntireColumn.Hidden = True
            Columns(9).EntireColumn.Hidden = True
            Columns(10).EntireColumn.Hidden = True
            Columns(11).EntireColumn.Hidden = True
            Columns(12).EntireColumn.Hidden = True
            FILE_NAME = "\" & Replace(Replace(Replace(COMPANY_NAME, ".", ""), "'", ""), "-", " ") & "_Invoice_for_" & Replace(Replace(Replace(CUSTOMER_COMPANY_NAME, ".", ""), "'", ""), "-", " ") & "_Invoice_" & INVOICE_NO & "_To_" & CUSTOMER_EMAIL & "_Due Date_" & Replace(DUE_DATE, "/", "-") & ".xlsx"
            ActiveWorkbook.Close True, PATH & FILE_NAME
            Update_Progress_Bar CSng(X)
        End If
    Next X
End Sub

'SAVE SUMMARY DATA IN SEPARATE SHEET
Sub Save_Summary(WS, PATH, WRK_DATE)
    WS.Copy
    ActiveWorkbook.Close True, PATH & "\Invoice_Data_for_" & WRK_DATE & "_Delivery.xlsx"
    Update_Progress_Bar Replace(UserForm1.FrameProgress.Caption, "%", "")
End Sub
