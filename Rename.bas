Attribute VB_Name = "Rename"
'Used for Table1 renaming
'Table 1 gets renamed by excel during Workbook.Close process, saving sheet and incrementing Table1 number
'Table 1 gets renamed by excel during Import
'We want to do it on open workbook
'Call it in ThisWorkbook
Sub Rename_Table1()
    Declare_Workbooks
    Set WB = ThisWorkbook
    Set WS = WB.Sheets("Invoice Data")
    For Each LIST_OBJ In WS.ListObjects
        LIST_OBJ.Name = "Table1"
    Next LIST_OBJ
    WB.Save
End Sub
