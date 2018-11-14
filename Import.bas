Attribute VB_Name = "Import"
'IMPORT INVOICE DATA
'macro attached to import button
'Captures workbook through filedialog
'copies Table1 range to Table1 range from solution output summary sheets into current worksheet
Sub Import_Invoice_Data()
    Declare_Workbooks
    Declare_Import
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    CHOICE = Application.FileDialog(msoFileDialogOpen).Show
    If CHOICE <> 0 Then
        PATHSELECTED = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    End If
    Delete_Table_Data ThisWorkbook.Sheets("Invoice Data").Name, "Table1"
    Set THIS_ACTIVE = ThisWorkbook
    Set THIS_ACTIVE_SHEET = THIS_ACTIVE.Sheets("Invoice Data")
    Set WB = Application.Workbooks.Open(PATHSELECTED)
    Set WS = WB.Sheets("Invoice Data")
    WS.ListObjects("Table1").Range.Copy Destination:=THIS_ACTIVE_SHEET.Range("A4")
    Rename_Table1
    MsgBox "Import complete!"
    WB.Close
End Sub
