# Notes
'
Autofit All Tables
Sub AutoFitAllTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        ' Option A: To fit to the Window/Page Margins
        tbl.AutoFitBehavior (wdAutoFitWindow)
        
        ' Option B: To fit to the Contents (Uncomment the line below to use)
        ' tbl.AutoFitBehavior (wdAutoFitContent)
    Next tbl
    MsgBox "All tables have been auto-fitted!"
End Sub
'
