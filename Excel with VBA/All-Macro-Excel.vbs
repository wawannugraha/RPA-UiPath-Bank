Sub SAVEasXLSB()
 Application.ScreenUpdating = False
 
 Dim fpath As String
 Dim wname As String
 
wname = Left(ActiveWorkbook.Name, (InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare) - 1))
fpath = ActiveWorkbook.Path & "\Output\" & wname
ActiveWorkbook.SaveAs Filename:=fpath & ".xlsb", FileFormat:= _
 xlExcel12, CreateBackup:=False
 Application.ScreenUpdating = True
End Sub



Sub AutofitColumn(ByVal Range As Rg)
Columns(rg).EntireColumn.AutoFit
End Sub




Sub FilterRows(ByVal Sheetname As sheetname)
With Worksheets(sheetname).Range("A1")
.AutoFilter field:=2, Criteria1:="07*", Operator:=xlOr, Criteria2:="08*"
.AutoFilter field:=5, Criteria1:="Not Exist"
End With


Sub SelectSheetsAndSaveAsPDF(ByVal in_FileLocation As FilePath, ByVal in_Sheet1 as Sheet1, ByVal in_Sheet2 as Sheet2 )

'Create and assign variables
Dim saveLocation As String
Dim sheetArray As Variant

saveLocation = FilePath
sheetArray = Array(Sheet1, Sheet2)

'Select specific sheets from workbook, the save all as PDF
Sheets(sheetArray).Select
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=saveLocation

End Sub