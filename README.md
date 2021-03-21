# excel-consolidate-using-vba
excel multiple workbook copy paste into single workbook using vba Coding
 <h2> below code copy paste your VBA Editor </h2>
 <br>
 <i>
Sub lume() <br>
Dim fd As FileDialog <br>
Set fd = Application.FileDialog(msoFileDialogOpen) <br>
fd.AllowMultiSelect = True <br>
fd.Filters.Add "Plese select Excel file only", "*.xl*", 1 <br>
fd.Show <br>
For i = 1 To fd.SelectedItems.Count <br>
Workbooks.Open fd.SelectedItems(i) <br>
frow = Cells(Rows.Count, 1).End(xlUp).Row <br>
fcol = Cells(1, Columns.Count).End(xlToLeft).Address <br>
findname = Left(fcol, Len(fcol) - 1) <br>
Range("a1:" & findname & frow).Copy <br>
Rowf = ThisWorkbook.Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row <br>
ThisWorkbook.Sheets("Sheet1").Range("a" & Rowf + 1).PasteSpecial <br>
ActiveWorkbook.Close <br>
Next <br>
MsgBox fd.SelectedItems.Count & " File Completed" <br>
End Sub <br>

</i>
  
