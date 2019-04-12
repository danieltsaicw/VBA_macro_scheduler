Dim args, objExcel

set args = WScript.Arguments
set objExcel = CreateObject("Excel.Application")

objExcel.Workbooks.Open args(0)
objExcel.Visible = True

objExcel.Run "macro_foo2"

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)
objExcel.Quit