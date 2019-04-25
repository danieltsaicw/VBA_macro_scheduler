Dim args, objExcel

set args = WScript.Arguments
set objExcel = CreateObject("Excel.Application")


Wscript.Echo "(1) Loading Bloomberg API"
WScript.Sleep 5000

Set wbks2 = objExcel.Workbooks 
wbks2.Open "C:\blp\api\Office Tools\BloombergUI.xla" 
wbks2("BloombergUI.xla").RunAutoMacros 1 'xlAutoOpen 

Wscript.Echo "(2) Done Loading Bloomberg API"
WScript.Sleep 5000


objExcel.Workbooks.Open args(0)
objExcel.Visible = True


Wscript.Echo "(3) Start merging new guys"
objExcel.Run "main"
Wscript.Echo "Done"
WScript.Sleep 1000

Wscript.Echo "(4) Start loading data from Bloomberg"
WScript.Sleep 30000

Wscript.Echo "(5) Start duplicating hard-coding copy"
objExcel.Run "duplicate"
Wscript.Echo "Done"
WScript.Sleep 1000

Wscript.Echo "(6) Start Ranking"
objExcel.Run "Rank"
Wscript.Echo "Done"
WScript.Sleep 1000

objExcel.ActiveWorkbook.Save
Wscript.Echo "AcenderQuality Saved"
WScript.Sleep 50000

objExcel.ActiveWorkbook.Close(0)
Set wbks2 = Nothing
objExcel.Quit