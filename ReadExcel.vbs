'ReadExcel Using Search                    
Set objExcel = Wscript.CreateObject("Excel.Application")   
Set objWorkbook = objExcel.Workbooks.Open("I:\VBSProj\Data.xlsx")   
objExcel.visible=False
rowCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
colCount=objExcel.ActiveWorkbook.Sheets(1).UsedRange.Columns.count  
Msgbox("Rows    :" &  rowCount & "   " & "Columns :" & colCount)
a=inputbox("Enter the Roll number","Search") 
  for intRow=2 to rowCount
     if (CStr(a) = CStr(objExcel.Cells(intRow, 1).Value)) then
       for intCol=2 to colCount
        c=c & objExcel.Cells(1, intCol) &" : "& objExcel.Cells(intRow, intCol).Value & vbCrLf        
        next
        MsgBox ("----Student Exam Status----" & vbCrLf & vbCrLf & c) 
     End if
        c=null
  next
objExcel.Quit
