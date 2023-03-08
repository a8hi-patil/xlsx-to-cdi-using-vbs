Dim objExcel,ObjWorkbook,objsheet
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("E:\Abhi\vb\dva\new\822_engg_book_new_nomenclature.xlsx")
set objsheet = objExcel.ActiveWorkbook.Worksheets("FeatureInfo")
Row=objsheet.UsedRange.Rows.Count
Col=objsheet.UsedRange.columns.count
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("E:\Abhi\vb\dva\new\822_engg_book_new_nomenclature.cdi",2,True)


for i=4 to Row  Step 1
      
      
      objFileToWrite.writeline("FeatureAttributeType=X_DEV")
      rowdata =  objsheet.cells(i,2).value
      objFileToWrite.writeline("AltFeatureLbl=" & rowData)
      objFileToWrite.writeline("MeasuredActual=0.00" )
      objFileToWrite.write(vbnewline )

      objFileToWrite.writeline("FeatureAttributeType=Y_DEV" )
      rowdata =  objsheet.cells(i,2).value 
      objFileToWrite.writeline("AltFeatureLbl=" & rowData)
      objFileToWrite.writeline("MeasuredActual=0.00" )
      objFileToWrite.write(vbnewline )

      objFileToWrite.writeline("FeatureAttributeType=Z_DEV" )
      rowdata2 =  objsheet.cells(i,2).value
      objFileToWrite.writeline("AltFeatureLbl=" & rowData)
      objFileToWrite.writeline("MeasuredActual=0.00" )
      objFileToWrite.write(vbnewline )
       

      
Next