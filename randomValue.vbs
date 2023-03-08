Dim objExcel,ObjWorkbook,objsheet
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("E:\Abhi\vb\dva\822_Measurement_plan_final.xlsx")
set objsheet = objExcel.ActiveWorkbook.Worksheets("VSA")
Row=objsheet.UsedRange.Rows.Count
Col=objsheet.UsedRange.columns.count
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("E:\Abhi\vb\dva\temp1.cdi",2,True)


for i=4 to Row  Step 1
      
      
      objFileToWrite.writeline("FeatureAttributeType=X_DEV")
      rowdata =  objsheet.cells(i,1).value
      lowerX = objsheet.cells(i,2).value
      UpperX = objsheet.cells(i,4).value
      Randomize
    '   randx = (UpperX-lowerX+1)*Rnd+lowerX
        range = UpperX - lowerX
        ' Generate a random number between 0 and 1
        randomNum = Rnd()

        ' Scale the random number to the range
        scaledNum = randomNum * range
        finalNum = scaledNum + lowerX
      randx = FormatNumber (finalNum, 2, , , vbFalse)
      objFileToWrite.writeline("AltFeatureLbl=" & rowData)
      objFileToWrite.writeline("MeasuredActual="&randx )
      objFileToWrite.write(vbnewline )

      objFileToWrite.writeline("FeatureAttributeType=Y_DEV" )
      rowdata =  objsheet.cells(i,1).value 
      lowerY = objsheet.cells(i,5).value
      UpperY = objsheet.cells(i,7).value
      Randomize
    '   randy = (UpperY-lowerY+1)*Rnd+lowerY
    range = UpperY - lowerY
        ' Generate a random number between 0 and 1
        randomNum = Rnd()

        ' Scale the random number to the range
        scaledNum = randomNum * range
        finalNum = scaledNum + lowerY
      randy = FormatNumber (finalNum, 2, , , vbFalse)
      objFileToWrite.writeline("AltFeatureLbl=" & rowData)
      objFileToWrite.writeline("MeasuredActual="&randy )
      objFileToWrite.write(vbnewline )

      objFileToWrite.writeline("FeatureAttributeType=Z_DEV" )
      rowdata2 =  objsheet.cells(i,1).value
      lowerZ = objsheet.cells(i,8).value
      UpperZ = objsheet.cells(i,10).value
      Randomize
    '   randz = (UpperZ-lowerZ+1)*Rnd+lowerZ
    range = UpperZ - lowerZ
        ' Generate a random number between 0 and 1
        randomNum = Rnd()

        ' Scale the random number to the range
        scaledNum = randomNum * range
        finalNum = scaledNum + lowerZ
      randz = FormatNumber (finalNum, 2, , , vbFalse)
      objFileToWrite.writeline("AltFeatureLbl=" & rowData)
      objFileToWrite.writeline("MeasuredActual="&randz )
      objFileToWrite.write(vbnewline )
       

      
Next