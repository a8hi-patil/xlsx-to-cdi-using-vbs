for j=1 to 5 Step 1
filename ="E:\Abhi\vb\dva\new\measurement_plan_"&j&".cdi"
'  Dim name As String = i
' Dim filename As String = $"Hi {name}"
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(filename,2,True)

Dim objExcel,ObjWorkbook,objsheet
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("E:\Abhi\vb\dva\new\822_engg_book_new_nomenclature.xlsx")
set objsheet = objExcel.ActiveWorkbook.Worksheets("VSA")
Row=objsheet.UsedRange.Rows.Count
Col=objsheet.UsedRange.columns.count
' Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("E:\Abhi\vb\gayathri dva\temp1.cdi",2,True)


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






Next