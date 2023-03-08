

for i=1 to 5 Step 1
filename ="E:\Abhi\vb\gayathri dva\temp00"&i&".cdi"
'  Dim name As String = i
' Dim filename As String = $"Hi {name}"
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(filename,2,True)
' MsgBox filename
Next