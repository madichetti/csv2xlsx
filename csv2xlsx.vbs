Option Explicit 
Dim colArgs,  outputFile, inputFile, inputName, inputFileList, fileName
Dim outputFullpath, inputFileFullpath, fileExists
Dim Objfso, ObjExcel, ObjWorkbook, ObjSheet,newSheet
Dim fileDelimiter, rawFile, currentRow, rowText, rowColumnValue, columnValue, lineArray

On Error Resume Next

' variable to hold arguments passed:
Set colArgs = WScript.Arguments.Named

' Arguments to be minimum 2.
If colArgs.Count < 2  Then
    Wscript.Echo " ******************************************************************** "
    WScript.Echo "Insufficient arguments supplied!."    
    Wscript.Echo "Required output (excel FileName). "
    Wscript.Echo "Required at-least one CSV file with comma delimited. "
    Wscript.Echo "Usage: /o:<File Name with extension xlsx> /i:<csv FileName>,<csv FileName>;...."
    Wscript.Echo "Example c:\temp>cscript.exe //nologo csv2xlsx.vbs /d:c /o:output.xlsx /i:s.csv,s2.csv,s3.csv,s4.csv"
    Wscript.Echo " ******************************************************************** "
    WScript.Quit
End If

' Get the delimiter for the csv Files
If colArgs.Exists("d") Then
    fileDelimiter = colArgs.Item("d")
End If

' Get the filename for the output excel
If colArgs.Exists("o") Then
    outputFile = colArgs.Item("o")
End If

' If the filename is blank or nothing
If (outputFile = Empty) Then
    Wscript.Echo "Usage: /o:<File Name with extension xlsx> is required."
    Wscript.Quit
End If

' If the output extension is missing, then add xlsx as extension.
If (Instr(outputFile,".xlsx")<=0) Then
    outputFile = outputFile & ".xlsx"
End If

' Get the Input Files comma seperated
If colArgs.Exists("i") Then
    inputFile = colArgs.Item("i")
End If

' If the input filename is blank or nothing
If (inputFile = Empty) Then
    Wscript.Echo "Usage: /i:<csv filename> is required. "
    Wscript.Quit
End If

' create an Object to get the current Filesystem path.
Set Objfso = CreateObject("Scripting.FileSystemObject")
outputFullpath = Objfso.GetAbsolutePathName(outputFile)

' Create Spreadsheet
Set ObjExcel = CreateObject("Excel.Application")
ObjExcel.Visible = false
ObjExcel.displayalerts=false

' Add the workbook to objExcel
Set ObjWorkbook = ObjExcel.Workbooks.Add()
' process the input files
inputFileList = Split(inputFile,",")
For Each rawFile in inputFileList
    If (Instr(rawFile,".csv") or Instr(rawFile,".tab") or Instr(rawFile,".pipe") ) Then
        inputFileFullpath = Objfso.GetAbsolutePathName(rawFile)
        
        ' ' Get the current CSV Data loaded into workbook
        ' 'expression.Open (FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, NotIfy, Converter, AddToMru, Local, CorruptLoad)
        ' 'Set xTempWb = ObjExcel.Workbooks.Open(inputFileFullpath)
        ' Set xTempWb = ObjExcel.Workbooks.Open(inputFileFullpath,0,False,xlDelimited)',,,,,, , ,  Array(Array(1, 4)))
        ' xTempWb.Sheets(1).UsedRange.NumberFormat = "@"
        ' xTempWb.Sheets(1).Move, ObjWorkbook.Sheets(ObjWorkbook.Sheets.Count)
        
        ' New workSheet
        Set ObjSheet = ObjWorkbook.Worksheets.Add(, ObjWorkbook.Worksheets(ObjWorkbook.Worksheets.Count))
        fileName = replace(Objfso.GetFileName(inputFileFullpath),".csv","")
        fileName = replace(fileName,".tab","")
        fileName = replace(fileName,".pipe","")
        ObjSheet.Name = fileName
     
        Set inputFile = Objfso.OpenTextFile(inputFileFullpath)
        currentRow = 1
        Do While inputFile.AtEndOfStream <> True
            If (UCase(fileDelimiter) ="T") Then
                lineArray = Split(inputFile.ReadLine,vbTab)
            ElseIf (UCase(fileDelimiter) ="P") Then
                lineArray = Split(inputFile.ReadLine,"|")
            Else
                lineArray = Split(inputFile.ReadLine,",")
            End If
            For each columnValue in lineArray
                rowColumnValue = Trim(columnValue)
                'Check 1st position in column for doubleQuotes
                If (Left(rowColumnValue,1) = """") Then
                    rowColumnValue = Right(rowColumnValue,Len(rowColumnValue)-1)
                End If 
                'Check last column position for doubleQuotes
                If (right(rowColumnValue,1) = """") Then
                    rowColumnValue = Left(rowColumnValue,Len(rowColumnValue)-1)
                End If 
                If Trim(rowText) ="" Then
                    rowText = Trim(rowColumnValue)
                Else
                    rowText = rowText + "," + rowColumnValue
                End If
            Next
            'wscript.Echo rowText
            lineArray = Split(rowText,",")
            rowText = Empty
            ObjSheet.Range(chr(65) & "" & currentRow, chr(65+ Ubound(lineArray))  & "" &  currentRow).NumberFormat = "@"
            ObjSheet.Range(chr(65) & "" & currentRow, chr(65+ Ubound(lineArray))  & "" &  currentRow) = lineArray
            currentRow = currentRow + 1
        Loop
        inputFile.Close
        ObjSheet.UsedRange.NumberFormat = "@"
        ObjSheet.UsedRange.EntireColumn.Autofit()
    End If
Next
ObjWorkbook.Sheets("sheet1").Delete

fileExists=Objfso.FileExists(outputFullpath)
If fileExists Then
    Objfso.DeleteFile(outputFullpath) 
end If

' Save Spreadsheet, 51 = Excel 2007-2010 
ObjWorkbook.SaveAs outputFullpath, 51

' Release Lock on Spreadsheet
ObjExcel.Quit()

Set ObjWorkbook = Nothing
Set ObjExcel = Nothing
Set Objfso = Nothing


WScript.Echo "Done!"