Sub ForecastAuto()

Dim dateval As Date
Dim datestr As String
Dim datestr2 As String

'Set Directory & File System Object



  FolderName = ("Y:\Forecast Summary Automation\alldatadump\")



                Set FileSys = CreateObject("Scripting.FileSystemObject")



                Set myFolder = FileSys.GetFolder(FolderName)







        dteFile = DateSerial(1900, 1, 1)



        For Each objFile In myFolder.Files



            If InStr(1, objFile.Name, ".csv") > 0 Then



                If objFile.DateLastModified > dteFile Then



                    dteFile = objFile.DateLastModified



                    strFilename2 = strFilename



                    strFilename = objFile.Name



                End If



            End If



        Next objFile



'opening of latest file in the folder



Set wb = Workbooks.Open(FolderName & Application.PathSeparator & strFilename)
'
''Clear data from existing forecast workbook
'Last_Row = ThisWorkbook.Sheets("Data").Range("A18").End(xlDown).Row
'
'ThisWorkbook.Sheets("Data").Range("A18:" & "R" & Last_Row).ClearContents


'copy and past data from alldata (appended data from the 'LIVE' folder)
Last_RowAlldata = wb.Worksheets(1).Range("A2").End(xlDown).Row

wb.Worksheets(1).Range("A2:" & "R" & Last_RowAlldata).Copy
Workbooks(1).Worksheets("Data").Range("A18").PasteSpecial Paste:=xlPasteValues

'copying formulas down
ThisWorkbook.Worksheets("Data").Range("S17:AU17").Select
Last_RowFormula = ThisWorkbook.Worksheets("Data").Range("A18").End(xlDown).Row
Selection.AutoFill Destination:=Range("S17:" & "AU" & Last_RowFormula)

'update cell on Summary tab F4 cell
datestr = ThisWorkbook.Worksheets("Data").Range("L1").Value
datestr2 = Mid(datestr, 10, 11)
dateval = datestr2
ThisWorkbook.Worksheets("Summary").Range("F4").Value = dateval
ThisWorkbook.RefreshAll

End Sub

