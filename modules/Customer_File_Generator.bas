
'This macro will help to split the Workbooks by customer. Taking as main Workbook, the UDM_Dispute file

Sub Customer_File_Generator()

    Dim wbPath As String
    Dim wbName As String
    Dim fullPath As String
    Dim sourceWB As Workbook, tempWB As Workbook
    Dim ws As Worksheet, newWS As Worksheet
    Dim rng As Range, cell As Range
    Dim AORList As Variant, AORSheetNames As Variant, AORPaths As Variant
    Dim AORCol As Long, custCol As Long
    Dim todayDate As String
    Dim uniqueCusts As Collection
    Dim cust As Variant
    Dim filePath As String, savePath As String
    Dim filteredRng As Range
    Dim col As Long, lastCol As Long
    Dim cleanCust As String
    Dim header As String 'new added

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

        ' This is the format of the Workbook, following today's date
    todayDate = Format(Date, "mm-dd-yy")

        ' Location of the Workbook
    wbName = "UDM_Dispute " & todayDate & ".xlsm"
    wbPath = "/sample_data/Output/"
    fullPath = wbPath & wbName

        ' Open the workbook
    Set sourceWB = Workbooks.Open(fullPath)
    Set ws = sourceWB.Sheets("Main")

        ' Here we define the filter Parameters
    AORList = Array("Africa", "Asia", "Central East", "Global", _
                    "Latam", "North America", "Oceania", "Region Integrated", "South America", "Miscellaneous")

    AORSheetNames = Array("Africa", "Asia", "Central East", "Global", _
                          "Latam", "North America", "Oceania", "Region Integrated", "South America", "Miscellaneous")

    AORPaths = Array( _
        "/sample_data/Output/Africa/", _
        "/sample_data/Output/Asia/", _
        "/sample_data/Output/Central East/", _
        "/sample_data/Output/Global/", _
        "/sample_data/Output/Latam", _
        "/sample_data/Output/North America/", _
        "/sample_data/Output/Oceania/", _
        "/sample_data/Output/Region Integrated/", _
        "/sample_data/Output/South America/", _
        "/sample_data/Output/Miscellaneous/" _
    )

        ' Get Column Index Numbers for Headers
    AORCol = 0
    custCol = 0
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        header = ws.Cells(1, col).Value
        header = Trim(Replace(header, Chr(160), " "))

        If header = "AOR" Then
            AORCol = col
        ElseIf header = "Customer Name" Then
            custCol = col
        End If
    Next col

    If AORCol = 0 Or custCol = 0 Then
        MsgBox "Could not find 'AOR' or 'Customer Name' in row 1 headers.", vbCritical
        Exit Sub
    End If

        ' Here is another Loop "for", through each AOR value
    Dim i As Long
    For i = LBound(AORList) To UBound(AORList)

            ' Clear existing filters
        If ws.AutoFilterMode Then ws.AutoFilterMode = False

            ' Here, we apply filters again
        ws.Range("A1").CurrentRegion.AutoFilter Field:=AORCol, Criteria1:=AORList(i)

        On Error Resume Next
        Set rng = ws.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rng Is Nothing Then
                
                ' Here, we create new sheets
            Set newWS = sourceWB.Sheets.Add(After:=sourceWB.Sheets(sourceWB.Sheets.Count))
            newWS.Name = AORSheetNames(i)
            rng.Copy Destination:=newWS.Range("A1")

                ' Get unique customers
            Set uniqueCusts = New Collection
            On Error Resume Next
            For Each cell In newWS.Range(newWS.Cells(2, custCol), newWS.Cells(newWS.Cells(Rows.Count, custCol).End(xlUp).Row, custCol))
                If Len(Trim(cell.Value)) > 0 Then
                    uniqueCusts.Add cell.Value, CStr(cell.Value)
                End If
            Next cell
            On Error GoTo 0

                ' For each customer, create a new workbook and save
            For Each cust In uniqueCusts
                
                    ' Filter by customer
                If newWS.AutoFilterMode Then newWS.AutoFilterMode = False
                newWS.Range("A1").CurrentRegion.AutoFilter Field:=custCol, Criteria1:=cust

                Set filteredRng = Nothing
                On Error Resume Next
                Set filteredRng = newWS.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)
                On Error GoTo 0

                If Not filteredRng Is Nothing Then
                    Set tempWB = Workbooks.Add
                    filteredRng.Copy Destination:=tempWB.Sheets(1).Range("A1")

                        'If syntax error, this code will clean invalid characters
                    cleanCust = Application.WorksheetFunction.Clean(cust)
                    cleanCust = Replace(cleanCust, "\", "")
                    cleanCust = Replace(cleanCust, "/", "")
                    cleanCust = Replace(cleanCust, ":", "")
                    cleanCust = Replace(cleanCust, "*", "")
                    cleanCust = Replace(cleanCust, "?", "")
                    cleanCust = Replace(cleanCust, """", "")
                    cleanCust = Replace(cleanCust, "<", "")
                    cleanCust = Replace(cleanCust, ">", "")
                    cleanCust = Replace(cleanCust, "|", "")
                    filePath = AORPaths(i) & cleanCust & ".xlsx"

                        ' Validation of the location
                    If Dir(AORPaths(i), vbDirectory) = "" Then MkDir AORPaths(i)

                        ' This step will ensure, we load correctly the file. If already exists, this code will overwrite existing file
                    If Dir(filePath) <> "" Then Kill filePath

                        ' Autofit columns
                    tempWB.Sheets(1).Cells.EntireColumn.AutoFit

                        ' Save workbook
                    tempWB.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
                    tempWB.Close SaveChanges:=False
                End If
            Next cust
        End If

        Set rng = Nothing
    Next i

        ' This is the final cleanup
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

        ' Final Output
    MsgBox "The customer information has been separated successfully"

End Sub