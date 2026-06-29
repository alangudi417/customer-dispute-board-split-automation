Option Explicit

'This macro will help to add the UDM_Comments, and add another column that determines the status of the Disputes.

Sub Data_Preparation()

    Dim todayStr As String
    Dim pathDisputes As String, pathNotes As String
    Dim wbDisputes As Workbook, wbNotes As Workbook
    Dim wsOpen As Worksheet, wsMain As Worksheet, wsNotes As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, lastColumn As Long
    Dim dictNotes As Object
    Dim notesData As Variant, mainData As Variant
    Dim cell As Range, col As Long
    Dim deleteCols As Variant
    Dim colMap As Object

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Excel Locations
    todayStr = Format(Date, "mm-dd-yy")
    pathDisputes = "/sample_data/Input/UDM_Dispute " & todayStr & ".xlsm"
    pathNotes = "/sample_data/Input/UDM_Notes " & todayStr & ".xlsx"

    ' Open Disputes workbook
    Set wbDisputes = Workbooks.Open(pathDisputes)
    Set wsOpen = wbDisputes.Sheets("Open")

    ' Delete old "Main" in case it does exist
    On Error Resume Next
    Application.DisplayAlerts = False
    wbDisputes.Sheets("Main").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create Main Sheet
    Set wsMain = wbDisputes.Sheets.Add(After:=wbDisputes.Sheets(wbDisputes.Sheets.Count))
    wsMain.Name = "Main"

    ' Copy data from Warboard Open Sheet, columns A:AJ to Main
    With wsOpen
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A1:AD" & lastRow).Copy Destination:=wsMain.Range("A1")
    End With

    ' Format of headers
    With wsMain
        lastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
        With .Range(.Cells(1, 1), .Cells(1, lastColumn)).Interior
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = -0.249977111117893
        End With
    End With

    ' Add headers for new columns
    wsMain.Range("AE1").Value = "Dispute Status"
    wsMain.Range("AF1").Value = "UDM Notes"

    ' Fill "Dispute Status" column
    mainData = wsMain.Range("A2:AD" & lastRow).Value
    ReDim categoryData(1 To UBound(mainData), 1 To 1)

    For i = 1 To UBound(mainData)
        If Trim(mainData(i, 21) = "") Then ' U = column 21
            categoryData(i, 1) = "New"
        Else
            categoryData(i, 1) = "In process"
        End If
    Next i
    wsMain.Range("AE2").Resize(UBound(categoryData), 1).Value = categoryData

    ' Load the notes in the workbook
    Set wbNotes = Workbooks.Open(pathNotes, ReadOnly:=True)
    Set wsNotes = wbNotes.Sheets(1)

    ' Load Notes data into dictionary for fast lookup
    Dim noteLastRow As Long
    noteLastRow = wsNotes.Cells(wsNotes.Rows.Count, "C").End(xlUp).Row
    notesData = wsNotes.Range("C2:AW" & noteLastRow).Value

    Set dictNotes = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(notesData)
        If Not dictNotes.exists(notesData(i, 1)) Then
            dictNotes(notesData(i, 1)) = notesData(i, 47) ' Column AW (47th column of full sheet)
        End If
    Next i

    ' Fill in UDM Notes via dictionary lookup
    ReDim notesResult(1 To UBound(mainData), 1 To 1)
    For i = 1 To UBound(mainData)
        If dictNotes.exists(mainData(i, 9)) Then ' I = 9th col (Dispute Key)
            notesResult(i, 1) = dictNotes(mainData(i, 9))
        Else
            notesResult(i, 1) = ""
        End If
    Next i
    wsMain.Range("AF2").Resize(UBound(notesResult), 1).Value = notesResult
    wsMain.Columns("AF").NumberFormat = "@"

    ' Rename the headers
    wsMain.Range("I1").Value = "Dispute"
    wsMain.Range("J1").Value = "PO"
    wsMain.Range("L1").Value = "Invoice"

    ' Autofit the columns
    wsMain.Range("A1:AF1").EntireColumn.AutoFit

    ' Change the width of column AF, and unwrap
    wsMain.Columns("AF").ColumnWidth = 60
    wsMain.Columns("AF").WrapText = False

    ' Delete the columns we do not need
    deleteCols = Array("A", "B", "C", "D", "E", "F", "R", "S", "U", "V", "W", "X", "AA", "AB", "AC")
    For i = UBound(deleteCols) To LBound(deleteCols) Step -1
        wsMain.Columns(deleteCols(i)).Delete
    Next i

    ' Close the workbooks
    wbNotes.Close SaveChanges:=False
    wbDisputes.Save
    wbDisputes.Close

    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "The UDM Dispute file has been updated"

End Sub