Attribute VB_Name = "Module1"
Sub TransferData()
    Dim wbProduction As Workbook
    Dim wsProduction As Worksheet
    Dim wsReport As Worksheet
    Dim lastRowProduction As Long
    Dim lastRowReport As Long
    Dim i As Long
    Dim j As Long
    Dim targetColumn As String
    Dim targetColumnNum As Integer

    ' Selection of production file by the user
    Dim fileNameProduction As String
    fileNameProduction = Application.GetOpenFilename("Excel Files (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", , "Select production file.")

    ' Check if file was selected
    If fileNameProduction = "False" Then
        MsgBox "Production file not selected"
        Exit Sub
    End If

    ' Providing the target column in the reporting file
    targetColumn = InputBox("Provide the letter of the target column in the reporting file:")

    ' Convert column letter to a number
    targetColumnNum = Range(targetColumn & "1").Column

    ' Open production file and use first sheet
    Set wbProduction = Workbooks.Open(fileNameProduction)
    Set wsProduction = wbProduction.Sheets(1)

    ' Reporting sheet is an active sheet
    Set wsReport = ThisWorkbook.ActiveSheet

    ' Find the last row in production and reporting file
    lastRowProduction = wsProduction.Cells(wsProduction.Rows.Count, "E").End(xlUp).Row
    lastRowReport = wsReport.Cells(wsReport.Rows.Count, "I").End(xlUp).Row

    ' Scan through the rows in production file
    For i = 2 To lastRowProduction ' First row is a header
        ' Check if value in column E has a matching value in column I
        For j = 2 To lastRowReport
            If wsProduction.Cells(i, 5).Value = wsReport.Cells(j, 9).Value Then
                ' Check if cell in X column in production file is empty
                If wsProduction.Cells(i, 24).Value = "" Then
                    ' If so, transfer the value from the X column, but from the row below
                    wsReport.Cells(j, targetColumnNum).Value = wsProduction.Cells(i + 1, 24).Value
                Else
                    ' Else transfer the data from X column for a matching row
                    wsReport.Cells(j, targetColumnNum).Value = wsProduction.Cells(i, 24).Value
                End If
                Exit For
            End If
        Next j
    Next i

    ' Communicate to the user
    MsgBox "Transfer completed successfully."

End Sub

