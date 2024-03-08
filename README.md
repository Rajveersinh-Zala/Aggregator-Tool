# Aggregator Tool
---
- [Key Features](#key-features)
- [How it Works](#how-it-works)
- [Benefits](#benefits)
- [VBA script](#vba-script)

This sophisticated Excel project addresses the challenge of efficiently managing a list of customer names, their associated orders, and corresponding dates. The primary objectives include the extraction of unique customer names, the creation of individualized sheets for each customer, detailed presentation of customer-specific information within these sheets, and the culmination of this data into a comprehensive dashboard.

### Key Features

- Utilizes VBA scripting to analyze and extract unique customer names from a dataset.
- Automatically generates dedicated sheets for each customer, ensuring a focused and organized data structure.
- Includes detailed information within each customer sheet, presenting order details and associated dates.
- Culminates in the creation of a dynamic dashboard, providing a holistic view of customer-related metrics and insights.

### How it Works

- Data Analysis: The project starts by analyzing a list of customer names, orders, and dates.
- Sheet Generation: Unique customer names trigger the automatic creation of dedicated sheets, streamlining data organization.
- Detail Integration: Within each customer sheet, detailed information about their orders and associated dates is presented comprehensively.
- Dashboard Creation: The project concludes by aggregating key metrics and insights into a visually appealing dashboard, offering a centralized overview.

### Benefits
- Efficiently organizes customer data for enhanced readability and accessibility.
- Minimizes manual effort by automating the process of sheet generation and data consolidation.
- Provides a dynamic dashboard for quick and informed decision-making.

### VBA script
demo of creating all unique sheets where assuming names are in Column A
```vba
Sub CreateSheetsFromTable()
    Dim ws As Worksheet
    Dim nameRange As Range
    Dim cell As Range
    Dim table As ListObject

    ' Assuming names start from cell A2 in the active sheet
    Set nameRange = ActiveSheet.Range("A2", ActiveSheet.Cells(Rows.Count, "A").End(xlUp))

    ' Loop through each cell in the range and create a sheet for each name
    For Each cell In nameRange
        ' Check if the sheet doesn't already exist
        If SheetExists(cell.Value) = False Then
            ' Create a new sheet
            Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
            ws.Name = cell.Value
            ' Update cell XFD1 in the active sheet with the sheet name
            ActiveSheet.Range("XFD2").Value = cell.Value
            ' Set cell XFD1 formatting (color, border, and bold)
            With ActiveSheet.Range("XFD2")
                .Font.Bold = True
                .Interior.Color = RGB(255, 217, 102)
            End With
            With ActiveSheet.Range("XFD2").Borders
                .LineStyle = xlContinuous
                .Color = RGB(0, 0, 0) 
                .Weight = xlThin
            End With
        End If
    Next cell
    Application.CutCopyMode = False ' Clear clipboard after copying
    Sheets("re").Select
    
    copyHeaderAndPaste
    
    giving_names_with_loop
    
End Sub

Function SheetExists(sheetName As String) As Boolean
    ' Check if a sheet with the given name already exists
    On Error Resume Next
    SheetExists = Not Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
```
